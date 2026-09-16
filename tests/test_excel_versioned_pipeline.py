"""XLSX -> immutable publication -> durable parents -> canonical retrieval.

SQLite and fake vector/model adapters deliberately avoid live infrastructure.
"""

from dataclasses import replace
import hashlib
import json
from types import SimpleNamespace

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from backend.core.errors import AppError
from backend.db.models import Base, User
from backend.documents.catalog import DocumentCatalog
from backend.documents.publication import DocumentPublication, DocumentPublicationConfig
from backend.documents.retrieval import DocumentRetrievalScope
from backend.indexing.document_loader import DocumentLoader
from backend.indexing.milvus_client import _format_retrieval_hit
from backend.indexing.milvus_writer import MilvusWriter
from backend.rag.evidence import pack_evidence
from backend.security.uploads import StoredUpload
from test_excel_structured_loader import workbook
from test_milvus_writer import FakeEmbeddingService, VersionedStore
from test_parent_chunk_store_versions import FakeCache, load_parent_store_module
from test_rag_retrieval_targets import Embedding, RoutedStore, TargetStore, _utils


def test_real_workbook_publication_and_retrieval_preserve_json_and_current_version(
    tmp_path,
):
    engine = create_engine(
        "sqlite://", connect_args={"check_same_thread": False}, poolclass=StaticPool
    )
    Base.metadata.create_all(engine)
    sessions = sessionmaker(bind=engine, expire_on_commit=False)
    with sessions.begin() as db:
        db.add(User(id=1, username="reader", password_hash="offline", role="user"))
    cache = FakeCache()
    parents = load_parent_store_module(sessions, cache).ParentChunkStore()
    catalog = DocumentCatalog(sessions)
    events = []
    store = VersionedStore(events, collection_name="documents_catalog_v1")
    store.verify_result = SimpleNamespace(exact=True)
    writer = MilvusWriter(
        milvus_manager=store, embedding_service=FakeEmbeddingService(events)
    )
    config = replace(
        DocumentPublicationConfig.from_runtime(),
        tenant_id="tenant-a",
        upload_dir=tmp_path,
        vector_collection=store.collection_name,
    )
    publication = DocumentPublication(
        catalog=catalog,
        loader=DocumentLoader(),
        parent_store=parents,
        writer=writer,
        config=config,
    )
    scope = DocumentRetrievalScope(catalog)

    def submit(name):
        path = tmp_path / name
        workbook(path)
        return publication.submit(
            StoredUpload(
                original_name="sample.xlsx",
                object_key=name,
                path=path,
                extension=".xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                size_bytes=path.stat().st_size,
                content_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
            ),
            owner_id=1,
        )

    try:
        candidate = submit("first.xlsx")
        assert scope.resolve(tenant_id="tenant-a").targets == ()
        outcome = publication.run(candidate.job.id)
        assert outcome.published
        snapshot = scope.resolve(tenant_id="tenant-a")
        assert snapshot.targets[0].document_version_ids == (candidate.version.id,)
        assert scope.resolve(tenant_id="tenant-b").targets == ()
        assert len(store.records) == 4
        assert all(record["chunk_level"] == 3 for record in store.records)
        hits = [
            _format_retrieval_hit({"id": index, "distance": 0.9, "entity": record})
            for index, record in enumerate(store.records)
        ]
        cache.values.clear()  # Force the actual SQLite ParentChunk reload path.
        utils = _utils(AUTO_MERGE_ENABLED="true", AUTO_MERGE_THRESHOLD="2")
        utils._document_retrieval_scope = scope
        utils._parent_chunk_store = parents
        utils._embedding_service = Embedding()
        retrieval_store = TargetStore(hybrid=hits)
        utils._milvus_manager = RoutedStore({store.collection_name: retrieval_store})
        result = utils.retrieve_documents(
            "苹果价格和季度数据",
            tenant_id="tenant-a",
            top_k=4,
            retrieval_snapshot=snapshot,
        )
        assert result["meta"]["auto_merge_applied"]
        assert len(result["docs"]) == 2
        assert all(doc["chunk_level"] == 2 for doc in result["docs"])
        tables = {
            json.loads(doc["text"])["sheet"]: json.loads(doc["text"])
            for doc in result["docs"]
        }
        assert tables["商品表"]["rows"][0]["values"]["价格_2"] == "6"
        assert tables["季度汇总"]["rows"][-1]["values"]["Q2"] == "=B3+12"
        evidence = pack_evidence(result["docs"], maximum_characters=12000)
        assert evidence.truncated_count == evidence.omitted_count == 0
        assert "商品表" in evidence.text
        assert all(
            doc["document_version_id"] == candidate.version.id
            for doc in evidence.documents
        )
        assert all(
            doc["section_id"].startswith("sheet:") and doc["content_hash"]
            for doc in evidence.documents
        )
        assert candidate.version.id in retrieval_store.hybrid_calls[0]["filter_expr"]

        # A parser failure must never replace the published workbook.
        path = tmp_path / "bad.xlsx"
        path.write_bytes(b"not an Excel archive")
        bad = publication.submit(
            StoredUpload(
                original_name="sample.xlsx",
                object_key=path.name,
                path=path,
                extension=".xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                size_bytes=path.stat().st_size,
                content_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
            ),
            owner_id=1,
        )
        with pytest.raises(AppError):
            publication.run(bad.job.id)
        assert scope.resolve(tenant_id="tenant-a").targets[0].document_version_ids == (
            candidate.version.id,
        )
        cache.values.clear()
        restored = parents.get_documents_by_ids(
            [doc["chunk_id"] for doc in result["docs"]]
        )
        assert [doc["text"] for doc in restored] == [
            doc["text"] for doc in result["docs"]
        ]
    finally:
        engine.dispose()


def test_structured_parser_changes_default_build_fingerprint(monkeypatch):
    monkeypatch.delenv("DOCUMENT_PARSER_VERSION", raising=False)
    monkeypatch.delenv("DOCUMENT_CHUNKER_VERSION", raising=False)
    config = DocumentPublicationConfig.from_runtime()
    old = replace(
        config,
        parser_version="document-loader-v2",
        chunker_version="three-level-800-100-v2",
    )
    assert config.build_profile.fingerprint != old.build_profile.fingerprint
    assert config.parser_version == "document-loader-v3-excel-json"
