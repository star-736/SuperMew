# Anthropic Frontend Design Skill

**Name**: `anthropic-frontend-design`  
**Description**: 创建独特的生产级前端界面，避免通用的"AI slop"美学。用于构建具有卓越细节关注和大胆创意选择的 UI 组件、页面、应用程序或界面。

---

## 核心理念：反 AI 垃圾 (Anti-AI Slop)

Claude（及所有 AI 代理）capable of extraordinary creative work, yet often default to safe, generic patterns. 此技能 **强制 (MANDATES)** 打破这些模式。

### 避免 (AVOID)
- Inter, Roboto, Arial, 系统字体
- 紫色白色渐变
- 模板化 SaaS 布局
- 表情符号作为图标

### 强制 (MANDATE)
- 独特排版
- 特定情境配色
- 有意境的动效
- 意外的空间构图
- 生产级功能代码

---

## 设计思考流程

编码前，理解上下文并承诺 **大胆 (BOLD)** 的美学方向：

1. **目的 (Purpose)**: 解决什么问题？面向谁？
2. **基调 (Tone)**: 选择极端方向——残酷极简、混乱最大化、复古未来、有机、奢华、playful、编辑风格等。
3. **差异化 (Differentiation)**: 什么让它 **令人难忘 (UNFORGETTABLE)**？

---

## 实施标准

### 1. 专业 UI 规则

| 规则 | 做 (Do) | 不做 (Don't) |
|------|---------|-------------|
| **图标** | 使用 SVG (Heroicons, Lucide, Simple Icons) | 使用表情符号如 🎨 🚀 ⚙️ 作为 UI 图标 |
| **排版** | 美观、独特的 Google/自定义字体 | Inter, Roboto, Arial, 系统字体 |
| **悬停** | 稳定过渡 (颜色/透明度/阴影) | 移位布局的缩放变换 |
| **光标** | 所有交互项添加 `cursor-pointer` | 按钮/卡片上保留默认光标 |
| **对比度** | 无障碍最低 4.5:1 | 不可读的低对比度"氛围" |

### 2. 动效与动画
- 尽可能优先使用 CSS 解决方案
- 聚焦高影响力时刻（页面加载时的交错显示）
- 微交互使用 150-300ms 持续时间

### 3. 空间构图
- 使用不对称、重叠或对角流打破标准网格
- 平衡慷慨的负空间或有意的密度

---

## 交付前检查清单

### 视觉质量
- [ ] 无表情符号用作图标 (仅 SVG)
- [ ] 排版有特色且非"AI 标准"
- [ ] 配色方案对情境独特 (无通用渐变)
- [ ] 悬停状态提供清晰、稳定的视觉反馈

### UX 与无障碍
- [ ] 所有交互元素都有 `cursor-pointer`
- [ ] 表单输入有标签；图片有 alt 文本
- [ ] 文本对比度满足最低 4.5:1 (测试亮/暗模式)
- [ ] 所有断点响应式 (375px, 768px, 1024px, 1440px)
- [ ] 移动端无水平滚动

---

## 美学方向参考

*承诺选择一个方向并完全执行——不要半途而废。*

| 方向 | 特点 |
|------|------|
| **Brutally Minimal** | 单色，极端白空间，稀疏排版 |
| **Maximalist Chaos** | 重叠元素，密集信息，图案混合 |
| **Retro-Futuristic** | chrome 效果，霓虹点缀，80 年代灵感 |
| **Luxury/Refined** | 金/暗点缀，衬线字体，慷慨间距 |
| **Playful/Toy-like** | 圆角，明亮粉彩，弹性动画 |
| **Editorial/Magazine** | 基于网格，大胆标题，清晰层级 |
| **Brutalist/Raw** | 等宽字体，强烈对比，工业风 |
| **Art Deco** | 锐角，金属点缀，华丽边框 |
