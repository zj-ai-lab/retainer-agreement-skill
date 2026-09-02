# Changelog

本仓库遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)；版本号与技能 `SKILL.md` frontmatter 一致。

## [2.0.1] - 2026-09-02

- README 改为面向非技术使用者的大白话版本，附可整段复制发给 AI 助手的安装 / 试跑指令、测试清单与常见问题。
- 新增 `examples/make_demo_sample.py`：生成一份虚构示例合同，没有现成模板也能试跑初始化。
- `init_template.py`：map 字段支持 `leftover_ok`（样本值是通用词、固定文字里还会再出现时不视为残留）。
- 新增测试反馈 issue 模板。

## [2.0.0] - 2026-09-02

首个公开版本（内部 v1.x 自 2026-05 起在律所实务中迭代，本版为面向公开发布的重构）。

### 特性
- **manifest 驱动的渲染引擎**：模板包 `templates/<slug>/{template.docx, manifest.json}`；字段清单、提问话术、默认值、多段字段与下划线模式全部来自 manifest，引擎不含任何律所的版式常量，几何量逐段从模板自身读取。
- **第 0 步模板初始化**：Agent 用自然语言向使用者要一份常用合同 Word，`inspect_template.py` 检视结构、人话确认变量、`init_template.py` 建包（动态识别"标签 + 制表符 + 下划线"的空白行与 tab leader 画线，规整为单 run 连续下划线方案；收费段自动判断金额下划线；隐私清扫拦截残留样本值），干跑一份让使用者在 Word 里验收。
- **收费条款润色 + 风险审查**：口语 → 正式条款（含大写金额）；A 代理范围 / B 外部触发兜底期限 / C 后续阶段衔接 / D 风险代理上限保底 / E 禁止风险代理案件红线（婚姻继承、刑事、行政、国家赔偿、群体性诉讼）；渲染前红线词扫描 WARN 兜底。
- 模板指纹（`template_sha256`）漂移 WARN 与 `relock`；产物 core.category 写入引擎版本与模板指纹。
- 回归测试 12 例（含合成非特定律所版式的端到端；依赖私有模板包的用例在本仓库自动跳过）。
