# templates/ — 模板包目录

本 skill **没有模板不能出件**。每个模板包一个子目录：

```
templates/
└── <slug>/
    ├── template.docx            占位符版模板（{{KEY}} 形式）
    ├── manifest.json            字段清单 / 提问话术 / 默认值 / 渲染方式 / 模板指纹
    └── contract.skeleton.json   contract.json 骨架（init 生成，agent 填数据用）
```

- **首次使用**：目录为空 → 按 SKILL.md 第 0 步，请使用者提供一份常用的委托代理合同 Word，
  `scripts/inspect_template.py` 看结构 → 与使用者用自然语言确认哪些地方是要填的 → `scripts/init_template.py init` 建包。
- **再加一种合同**（刑事 / 常年顾问 / 非诉专项……）：同一套流程，再建一个子目录；多包时渲染需指定 `--template <slug>`。
- **在 Word 里改过 template.docx**：跑 `scripts/init_template.py relock templates/<slug>` 重新锁定指纹，否则渲染时会 WARN 模板漂移。
- **隐私**：init 只把样本里的当事人信息替换成占位符，样本原件不复制进来；manifest 不保存任何样本值；
  init 的隐私清扫会拦截残留的样本值。模板包若含本所固定信息（所名 / 地址 / 账号）属于律所资料，**不要提交到公开仓库**。
