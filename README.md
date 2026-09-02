# retainer-agreement · 委托代理合同 Agent Skill

给 AI 编程 / 办公 Agent（Claude Code、Codex、Kimi Code 等支持 SKILL.md 的运行时）用的技能：
**用你们所自己的合同模板**，在面谈现场对话式出一份《委托代理合同》Word（.docx）。

- **模板是你的**：技能本身不带任何律所的合同。第一次使用时 Agent 会用一句话向你要一份常用的合同 Word，
  自动找出每次要填的位置（当事人、案由、承办律师、收费条款……），装成模板包；版式、字体、下划线全按你的原样。
- **对话式填写**：Agent 按模板里的空位一项项口语化地问，不用填表；日期默认当天，合同编号留给律所系统回填。
- **收费条款润色 + 风险审查**：把"签合同付三万、开庭前付两万"这类口述改写成正式条款（含大写金额），
  并主动追问代理范围、外部触发条件的兜底期限、后续阶段衔接、风险代理上限；婚姻继承 / 刑事等禁止风险代理的案件命中即改写。
- **本地运行、不上传**：全部在你的机器上跑，只依赖 `python-docx`。

## 安装

```bash
git clone https://github.com/zj-ai-lab/retainer-agreement-skill.git
pip install python-docx

# Claude Code
ln -s "$PWD/retainer-agreement-skill" ~/.claude/skills/retainer-agreement
# Codex
cp -R retainer-agreement-skill ~/.codex/skills/retainer-agreement
```

其他运行时：把仓库目录放进它加载 skill 的位置即可（目录内含 `SKILL.md`）。

## 第一次使用：装模板

对 Agent 说"做一份委托合同"。它发现还没有模板，会请你**给一份你们所常用的合同 Word**——签过的旧合同也行。
然后它会用人话跟你确认："我准备把这几处改成每次填写的位置：委托人姓名、身份证号……律师费那两段按分期逐段生成，
有没有漏的或不该动的？"确认后它生成模板包，并干跑一份示例合同请你在 Word 里过目。看过点头就装好了。

- 旧合同里的当事人信息只用来定位空位，**不会保存进技能**；生成的模板里全是占位符。
- 刑事辩护、常年顾问等其他合同：再给一份对应的 Word，同样流程装第二个模板；以后出件时 Agent 会先问用哪份。

## 之后每次出件

"做个委托合同，客户来了" → Agent 按模板问字段 → 你口述收费方式 → 它润色成条款并做风险审查 → 出 .docx。

## 数据与隐私

| 事项 | 说明 |
|---|---|
| 运行位置 | 本机；脚本不联网 |
| 你的合同模板 | 在 `templates/<slug>/`，已被 `.gitignore` 排除，不会被提交到任何仓库 |
| 样本合同原件 | 不复制进技能；初始化的隐私清扫会拦截残留的当事人信息 |
| 本仓库 | 不含任何律所的模板；示例数据全部脱敏（张三 / 李四 / 13800138000） |

## 目录

```
SKILL.md                    Agent 工作流（初始化 → 收集字段 → 收费条款润色与风险审查 → 渲染 → 自检 → 交付）
scripts/render.py           manifest 驱动的渲染引擎
scripts/inspect_template.py 检视合同 Word 的结构（供 Agent 判断变量位置）
scripts/init_template.py    合同样本 → 模板包；relock 重锁指纹
scripts/num2cn.py           金额大写
templates/                  你的模板包放这里（见 templates/README.md）
examples/example.json       脱敏示例数据
tests/test_render.py        回归测试（python3 tests/test_render.py；无私有模板包的用例自动跳过）
```

## 许可

- 代码（`scripts/`、`tests/`）：[Apache-2.0](LICENSE)
- 文本（`SKILL.md`、`README.md`、`templates/README.md`）：[CC BY-SA 4.0](LICENSE-CC-BY-SA-4.0)

详见 [NOTICE](NOTICE)。本仓库由 [zj-ai-lab](https://github.com/zj-ai-lab) 维护；欢迎 issue，PR 请先开 issue 讨论。
