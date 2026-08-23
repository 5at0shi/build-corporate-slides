# build-corporate-slides

社内説明・企画・提案・報告向けの編集可能なPowerPointを、PLAN・CREATE・REVISEの3モードで生成するClaude Codeスキル。

- スキル本体: [`.claude/skills/build-corporate-slides`](.claude/skills/build-corporate-slides)（現在 v0.8.0）
- 開発・検証用ワークスペース: `input/` `work/` `output/`（[`.slide-skill-config.example.yaml`](.slide-skill-config.example.yaml)を`.slide-skill-config.yaml`としてコピーして使う）

## 構成

```text
.claude/skills/build-corporate-slides/
├── SKILL.md              # スキルの入口
├── references/           # 設計原則・renderer catalog・レイアウト例
├── runtime/python/slidekit/  # PowerPoint生成エンジン
└── scripts/               # 検証・レンダリング・自己テスト
```

## 開発時の確認

```bash
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/self_test.py
./.venv/bin/python work/generate_pptx.py
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/validate_pptx.py output/deck.pptx
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/render_and_check.py output/deck.pptx
```
