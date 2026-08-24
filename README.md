# build-corporate-slides

社内説明・企画・提案・報告向けの編集可能なPowerPointを、PLAN・CREATE・REVISEの3モードで生成するClaude Codeスキル。

- スキル本体: [`.claude/skills/build-corporate-slides`](.claude/skills/build-corporate-slides)（現在 v0.9.0）
- 開発・検証用ワークスペース: `slides/input/` `slides/work/` `slides/output/`（[`.slide-skill-config.example.yaml`](.slide-skill-config.example.yaml)を`.slide-skill-config.yaml`としてコピーして使う）
- 使い方の早見表: [`user-guide/getting-started.md`](.claude/skills/build-corporate-slides/user-guide/getting-started.md)（設定ファイルの各項目、PLAN/CREATE/REVISEの流れ）
- モードの見た目比較: [`user-guide/mode-guide.pdf`](.claude/skills/build-corporate-slides/user-guide/mode-guide.pdf)（business/dense/large-roomの実寸サンプル）
- 部位の名称一覧: [`user-guide/component-map.pdf`](.claude/skills/build-corporate-slides/user-guide/component-map.pdf)（修正依頼時に部位名を視覚的に参照できる）

## 構成

```text
.claude/skills/build-corporate-slides/
├── SKILL.md              # スキルの入口
├── references/           # 設計原則・renderer catalog・レイアウト例（AIが読む）
├── user-guide/            # 使い方・モード比較・部位名称一覧（人間が読む）
├── runtime/python/slidekit/  # PowerPoint生成エンジン
└── scripts/               # 検証・レンダリング・自己テスト
```

## 開発時の確認

```bash
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/self_test.py
./.venv/bin/python slides/work/generate_pptx.py
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/validate_pptx.py slides/output/deck.pptx
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/render_and_check.py slides/output/deck.pptx
```
