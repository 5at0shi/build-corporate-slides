# build-corporate-slides

社内説明・企画・提案・報告向けの編集可能なPowerPointを、PLAN・CREATE・REVISEの3モードで生成するClaude Codeスキル。

- スキル本体: [`.claude/skills/build-corporate-slides`](.claude/skills/build-corporate-slides)（現在 v0.9.1）
- 開発・検証用ワークスペース: `build_slides/input/` `build_slides/work/` `build_slides/output/`（このリポジトリ自身の動作確認用。実際にスキルを使う側は[`user-guide/.slide-skill-config.example.yaml`](.claude/skills/build-corporate-slides/user-guide/.slide-skill-config.example.yaml)を自分のプロジェクトへ`.slide-skill-config.yaml`としてコピーする）
- 使い方の早見表: [`user-guide/getting-started.md`](.claude/skills/build-corporate-slides/user-guide/getting-started.md)（設定ファイルの各項目、PLAN/CREATE/REVISEの流れ）
- 作れるものの実物見本・部位の名称一覧: [`user-guide/capability-showcase.pdf`](.claude/skills/build-corporate-slides/user-guide/capability-showcase.pdf)（配色・コンポーネント・アイコン・チャート・renderer各種を、使用したPython関数名付きで一覧できる）
- モードの見た目比較: [`user-guide/mode-guide.pdf`](.claude/skills/build-corporate-slides/user-guide/mode-guide.pdf)（business/dense/large-roomの実寸サンプル）

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
# slidekit全体の自己テスト（全renderer typeを最小構成で1枚ずつ生成・検証）
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/self_test.py

# 過去の不具合を再現する回帰確認デッキを生成し、検証・レンダリングする例
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_regression_check.py
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/validate_pptx.py build_slides/output/skill-regression-check.pptx
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/render_and_check.py build_slides/output/skill-regression-check.pptx

# 参考資料（capability-showcase.pdf等）を再生成する場合
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_capability_showcase.py
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_mode_guide.py
```
