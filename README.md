# build-corporate-slides

社内説明・企画・提案・報告向けの編集可能なPowerPointを、PLAN・CREATE・REVISEの3モードで生成するClaude Codeスキル。

- スキル本体: [`.claude/skills/build-corporate-slides`](.claude/skills/build-corporate-slides)（現在 v0.9.3）
- 生成したデッキの実例: [`samples/`](samples/)（YAMLとPDFの対。内容のある資料としての組み立て方）
- 開発・検証用ワークスペース: `build_slides/input/` `build_slides/work/` `build_slides/output/`（このリポジトリ自身の動作確認用。中身は追跡しない）
- 使い方の早見表: [`user-guide/getting-started.md`](.claude/skills/build-corporate-slides/user-guide/getting-started.md)（作れるものの見つけ方、設定ファイルの各項目。PLAN/CREATE/REVISEの詳細な流れは`SKILL.md`）
- 作れるものの実物見本・部位の名称一覧: [`user-guide/capability-showcase.pdf`](.claude/skills/build-corporate-slides/user-guide/capability-showcase.pdf)（配色・タイポグラフィ（3モード実寸）・コンポーネント・アイコン・チャート・renderer各種を、使用したPython関数名付きで一覧できる）

## 自分のプロジェクトで使う

スキルは`.claude/skills/build-corporate-slides/`の中で完結している。使う側のプロジェクトへは、**このディレクトリごとコピーする**。

```bash
mkdir -p <あなたのプロジェクト>/.claude/skills
cp -r .claude/skills/build-corporate-slides <あなたのプロジェクト>/.claude/skills/
```

`.claude/`ごとコピーしてもこのリポジトリでは同じ結果になる（他に何も置いていないため）が、コピー先に既存の`.claude/`がある場合は上書きしてしまうので、スキルのディレクトリだけを移すほうが安全。

スキルの外にあって別途必要なものは次の3つ。

| | 必要性 |
|---|---|
| `.slide-skill-config.yaml`（プロジェクト直下） | 無くても生成はできるが、部署名・開示区分・ロゴが入らない。[`.slide-skill-config.example.yaml`](.claude/skills/build-corporate-slides/user-guide/.slide-skill-config.example.yaml)をコピーして編集する |
| python-pptx / PyYAML / Pillow | 必須。[`runtime/python/requirements.txt`](.claude/skills/build-corporate-slides/runtime/python/requirements.txt)が基準 |
| LibreOffice等 | PDF化と目視確認にのみ必要。PPTXの生成自体には不要 |

`build_slides/`はスキルが必要に応じて自動で作るため、事前に用意しなくてよい。

## 構成

```text
.claude/skills/build-corporate-slides/
├── SKILL.md              # スキルの入口
├── references/           # 設計原則・renderer catalog・レイアウト例（AIが読む）
├── user-guide/            # 使い方・モード比較・部位名称一覧（人間が読む）
├── runtime/python/slidekit/  # PowerPoint生成エンジン
└── scripts/               # 検証・レンダリング・自己テスト

samples/                   # 生成したデッキの実例（YAML＋PDF）
build_slides/              # 開発・検証用の作業領域（追跡しない）
```

## 開発時の確認

```bash
# slidekit全体の自己テスト（全renderer typeを最小構成で1枚ずつ生成・検証）
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/self_test.py

# 過去の不具合を再現する回帰確認デッキを生成し、検証・レンダリングする例
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_regression_check.py
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/validate_pptx.py build_slides/output/skill-regression-check.pptx
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/render_and_check.py build_slides/output/skill-regression-check.pptx

# 参考資料（capability-showcase.pdf）を再生成する場合
./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_capability_showcase.py

# レンダリングを変更したら、samples/のPDFを最新の実装へ揃える
./.venv/bin/python samples/render_samples.py
```
