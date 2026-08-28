---
name: build-corporate-slides
description: 社内向けPowerPoint資料の構成設計、作成、レビュー指摘を踏まえた修正に使用する。組織のスライドデザイン原則に従い、情報の伝わりやすさ、必要十分な情報密度、PowerPoint上での編集・引継ぎやすさを重視して資料を作成する。
---

# Corporate Slides

社内説明、企画、提案、報告のための編集可能なPowerPointを、PLAN・CREATE・REVISEの3モードで扱う。デザインの目的は、モダンでスマートな印象と、情報構造に沿った明快な強弱である。「AIらしさを消す」こと自体を目的にせず、意味のない反復、過剰装飾、強弱の欠如による違和感を防ぐ。

## 最初に行うこと

1. ワークスペース直下の `.slide-skill-config.yaml` を読む。なければ [`.slide-skill-config.example.yaml`](user-guide/.slide-skill-config.example.yaml) を基準にし、既定値を使う。
2. `paths` をワークスペース基準で解決し、必要なディレクトリだけ作る。
3. 依頼が構成検討、実制作、既存資料の修正のどれかを判断する。

設定済みの `python.executable` がある場合は必ずそれを優先する。別のvenvを作成・選択したり、依存関係を勝手にインストールしたりしない。不足時は [`runtime/python/requirements.txt`](runtime/python/requirements.txt) と照合して報告する。標準利用環境は個人PCでの閲覧、事前配布、オンライン会議の画面共有であり、遠距離投影を暗黙に想定しない。

CREATE前に、設定されたPythonで次を実行する。不足設定・依存があれば生成を始めず、差分を報告する。`check_config.py` は検証するconfigのパスを引数に取る。

```bash
<python> .claude/skills/build-corporate-slides/scripts/check_config.py .slide-skill-config.yaml
<python> .claude/skills/build-corporate-slides/scripts/check_environment.py
```

## Workspace Contract

設定値の既定は `build_slides/input/`、`build_slides/work/`、`build_slides/output/` とする。既存プロジェクトのディレクトリ内で作業する場合でも、スライド関連一式が `build_slides/` 配下にまとまるため他の作業内容と混ざらない。

- `build_slides/input/`: 人間からAIへ渡されたbrief、画像、図表、参考資料
- `build_slides/work/`: AIの制作コードとQA用中間成果物
- `build_slides/output/`: 人間が利用する最終成果物

標準位置は次のとおり。

```text
build_slides/
├── input/
├── work/
│   ├── slide_content.yaml
│   ├── generate_pptx.py
│   └── render/
│       ├── deck.pdf
│       └── pages/
│           ├── slide-01.png
│           └── ...
└── output/
    └── deck.pptx
```

`render_dir` や `script_dir` はconfigに増やさず、常に `work_dir` から導出する。PDFが納品物として明示された場合のみ `output_dir` にも出力する。

1ワークスペース＝1デッキを既定とする。同じワークスペースで複数のデッキを扱う場合は、内容・生成スクリプト・成果物の3つに同じ識別子を付けて対応を保つ（例: `work/pricing.yaml`、`work/pricing.generate.py`、`output/pricing.pptx`）。どれか一つだけ名前を変えると、どのスクリプトがどの成果物を作ったのかが追えなくなる。

生成は必ず `work_dir` に残るスクリプトから行う。その場限りのコマンドやインラインのスクリプトで生成すると、再現も引き継ぎもできない状態になり、完了条件を満たさない。

## PLAN

すぐにPPTXを作らず、資料の目的、読み手、意思決定、ストーリーを整理する。各スライドについて、タイトル、中心メッセージ、主な内容、最適な表現・レイアウトを示す。タイトルと箇条書きだけの計画にしない。

詳しくは [slide-planning.md](references/slide-planning.md) を読む。通常は `business` を使う。情報量の多いページだけ `dense` の判断を許容し、`large-room` は大会議室など遠距離投影が明示された場合だけ使う。

## CREATE

1. 確定したSlide Planを基に、内容と意味構造を `build_slides/work/slide_content.yaml`、描画の入口を `build_slides/work/generate_pptx.py` に分けて維持する。新規作成では [`.slide-content.example.yaml`](user-guide/.slide-content.example.yaml) と [`assets/generate_pptx.py`](assets/generate_pptx.py) を出発点にできる。
2. 各ページを [`renderer-catalog.md`](references/renderer-catalog.md) へ照合する。該当する意味ベースrendererを優先し、内容に合わないページだけ `DeckBuilder` とlayout primitivesで個別構築する。rendererへ無理に押し込まない。個別構築したページは、描画関数を `generate_pptx.py` に置き `render_deck(..., extra_renderers={...})` へ渡して同じデッキに含める。内容は他のページと同じくYAMLへ残す（[renderer-catalog.md](references/renderer-catalog.md)の「個別構築ページをYAMLと同居させる」）。
3. `DeckBuilder.from_workspace(ROOT)` を入口にし、config、パス、部署名、開示範囲、ロゴ、フォント、標準modeを自動反映する。config項目を生成コードで重複管理しない。
4. 編集可能なテキスト、表、図形、チャートを優先し、原則としてスライド全体を画像化しない。
   - 同じ見出し配下の箇条書きは、複数段落を持つ一つのtextboxへまとめる。
   - 番号・項目名・説明が一緒に動く場合は、一つの意味単位としてまとめる。
   - 見た目上の行ごとにtextboxを分けない。独立移動、別背景、別整列など具体的な理由がある場合だけ分ける。
5. YAMLの事前診断を通し、文字縮小より先に重複削除、統合、表化、ページ分割を検討する。
6. `build_slides/output/deck.pptx` を生成する。
7. [`scripts/validate_pptx.py`](scripts/validate_pptx.py) で構造・編集性を検証する。
8. [`scripts/render_and_check.py`](scripts/render_and_check.py) で `build_slides/work/render/deck.pdf` とページPNGを作る。
9. 全ページを個別表示で確認し、デッキ全体の反復は一覧でも確認する。必要なら同じYAMLと生成スクリプトを修正して再生成する。

生成方式の選択、依存関係、検証は [powerpoint-production.md](references/powerpoint-production.md)、YAMLの責務は [content-model.md](references/content-model.md)、renderer選択は [renderer-catalog.md](references/renderer-catalog.md)、視覚判断は [visual-quality.md](references/visual-quality.md)、色・Componentの具体的な選択は [design-system.md](references/design-system.md) を読む。

レイアウトを決める前に、内容に該当する場合だけ [layout-patterns.md](references/examples/layout-patterns.md) を読む。例は完成テンプレートではなく、選択基準とDesign DNAを示す。表紙を作る場合は [cover-modern-brand-field.md](references/examples/cover-modern-brand-field.md)、セクション見出しとグラフを組む場合は [section-lead-chart.md](references/examples/section-lead-chart.md) も読む。画像、ロゴ、PNGグラフを使う場合は [image-handling.md](references/image-handling.md) も読む。

## REVISE

文言、項目、数値、順序は原則として既存の `build_slides/work/slide_content.yaml` を修正し、配色、余白、配置、図解構造は `build_slides/work/generate_pptx.py` またはslidekitを修正して再生成する。別名の生成スクリプトを増やさず、履歴はGitへ任せる。既存スクリプトがない、再現不能、または直接編集の方が明らかに安全な場合だけ例外とし、その理由を伝える。

指摘箇所と、それを成立させるために必要な周辺だけを変更する。再生成後はCREATEと同じvalidate・render・Visual QAを行う。関係のないページの配色、文字組み、レイアウトを刷新しない。

## 完了条件

- 最終PPTXが `output_dir`（既定 `build_slides/output`）`/deck.pptx` にある
- 構造検証に重大エラーがない
- レンダリング可能な環境では全ページを目視確認済み
- 生成スクリプトが `work_dir/generate_pptx.py` に残る
- YAML方式を使った場合は `work_dir/slide_content.yaml` が生成内容と一致している
- 箇条書きや連続項目が無理由に別textboxへ分割されていない
- configで有効なロゴ、部署名、開示範囲が表紙へ反映されている
- YAML事前診断の警告を確認し、無視した場合は理由がある
- 未実施のQA、フォント代替、レンダリング不可などの制約を明示している
