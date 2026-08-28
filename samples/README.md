# サンプル

`build-corporate-slides`スキルで実際に生成したデッキ。YAMLと、そこから生成したPDFを対にして置いてある。

スキルの見本（[capability-showcase.pdf](../.claude/skills/build-corporate-slides/user-guide/capability-showcase.pdf)）が構成要素の最小構成を並べたものであるのに対し、こちらは**内容のある資料としてどう組み立てるか**の実例。YAMLをそのまま書き換えて別テーマの資料の出発点にもできる。

| # | 資料 | ページ | 使っている構成 |
|---|---|---|---|
| 01 | [全社DX推進 年度総括](sample-01-dx-annual-review.pdf) | 4 | 8行×6列の表、左右5項目ずつの比較、4フェーズ×各4項目の工程 |
| 02 | [海外拠点 統合再編](sample-02-apac-consolidation.pdf) | 4 | 3×3マトリクス、5列×8行の損益表、左右非対称の比較 |
| 03 | [人事制度改定](sample-03-hr-system-revision.pdf) | 6 | 対象範囲と対象外、責任階層、優先度付き課題、実績数値、根拠と判断 |
| 04 | [NVIDIA 直近3期の業績と要因分析](sample-04-nvidia-performance.pdf) | 7 | 増減要因の分解（ブリッジ図）、四半期推移のグラフ、論点の階層分解 |
| 05 | [Claude Code 導入経路の比較](sample-05-claude-code-bedrock.pdf) | 7 | 9行の機能比較表、導入時の課題と対応方針 |
| 06 | [GitHub Copilot と Claude Code の比較](sample-06-copilot-vs-claude-code.pdf) | 8 | 用途の2軸整理、段階ごとに箇条書きを持つ進行図 |
| 07 | [AIエージェントの社内活用](sample-07-skills-and-subagents.pdf) | 8 | 表＋気づきの併記、期間の帯（タイムライン）、推進体制 |

01〜03は情報量の上限を試した高密度サンプル、04は全ページを高密度にした調査資料、05〜07は情報量の多いページと少ないページを混在させたもの。1ページあたりの文字量は70〜626字の幅がある。

renderer 20種のうち17種がいずれかのサンプルに登場する。残る`section_divider`・`funnel`・`cycle`を含め、20型すべてを1つのデッキに収めた例は[.slide-content.example.yaml](../.claude/skills/build-corporate-slides/user-guide/.slide-content.example.yaml)にある。

## 内容について

04〜07は公開情報を調べて作成した。作成時点（2026年8月）の情報であり、**スキルの出力例としてのサンプルであって、内容の正確性を保証するものではない**。数値を引用する場合は一次情報を確認すること。

## 作り直し

スキル側のレンダリングを変更したら、PDFを作り直して最新の実装へ揃える。

```bash
python3 samples/render_samples.py          # 全件
python3 samples/render_samples.py 04 05    # 番号で絞る
```

各サンプルのYAML事前診断と構造検証も同時に走る。PPTXは中間物として`build_slides/output/`へ出力し、追跡しない（YAMLから何度でも作り直せるため）。
