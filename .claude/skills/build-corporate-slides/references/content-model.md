# Content Model

`slides/work/slide_content.yaml` は文章、数値、順序、意味構造を保持する。座標、寸法、フォントサイズ、色コードなどの描画詳細は `slides/work/generate_pptx.py` またはslidekitへ置き、YAMLを別の描画言語にしない。

新規作成時は [`.slide-content.example.yaml`](../.slide-content.example.yaml) を出発点にできる。

```yaml
deck:
  title: "生成AIの社内導入"
  mode: business

slides:
  - id: value_and_risk
    type: comparison
    density: standard
    title: "なぜ今、生成AIを検証するのか"
    primary_message: "PoCで会社としての使い方を決める"
    left:
      heading: "期待される業務価値"
      items:
        - title: "作業時間の短縮"
          body: "文書作成や要約にかかる時間を削減"
    right:
      heading: "放置した場合の経営リスク"
      items:
        - title: "個別利用の拡大"
          body: "管理されない利用が先行する"
```

## Rules

- `primary_message` は原則一つ。描画時の最強要素を決める。
- 項目を強調したい場合も、各項目へ装飾命令を列挙しない。現状`emphasis`相当のフラグは`matrix_2x2`の`quadrants[].emphasis`（真偽値、強調象限を`brand-soft`にする）にのみ実装されている。他のrendererには汎用の`emphasis`フィールドはまだ無く、強調は`primary_message`や項目の並び順、`top_priority`（`priority_actions`）のような個別フィールドで表現する。
- `density` は `standard` / `dense`。denseでも文章縮小より重複削除と構造化を優先する。
- 標準構造に合わない図解はPythonで個別構築できる。すべてを固定スキーマへ押し込まない。
- REVISEでは文言・数値・項目順をYAML、レイアウトとDesign DNAをPythonまたはslidekitで修正する。
- `type` は見た目ではなく、伝える関係から選ぶ。対応typeとvariantは [renderer-catalog.md](renderer-catalog.md) を読む。
- renderer固有フィールドを使わない特殊ページはPythonで構築してよいが、YAMLには少なくともtitle、primary_message、density、主要データを残す。
