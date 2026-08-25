# Renderer Catalog

rendererは完成テンプレートではなく、頻出する意味構造を編集可能なPowerPointへ安定して変換する入口である。見た目が似ていることではなく、情報の関係が一致するときに使う。

| type | 使う状況 | 主なvariant | 編集構造 |
|---|---|---|---|
| `cover` | 表紙 | brand side / shape | メタ情報とロゴをconfigから自動反映 |
| `comparison` | 二つの観点、価値とリスク、現状と将来 | `balanced`, `asymmetric` | 各列の項目群を一つの複数段落textboxにする |
| `evidence_and_decision` | 根拠から一つの推奨判断を導く | 標準 | 根拠群＋独立した判断領域 |
| `scope_and_exclusions` | 対象範囲と対象外を同時に示す | 標準 | 対象の意味単位＋対象外の一括リスト |
| `process_with_gates` | フェーズ、作業、判断時点 | 標準 | フェーズとゲートを別レイヤーにする |
| `table_with_conclusion` | 条件、比較、評価を表で読む | 標準 | ネイティブPowerPoint table＋結論 |
| `chart_with_insight` | グラフと読み取りを示す | `standard`, `conclusion-led` | `chart`指定でネイティブPowerPointグラフ（数値・系列名を編集可）、`image`指定でPNG画像 |
| `org_layers` | 意思決定・運営など縦の責任階層と、横に並ぶ実行部門 | 標準 | 階層バンド＋実行部門ごとのCard |
| `priority_actions` | 優先度付きの課題と、対応する方針 | 標準 | 優先度付きリスト＋淡色の対応方針パネル |
| `stage_track` | 現在から将来への段階的な進行（ロードマップ等） | 標準 | 同格のステージCard群 |
| `numbered_list` | アジェンダ、依頼事項など番号付きの単列項目 | 標準 | 番号付き行の集合。上部導入文／下部結論のどちらかを選べる |
| `section_divider` | 複数テーマを扱う資料の章区切り | 標準 | 通常ページのヘッダーを使わない単独ページ |
| `matrix_2x2` | 2軸で選択肢を4象限に整理する（ポートフォリオ分析等） | 標準 | 象限ごとのBackground Zone。各象限は独立編集可能 |
| `stat_highlight` | 単一の実績数値を主役に、補足指標と結論を示す | 標準 | 主指標＋補足指標Cardの集合 |

## Selection

- 両側を公平に読むなら `comparison: balanced`。
- 片側を補助領域として弱めるなら `comparison: asymmetric`。
- 結論が一つで根拠が複数なら `evidence_and_decision`。
- 順番だけでなく承認時点が重要なら `process_with_gates`。
- セル編集と行追加が想定される情報は `table_with_conclusion`。
- グラフ全体を公平に読むなら `chart_with_insight: standard`、一つの主張が主役なら `conclusion-led`。
  - 実データから作る棒・折れ線・円グラフは`chart`（`type: column|bar|line|pie`、`categories`、`series`）を使い、ネイティブPowerPointグラフにする（数値をPowerPoint上で直接編集できる）。パワーポイント生成側で用意したPNGしかない場合だけ`image`を使う。原則として`chart`を優先する。
  - `add_native_chart`は生成時に不整合を検知して止める（categoriesとvalues数の不一致、`pie`への複数系列指定、`categories`が空、など）。負の値を含む棒・列グラフはラベル位置を自動調整し、カテゴリ数が多い（8件超）折れ線・棒グラフは点ごとのラベルを自動的に省略する（軸・目盛線での判読に任せる）。
- 縦の責任階層と横の実行部門を両方示すなら `org_layers`。`layers`は2件までを目安にする（3件以上は各層の本文が収まらずpreflightで警告する。業務実行層は`execution`で別途横分割できるため、階層自体は「意思決定・運営」の2段+実行、のように抑える）。
- 課題と対応策を優先度付きで左右に並べるなら `priority_actions`。
- 現在から将来への段階的な広がりを示すなら `stage_track`。
- アジェンダやNext Stepなど、番号付きの単列項目だけで構成されるページは `numbered_list`。
- 複数テーマを扱う資料で章を区切るなら `section_divider`。表紙と混同しない（ロゴ・部署名・開示区分は表紙のみ）。
- 2軸で選択肢の位置づけを比較するなら `matrix_2x2`。軸上の座標が厳密な数値でなく、4象限のどれに属するかが要点の場合に使う。散布図のように正確な座標が要点の場合はPNGを`chart_with_insight`で扱う。
- 実績・効果を1つの数値で語るなら `stat_highlight`。複数の観点を対等に比較する場合は`table_with_conclusion`を優先する。

## Escape Hatch

一致するrendererがない場合は無理に近い型へ入れない。`DeckBuilder.add_slide()` と `Region.columns()` / `rows()` を使って個別構築し、既存のtypography、semantic color、余白、編集性規則は維持する。新しいrendererを追加するのは、同じ意味構造が複数回現れ、決定論的な実装が再利用できる場合だけとする。

項目数が少ないページを個別構築する場合は、`add_item_list` や `add_paragraph_textbox(vertical_anchor=...)` を使い、上詰めで余白が偏らないよう領域内で縦方向に配置を検討する。
