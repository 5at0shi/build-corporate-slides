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
| `table_with_insight` | 表から複数の気づきを箇条書きで示す（chart_with_insightの表版） | 標準 | ネイティブPowerPoint table＋読み取れることの箇条書き＋結論 |
| `chart_with_insight` | グラフと読み取りを示す | `standard`, `conclusion-led` | `chart`指定でネイティブPowerPointグラフ（数値・系列名を編集可）、`image`指定でPNG画像 |
| `org_layers` | 意思決定・運営など縦の責任階層と、横に並ぶ実行部門 | 標準 | 階層バンド＋実行部門ごとのCard |
| `priority_actions` | 優先度付きの課題と、対応する方針 | 標準 | 優先度付きリスト＋淡色の対応方針パネル |
| `stage_track` | 現在から将来への段階的な進行（ロードマップ等） | 標準 | 同格のステージCard群。既定でCard間を矢印で繋ぐ（`connectors: false`で非表示） |
| `numbered_list` | アジェンダ、依頼事項など番号付きの単列項目 | 標準 | 番号付き行の集合。上部導入文／下部結論のどちらかを選べる |
| `section_divider` | 複数テーマを扱う資料の章区切り | 標準 | 通常ページのヘッダーを使わない単独ページ |
| `matrix` | 2軸で選択肢をマス目に整理する（ポートフォリオ分析等）、または軸のない固定カテゴリの整理（SWOT等） | `rows`/`cols`（既定2×2）、`x_axis`/`y_axis`の有無 | マスごとのBackground Zone。各マスは独立編集可能 |
| `stat_highlight` | 単一の実績数値を主役に、補足指標と結論を示す。`stat`省略時は複数指標を均等グリッドで一覧するKPIダッシュボードになる | 標準 | 主指標＋補足指標Cardの集合 |
| `funnel` | 順を追って絞り込まれていく推移（ファネル分析、市場規模のTAM/SAM/SOM等） | 標準、`insights`指定で気づきの箇条書きを併記 | 値に応じた幅の帯を積む。各段は独立編集可能 |
| `cycle` | 繰り返し・循環するプロセス（PDCA等） | 標準（stepsは4〜6件目安） | 円周上のCard群＋隣接する矢印（最後尾から先頭へも繋ぎ輪にする） |

## Selection

- 両側を公平に読むなら `comparison: balanced`。
- 片側を補助領域として弱めるなら `comparison: asymmetric`。
- 結論が一つで根拠が複数なら `evidence_and_decision`。
- 順番だけでなく承認時点が重要なら `process_with_gates`。
- セル編集と行追加が想定される情報は `table_with_conclusion`。結論を1文に絞れず、表から複数の気づきを箇条書きで示したいなら `table_with_insight`（`chart_with_insight`のグラフを表に差し替えたもの。骨格は共通で、insightsに複数の気づきを列挙し、primary_messageで最後に結論を1文添える）。
- グラフ全体を公平に読むなら `chart_with_insight: standard`、一つの主張が主役なら `conclusion-led`。
  - 実データから作る棒・積み上げ棒・折れ線・円・散布図は`chart`（`type: column|stacked_column|bar|line|pie|scatter`）を使い、ネイティブPowerPointグラフにする（数値をPowerPoint上で直接編集できる）。パワーポイント生成側で用意したPNGしかない場合だけ`image`を使う。原則として`chart`を優先する。
    - `column`/`stacked_column`/`bar`/`line`/`pie`は`categories`と`series`（`[{"name": ..., "values": [...]}]`）を組み合わせるカテゴリ型。
    - `scatter`は`categories`不要で、`series`を`[{"name": ..., "points": [{"x": ..., "y": ...}, ...]}]`というXY座標形式にする（マーカーのみ、線で結ばない）。ただし実務でのグラフ調整は細部の作り込みを伴うことが多く、このrendererが担うのは「用意した図を説明に合わせて配置する」という基本的な作図まで。込み入った散布図・複合グラフは事前生成PNGを`image`で使う。
  - `add_native_chart`は生成時に不整合を検知して止める（categoriesとvalues数の不一致、`pie`への複数系列指定、`categories`が空、`scatter`のpointsが空またはx/y欠落、など）。負の値を含む棒・列グラフはラベル位置を自動調整し、カテゴリ数が多い（8件超）折れ線・棒グラフは点ごとのラベルを自動的に省略する（軸・目盛線での判読に任せる）。`stacked_column`のラベルはセグメント内（CENTER）に置く。
- 縦の責任階層と横の実行部門を両方示すなら `org_layers`。`layers`は2件までを目安にする（3件以上は各層の本文が収まらずpreflightで警告する。業務実行層は`execution`で別途横分割できるため、階層自体は「意思決定・運営」の2段+実行、のように抑える）。
- 課題と対応策を優先度付きで左右に並べるなら `priority_actions`。
- 現在から将来への段階的な広がりを示すなら `stage_track`。
- アジェンダやNext Stepなど、番号付きの単列項目だけで構成されるページは `numbered_list`。
- 複数テーマを扱う資料で章を区切るなら `section_divider`。表紙と混同しない（ロゴ・部署名・開示区分は表紙のみ）。
- 選択肢の位置づけをマス目に整理するなら `matrix`。軸上の座標が厳密な数値でなく、どのマスに属するかが要点の場合に使う。正確な座標が要点の場合は`chart_with_insight`の`chart.type: scatter`を使う（込み入った散布図はPNGの`image`）。`x_axis`/`y_axis`を指定すると連続軸上の位置づけ（BCGマトリクス等）、省略するとSWOT等「軸を持たない固定カテゴリ」になる（`cells`の構造はどちらも同じで、軸ラベル分の余白の有無だけが違う。x_axis/y_axisは両方指定するか両方省略する）。`rows`/`cols`は既定2×2で、GE-McKinseyの3×3等2×2を超えるマトリクスも同じ構造で指定できる（`cells`はrows×cols件ちょうど必要）。`table_with_conclusion`とは似て見えるが役割が違う: `matrix`は行・列の数自体に意味がある少数の軸/カテゴリへの位置づけ・分類（各マスはラベル＋見出し＋本文の塊）、`table_with_conclusion`は可変長の項目×評価基準の一覧比較（各セルは短い値）。項目数が増減しても構造の意味が変わらないなら`table_with_conclusion`、軸やカテゴリの数自体が意味を持つなら`matrix`を選ぶ。
- 実績・効果を1つの数値で語るなら `stat_highlight`（`stat`を指定）。複数の観点を対等に比較する場合は`table_with_conclusion`を優先する。複数指標を対等な重みでまとめて一覧したい（KPIダッシュボード）場合は、`stat`を省略し`supporting`だけを指定する（heroが無いだけで構造はstat_highlightと同じため、別rendererにしない）。数値が良い結果か悪い結果かを色でも示したい場合は、各項目に`tone: "positive"/"negative"/"warning"`を指定する（数値の符号では自動判定しない。詳細は[design-system.md](design-system.md)の「符号・重大度の色」）。
- 順を追って絞り込まれていく推移（リード獲得のファネル、市場規模のTAM/SAM/SOM等）を示すなら `funnel`。`stages`は値の大きい順に並べる。帯の幅はおおよその絞り込み具合を示す構造表現で、正確な比率を厳密に伝えたい場合は`chart_with_insight`の棒グラフを使う。段の値だけでなく、そこから読み取れる複数の気づきも合わせて示したい場合は`insights`を指定する（帯が右側の箇条書き分だけ幅を譲る。table_with_insight/chart_with_insightと同じ骨格）。`insights`を省略すると帯が全幅を使う。
- 一方向に進むのではなく繰り返し・循環するプロセス（PDCA、OODAループ等）を示すなら `cycle`。`steps`を円周上に時計回りで配置し、最後尾から先頭へも矢印で結んで輪にする（`stage_track`は一方向の進行、`cycle`は繰り返しが要点、という使い分け）。矢印は直線のみ（曲線の弧は環境間で見え方がぶれやすいため使わない）。`steps`は4〜6件を目安にする（多いと隣接するCard同士の間隔が狭くなり、preflightが警告する）。

## Escape Hatch

一致するrendererがない場合は無理に近い型へ入れない。`DeckBuilder.add_slide()` と `Region.columns()` / `rows()` を使って個別構築し、既存のtypography、semantic color、余白、編集性規則は維持する。新しいrendererを追加するのは、同じ意味構造が複数回現れ、決定論的な実装が再利用できる場合だけとする。slidekit内部のLayout/Atom/Fragment/Rendererという層構成は[`runtime/python/slidekit/ARCHITECTURE.md`](../runtime/python/slidekit/ARCHITECTURE.md)を参照。

項目数が少ないページを個別構築する場合は、`add_item_list` や `add_paragraph_textbox(vertical_anchor=...)` を使い、上詰めで余白が偏らないよう領域内で縦方向に配置を検討する。
