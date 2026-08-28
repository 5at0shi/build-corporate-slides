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
| `stage_track` | 現在から将来への段階的な進行（期間の幅は問わない） | 標準 | 同格のステージCard群。既定でCard間を矢印で繋ぐ（`connectors: false`で非表示）。各段の中身は`body`（文章）と`items`（箇条書き）のどちらでも書ける |
| `numbered_list` | アジェンダ、依頼事項など番号付きの単列項目 | 標準 | 番号付き行の集合。上部導入文／下部結論のどちらかを選べる |
| `section_divider` | 複数テーマを扱う資料の章区切り | 標準 | 通常ページのヘッダーを使わない単独ページ |
| `matrix` | 2軸で選択肢をマス目に整理する（ポートフォリオ分析等）、または軸のない固定カテゴリの整理（SWOT等） | `rows`/`cols`（既定2×2）、`x_axis`/`y_axis`の有無 | マスごとのBackground Zone。各マスは独立編集可能 |
| `stat_highlight` | 単一の実績数値を主役に、補足指標と結論を示す。`stat`省略時は複数指標を均等グリッドで一覧するKPIダッシュボードになる | 標準 | 主指標＋補足指標Cardの集合 |
| `funnel` | 順を追って絞り込まれていく推移（ファネル分析、市場規模のTAM/SAM/SOM等） | 標準、`insights`指定で気づきの箇条書きを併記 | 値に応じた幅の帯を積む。各段は独立編集可能 |
| `cycle` | 繰り返し・循環するプロセス（PDCA等） | 標準（stepsは4〜6件目安） | 円周上のCard群＋隣接する矢印（最後尾から先頭へも繋ぎ輪にする） |
| `timeline` | 開始と終了が異なる複数の取り組みが並走する計画（ロードマップ、ガント） | 標準 | 共通の時間軸上に、期間ぶんの幅を持つ帯を並べる。帯ごとに独立編集可能 |
| `waterfall` | AからBへの変化を増減要因に分解する（ブリッジ図、EBITDAウォーク、予算差異） | 標準 | 基準値の棒＋増減の棒＋累計の連結線。棒ごとに独立編集可能 |
| `issue_tree` | 論点をMECEに分解する（イシューツリー、ロジックツリー） | 標準（3階層固定） | 根＋枝Card＋内訳の行。枝・内訳ごとに独立編集可能 |

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
- 現在から将来への段階的な広がりを示すなら `stage_track`。各段の中身は`body`（1つの文章）でも`items`（箇条書き）でも書ける（両方指定するとbodyが上、itemsがその下に並ぶ）。段階ごとに「やること」を複数並べる計画表は`items`を使う（文章に詰め込むより読み取りやすく、後からPowerPoint上で1行だけ足す編集もしやすい）。段階の**期間の幅**そのものを示したい場合は`timeline`を使う。
- アジェンダやNext Stepなど、番号付きの単列項目だけで構成されるページは `numbered_list`。
- 複数テーマを扱う資料で章を区切るなら `section_divider`。表紙と混同しない（ロゴ・部署名・開示区分は表紙のみ）。
- 選択肢の位置づけをマス目に整理するなら `matrix`。軸上の座標が厳密な数値でなく、どのマスに属するかが要点の場合に使う。正確な座標が要点の場合は`chart_with_insight`の`chart.type: scatter`を使う（込み入った散布図はPNGの`image`）。`x_axis`/`y_axis`を指定すると連続軸上の位置づけ（BCGマトリクス等）、省略するとSWOT等「軸を持たない固定カテゴリ」になる（`cells`の構造はどちらも同じで、軸ラベル分の余白の有無だけが違う。x_axis/y_axisは両方指定するか両方省略する）。`rows`/`cols`は既定2×2で、GE-McKinseyの3×3等2×2を超えるマトリクスも同じ構造で指定できる（`cells`はrows×cols件ちょうど必要）。**`cells`の並び順は左上から右へ、次の行へ、という読み順で、先頭の行が`y_axis.high`側（上）、左の列が`x_axis.low`側になる。** 2×2なら`cells[0]`が「y高×x低」、`cells[1]`が「y高×x高」、`cells[2]`が「y低×x低」、`cells[3]`が「y低×x高」。最重要の象限（多くのフレームワークで「両方が高い」）は`y_axis.high`側かつ`x_axis.high`側なので、2×2では`cells[1]`になる。上下を取り違えると意味が反転するため、`emphasis`を付ける前に必ず確認する。`table_with_conclusion`とは似て見えるが役割が違う: `matrix`は行・列の数自体に意味がある少数の軸/カテゴリへの位置づけ・分類（各マスはラベル＋見出し＋本文の塊）、`table_with_conclusion`は可変長の項目×評価基準の一覧比較（各セルは短い値）。項目数が増減しても構造の意味が変わらないなら`table_with_conclusion`、軸やカテゴリの数自体が意味を持つなら`matrix`を選ぶ。
- 実績・効果を1つの数値で語るなら `stat_highlight`（`stat`を指定）。複数の観点を対等に比較する場合は`table_with_conclusion`を優先する。複数指標を対等な重みでまとめて一覧したい（KPIダッシュボード）場合は、`stat`を省略し`supporting`だけを指定する（heroが無いだけで構造はstat_highlightと同じため、別rendererにしない）。数値が良い結果か悪い結果かを色でも示したい場合は、各項目に`tone: "positive"/"negative"/"warning"`を指定する（数値の符号では自動判定しない。詳細は[design-system.md](design-system.md)の「符号・重大度の色」）。
- 順を追って絞り込まれていく推移（リード獲得のファネル、市場規模のTAM/SAM/SOM等）を示すなら `funnel`。`stages`は値の大きい順に並べる。帯の幅はおおよその絞り込み具合を示す構造表現で、正確な比率を厳密に伝えたい場合は`chart_with_insight`の棒グラフを使う。段の値だけでなく、そこから読み取れる複数の気づきも合わせて示したい場合は`insights`を指定する（帯が右側の箇条書き分だけ幅を譲る。table_with_insight/chart_with_insightと同じ骨格）。`insights`を省略すると帯が全幅を使う。
- 開始と終了が異なる複数の取り組みが**並走する**計画を示すなら `timeline`。`periods`（時間軸の目盛り）と`rows`（`start`/`end`で占める区間を指定、1始まりでendを含む）で書く。`stage_track`・`process_with_gates`との使い分けは「帯の長さに意味があるか」: `timeline`だけが幅＝期間を表し、`stage_track`は段階の順序、`process_with_gates`は判断を下す時点を示す（どちらも長さに意味を持たせない）。判断ポイントも示したい場合は`process_with_gates`のページを別に立てる（1ページに両方を詰め込まない）。`periods`は8件までを目安にする（超えるとpreflightが警告する）。帯に収まらない長さの`title`は自動的に帯の右側へ回る。各段の中に複数の作業を並べたいだけで期間の幅が要らないなら、`stage_track`の`items`を使う。
- 「AからBへ、なぜその差が生まれたか」を要因に分解するなら `waterfall`（ブリッジ図）。`bars`を左から右へ順に読む形で並べ、開始値・終了値・小計には`kind: total`を付ける（基準線0からの絶対値として描く）。`kind`を省略した棒は直前までの累計に対する増減として宙に浮き、既定では`value`の符号で色が決まる（増=positive、減=negative）。コストのブリッジのように「減ることが良い結果」の場合は符号と意味が逆になるため、棒ごとに`tone: "positive"/"negative"/"warning"`で色を明示する（`stat_highlight`の`tone`と同じ語彙。数値の符号だけで良し悪しを決めない、というスキル共通の考え方）。`bars`は3件以上必須（開始・増減・終了が揃って初めて「変化の分解」になる。2件以下は単なる比較なので`comparison`や`chart_with_insight`を使う）。`value_label`を省略すると`value`から自動生成する（増減には符号を付ける）。単位付きの表記にしたい場合は`value_label`で指定する。合計の内訳を並べたいだけなら`chart_with_insight`の積み上げ棒を使う。
- 論点を階層的に分解して示すなら `issue_tree`。`root`（分解する問い）＋`branches`（第1階層の分解軸）＋各枝の`items`（内訳）で書く。階層は3段固定で、それ以上深い分解は章を分けて複数ページにする（1ページの幅に収まらず、読み手も枝を追えなくなるため）。`items`を1つも持たない場合は2段のツリーとして描き、枝を広く使う。線に矢印は付けない（時間や因果の流れではなく包含関係を示すため）。`branches`は5件まで、`items`の合計は12件までを目安にする（超えるとpreflightが警告する）。分解の結果として「どれを優先するか」まで示すなら`priority_actions`、2軸で位置づけるなら`matrix`へ切り替える。
- 一方向に進むのではなく繰り返し・循環するプロセス（PDCA、OODAループ等）を示すなら `cycle`。`steps`を円周上に時計回りで配置し、最後尾から先頭へも矢印で結んで輪にする（`stage_track`は一方向の進行、`cycle`は繰り返しが要点、という使い分け）。矢印は直線のみ（曲線の弧は環境間で見え方がぶれやすいため使わない）。`steps`は4〜6件を目安にする（多いと隣接するCard同士の間隔が狭くなり、preflightが警告する）。

## 共通フィールド

全typeで使える。個別typeのフィールドは[`.slide-content.example.yaml`](../user-guide/.slide-content.example.yaml)に全20型の実例がある。

| フィールド | 意味 | 既定 |
|---|---|---|
| `type` | renderer type（上表のいずれか） | 必須 |
| `title` | ページタイトル | 必須 |
| `primary_message` | そのページの中心メッセージ。`cover`と`section_divider`以外は**必須**（結論用の領域を常に確保するため、欠けているとpreflightがエラーで止める） | 必須 |
| `density` | `standard` / `dense` | `standard` |
| `id` | YAML上の識別子。描画には使わない | 任意 |
| `message_style` | `primary_message`の見た目。`editorial`（左に短い青線）/ `subtle`（淡いグレー面）/ `solid`（濃紺の面）/ `card`（枠線付き）/ `plain`（装飾なし） | typeごとの既定 |

`cover`のみ次も使える。通常はconfigの値が自動で入るため、そのページだけ変えたい場合に指定する。

| フィールド | 意味 |
|---|---|
| `classification` | 開示区分。configの`organization.classification`を上書きする（空文字でバッジ非表示） |
| `created` | 表紙の日付。省略時は生成日 |
| `eyebrow` / `subtitle` / `brand_side` / `brand_shape` / `brand_width` | 表紙の見出し上ラベル・副題・ブランド面の位置と形状 |

見出しラベルは各typeで上書きできる（省略時は括弧内の既定値）。

| type | フィールド（既定値） |
|---|---|
| `evidence_and_decision` | `evidence_heading`（`判断の根拠`）/ `decision_heading`（`推奨方針`）/ `decision_detail`（判断の補足文） |
| `scope_and_exclusions` | `scope_heading`（`対象範囲`）/ `exclusions_heading`（`対象外`）/ `period`（期間などの補足を1行で添える） |
| `org_layers` | `execution_heading`（`業務実行`） |
| `priority_actions` | `issues_heading`（`想定される課題`）/ `actions_heading`（`対応方針`）/ `top_priority`（`最優先`。この文字列と一致する`priority`を持つ項目を強調色にする） |
| `chart_with_insight` / `table_with_insight` / `funnel` | `insight_heading`（`読み取れること`） |
| `numbered_list` | `message_position`（`top`。`primary_message`を導入文として上に置くか、`bottom`で結論として下に置くか） |
| `stage_track` | `connectors`（`true`。段の間を矢印で結ぶ。`false`で矢印なしのCard群になる）/ `note`（帯の下に添える注記） |
| `timeline` | 各`rows`の`tone`（`brand`（既定は`brand`/`teal`/`neutral`の循環）/ `teal` / `neutral` / `positive` / `negative` / `warning`。帯の色を意味で選ぶ。未定義の名前は既定色へ落として生成は止めない） |
| `waterfall` | 各`bars`の`kind`（省略時は増減。`total`で基準線0からの絶対値）/ `value_label`（省略時は`value`から自動生成）/ `tone`（`positive`/`negative`/`warning`。既定は`value`の符号） |
| `issue_tree` | `root.label`（根の上に置く小見出し。省略可）/ `root.body`・`branches[].body`（補足文。省略可） |

## Escape Hatch

一致するrendererがない場合は無理に近い型へ入れない。`DeckBuilder.add_slide()` と `Region.columns()` / `rows()` を使って個別構築し、既存のtypography、semantic color、余白、編集性規則は維持する。新しいrendererを追加するのは、同じ意味構造が複数回現れ、決定論的な実装が再利用できる場合だけとする。slidekit内部のLayout/Atom/Fragment/Rendererという層構成は[`runtime/python/slidekit/ARCHITECTURE.md`](../runtime/python/slidekit/ARCHITECTURE.md)を参照。

項目数が少ないページを個別構築する場合は、`add_item_list` や `add_paragraph_textbox(vertical_anchor=...)` を使い、上詰めで余白が偏らないよう領域内で縦方向に配置を検討する。

### 個別構築ページをYAMLと同居させる

個別構築するページも、内容はYAMLへ残す（[content-model.md](content-model.md)）。生成スクリプトで描画関数を定義し、`render_deck`の`extra_renderers`へ渡す。

```python
def render_custom_diagram(builder, spec, page):
    slide, area = builder.add_slide(spec["title"], density=spec.get("density"), page=page)
    body, conclusion = area.rows([4.7, 0.62], gap=Inches(0.24))
    ...  # Region.columns()/rows()とadd_*で組み立てる
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"], style="subtle")

output, warnings = render_deck(content, ROOT,
                               extra_renderers={"custom_diagram": render_custom_diagram})
```

YAML側は他のページと同じ形で書く。

```yaml
  - id: special
    type: custom_diagram
    title: "独自の図解"
    primary_message: "このページの中心メッセージ"
    nodes: ["入力", "処理", "出力"]
```

渡したtypeはpreflightでも許容され、共通の契約（`title`・`primary_message`・`density`・文字量・項目数）は他のページと同じだけ検査される。個別構築ページのためにpreflightごと外す必要はない。

既存typeと同じ名前は受け付けない（`ValueError`になる）。デッキ単位で標準rendererの意味が黙って差し替わると、同じtypeが資料によって別の見た目になるため。renderer自体を変えたい場合はslidekit側を直す。
