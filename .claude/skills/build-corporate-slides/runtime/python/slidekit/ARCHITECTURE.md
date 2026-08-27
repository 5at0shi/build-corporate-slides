# slidekit の層構成

`slidekit`は4層で構成する。下から上へ **Layout → Atom → Fragment → Renderer** の順に積み上がり、各層は自分より下の層だけに依存する（上の層を知らない）。層の名前はこの4つだけで、同義語を作らない。

## 4つの層

| 層 | 何を持つか | ビジネス語彙 | ファイル |
|---|---|---|---|
| **Layout** | 矩形領域(`Region`)の分割・余白計算。描画は一切しない | 持たない | `layout.py` |
| **Atom** | 1ページを構成する部品。単独の図形・文字・アイコン・グラフ・表・画像と、それらの名前付きプリセット・定型的な組み合わせ | 持たない（`add_key_message`等の部品名のみ） | `atoms.py` `typography.py` `icons.py` `charts.py` `tables.py` `images.py` `components.py` |
| **Fragment** | Region＋項目リストからN個の配置を計算し、中身を埋めるためのRegionを複数返す（中身は書かない） | 持たない | `fragments.py` |
| **Renderer** | 上記すべてを組み合わせ、意味を持った1ページを完成させる | **持つ**（「比較」「結論」「ゲート」等） | `renderers.py` `pageframe.py` |

### Layout

`Region`は矩形領域そのものを表すデータで、`rows()` / `columns()` / `inset()` で分割・余白計算をするだけ。色も線も文字も持たない。Atom以上のすべての層はRegion（またはx/y/w/h）を受け取って初めて描画位置を持つ。

### Atom

「1ページを構成する部品」の層。**この層の中に階層はない**。以下はすべて対等なAtomである。

- **単独の描画**: `Box` `Connector` `Marker` `add_hairline`（`atoms.py`）、`add_textbox` `add_paragraph_textbox`（`typography.py`）、`add_icon`、`add_native_chart`、`add_data_table`、`add_image_contain`
- **Atomの名前付きプリセット**: `add_card` / `add_background_zone` / `add_focus_panel` は、いずれも`Box`へ引数を変えて委譲するだけ（`components.py`）
- **少数のAtomの定型的な組み合わせ**: `add_section_lead`（Marker＋テキスト）、`add_key_message`（罫線またはBox＋テキスト）、`Tag`（Box＋テキスト）、`Stat`、`add_text_list`

ファイルの分け方は層ではなく**扱う媒体**による。図形は`atoms.py`、文字は`typography.py`、アイコンは`icons.py`、グラフは`charts.py`、表は`tables.py`、画像は`images.py`。図形と文字にまたがるもの、およびデザインシステム上の決定（tone / style）を名前へ畳み込んだものは`components.py`（理由は後述）。

### Fragment

**1つのRegionと項目リストを受け取り、N個の配置を計算して「中身を埋めるためのRegion」を複数返す**層（`BoxGrid` `ProportionalStack` `RadialLayout` `MarkerOverlay`）。配置だけを担当し、各項目の中身は書かない（書くのはrenderer）。複数のrendererで同じ配置の計算が必要になったものをここへ集約する。

AtomとFragmentの境界は、次の一点だけで判定する。

| | 引数 | 戻り値 | 例 |
|---|---|---|---|
| **Atom** | 1つの位置 `(x, y, w, h)` | 完成した部品（図形など） | `add_card` `add_key_message` |
| **Fragment** | Region＋項目リスト | 配置後のRegion **複数** | `BoxGrid` `ProportionalStack` |

Atom層の`add_card`と`Box`が同じ層なのは「引数を変えて委譲するだけで操作が同一」だから。対してFragmentは「N個の配置を計算する」という別種の操作なので層が分かれる（`fragments.py`は`components.py`をimportするが逆はない）。

判定が紛らわしい例:
- `add_panel`はRegionを返すがFragmentではない。1つしか返さず、項目リストも取らない（1:1であって1:Nではない）。
- `add_item_list` / `add_icon_list`は項目リストを取るがFragmentではない。中身を自分で描き切り、Regionを返さない。

命名にビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。形だけで再利用できることがこの層の価値。

### Renderer

Layout / Atom / Fragmentを組み合わせ、意味を持った1ページを完成させる最上層。**ここで初めてビジネス語彙が現れる。** 1関数 = 1 renderer type = YAMLの`type:`フィールド1つに対応し、これがユーザー・AIが実際に触るインターフェース（[`renderer-catalog.md`](../../../references/renderer-catalog.md)）。

`pageframe.py`（表紙・章扉・ページヘッダー）も、content_region()の内側に置く部品ではなくスライド1枚の外枠そのものを描くため、この層に属する。`DeckBuilder`経由でのみ呼ぶ（renderer層が直接importしないのは循環importを避けるため）。

## 名前の可視範囲（層とは別の軸）

層は「何の上に積まれているか」を決めるだけで、その名前を誰が呼んでよいかは決めない。可視範囲は次の3段で、いずれもコードから機械的に判定できる。

| 段 | 判定方法 | 意味 | 例 |
|---|---|---|---|
| モジュール内部 | `_`始まり | 同じファイルの中からのみ呼ぶ | `_filled_shape` `_flat` `_type_for` `_list_item_segments` |
| パッケージ内部 | `_`なし・`__all__`外 | slidekit内で層をまたいで呼ぶが、外へは出さない | `Box` `Connector` `Marker` `Tag` `Stat` `BoxGrid` `ProportionalStack` |
| パッケージ公開 | `__all__`収録（44名） | 生成スクリプト（`work/generate_pptx.py`）から呼んでよい | `add_card` `add_background_zone` `add_textbox` `Region` `DeckBuilder` `RENDERERS` |

renderer-catalog.mdの Escape Hatch（該当rendererが無いページを個別構築する）で使ってよいのは、最下段の`__all__`収録名だけ。`Box`や`BoxGrid`を生成スクリプトから直接呼ばないのは、この線を越えるため。

**capability-showcase.pdfに実物を載せているのはこれとは別で、`__all__`のうち「見た目として名指しできる部品」だけ**を選んでいる（`add_card`は載せるが`Region`や`load_workspace_config`は載せない）。これは人が「4ページ目のBackground Zoneを広げて」と言うための語彙を示す**掲載範囲の判断**であって、設計上の制約ではない。Atom層・Fragment層のカタログを載せていないのも同じ理由（`BoxGrid`を名指しして指示する場面が無い）で、層の性質から来るものではない。

## `atoms.py`と`components.py`を分けている理由

どちらも同じAtom層だが、2つの理由で別ファイルにする。

1. **import順の制約（技術的な理由）**: `typography.py`が`atoms.py`をimportしている。`components.py`は図形と文字の両方を使う（例: `add_key_message` = 罫線またはBox＋テキスト）ため`typography.py`もimportする必要があり、`atoms.py`へは統合できない（循環importになる）。
2. **呼び出し方が違う（意味的な理由）**: `atoms.py`は**見た目の値**を引数で受け取る（`Box(..., fill=..., radius=..., line=...)`）。`components.py`は**意味の名前**で呼ぶ（`add_background_zone(tone="brand-soft")` / `add_key_message(style="editorial")`）。後者はデザインシステム上の決定を名前へ畳み込んだもので、値の羅列では置き換えられない。

`add_card`と`Box`が同じ層なのは変わらない。`add_card`は`Box`へ引数を変えて委譲するだけで、上下関係ではない。

## 層に属さない補助モジュール

以下は4層のどこにも属さない、横断的な支援モジュール。

- `theme.py`: `PALETTE`（色）、`TYPE`系（タイポグラフィスケール）などのデザイントークン。全層から参照される。
- `builder.py`: `DeckBuilder`。config読み込み・パス解決・presentation生成と、`pageframe.py`の呼び出しを担うオーケストレーション。
- `config.py` / `preflight.py`: workspace設定の読み込みと、content YAMLの事前検証。
- `textmetrics.py`: 文字幅・折り返し行数の概算ヒューリスティック。Atom層とRenderer層のレイアウト計算から呼ばれる。

## この文書の位置づけ

ここは`runtime/python/slidekit/`の実装アーキテクチャを説明する開発者向け文書であり、`references/`配下（PLAN/CREATE/REVISEでAIがcontent-authoringのために読むガイド）とは読者が異なる。層構成を変えた場合は、この文書と該当モジュールのdocstringの両方を更新する。
