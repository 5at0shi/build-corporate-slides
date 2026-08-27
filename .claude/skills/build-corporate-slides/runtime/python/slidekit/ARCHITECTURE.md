# slidekit の層構成

`slidekit`は4層で構成する。下から上へ **Layout → Atom → Fragment → Renderer** の順に積み上がり、各層は自分より下の層だけに依存する（上の層を知らない）。層の名前はこの4つだけで、同義語を作らない。

## 4つの層

| 層 | 何を持つか | ビジネス語彙 | ファイル |
|---|---|---|---|
| **Layout** | 矩形領域(`Region`)の分割・余白計算。描画は一切しない | 持たない | `layout.py` |
| **Atom** | 1ページを構成する部品。単独の図形・文字・アイコン・グラフ・表・画像と、それらの名前付きプリセット・定型的な組み合わせ | 持たない（`add_key_message`等の部品名のみ） | `atoms.py` `typography.py` `icons.py` `charts.py` `tables.py` `images.py` `components.py` |
| **Fragment** | AtomとLayoutを組み合わせた、意味的にはまだ完結しない再利用可能な構造パターン | 持たない | `fragments.py` |
| **Renderer** | 上記すべてを組み合わせ、意味を持った1ページを完成させる | **持つ**（「比較」「結論」「ゲート」等） | `renderers.py` `pageframe.py` |

### Layout

`Region`は矩形領域そのものを表すデータで、`rows()` / `columns()` / `inset()` で分割・余白計算をするだけ。色も線も文字も持たない。Atom以上のすべての層はRegion（またはx/y/w/h）を受け取って初めて描画位置を持つ。

### Atom

「1ページを構成する部品」の層。**この層の中に階層はない**。以下はすべて対等なAtomである。

- **単独の描画**: `Box` `Connector` `Marker` `add_hairline`（`atoms.py`）、`add_textbox` `add_paragraph_textbox`（`typography.py`）、`add_icon`、`add_native_chart`、`add_data_table`、`add_image_contain`
- **Atomの名前付きプリセット**: `add_card` / `add_background_zone` / `add_focus_panel` は、いずれも`Box`へ引数を変えて委譲するだけ（`components.py`）
- **少数のAtomの定型的な組み合わせ**: `add_section_lead`（Marker＋テキスト）、`add_key_message`（罫線またはBox＋テキスト）、`Tag`（Box＋テキスト）、`Stat`、`add_text_list`

ファイルの分け方は層ではなく**扱う媒体**による。図形は`atoms.py`、文字は`typography.py`、アイコンは`icons.py`、グラフは`charts.py`、表は`tables.py`、画像は`images.py`、そして**公開している部品の名前**は`components.py`にまとめる。

### Fragment

AtomとLayoutを組み合わせた、意味的にはまだ完結しない再利用可能な構造パターン（`BoxGrid` `ProportionalStack` `RadialLayout` `MarkerOverlay`）。複数のrendererで同じ形が繰り返し必要になった構造をここへ集約する。命名にビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。形だけで再利用できることがこの層の価値。

### Renderer

Layout / Atom / Fragmentを組み合わせ、意味を持った1ページを完成させる最上層。**ここで初めてビジネス語彙が現れる。** 1関数 = 1 renderer type = YAMLの`type:`フィールド1つに対応し、これがユーザー・AIが実際に触るインターフェース（[`renderer-catalog.md`](../../../references/renderer-catalog.md)）。

`pageframe.py`（表紙・章扉・ページヘッダー）も、content_region()の内側に置く部品ではなくスライド1枚の外枠そのものを描くため、この層に属する。`DeckBuilder`経由でのみ呼ぶ（renderer層が直接importしないのは循環importを避けるため）。

## 「コンポーネント」は層ではない

`components.py`の`add_card`や`add_background_zone`を**コンポーネント**と呼ぶが、これは層の名前ではなく、**Atom層のうちユーザー・AIへ公開している部品の呼び名**である。層（構造上の位置）と公開範囲（名前で指示できるか）は別の軸で、次のように直交する。

| | 公開する（コンポーネント） | 公開しない |
|---|---|---|
| **Atom層** | `add_card` `add_background_zone` `add_key_message` `add_icon`（14種）`add_native_chart`（6種）など | `Box` `Connector` `Marker` `set_run` など |
| **Fragment層** | なし | `BoxGrid` `ProportionalStack` など全部 |

公開しているものだけをcapability-showcaseに実物として載せ、「4ページ目のBackground Zoneを広げて」のように名指しできる語彙とする。`Box`や`BoxGrid`のような実装内部の名前をユーザーが指示に使うことは想定しない。これは意図した選択であり、抜け漏れではない。

なお`add_card`と`Box`は同じAtom層にある。両者の違いは「公開名を持つプリセットか、素のプリミティブか」だけで、上下関係ではない。

## 層に属さない補助モジュール

以下は4層のどこにも属さない、横断的な支援モジュール。

- `theme.py`: `PALETTE`（色）、`TYPE`系（タイポグラフィスケール）などのデザイントークン。全層から参照される。
- `builder.py`: `DeckBuilder`。config読み込み・パス解決・presentation生成と、`pageframe.py`の呼び出しを担うオーケストレーション。
- `config.py` / `preflight.py`: workspace設定の読み込みと、content YAMLの事前検証。
- `textmetrics.py`: 文字幅・折り返し行数の概算ヒューリスティック。Atom層とRenderer層のレイアウト計算から呼ばれる。

## この文書の位置づけ

ここは`runtime/python/slidekit/`の実装アーキテクチャを説明する開発者向け文書であり、`references/`配下（PLAN/CREATE/REVISEでAIがcontent-authoringのために読むガイド）とは読者が異なる。層構成を変えた場合は、この文書と該当モジュールのdocstringの両方を更新する。
