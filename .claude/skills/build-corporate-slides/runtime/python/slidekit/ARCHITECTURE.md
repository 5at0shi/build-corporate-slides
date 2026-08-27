# slidekit の層構成

`slidekit`は5層で構成する。下から上へ、Layout → Atom → Component → Fragment → Renderer の順に積み上がり、各層は自分より下の層だけに依存する（上の層を知らない）。この文書はその全体像を一箇所にまとめる唯一の場所。各モジュールのdocstringは自分の層だけを説明し、ここへ相互参照する。

## 層の一覧

| 層 | ファイル | 何を持つか | ビジネス語彙 |
|---|---|---|---|
| Layout | `layout.py` | `Region`。矩形の分割・inset計算のみ。描画は一切しない | 持たない |
| Atom | `atoms.py`、`typography.py`の一部、`icons.py`、`charts.py`、`tables.py`、`images.py` | これ以上分解できない描画の最小単位（1個の図形、1個のアイコン、1個のグラフ、1個の表、1枚の画像） | 持たない |
| Component | `components.py` | Atomのプリミティブ（主にBox）を人が指示に使える意味の通る名前でラップした部品群 | 部位名として持つ（「Background Zone」等） |
| Fragment | `fragments.py` | AtomとLayoutを組み合わせた、意味的にはまだ完結しない再利用可能な構造パターン（BoxGrid、ProportionalStack等） | 持たない（形だけの再利用） |
| Renderer | `renderers.py` | 上記すべてを組み合わせ、意味を持った1ページを完成させる17種の関数 | 持つ（「比較」「結論」「ゲート」等） |

## 各層の役割

**Layout** — `Region`は矩形領域そのものを表すデータであり、`rows()`/`columns()`/`inset()`で分割・余白計算をするだけ。色も線も文字も持たない。Atom以上のすべての層はRegion（またはx/y/w/h）を受け取って初めて描画位置を持てる。

**Atom** — 「これ以上分解できない」描画の最小単位。ビジネス語彙を持たず、renderer/Fragment/Componentのコードから呼ばれる実装内部の概念で、ユーザー/AIが直接名指しする対象ではない。`atoms.py`のBox/Marker/Connector/add_hairlineが中核だが、1個のアイコン（`icons.py`のadd_icon）、1個のネイティブグラフ（`charts.py`のadd_native_chart）、1個の表（`tables.py`のadd_data_table）、1枚の画像配置（`images.py`のadd_image_contain）も同じ意味でAtomに含める——複数の部品を組み合わせた「構造」ではなく、単一の描画操作である点が共通するため。`typography.py`の`add_textbox`/`set_run`/`style_text_frame`、`Tag`/`Stat`も同様にAtom。

**Component** — Atomのプリミティブを、人が指示に使える意味の通る名前でラップした部品群（`add_card`/`add_background_zone`/`add_focus_panel`/`add_key_message`/`add_section_lead`/`add_slide_title`/`add_cover`/`add_section_divider`/`add_panel`/`add_item_list`/`add_icon_list`）。「4ページ目のBackground Zoneをもう少し広げて」のようにユーザーが名指しする対象であり、capability-showcaseに実物を載せて公開しているのはこの層まで（Atom/Fragmentは公開しない。理由は下記「非公開の層」）。

例外的な事情として、`add_item_list`/`add_icon_list`は実体を`typography.py`の`add_text_list`（複数項目をまとめて描く、より複雑な処理）へ委譲している。つまり`typography.py`はAtom層の最小プリミティブとComponent層寄りの複合処理の両方を含んでおり、ファイル境界が層境界と完全には一致しない。これは意図的な設計ではなく、テキスト関連の処理をtypography.pyへ集約した結果生じた既知のずれとして残している。

**Fragment** — AtomとLayoutを組み合わせた、意味的にはまだ完結しない再利用可能な構造パターン（BoxGrid、ProportionalStack、RadialLayout、MarkerOverlay）。複数のrendererで同じ形が繰り返し必要になった構造をここに集約する。命名はビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。形だけで再利用できることがFragment層の価値であり、Componentと同様ユーザー/AIが直接名指しする対象ではない。

**Renderer** — Layout/Atom/Component/Fragmentを組み合わせ、意味を持った1ページを完成させる最上層。ここで初めてビジネス語彙が現れる。1関数=1renderer type=YAMLの`type:`フィールド1つに対応し、これがユーザー/AIが実際にcontent-authoringで触るインターフェース（[`renderer-catalog.md`](../../../references/renderer-catalog.md)）。

## 非公開の層（Atom / Fragment）

Atom層とFragment層は実装内部の語彙であり、user-guide/capability-showcaseにもreferences/にも実物を載せて公開しない。理由は、ユーザー/AIが実際に指示で使う語彙は「renderer type + そのフィールド」と「Componentの部位名」の2つに限られ、AtomやFragmentの名前（Box、BoxGrid等）を直接指示することは想定しないため。これは意図的な選択であり、抜け漏れではない。

## 層に属さない補助モジュール

以下はこの5層のどこにも属さない横断的な支援モジュール。
- `theme.py`: PALETTE（色）、TYPE系（タイポグラフィスケール）などのデザイントークン。全層から参照される。
- `builder.py`: `DeckBuilder`。config読み込み・パス解決・presentation生成を担う、renderer層の入口（オーケストレーション）。
- `config.py` / `preflight.py`: workspace設定の読み込み、content YAMLの事前検証。
- `textmetrics.py`: 文字幅・折り返し行数の概算ヒューリスティック。Component/Renderer層のレイアウト計算から呼ばれる。

## この文書の位置づけ

ここはコード（`runtime/python/slidekit/`）の実装アーキテクチャを説明する開発者向け文書であり、`references/`配下（PLAN/CREATE/REVISEでAIがcontent-authoringのために読むガイド）とは読者が異なる。層構成を変更した場合はこの文書と、変更した層のモジュールdocstringの両方を更新する。
