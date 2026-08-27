# はじめに

`build-corporate-slides`スキルの簡易ガイド。詳細な設計原則は`SKILL.md`と他のreferencesを参照。ここは「最初に何をすればよいか」「何が作れるか」「設定ファイルの各項目は何か」だけをまとめた早見表。

## まず見る: このスキルで作れるもの

文章より先に実物を見るのが早い。[`capability-showcase.pdf`](capability-showcase.pdf)に、配色・タイポグラフィ・基本コンポーネント・アイコン（14種）・ネイティブチャート（6種：棒・積み上げ棒・横棒・折れ線・円・散布図）・スライド構成（renderer、16種）の実物見本を1つのPowerPointとしてまとめてある。各ページには使用したPython関数名も添えてあるので、修正を頼むときの手がかりにもなる。

スライド構成（1ページ＝1つの意味構造）の一覧と使い分けは[`renderer-catalog.md`](../references/renderer-catalog.md)を見る。「比較のページを作りたい」のような依頼はそこに載っている型に対応する。

## 全体の流れ

1. **PLAN** — 資料の目的・読み手・構成をAIと一緒に決める。まだPowerPointは作らない。
2. **CREATE** — 確定した構成をもとにPowerPointを生成する。`slides/work/slide_content.yaml`（内容）と`slides/work/generate_pptx.py`（生成コード）が残る。
3. **REVISE** — 修正はYAMLの文言・数値を直すか、生成コードのレイアウトを直して再生成する。手作業でPowerPointを直接編集することは基本的にしない（再生成すると消えるため）。

修正を頼むときは「4ページ目のBackground Zoneをもう少し広げて」のように、部位の名称を使うと伝わりやすい。名称の一覧は[`capability-showcase.pdf`](capability-showcase.pdf)の各部品のキャプション（関数名）を参照する。モードの違い（後述）は[`mode-guide.pdf`](mode-guide.pdf)で実物のサイズ比較ができる。

## 設定ファイル（`.slide-skill-config.yaml`）

ワークスペース直下に置く。存在しない場合は[`.slide-skill-config.example.yaml`](.slide-skill-config.example.yaml)をコピーして使う。

| キー | 意味 | 例 |
|---|---|---|
| `python.executable` | 生成に使うPythonの実行ファイル | `./.venv/bin/python`（Windowsは`./.venv/Scripts/python.exe`） |
| `paths.input_dir` | 人間が渡す素材（グラフ画像・参考資料）の置き場所 | `./slides/input` |
| `paths.work_dir` | 生成コードと中間成果物（YAML、レンダリング結果） | `./slides/work` |
| `paths.output_dir` | 最終成果物（`deck.pptx`）の置き場所 | `./slides/output` |
| `organization.department` | 表紙・各ページに出す部署名 | `企画部` |
| `organization.classification` | 開示区分。表紙右上のバッジに表示 | `部外秘`、`社外秘`、空文字で非表示 |
| `deck.mode` | 資料全体の既定モード。詳しくは次項 | `business`（通常はこれ） |
| `typography.headline_font` / `body_font` / `editorial_font` | 利用端末に導入済みの日本語フォント名 | `Hiragino Sans`、`Yu Gothic`など |
| `branding.logo.enabled` | ロゴを表紙に出すか | `true` / `false` |
| `branding.logo.path` | 正式ロゴ画像のパス。未指定時はskill同梱の仮ロゴを使う | `./slides/input/company-logo.png` |

`deck.mode`を変更しても個別ページのモードは変わらない。ページ単位で情報量が多い場合は、YAML側のそのスライドだけに`density: dense`を指定する（`deck.mode`自体は`business`のまま）。

## モードについて（3種類）

- **business（標準）**: 通常の社内資料はすべてこれ。個人PC閲覧・事前配布・画面共有が前提。
- **dense**: 表や比較など情報量が多い"ページだけ"に使う。資料全体をdenseにはしない。
- **large-room**: 大会議室・講演など遠距離投影が明示された場合だけ使う。通常の社内資料では使わない。

実際の文字サイズ・見た目の違いは[`mode-guide.pdf`](mode-guide.pdf)で確認できる。本文（body）は`business`で12ptを基準にしている（[`design-system.md`](../references/design-system.md)参照）。

## うまくいかないときは

- 生成前に`scripts/check_config.py`・`scripts/check_environment.py`で設定・依存関係を確認する。
- 生成後は`scripts/validate_pptx.py`で構造・文字サイズ・はみ出しを機械チェックできる。
- `scripts/render_and_check.py`でPDF・ページ画像を作り、実際の見た目を確認する。
