# はじめに

`build-corporate-slides`スキルの早見表。PLAN/CREATE/REVISEの流れなど使い方の詳細は`SKILL.md`を読む。ここでは「何が作れるか」と設定ファイルの中身だけをまとめる。

## まず見る: このスキルで作れるもの

文章より先に実物を見るのが早い。[`capability-showcase.pdf`](capability-showcase.pdf)に、配色・タイポグラフィ（3モード実寸）・基本コンポーネント・アイコン（14種）・ネイティブチャート（6種）・スライド構成（renderer、20種）の実物見本を1つのPowerPointとしてまとめてある。各ページには使用したPython関数名も添えてあるので、修正を頼むときの手がかりにもなる（例:「4ページ目のBackground Zoneをもう少し広げて」）。

見本に載っているのは各構成要素の最小構成の例であり、これだけが作れる形ではない。文言・配色・件数はもちろん、内容を具体的に説明すればAIが個別に対応できることが多い。

スライド構成（1ページ＝1つの意味構造）の一覧と使い分けは[`renderer-catalog.md`](../references/renderer-catalog.md)を見る。「比較のページを作りたい」のような依頼はそこに載っている型に対応する。

## 置き場所

スキルはプロジェクト直下（`.claude/skills/build-corporate-slides/`）にも、ホーム直下（`~/.claude/skills/build-corporate-slides/`）にも置ける。後者なら全プロジェクトから使える（同名が両方にある場合はホーム側が優先される）。

スキル本体の置き場所と作業領域は独立している。ホーム直下に1つ置いた場合でも、設定ファイルと`build_slides/`はプロジェクト側に作られるため、部署名・開示区分・ロゴ・生成物はプロジェクトごとに別になる。

## 設定ファイル（`.slide-skill-config.yaml`）

ワークスペース直下に置く。存在しない場合は[`.slide-skill-config.example.yaml`](.slide-skill-config.example.yaml)をコピーして使う。

| キー | 意味 | 例 |
|---|---|---|
| `python.executable` | 生成に使うPython。venvへのパス、またはPATH上のコマンド名 | `./.venv/bin/python`（Windowsは`./.venv/Scripts/python.exe`）、`python3` |
| `paths.input_dir` | 人間が渡す素材（グラフ画像・参考資料）の置き場所 | `./build_slides/input` |
| `paths.work_dir` | 生成コードと中間成果物（YAML、レンダリング結果） | `./build_slides/work` |
| `paths.output_dir` | 最終成果物（`deck.pptx`）の置き場所 | `./build_slides/output` |
| `organization.department` | 表紙・各ページに出す部署名 | `企画部` |
| `organization.classification` | 開示区分。表紙右上のバッジに表示 | `部外秘`、`社外秘`、空文字で非表示 |
| `deck.mode` | 資料全体の既定モード。`business`（標準）/ `dense`（情報量が多いページだけ）/ `large-room`（遠距離投影が明示された場合だけ）。実寸比較は[`capability-showcase.pdf`](capability-showcase.pdf)3ページ目 | `business`（通常はこれ） |
| `typography.headline_font` / `body_font` / `editorial_font` | 利用端末に導入済みの日本語フォント名 | `Hiragino Sans`、`Yu Gothic`など |
| `branding.logo.enabled` | ロゴを表紙に出すか | `true` / `false` |
| `branding.logo.path` | 正式ロゴ画像のパス。通常は指定不要（次段落） | `./build_slides/input/company-logo.png` |

正式ロゴは`branding.logo.path`を編集しなくても、`paths.input_dir`直下に`company-logo.png`という名前で置くだけで自動的に使われる。後で差し替える場合も同じファイル名へ上書きするだけでよい（`branding.logo.path`はこの既定の名前・場所を使いたくない場合だけ指定する）。どちらも無ければskill同梱の仮ロゴを使う。

`deck.mode`を変更しても個別ページのモードは変わらない。ページ単位で情報量が多い場合は、YAML側のそのスライドだけに`density: dense`を指定する（`deck.mode`自体は`business`のまま）。
