# PowerPoint Production

## 実行環境

1. `.slide-skill-config.yaml` の `python.executable`
2. 設定がなければワークスペース直下の一般的なvenv
3. 現在利用可能なPython

この順で解決する。設定された実行ファイルが存在しない場合は別環境へ黙って切り替えず報告する。Python方式では `runtime/python/requirements.txt` を依存関係の基準にする。インストール許可が明示されていない限り、venv作成やpip installを行わない。

特定のプレゼン生成機能が利用できないことだけを理由に停止しない。利用可能な環境で編集可能なPPTXを安定して生成できる方法を選ぶ。設定済みPythonとpython-pptxが利用可能なら標準実装とする。

## 生成コード

各資料専用の `work_dir/generate_pptx.py` を残す。標準生成は `render_deck(content, ROOT)` を使い、configとDesign DNAを `DeckBuilder` から受け取る。意味ベースrendererに合わないページだけ、`DeckBuilder.add_slide()` が返す `Region` をcolumns/rowsへ分割して個別構築する。低水準の座標指定を資料全体で繰り返さない。

生成コードは次を満たす。

- `DeckBuilder.from_workspace()` または `render_deck()` でconfigとパスを解決する
- `runtime/python` をimport pathへ加え、slidekitを利用する
- 標準は `business`。ページ単位の `density: dense` は許容するが、資料全体の `large-room` は遠距離投影が明示された場合だけ使う
- 内容と意味構造は原則 `build_slides/work/slide_content.yaml`、座標・配色・描画判断は `build_slides/work/generate_pptx.py` に置く
- 出力先を `paths.output_dir/deck.pptx` とする
- 再実行して同じ成果物を上書き生成できる
- 一つの意味単位を一つの編集可能オブジェクトにする。箇条書き群や連続項目を見た目上の行ごとにboxへ分割しない
- `add_paragraph_textbox` などを使い、一つのtextbox内の段落とrunで階層を作る
- 表は可能な限りPowerPoint table、図は編集可能な図形、グラフは可能な限りネイティブchartにする（`chart_with_insight`の`chart`フィールド。`add_native_chart`が実体で、column/stacked_column/bar/line/pie/scatterに対応。事前生成PNGしかない場合だけ`image`を使う）

表紙では `DeckBuilder.add_cover()` を使い、configの部署名、開示範囲、日付、ロゴを自動反映する。ロゴをAI判断で省略しない。

```python
from slidekit import DeckBuilder

deck = DeckBuilder.from_workspace(workspace_root)
deck.add_cover(title, subtitle=subtitle)
```

`branding.logo.enabled: true` の場合の解決順は (1) `branding.logo.path` の明示指定 (2) `input_dir` 直下の `company-logo.png`（configを編集せずこの名前で置くだけで自動的に使われる） (3) Skill同梱の仮ロゴ。詳細は [image-handling.md](image-handling.md) を読む。

## 検証とレンダリング

`validate_pptx.py` はファイル破損、スライド数、空ページ、スライド外オブジェクト、極小テキスト（本文・表セルとも）、テキストが枠や表の行高からはみ出す可能性、薄い装飾図形（罫線・アクセントバー等）がテキストに重なっている可能性、短いtextboxの過剰分割を検出する。Visual QAの代替ではない。Skill更新時は `scripts/self_test.py` も実行する。

`render_and_check.py` はmacOSではKeynote、その他ではLibreOfficeを優先してPDF化し、Popplerの `pdftoppm` でページPNGを作る。WindowsではPowerPoint COMも利用できる。Keynoteの自動操作権限がない場合はLibreOfficeへフォールバックする。レンダラーがなければPPTX生成を無効扱いにせず、Visual QA未実施と明示する。
