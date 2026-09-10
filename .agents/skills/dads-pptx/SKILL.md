---
name: dads-pptx
description: デジタル庁デザインシステム（DADS）に準拠したPowerPoint資料（.pptx）をpython-pptxで生成する。「DADS風のスライド」「デジタル庁デザインの資料」「行政・公共向けの提案書/勉強会資料をPPTXで」といった依頼、およびDADSのトークン（カラー・タイポグラフィ・余白・角の形状）に沿ったスライド作成時に使う。Generate DADS-compliant PPTX decks with verified contrast and layout.
---

# DADS準拠のPPTX作成

デジタル庁デザインシステム（DADS）β版の基本デザインをPowerPointへ写像し、
読みやすさとコントラストに配慮したスライドを生成する。
同梱トークンは `@digital-go-jp/design-tokens` v2.0.1 を基準とし、最新仕様への完全準拠や
PowerPoint上での表示・アクセシビリティを自動検証だけで保証しない。

以下の同梱ファイルのパスは、この `SKILL.md` があるディレクトリを基準に解決する。

同梱物：
- `scripts/dads_theme.py` … トークン定義と描画ヘルパ（**これをimportして使う。書き直さない**）
- `scripts/render_check.py` … 生成物を再読込して近似レンダリングし、はみ出しを検出
- `references/design-rules.md` … 守るべき設計ルールとレイアウトパターン集
- `references/dads-tokens.css` … カラートークンの実値（`@digital-go-jp/design-tokens` v2.0.1）
- `examples/build_example.py` … 4枚の最小構成サンプル（そのまま動く）

## 手順

### 0. 環境を用意する

既存のPython環境で `pptx` と `PIL` をimportできる場合は再利用する。未準備の場合のみ：

```bash
python3 -m venv .venv
.venv/bin/pip install python-pptx pillow
```

作業ディレクトリに `scripts/` と `out/` を作り、同梱の `scripts/dads_theme.py` と
`scripts/render_check.py` を作業先の `scripts/` にコピーする。既存ファイルがある場合は差分を確認し、
利用者の変更を上書きしない。同梱の `dads_theme.py` を再実装しない。

プレビュー用の `render_check.py` は次のLinuxフォントパスを参照する：

- `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc`
- `/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc`

実行前に存在を確認する。他の環境では作業先コピーの `FONT_R` / `FONT_B` を、
利用可能な日本語フォントの実パスに合わせる。PPTX側の指定書体は `Noto Sans JP`。

動作確認には同梱の `examples/build_example.py` を実行できる。
出力先を明示し、スキル内に生成物を書き込まない：

```bash
# リポジトリルートから実行する例。別の配置先ではスキルの実パスに置き換える。
mkdir -p out
.venv/bin/python .agents/skills/dads-pptx/examples/build_example.py out/example.pptx
```

### 1. 構成を決めてから書く

先に「1枚 = 1メッセージ」で見出しと要素を列挙する。1枚に詰め込みすぎない。
本文が読み物として長くなる場合はスライドを分ける。依頼の枚数・用途に合わせて、表紙・目次・本編・まとめを選ぶ。

**出典スライドを付ける**。末尾に次の2行を必ず入れる：

```
出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ のコンテンツを加工して作成
本資料はデジタル庁が作成・監修したものではありません。
```

### 2. 生成スクリプトを書く

`references/design-rules.md` のレイアウトパターンを組み合わせる。
座標は 1280 × 720 CSS px のキャンバスで指定し、`px()` がEMUへ変換する。

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))
from pptx import Presentation
from dads_theme import *

prs = Presentation()
prs.slide_width, prs.slide_height = px(CANVAS_W), px(CANVAS_H)
BLANK = prs.slide_layouts[6]
```

### 3. 検証する（省略しない）

```bash
.venv/bin/python scripts/build_pptx.py
.venv/bin/python scripts/render_check.py out/<名前>.pptx out/preview
```

`OVERFLOW` が出たら枠の高さか文字数を直し、再生成・再検証する。
このスクリプトは警告があっても終了コード0になるため、標準出力の警告件数を確認する。
フォント不足などで検証できない場合は原因と未確認事項を報告する。

### 4. 画像を見て直す

`out/preview/slideNN.png` をCodexの画像表示ツール（`view_image` など）で開き、全スライドを目視確認する。数値チェックでは分からない
「下半分が空く」「記号が豆腐になる」「角丸が潰れる」といった問題も調べる。
Pillowによる近似描画なので、PowerPointの実表示とは差があり得る。画像表示ツールが使えなければ、
目視確認済みとは書かず、プレビューの保存先と未確認事項を伝える。

納品時はPPTX、プレビュー、再生成用スクリプトの保存先と検証結果を簡潔に伝える。

## 必ず守るルール

| 項目 | ルール |
| --- | --- |
| テキストのコントラスト | 背景に対し 4.5:1 以上（白地なら Gray-536 `#767676` が下限） |
| 非テキストのコントラスト | 罫線・境界は 3:1 以上（白地なら Gray-420 `#949494` が下限。Gray-300は意味を持つ境界には**不可**（装飾用の補助線は可）） |
| 文字サイズ | 14 CSS px 未満は使わない。本文・UIは 16 CSS px 以上 |
| 書体 | `Noto Sans JP`（`dads_theme.FONT`）。ラテン・和文（`a:ea`）両方に適用済み |
| 色だけに依存しない | 意味は必ずラベル・記号・位置でも示す |
| エレベーション | 既定は高さレベル0。ドロップシャドウで境界を作らない（`shadow.inherit = False` 済み） |
| 出典表記 | 末尾スライドに出典と「デジタル庁が作成したものではない」旨を記載 |

色を追加・変更したら `contrast()` で必ず検算する：

```python
assert contrast(fg, bg) >= 4.5   # テキスト
assert contrast(line, bg) >= 3.0 # 罫線・境界
```

## ヘルパAPI

| 関数 | 用途 |
| --- | --- |
| `px(v)` / `fpt(css_px)` | CSS px → EMU / font-size CSS px → pt |
| `block(slide, x, y, w, h, lines, anchor)` | テキストブロック。`lines` は `dict(text=, style=, color=, align=, bullet=, space_before=, space_after=)` のリスト |
| `para(tf, text, style, ...)` | 既存テキストフレームへの段落追加 |
| `rect(slide, x, y, w, h, fill, line, line_px, radius)` | 矩形／角丸矩形 |
| `circle` / `rule` / `badge` / `arrow_right` | 円・罫線・チップ・三角矢印 |
| `contrast(a, b)` | コントラスト比の計算 |
| `STYLE` | DADSのテキストスタイル辞書（`"Std-32B-150"` など） |

`style` は必ず `STYLE` にあるトークン名を使う。無い組み合わせが必要なら
DADSの定義を確認し、生成スクリプト内で `STYLE` に**定義に存在するものだけ**追加する。
同梱トークンの実値は変更しない。

## よくある失敗

- **記号が豆腐になる**：`▸` `✕` `✓` はNoto Sans JPに字形がない。矢印は `arrow_right()`（図形）、
  バツは `×`（U+00D7）を使う。`▶` `※` `・` `→` は使用可。
- **角丸が潰れる**：カードの端に接するアクセントバー（左端・上端の帯）を置くなら、
  そのカードは `radius=RADIUS["none"]` にする。角丸カードに直角の帯を重ねると角だけ直角になる。
- **下半分が空く**：カード高さと行間を広げて紙面を使い切る。余白は意図的に、偏りは事故。
- **枠からはみ出す**：日本語は折り返し幅の見積りを誤りやすい。`render_check.py` の結果で調整する。
- **フォントが効かない**：`font.name` はラテンのみ。和文は `a:ea` が必要（ヘルパは対応済み）。
  自前でrunを組む場合は `from dads_theme import _apply_font` と明示的にimportして使う。
- **表を `add_table` で作らない**：DADSの表は横罫線ベース。`design-rules.md` の `datatable` パターンを使う。

詳細な設計値とレイアウトパターンは `references/design-rules.md` を読むこと。
