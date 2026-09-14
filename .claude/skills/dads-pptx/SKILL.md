---
name: dads-pptx
description: デジタル庁デザインシステム（DADS）に準拠したPowerPoint資料（.pptx）をpython-pptxで生成する。「DADS風のスライド」「デジタル庁デザインの資料」「行政・公共向けの提案書/勉強会資料をPPTXで」といった依頼、およびDADSのトークン（カラー・タイポグラフィ・余白・角の形状）に沿ったスライド作成時に使う。「アニメーション付き」「クリックで順に表示」といったPPTXのアニメーション版の依頼にも使う（フェードのみ・クリック進行）。Generate DADS-compliant PPTX decks with verified contrast and layout, optionally with fade-in click animations.
---

# DADS準拠のPPTX作成

デジタル庁デザインシステム（DADS）β版の基本デザインをPowerPointへ写像し、
アクセシビリティ要件を満たしたスライドを生成する。

同梱物：
- `scripts/dads_theme.py` … トークン定義と描画ヘルパ（**これをimportして使う。書き直さない**）
- `scripts/render_check.py` … 生成物を再読込して近似レンダリングし、はみ出しを検出
- `references/design-rules.md` … 守るべき設計ルールとレイアウトパターン集
- `references/dads-tokens.css` … カラートークンの実値（`@digital-go-jp/design-tokens` v2.0.1）
- `examples/build_example.py` … 4枚の最小構成サンプル（そのまま動く）
- `scripts/dads_anim.py` … アニメーション（フェード・クリック進行）の書き込みと検査（アニメーション版のみ）
- `references/animation.md` … アニメーション版のルール・API・構成パターン
- `examples/build_anim_example.py` … アニメーション付き4枚のサンプル

## 手順

### 0. 環境を用意する

```bash
python3 -m venv .venv && .venv/bin/pip install python-pptx pillow
```

作業ディレクトリに `scripts/` を作り、スキルの `dads_theme.py` と `render_check.py` をコピーする。
`dads_theme.py` は編集しない（トークンとコントラスト検証の根拠が壊れる）。

### 1. 構成を決めてから書く

先に「1枚 = 1メッセージ」で見出しと要素を列挙する。1枚に詰め込みすぎない。
本文が読み物として長くなる場合はスライドを分ける。表紙・目次・本編・まとめ・出典で構成する。

**出典スライドは必須**（DADSの利用上の注意事項）。末尾に次の2行を必ず入れる：

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

出力は資料ごとに `out/<資料名>/` フォルダを作り、本体PPTX・付属ファイル（JSON・PDFなど）・`preview/` をまとめて置く。

```bash
.venv/bin/python scripts/build_pptx.py
.venv/bin/python scripts/render_check.py out/<資料名>/<資料名>.pptx out/<資料名>/preview
```

`OVERFLOW` が出たら枠の高さか文字数を直し、0件になるまで繰り返す。

### 4. 画像を見て直す

`out/<資料名>/preview/slideNN.png` をReadで**必ず目視確認**する。数値チェックでは分からない
「下半分が空く」「記号が豆腐になる」「角丸が潰れる」といった問題はここでしか見つからない。

## アニメーション版を作る場合

「アニメーション付きで」「クリックで順に出したい」と依頼されたときだけ行う。先に `references/animation.md` を読む。

1. 同梱の `scripts/dads_anim.py` も作業先の `scripts/` にコピーする（`dads_theme.py` と同じ場所に置く。編集しない）。
2. 効果は**フェードのみ・クリック進行**。飛び込み・回転・点滅・退場は使わない。
   全要素を表示した最終状態だけで内容が伝わるように作る（プレビューPNG・印刷は最終状態になる）。
3. 生成スクリプトで、見せる順に図形を `grab` で集めて `Timeline.click()` に渡し、各スライドの最後に `apply()` する。
   アニメーションのないスライドも `Timeline(slide).apply()` で画面切り替えを付けて揃える。

```python
from dads_anim import grab, Timeline, code_block

tl = Timeline(s)
with grab(s) as g:
    rect(s, ...); block(s, ...)
tl.click("カード1を表示", g.shapes)
tl.apply()
```

4. 通常の検証に加えてアニメーションを検査する。`ERROR` が出たら直す（終了コード1）。

```bash
.venv/bin/python scripts/dads_anim.py out/<資料名>/<資料名>.pptx
```

5. PNGは最終状態しか映らない。出力されたクリック数・発表者ノートのクリック手順と画像を突き合わせて順序を確認する。
   PowerPointで再生して動きを確認できない場合は「動作は未確認」と報告する。

サンプル：

```bash
.venv/bin/python .claude/skills/dads-pptx/examples/build_anim_example.py out/anim_example.pptx
.venv/bin/python .claude/skills/dads-pptx/scripts/dads_anim.py out/anim_example.pptx
```

## 必ず守るルール

| 項目 | ルール |
| --- | --- |
| テキストのコントラスト | 背景に対し 4.5:1 以上（白地なら Gray-536 `#767676` が下限） |
| 非テキストのコントラスト | 罫線・境界は 3:1 以上（白地なら Gray-420 `#949494` が下限。Gray-300は**不可**） |
| 文字サイズ | 14 CSS px 未満は使わない。本文・UIは 16 CSS px 以上 |
| 書体 | `Noto Sans JP`（`dads_theme.FONT`）。ラテン・和文（`a:ea`）両方に適用済み |
| 色だけに依存しない | 意味は必ずラベル・記号・位置でも示す |
| エレベーション | 既定は高さレベル0。ドロップシャドウで境界を作らない（`shadow.inherit = False` 済み） |
| アニメーション | 使う場合はフェードのみ・クリック進行。最終状態で内容が完結すること |
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
`dads_theme.py` の `STYLE` に**DADSの定義に存在するものだけ**追加する。

## よくある失敗

- **記号が豆腐になる**：`▸` `✕` `✓` はNoto Sans JPに字形がない。矢印は `arrow_right()`（図形）、
  バツは `×`（U+00D7）を使う。`▶` `※` `・` `→` は使用可。
- **角丸が潰れる**：カードの端に接するアクセントバー（左端・上端の帯）を置くなら、
  そのカードは `radius=RADIUS["none"]` にする。角丸カードに直角の帯を重ねると角だけ直角になる。
- **下半分が空く**：カード高さと行間を広げて紙面を使い切る。余白は意図的に、偏りは事故。
- **枠からはみ出す**：日本語は折り返し幅の見積りを誤りやすい。`render_check.py` の結果で調整する。
- **フォントが効かない**：`font.name` はラテンのみ。和文は `a:ea` が必要（ヘルパは対応済み）。
  自前でrunを組む場合は `_apply_font()` を使う。
- **表を `add_table` で作らない**：DADSの表は横罫線ベース。`design-rules.md` の `datatable` パターンを使う。

詳細な設計値とレイアウトパターンは `references/design-rules.md` を読むこと。
