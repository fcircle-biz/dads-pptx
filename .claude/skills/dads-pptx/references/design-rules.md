# DADS → PPTX 設計ルールとレイアウトパターン

出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ のコンテンツを加工して作成

## 1. 座標系

- キャンバスは **1280 × 720 CSS px**（16:9）。`prs.slide_width = px(CANVAS_W)`。
- 1 CSS px = 1/96 inch。`px()` がEMUへ変換する。
- `font-size` の CSS px は **pt に 0.75 倍**で対応（16 px = 12 pt）。`fpt()` が変換する。
- 左右マージン 64（基準単位8の8倍）、コンテンツ幅 1152。

標準の縦配置（コンテンツスライド）:

| 要素 | y | 高さ |
| --- | --- | --- |
| カテゴリーラベル（キッカー） | 52 | 26 |
| 見出し | 84 | 48 |
| ディバイダー（2px, Gray-420） | 140 | 2 |
| リード文（任意） | 152 | 30 |
| 本文領域 | 190〜648 | |
| フッター罫線 / フッター | 660 / 672 | 1 / 20 |

## 2. カラー

キーカラーはプリミティブカラーの1色相から選ぶ。既定はBlue。

| 役割 | トークン | 値 | 用途 |
| --- | --- | --- | --- |
| プライマリー | Blue-900 | `#0017C1` | 見出しアクセント、強調、CTA相当 |
| セカンダリー | Blue-500 | `#4979F5` | 非テキストの副次的強調（テキストには使わない） |
| ターシャリー | Blue-1000 | `#00118F` | 濃い面、表紙背景 |
| バックグラウンド | Blue-50 | `#E8F1FE` | 注記・強調ブロックの地色 |
| 本文 | Gray-900 | `#1A1A1A` | 本文テキスト |
| 補足 | Gray-700 / Gray-536 | `#4D4D4D` / `#767676` | 補足・フッター（536が4.5:1の下限） |
| 境界 | Gray-420 | `#949494` | 罫線・カード枠（3:1の下限） |
| 補助線 | Gray-300 | `#B3B3B3` | 意味を持たない装飾のみ |

セマンティック：Success `#197A4B`(green-800) / Error `#EC0000`(red-800) /
Warning `#C74700`(orange-800) `#B78F00`(yellow-700)。

他の色相に変える場合は `references/dads-tokens.css` の実値を使い、
`contrast()` で 4.5:1 / 3:1 を検算してから採用する。

## 3. タイポグラフィ

書体は `Noto Sans JP`（等幅は `Noto Sans Mono`）。太さは N(400) と B(700) のみ。
スタイル名は `<種別>-<size><太さ>-<line-height>`。

| 用途 | トークン |
| --- | --- |
| 表紙タイトル | `Dsp-64B-140` / `Dsp-48B-140` |
| スライド見出し | `Std-32B-150` |
| カード見出し | `Std-24B-150` / `Std-20B-150` |
| 本文 | `Std-17N-170` / `Std-16N-170` |
| 情報密度優先 | `Dns-16N-130` / `Dns-14N-130` |
| ラベル・チップ | `Oln-16B-100` / `Oln-14B-100` |

14 CSS px 未満は使わない。本文・UIは16 CSS px以上。

## 4. 余白・角の形状・エレベーション

- 余白は基準単位 **8** の倍率（8 / 16 / 24 / 32 / 64）。カード内パディングは24が標準。
- 角丸：なし(0) / スモール(8) / ミディアム(正方形16・横長12) / ラージ(正方形32・横長24)。
  同じ形状スタイルでも図形サイズで印象が変わるため、横長は小さめの半径にする。
- エレベーションは既定で高さレベル0。影は使わず、境界はコントラスト比で確保する。

## 5. レイアウトパターン

### 見出しブロック + フッター

```python
def header(slide, kicker, title, lead=None):
    rect(slide, MARGIN_X, 52, 6, 26, fill=PRIMARY)
    block(slide, MARGIN_X + 16, 52, CONTENT_W - 16, 26,
          [dict(text=kicker, style="Oln-14B-100", color=PRIMARY)], anchor="m")
    block(slide, MARGIN_X, 84, CONTENT_W, 48,
          [dict(text=title, style="Std-32B-150", color=INK)])
    rule(slide, MARGIN_X, 140, CONTENT_W, 2, BORDER)
    if lead:
        block(slide, MARGIN_X, 152, CONTENT_W, 30,
              [dict(text=lead, style="Std-17N-170", color=INK_SUB)])

def footer(slide, no, deck_title):
    rule(slide, MARGIN_X, 660, CONTENT_W, 1, BORDER_SOFT)
    block(slide, MARGIN_X, 672, 700, 20,
          [dict(text=deck_title, style="Dns-14N-130", color=INK_MUTE)])
    block(slide, CANVAS_W - MARGIN_X - 200, 672, 200, 20,
          [dict(text=str(no), style="Dns-14B-130", color=INK_MUTE, align="r")])
```

### 表紙

濃い面（`TERTIARY`）に白文字。左端に `SECONDARY` の8px帯、タイトル下に120×4のアクセント線、
下部に細い罫線と補足情報。

### カード（横並び 2〜4列）

```python
cw = (CONTENT_W - 16 * (n - 1)) / n
x = MARGIN_X + i * (cw + 16)
rect(slide, x, y, cw, h, fill=SURFACE, line=BORDER, radius=RADIUS["m"])
```

上端・左端に帯を重ねる場合は `radius=RADIUS["none"]` にする。

### 表（横罫線ベース）

```python
def datatable(slide, x, y, w, cols, rows, row_h=56, head_h=48):
    widths = [w * c for c in cols["ratio"]]
    xs, acc = [], x
    for cw in widths:
        xs.append(acc); acc += cw
    rect(slide, x, y, w, head_h, fill=BG_KEY)
    for i, label in enumerate(cols["labels"]):
        block(slide, xs[i] + 16, y, widths[i] - 32, head_h,
              [dict(text=label, style="Dns-14B-130", color=PRIMARY)], anchor="m")
    rule(slide, x, y + head_h, w, 2, PRIMARY)
    cy = y + head_h + 2
    for r in rows:
        for i, cell in enumerate(r):
            block(slide, xs[i] + 16, cy, widths[i] - 32, row_h,
                  [dict(text=cell,
                        style=("Std-16B-170" if i == 0 else "Std-16N-170"),
                        color=(INK if i == 0 else INK_SUB))], anchor="m")
        rule(slide, x, cy + row_h, w, 1, BORDER_SOFT)
        cy += row_h + 1
    return cy
```

`python-pptx` の `add_table` は使わない（DADSの罫線設計と合わない）。

### プロセスフロー

見出し帯付きの角丸なしカードを横に並べ、間に `arrow_right(slide, cx, cy, size, BORDER)` を置く。

### 注記・強調ブロック

`fill=BG_KEY, line=None` のブロックに `TERTIARY` の本文。警告色を使う場合は
白地＋左端6pxのセマンティックバー＋枠線（`radius=RADIUS["none"]`）。

### 箇条書き

`block()` の各行に `bullet="・"` を渡す。ぶら下げインデントが設定され、
折り返し2行目も揃う。段落間を空けたいときは `space_after=8`。

## 6. 検証

```bash
.venv/bin/python scripts/render_check.py out/deck.pptx out/preview
```

- `OVERFLOW slideNN used=..px box=..px` は枠に対して本文が高い。枠を広げるか文字を削る。
- 0件でも必ずPNGを目視する。空きの偏り・記号の欠落・角丸の潰れは画像でしか分からない。
