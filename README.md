# dads-pptx — デジタル庁デザインシステムでPowerPoint資料を作る

[デジタル庁デザインシステム（DADS）](https://design.digital.go.jp/dads/)β版の基本デザインに沿った
PowerPoint資料（.pptx）を、python-pptx で生成するためのツールキットとAIエージェント用スキルです。

- **DADSのトークンをそのまま使う** … カラー・タイポグラフィ・余白・角の形状を `dads_theme.py` に定義
- **アクセシビリティを検算しながら作る** … テキスト 4.5:1 / 非テキスト 3:1 のコントラスト比を `contrast()` で確認
- **はみ出しを検出する** … 生成したPPTXを近似描画し、枠からあふれたテキストを警告（`render_check.py`）
- **アニメーション版も作れる** … フェードのみ・クリック進行の表示を書き込み、検査する（`dads_anim.py`）
- **AIに頼んで作れる** … Claude Code / Codex 用のスキルを同梱

## スキルで作る（おすすめ）

### Claude Code

`/dads-pptx` で呼び出します。

```text
/dads-pptx 行政向けの説明資料を10枚程度で作ってください
/dads-pptx SQLの基本をクリックで順に表示するアニメーション付きで作ってください
```

スキルは `.claude/skills/dads-pptx/` にあります。追加直後に認識されない場合はセッションを開き直してください。
他のプロジェクトでも使う場合は、ユーザースコープへコピーします。

```bash
cp -r .claude/skills/dads-pptx ~/.claude/skills/
```

### Codex

`.agents/skills/dads-pptx/` に配置しています。

```text
$dads-pptx を使って、行政向けの説明資料をPowerPointで作成してください。
```

認識されない場合はCodexを再起動してください。配置先と呼び出し方は
[Codexのスキル仕様](https://learn.chatgpt.com/docs/build-skills)に基づきます。

## スクリプトで直接作る

### 環境

```bash
python3 -m venv .venv
.venv/bin/pip install python-pptx pillow
```

Windowsでは `.venv/bin/python` を `.venv\Scripts\python` に読み替えてください。

### 最小例

`scripts/` をimportパスに通し、出力先フォルダ（ここでは `out/sample/`）を作ってから実行します。

```python
import sys; sys.path.insert(0, "scripts")
from pptx import Presentation
from dads_theme import *          # scripts/dads_theme.py

prs = Presentation()
prs.slide_width, prs.slide_height = px(CANVAS_W), px(CANVAS_H)   # 1280 × 720 CSS px
s = prs.slides.add_slide(prs.slide_layouts[6])

rect(s, 0, 0, CANVAS_W, CANVAS_H, fill=WHITE)
block(s, MARGIN_X, 84, 1152, 48, [dict(text="見出し", style="Std-32B-150", color=INK)])
rule(s, MARGIN_X, 140, 1152, 2, BORDER)
prs.save("out/sample/sample.pptx")
```

動くサンプルはスキルに同梱しています。

```bash
mkdir -p out/example
.venv/bin/python .claude/skills/dads-pptx/examples/build_example.py out/example/example.pptx
.venv/bin/python .claude/skills/dads-pptx/examples/build_anim_example.py out/example/anim_example.pptx
```

### 検証

```bash
.venv/bin/python scripts/render_check.py out/<資料名>/<資料名>.pptx out/<資料名>/preview   # はみ出し検出＋プレビューPNG
.venv/bin/python scripts/dads_anim.py out/<資料名>/<資料名>.pptx                           # アニメーション版のみ
```

- `render_check.py` は警告があっても終了コード0です。標準出力の `OVERFLOW` の件数を確認してください。
- `render_check.py` は Linux のフォントパス（`/usr/share/fonts/opentype/noto/NotoSansCJK-*.ttc`）を参照します。
  他の環境では `FONT_R` / `FONT_B` を手元の日本語フォントに合わせてください。
- プレビューはPillowによる近似描画です。PowerPointの実表示と差が出ることがあるため、PNGを目で確認してください。

## 構成

| パス | 内容 |
| --- | --- |
| `scripts/dads_theme.py` | DADSのデザイントークンとPPTX描画ヘルパ（`px` / `block` / `rect` / `badge` / `contrast` など） |
| `scripts/render_check.py` | 生成物の近似描画とテキストのはみ出し検出 |
| `scripts/dads_anim.py` | アニメーション（フェード・クリック進行）の書き込みと検査 |
| `scripts/dads-tokens.css` | 参照したカラートークン（`@digital-go-jp/design-tokens` v2.0.1） |
| `scripts/build_*.py` | サンプル資料の生成スクリプト |
| `.claude/skills/dads-pptx/` | Claude Code用スキル（手順・設計ルール・ヘルパ・サンプル） |
| `.agents/skills/dads-pptx/` | Codex用スキル（同上＋UIメタデータ） |
| `docs/dads/` | DADSドキュメントのMarkdown版（AIの参照用） |
| `out/<資料名>/` | 生成した資料（PPTX・付属ファイル・`preview/`） |

## サンプル資料

このツールキットで作成した資料です。`out/` にあり、スクリプトから再生成できます。

| 資料 | 枚数 | 生成スクリプト | 特徴 |
| --- | --- | --- | --- |
| `out/AI駆動開発ガイド/` | 20 | `scripts/build_ai_development_guide.py` | 原稿Markdown・検証JSON・PDFも出力 |
| `out/SQLクロス集計/` | 13 | `scripts/build_sql_pivot.py` | アニメーション付き（コードの行強調・表の段階表示） |

## デザインの対応関係

- **キャンバス**: 1280 × 720 CSS px（1 CSS px = 1/96 inch）。`font-size` の CSS px は pt に 0.75 倍で対応（16 px = 12 pt）。
- **カラー**: プリミティブカラー Blue をキーカラーに設定（プライマリー Blue-900 / セカンダリー Blue-500 / ターシャリー Blue-1000 / 背景 Blue-50）。共通カラーは Solid Gray。
- **コントラスト**: テキストは 4.5:1 以上（本文 Gray-900、補足 Gray-536）、罫線などの非テキスト要素は 3:1 以上（Gray-420）。
- **タイポグラフィ**: Noto Sans JP。テキストスタイルは Dsp / Std / Dns / Oln のトークン定義（例: `Std-32B-150`）に対応。
- **余白**: 基準単位 8 CSS px の倍率スケール。左右マージンは 64（8 × 8）。
- **角の形状**: 角丸なし(0) / スモール(8) / ミディアム(12) を使用。端に接する装飾バーを持つカードは角丸なしに統一。
- **エレベーション**: 既定は高さレベル0。ドロップシャドウは使わず、境界はコントラスト比で確保。
- **アニメーション**: 使う場合はフェードのみ・クリック進行。全要素を表示した状態だけで内容が伝わるように構成。

詳しい設計ルールは [design-rules.md](.claude/skills/dads-pptx/references/design-rules.md) と
[animation.md](.claude/skills/dads-pptx/references/animation.md) を参照してください。

## ライセンス

本リポジトリのコードとドキュメントは [MIT License](LICENSE) で提供します。

ただし、次のものは各権利者の利用条件に従います（MIT Licenseの対象外）。

- `docs/dads/` および DADS のデザイントークン（`dads-tokens.css`）… デジタル庁デザインシステムの利用条件
- Noto Sans JP … SIL Open Font License 1.1

## 出典

出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ のコンテンツを加工して作成

本リポジトリはデジタル庁が作成・監修したものではありません。
