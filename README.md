# AI駆動開発 プレゼンテーション（DADS準拠）

デジタル庁デザインシステム（DADS）β版の基本デザインに沿って、
「AI駆動開発」をテーマにしたPPTX（16:9・全15枚）を生成します。

## 成果物

- `out/AI駆動開発_DADS.pptx` … 本体
- `out/preview/slideNN.png` … レイアウト確認用のレンダリング画像（近似）

## 構成

| ファイル | 内容 |
| --- | --- |
| `scripts/dads_theme.py` | DADSのデザイントークン（カラー／タイポグラフィ／余白／角丸）とPPTX描画ヘルパ |
| `scripts/build_pptx.py` | スライドの内容とレイアウト |
| `scripts/render_check.py` | 生成物をPillowで近似描画し、テキストのはみ出しを検出 |
| `scripts/dads-tokens.css` | 参照したカラートークン（`@digital-go-jp/design-tokens` v2.0.1） |
| `docs/dads/` | DADSドキュメント（Markdown版）の展開先 |
| `.claude/skills/dads-pptx/` | DADS準拠PPTXを作るためのスキル（手順・設計ルール・ヘルパ・サンプル） |
| `.agents/skills/dads-pptx/` | Codex用スキル（手順・設計ルール・ヘルパ・サンプル・UIメタデータ） |

## 再生成

```bash
python3 -m venv .venv
.venv/bin/pip install python-pptx pillow
.venv/bin/python scripts/build_pptx.py
.venv/bin/python scripts/render_check.py out/AI駆動開発_DADS.pptx out/preview
```

## スキル

Codex用は `.agents/skills/dads-pptx/` に配置しています。このリポジトリで次のように依頼できます。

```text
$dads-pptx を使って、行政向けの説明資料をPowerPointで作成してください。
```

認識されない場合はCodexを再起動してください。配置先と呼び出し方は
[Codexのスキル仕様](https://learn.chatgpt.com/docs/build-skills)に基づきます。
生成手順と依存関係は [Codex用SKILL.md](.agents/skills/dads-pptx/SKILL.md) を参照してください。

`.claude/skills/dads-pptx/` に、DADS準拠のPPTXを作る手順をスキルとして切り出しています。
`/dads-pptx` で呼び出せます（追加直後のセッションでは認識されないことがあるため、その場合はセッションを開き直してください）。

他のプロジェクトでも使う場合は、ユーザースコープへコピーします。

```bash
cp -r .claude/skills/dads-pptx ~/.claude/skills/
```

## デザインの対応関係

- **キャンバス**: 1280 × 720 CSS px（1 CSS px = 1/96 inch）。`font-size` の CSS px は pt に 0.75 倍で対応（16 px = 12 pt）。
- **カラー**: プリミティブカラー Blue をキーカラーに設定（プライマリー Blue-900 / セカンダリー Blue-500 / ターシャリー Blue-1000 / 背景 Blue-50）。共通カラーは Solid Gray。
- **コントラスト**: テキストは 4.5:1 以上（本文 Gray-900、補足 Gray-536）、罫線などの非テキスト要素は 3:1 以上（Gray-420）。`dads_theme.contrast()` で検証可能。
- **タイポグラフィ**: Noto Sans JP。テキストスタイルは Dsp / Std / Dns / Oln のトークン定義（例: `Std-32B-150`）に対応。
- **余白**: 基準単位 8 CSS px の倍率スケール。左右マージンは 64（8 × 8）。
- **角の形状**: 角丸なし(0) / スモール(8) / ミディアム(12) を使用。端に接する装飾バーを持つカードは角丸なしに統一。
- **エレベーション**: 既定は高さレベル0。ドロップシャドウは使わず、境界はコントラスト比で確保。

## 出典

出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ のコンテンツを加工して作成

本リポジトリの成果物はデジタル庁が作成・監修したものではありません。
