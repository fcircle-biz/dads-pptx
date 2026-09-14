# AGENTS.md

デジタル庁デザインシステム（DADS）β版に準拠したPPTXをpython-pptxで生成するリポジトリ。
このファイルはエージェント共通の作業ルール。Claude Code固有の事項は `CLAUDE.md` を参照。

## 構成

| パス | 役割 |
| --- | --- |
| `scripts/dads_theme.py` | DADSトークン（カラー/タイポ/余白/角丸）とPPTX描画ヘルパ。**単一の正** |
| `scripts/render_check.py` | 生成PPTXをPillowで近似描画し、はみ出しを検出 |
| `scripts/dads_anim.py` | アニメーション（フェード・クリック進行）の書き込みと検査 |
| `scripts/build_pptx.py` | 「AI駆動開発」15枚 → `out/AI駆動開発/` |
| `scripts/build_ai_development_guide.py` | 「AI駆動開発ガイド」20枚 → `out/AI駆動開発ガイド/` |
| `scripts/build_sql_pivot.py` | 「SQLクロス集計」13枚（フェードのアニメーション付き）→ `out/SQLクロス集計/` |
| `scripts/dads-tokens.css` | 参照したカラートークン実値（`@digital-go-jp/design-tokens` v2.0.1） |
| `.claude/skills/dads-pptx/` | Claude Code用スキル |
| `.agents/skills/dads-pptx/` | Codex用スキル（`agents/openai.yaml` 付き） |
| `docs/dads/` | DADS公式サイトのMarkdown版（139ファイル、AI参照用） |
| `out/<資料名>/` | 生成物。資料ごとにフォルダを分け、pptx / pdf / json と `preview/` を置く。コミット対象 |
| `work/` | 作業用。gitignore |

## 環境

`.venv/` に構築済み（Python 3.12 / python-pptx 1.0.2 / Pillow 12.3.0）。gitignoreなので無ければ作る。

```bash
python3 -m venv .venv && .venv/bin/pip install python-pptx pillow
```

`render_check.py` は次のフォント実パスに依存する。存在を確認してから実行する。

- `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc`
- `/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc`

## 生成と検証

```bash
.venv/bin/python scripts/build_pptx.py
.venv/bin/python scripts/render_check.py out/AI駆動開発/AI駆動開発_DADS.pptx out/AI駆動開発/preview
```

検証は省略しない。手順は3段。

1. **生成** … ビルドスクリプトを実行
2. **数値チェック** … `render_check.py`。`OVERFLOW` が出たら枠の高さか文字数を直し、0件になるまで繰り返す。
   **警告があっても終了コードは0**なので、標準出力の警告件数を必ず読む。
3. **目視確認** … `out/<資料名>/preview/slideNN.png` を全枚数開く。「下半分が空く」「記号が豆腐になる」
   「角丸が潰れる」は数値チェックでは検出できない。Pillowの近似描画なのでPowerPoint実表示とは差がある。

画像を開けない場合は「目視確認済み」と書かず、保存先と未確認事項を報告する。

## 守るルール

| 項目 | ルール |
| --- | --- |
| テキストのコントラスト | 4.5:1 以上（白地の下限は Gray-536 `#767676`） |
| 非テキストのコントラスト | 罫線・境界は 3:1 以上（白地の下限は Gray-420 `#949494`。Gray-300は不可） |
| 文字サイズ | 14 CSS px 未満は使わない。本文・UIは 16 CSS px 以上 |
| 書体 | `Noto Sans JP`。`font.name` はラテンのみなので和文は `a:ea` が必要（ヘルパは対応済み） |
| 色だけに依存しない | 意味はラベル・記号・位置でも示す |
| エレベーション | 高さレベル0。ドロップシャドウで境界を作らない |
| 出典表記 | 末尾スライドに出典と「デジタル庁が作成・監修したものではない」旨を記載（必須） |

座標は 1280 × 720 CSS px キャンバスで指定し `px()` でEMUへ変換。font-size は CSS px × 0.75 = pt。
左右マージン 64、余白は 8 の倍数。色を追加したら `contrast()` で検算する。

```python
assert contrast(fg, bg) >= 4.5   # テキスト
assert contrast(line, bg) >= 3.0 # 罫線・境界
```

## コード規約

- **`dads_theme.py` は編集しない。** トークン実値とコントラスト検証の根拠が壊れる。
  スタイルを足す必要があれば生成スクリプト側で `STYLE` に追加し、DADSの定義に存在する組み合わせだけにする。
- `dads_theme.py` / `render_check.py` / `dads_anim.py` / `dads-tokens.css` は `scripts/`・`.claude/skills/`・`.agents/skills/`
  の**3箇所に同一内容で存在する**（現在バイト一致）。片方だけ直すと乖離するので、変更したら全部に反映する。
- `design-rules.md` / `animation.md` は `.claude` / `.agents` の2箇所に同一内容で存在する。同様に同期する。
- アニメーションはフェードのみ・クリック進行。生成後に `scripts/dads_anim.py <pptx>` で検査する（`ERROR` は終了コード1）。
- 表は `add_table` で作らない。DADSの表は横罫線ベース。`design-rules.md` の `datatable` パターンを使う。
- 記号は `▸` `✕` `✓` を使わない（Noto Sans JPに字形がない）。矢印は `arrow_right()`、
  バツは `×`（U+00D7）。`▶` `※` `・` `→` は可。
- 端に接するアクセントバーを持つカードは `radius=RADIUS["none"]` にする。角丸と直角の帯は混ぜない。
- コメント・変数名・出力文字列は既存コードに合わせて日本語で書く。

## ヘルパAPI

`px` / `fpt` / `block` / `para` / `rect` / `circle` / `rule` / `badge` / `arrow_right` / `contrast` / `STYLE` / `RADIUS`。
詳細は `.agents/skills/dads-pptx/references/design-rules.md` と `scripts/dads_theme.py` を読む。

## DADS仕様の参照先

`docs/dads/` に公式サイトのMarkdown版がある。トークンやコンポーネント仕様を確認するときはここを引く。
スキル本体（SKILL.md / design-rules.md）は `docs/dads/` を参照していないので、
仕様の裏取りが必要な場合は明示的に開く。

- `docs/dads/MANIFEST.md` … 全ファイルの索引
- `docs/dads/foundations/` … カラー・タイポグラフィ・余白・角の形状・エレベーション
- `docs/dads/components/` … コンポーネント別仕様

## 出典

出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ のコンテンツを加工して作成

本リポジトリの成果物はデジタル庁が作成・監修したものではありません。
