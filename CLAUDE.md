# CLAUDE.md

このリポジトリの作業ルールは **[AGENTS.md](AGENTS.md)** に集約している。まずそれを読むこと。
構成・環境・生成と検証の手順・DADSの遵守ルール・コード規約はすべてそこにある。

以下はClaude Code固有の補足のみ。

## スキル

DADS準拠PPTXの作成は `/dads-pptx`（`.claude/skills/dads-pptx/`）を使う。
スキルはトークン定義・設計ルール・ヘルパ・サンプルを同梱しており、`dads_theme.py` を書き直さない前提で作られている。

`.agents/skills/dads-pptx/` はCodex用の同等スキル。**片方だけ更新しない** — 内容を変えたら両方に反映する
（`scripts/dads_theme.py` / `render_check.py` / `references/dads-tokens.css` は3箇所で、
`references/design-rules.md` は2箇所でバイト一致を保っている）。

## プレビューの目視確認

`render_check.py` の数値チェックだけでは足りない。生成後は `out/preview/slideNN.png` を
**Readツールで全枚数開いて目視確認する**。レイアウトの空き・豆腐文字・角丸の潰れはここでしか見つからない。
Readで開いていないなら「確認済み」と書かない。

## 出力先

生成物は `out/` に置く。`work/` は作業用でgitignore、`.venv/` もgitignore。
スキルディレクトリ内に生成物を書き込まない。
