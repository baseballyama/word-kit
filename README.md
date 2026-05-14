# word-kit (working title)

OOXML (ECMA-376) に準拠した PowerPoint (`.pptx`) 編集ライブラリ。
ブラウザと Node.js の双方で動作し、空のプレゼンテーション生成、テンプレート読込、
DOM ライクな自由編集をサポートすることを目指します。

> 📌 リポジトリ名は `word-kit` ですが、当面のスコープは PresentationML (`.pptx`) です。
> WordprocessingML / SpreadsheetML への展開は後続フェーズで検討します
> (OPC / DrawingML 層は共通化できるため、設計はマルチフォーマットを前提にします)。

## ステータス

設計フェーズ。`PLAN.md` を参照してください。

## ディレクトリ

- `references/` — 参考にする OSS / 仕様のサブモジュール群
- `packages/` — (未着手) モノレポのワークスペース。`PLAN.md` のパッケージ分割に従って配置予定
- `PLAN.md` — 実装計画
