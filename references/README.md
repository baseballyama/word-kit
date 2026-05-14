# references/

word-kit 実装中に「読む」ために用意した、外部 OSS と参考資料の集積場所。
原則として **git submodule** (shallow, `--depth 1`) として配置する。

> **public パッケージには一切含めない。**
> `references/` 配下は lint / format / build / publish の対象外。
> `packages/` 配下のコードから **import してはならない** (clean-room 実装の原則)。

## 採用している submodule

| パス                                            | ライセンス     | バージョン (pinned) | なぜ参考にするか                                                                                                                                            |
| ----------------------------------------------- | -------------- | ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`python-docx/`](./python-docx)                 | MIT            | v1.2.0              | **docx の最も成熟した実装の 1 つ**。Document/Section/Paragraph/Run/Table API の設計、styles と numbering の解決アルゴリズム、image 取り回しが宝庫           |
| [`dolanmiu-docx/`](./dolanmiu-docx)             | MIT            | master              | **TS で docx を「書く」設計の最良例**。XML builder のクラス階層、依存ゼロ志向、Browser+Node ビルド、テスト構成を参考にする。読み取り側は無いので補完が必要 |
| [`docxtemplater/`](./docxtemplater)             | MIT            | master              | **JS でテンプレ docx を編集する事実上のデファクト**。{{placeholder}} 置換、loop、conditional の実装、image module 連携など run-spanning 編集の現実解        |
| [`mammoth.js/`](./mammoth.js)                   | BSD-2-Clause   | master              | **docx を読んで意味解釈する**先行例 (HTML 化が出口)。styles の "ユーザー目線" 解釈、numbering の表現、track changes の扱いを参照                            |
| [`docx-preview/`](./docx-preview)               | MIT            | master              | **docx をブラウザで描画する**先行例。要素ごとのスタイル適用ロジック (CSS マッピング) と、何が見えているかの対応表として使う                                 |
| [`officegen/`](./officegen)                     | MIT            | v0.6.5              | 古めだが docx/pptx/xlsx 一括対応の Node ライブラリ。Stream ベース書き出しと relationship 採番の素朴な実装が読みやすい                                       |
| [`jszip/`](./jszip)                             | MIT/GPLv3 dual | main                | ブラウザ+Node の ZIP 実装の事実上のリファレンス (本体は **fflate** を採用予定だが、エッジケース処理を参考にする)                                            |
| [`Open-XML-SDK/`](./Open-XML-SDK)               | MIT            | main                | **Microsoft 公式** の OOXML SDK (.NET / C#)。スキーマ駆動の型生成 (`gen/`, `generated/`)、XSD への準拠、Word 互換性パッチが宝庫                             |

合計 ~235MB (shallow)。フル clone はせず、`--depth 1` で履歴を切り詰めている。

## 採用しなかった候補と理由

| 候補                              | 採用しなかった理由                                                                  |
| --------------------------------- | ----------------------------------------------------------------------------------- |
| Apache POI (Java)                 | 巨大 (>500MB)、Java 文化圏で読みづらい。Open-XML-SDK で代替可能                     |
| docx4j (Java)                     | 同上、docx 寄りだが API が複雑                                                      |
| Aspose.Words                      | クローズドソース                                                                    |
| docx-templates (guigrpa)          | docxtemplater で十分。後で必要なら追加                                              |
| python-docx-template (Jinja2)     | テンプレ DSL は本ライブラリの責務外                                                 |
| node-docx / docx-builder 等       | 更新停止、または独自情報量が薄い                                                    |

必要が出てきたら後で追加する。

## どう読むか (推奨)

WordprocessingML を初めて触るなら、次の順で読むと早い:

1. **`python-docx/docs/`** — docx の概念モデル (Document / Section / Paragraph / Run / Table / Style / Numbering) を最初に頭に入れる
2. **`python-docx/src/docx/oxml/`** — 実際の XML 要素クラスのカタログ。要素名と Python 属性の対応が引ける
3. **`Open-XML-SDK/generated/DocumentFormat.OpenXml.Wordprocessing/`** — XSD 由来の正式な型と名前空間を確認
4. **`dolanmiu-docx/src/`** — TS でクラス階層 + builder を組む設計の参考 (生成側)
5. **`docxtemplater/es6/`** — テンプレ編集 (run-spanning 置換、loop) の実装パターン
6. **`mammoth.js/lib/`** — docx → HTML の解釈で、style chain や numbering の意味的取り回しを見る
7. **`docx-preview/src/`** — docx を実際にレンダリングするときの "何を CSS にマッピングするか"

## ライセンス上の注意

- これらのリポジトリは **「読む」専用**。アルゴリズムやアーキテクチャから学ぶのは構わないが、
  **コード片を `packages/` に貼り付けない**。clean-room 実装 (仕様 + 自分の設計だけを根拠に書く) を貫く。
- 各リポジトリのライセンス全文は配下の `LICENSE` を参照。
- `mammoth.js` は **BSD-2-Clause**。他は MIT 系。

## サブモジュール操作のメモ

```bash
# 全 submodule を pinned commit に同期
git submodule update --init --recursive

# 個別 submodule を最新へ追従
git -C references/python-docx fetch origin
git -C references/python-docx checkout origin/master  # 各 repo の default branch

# 外す
git submodule deinit -f references/<name>
git rm -f references/<name>
rm -rf .git/modules/references/<name>
```

`--depth 1` を維持しているので、ローカルでは履歴を遡れない。
特定コミットの履歴を見たい場合は `git -C references/<name> fetch --unshallow` で深堀りする。
