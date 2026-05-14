# docs/specs

word-kit が実装するに当たり参照する OOXML 仕様の「手元コピー」と、
プロジェクト内に書き溜める **実用的なメモ** を置くフォルダ。

仕様本体は非常に大きい (PDF だけで合計数千ページ) ため、ここでは:

- 対応する要素 / 属性のサブセット
- Word が実際に吐く XML パターン (理論より実装ベース)
- 意図的にサポートしない部分
- ECMA-376 PDF のどのセクションに対応するかの索引

を整理する。

スコープは `.docx` (WordprocessingML) のみ。pptx/xlsx は `CLAUDE.md` で Out of scope と宣言済み。

## 公式ソースの取得

`scripts/fetch-specs.sh` で ECMA-376 の PDF と XSD を取得する (本体は `.gitignore`)。

```bash
./scripts/fetch-specs.sh        # 全部
./scripts/fetch-specs.sh ecma   # ECMA-376 PDF のみ
./scripts/fetch-specs.sh xsd    # XSD のみ
./scripts/fetch-specs.sh ms     # Microsoft Open Specs の索引のみ
```

ダウンロード先:

| パス                       | 中身                                          |
| -------------------------- | --------------------------------------------- |
| `docs/specs/ecma-376/`     | ECMA-376 Part 1-4 (PDF + zip 内 XSD)          |
| `docs/specs/xsd/`          | XSD (Part 1+2 zip から展開)                   |
| `docs/specs/ms-openspecs/` | Microsoft Open Specifications の参照 URL 索引 |

## ECMA-376 ざっくり地図

| Part   | 内容                                       | word-kit からの主用途                                              |
| ------ | ------------------------------------------ | ------------------------------------------------------------------ |
| Part 1 | Fundamentals and Markup Language Reference | WordprocessingML / DrawingML の正式定義                            |
| Part 2 | Open Packaging Conventions (OPC)           | `[Content_Types].xml` / `.rels` / ZIP 規約                         |
| Part 3 | Markup Compatibility (MCE)                 | `mc:AlternateContent`, `mc:Ignorable` 等 (round-trip 完全性に必須) |
| Part 4 | Transitional Migration Features            | strict ↔ transitional の差分 (Word は通常 transitional)            |

実装に最も効くのは **Part 1** と **Part 2**。Part 3 は round-trip 完全性に必須。

### WordprocessingML 主要章 (Part 1 内)

- `WordprocessingML` 章 — `document.xml`, `styles.xml`, `numbering.xml`, `settings.xml`, `fontTable.xml`, `header*.xml`, `footer*.xml`, `comments.xml`, `footnotes.xml`, `endnotes.xml`
- `DrawingML - Main` 章 — `a:*` 要素全般 (色, 図形, ブリップ画像参照)
- `DrawingML - WordprocessingDrawing` 章 — `wp:inline`, `wp:anchor` (Word 固有の画像配置)
- `DrawingML - Charts` 章 — Word 内のグラフ (低頻度、後回し)

### XSD の使い方

`docs/specs/xsd/` の中の主要 XSD:

- `wml.xsd` — WordprocessingML (最重要)
- `dml-main.xsd` — DrawingML 本体
- `dml-wordprocessingDrawing.xsd` — Word 内 Drawing 配置
- `dml-chart.xsd`, `dml-chartDrawing.xsd` — チャート系 (後回し)
- `shared-*.xsd` — 共通型 (色, 単位, 関係)
- `opc-*.xsd` — OPC (Part 2)
- `vml-*.xsd` — VML (旧形式、`<w:pict>` フォールバックで使われる)

XSD は **正式な根拠** として使う (`unknown` 要素を判定するときの基礎ボキャブラリ、および XSD → TS 型 codegen の元)。

## Microsoft Open Specifications

ECMA-376 に無い、または Microsoft 製品が事実上拡張している要素は以下を参照:

- **MS-OE376**: Office Implementation Information for ECMA-376 (全体の補足)
- **MS-DOCX**: Word Extensions — `commentsExtended.xml` / `commentsIds.xml` / `peopleReferences.xml` 等の Word 拡張
- **MS-ODRAWXML**: Drawing 拡張
- **MS-OFFCRYPTO**: 暗号化 (現状サポート外)

(詳細 URL は `ms-openspecs/INDEX.md`。)

## 実装メモを増やす場所

実装中に「Word 特有のクセ」を発見したら以下に書き溜める:

```
docs/specs/
  notes/
    style-resolution.md      # docDefaults → basedOn chain → direct formatting の優先順位
    toggle-properties.md     # bold/italic 等の XOR セマンティクス
    numbering-resolution.md  # abstractNum + num + lvlOverride
    run-spanning-text.md     # run を跨ぐテキスト検索置換
    fields.md                # fldChar / instrText / fldSimple の挙動
    sections.md              # inline sectPr と末尾 sectPr
    headers-footers.md       # default/first/even の対応
    tracked-changes.md       # ins/del/move/*Change の accept/reject
    fonts-cjk.md             # rFonts の ascii/eastAsia/hAnsi/cs と hint
    mc-altcontent.md         # AlternateContent の取り回しと Ignorable
    twip-emu-units.md        # 単位系チートシート (Twip / HalfPoint / EMU)
    ...
```

体系的にではなく **遭遇順に** 書く。これが実装の生きた手帳になる。
