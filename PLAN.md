# word-kit 実装計画 (v1)

> ステータス: 設計フェーズ。実装着手前。
> スコープは **`.docx` (WordprocessingML)** のみ。pptx / xlsx は Out of scope (`CLAUDE.md` 参照)。
> OPC / DrawingML 層は **将来 pptx/xlsx へ転用可能** な形で設計するが、当面は docx を成立させることに集中する。

---

## 1. 目的と差別化

### 目的

**OOXML (ECMA-376 Part 1 — WordprocessingML) に準拠した、ブラウザと Node.js の両方で動く Word 編集ライブラリ** を作る。
最終的に Word の主要機能 (paragraph / run / table / styles / numbering / header / footer / section / image / comments / footnotes / fields / tracked changes / SDT / hyperlinks) を **読み書き** できることを目指す。

### 既存ライブラリとのギャップ

| 既存          | 言語   | 環境         | 生成 | テンプレ編集 | 自由編集 (DOM-like) | フル仕様 |
| ------------- | ------ | ------------ | ---- | ------------ | ------------------- | -------- |
| python-docx   | Python | サーバ       | ◎    | ○            | ○                   | ○        |
| dolanmiu/docx | TS     | Browser+Node | ◎    | ×            | ×                   | △        |
| docxtemplater | JS     | Browser+Node | △    | ◎ (テンプレ) | ×                   | △        |
| mammoth.js    | JS     | Browser+Node | ×    | × (読み取り) | ×                   | △        |
| docx-preview  | TS     | Browser      | ×    | × (描画)     | ×                   | △        |
| officegen     | JS     | Node のみ    | ○    | ×            | ×                   | ×        |
| Open-XML-SDK  | C#     | .NET         | ◎    | ◎            | ◎                   | ◎        |

→ **「ブラウザ+Node の両対応で、既存 .docx を壊さず読み込み、DOM ライクに自由編集 (find/replace, insert, style 適用, table 編集, header/footer, tracked changes) できる TypeScript ライブラリ」** はまだ無い。

具体的に既存に無いのは:

- **dolanmiu/docx**: 生成は強いが、**既存 docx を読み込めない** (open API なし)。
- **docxtemplater**: テンプレ {{placeholder}} 置換特化で、ブロック挿入・table 行追加などの汎用編集は弱い。
- **mammoth.js / docx-preview**: 読みっぱなし (HTML 化 / 描画) で書き戻せない。
- **python-docx**: 仕様カバレッジは高いが Python (ブラウザ不可)。

### 非ゴール

**現時点で外す:**

- レンダリング (PDF / HTML / 画像 出力)。`mammoth.js` / `docx-preview` の領分。
- docx 暗号化/復号 (MS-OFFCRYPTO)。後続フェーズ。
- バイナリ `.doc` (Word 97-2003) 互換。
- マクロ実行 (`vbaProject.bin` は pass-through で温存するが、解釈はしない)。

**永続的に外す:**

- Word アプリとの自動化 (COM 等)。
- `.pptx` / `.xlsx` 対応 (OPC/DrawingML は共有しても、上位 API は別ライブラリの責務)。

---

## 2. 設計原則

1. **Lossless round-trip を最優先する。**
   読み込んで何もせず書き出した結果が Word の「修復しますか?」を発生させてはならない。
   理想は XML レベルで意味的に等価。可能なら触れていないパートの **bytes ハッシュも一致** させる ("untouched parts" の検出)。
2. **Pass-through (素通し) を前提とする。**
   ライブラリが理解できない要素は raw XML ノードとして保持し、書き戻し時に元のまま吐く。
   これが無いとテンプレ docx の独自プロパティ (Word の MS 拡張、SDT のカスタム XML、サードパーティアドインの注釈など) が消える。
3. **AST はレイヤ化する。**
   - **L0: Raw XML AST** — 順序・名前空間プレフィックス・属性順・空白を保持する low-fidelity XML 表現。round-trip の番人。
   - **L1: Semantic AST** — `Document` / `Paragraph` / `Run` / `Table` / `Section` 等の意味モデル。編集 API はここを触る。
   - L1 ⇔ L0 の双方向変換を保つ。L1 に乗らない要素は L0 のまま L1 ノードに参照として保持。
4. **環境抽象化は I/O 境界だけ。**
   `fs` は `io-node`、`Blob`/`File` は `io-browser` に閉じる。コアは `Uint8Array` で完結。
5. **依存ゼロを志向する。**
   外部 npm 依存はミニマル (zip と XML 周辺のみ)。`Buffer`、`stream`、`path` などの Node 専用 API はコアから禁ずる。
6. **ライセンス的にクリーン。**
   参考実装はあくまで `references/` で読むだけ。コードコピーは禁止 (clean-room 実装)。

---

## 3. 技術スタック

| 領域           | 選定                                                                         | 理由                                        |
| -------------- | ---------------------------------------------------------------------------- | ------------------------------------------- |
| 言語           | TypeScript (strict)                                                          | 型で OOXML スキーマを表現する旨味が大きい   |
| パッケージ管理 | pnpm workspaces                                                              | monorepo 性能と決定的解決 (既存 scaffold)   |
| ビルド         | tsup (esbuild ラップ)                                                        | ESM-only, dts 同梱, 速い (既存 scaffold)    |
| テスト         | vitest                                                                       | ESM ネイティブ, in-browser テスト可         |
| Lint/Format    | oxlint + oxfmt (既存 scaffold)                                               | 高速、設定軽量                              |
| バージョニング | changesets (既存 scaffold)                                                   | monorepo の semver と changelog 自動化      |
| ZIP            | **fflate** (MIT)                                                             | 軽量・isomorphic。JSZip (GPLv3 dual) を回避 |
| XML            | **fast-xml-parser** を初期採用、将来 custom SAX に差し替え可能な抽象を設ける | 速度と round-trip 性のバランス              |
| ターゲット     | Node 20/22/24, モダンブラウザ (ES2022)                                       | Buffer 非依存・WebStreams 利用              |

---

## 4. パッケージ分割

```
packages/
  core/                @word-kit/core         # 公開エントリ。Docx.open / Docx.create
  opc/                 @word-kit/opc          # ZIP, [Content_Types].xml, .rels, parts
  ooxml-xml/           @word-kit/ooxml-xml    # 名前空間つき XML パーサ/シリアライザ
  ast/                 @word-kit/ast          # Raw + Semantic AST 型と traversal
  drawingml/           @word-kit/drawingml    # DrawingML: image/shape/transform/color (共通基盤)
  wordprocessingml/    @word-kit/wml          # WordprocessingML: document/styles/numbering/headers/footers/sections/comments
  parser/              @word-kit/parser       # OPC + XML -> AST 変換
  writer/              @word-kit/writer       # AST -> XML + OPC 書き戻し
  editor/              @word-kit/editor       # 高レベル編集 API (find/replace, insertParagraph, addImage, table 編集等)
  validator/           @word-kit/validator    # XSD/構造/Word 互換性チェック
  io-node/             @word-kit/io-node      # fs adapter
  io-browser/          @word-kit/io-browser   # Blob/File adapter
  fixtures/            (private)              # テスト用 .docx 集
  testing/             (private)              # round-trip helpers, snapshot 比較
```

依存関係:

```
core → editor → parser/writer → ast → ooxml-xml + opc
                                  ↑
                          drawingml + wordprocessingml が AST を拡張定義
io-node/io-browser は core から動的に参照される (peer 的)
```

`drawingml/` は **当面 docx 内 `<w:drawing>` で使う分だけ** 実装する。将来 pptx/xlsx に再利用できるよう、`wordprocessingml/` 側に閉じた DrawingML 拡張は置かない。

---

## 5. AST 設計の要点

### L0 (Raw XML AST)

```ts
type XmlNode =
  | {
      kind: "element";
      name: QName;
      attrs: Attr[];
      children: XmlNode[];
      preservedWhitespace?: boolean; // xml:space="preserve"
    }
  | { kind: "text"; value: string }
  | { kind: "cdata"; value: string }
  | { kind: "comment"; value: string }
  | { kind: "pi"; target: string; data: string };

type QName = { uri: string; local: string; prefix?: string };
type Attr = { name: QName; value: string };
```

ポイント:

- **要素の出現順は保持する** (WordprocessingML はスキーマ的に sequence 制約あり)。
- **属性順とプレフィックスは保持する** (`w:` `w14:` `w15:` `mc:` 等を保つと round-trip が安定する)。
- 名前空間は URI で正規化、prefix は表示用に保持。
- `xml:space="preserve"` の付いた `<w:t>` は **空白除去をしない**。

### L1 (Semantic AST: Document 抜粋)

```ts
type Document = {
  body: Body;
  styles: StylesPart; // word/styles.xml
  numbering?: NumberingPart; // word/numbering.xml
  settings: SettingsPart; // word/settings.xml
  webSettings?: WebSettingsPart;
  fontTable: FontTablePart;
  theme?: ThemePart; // word/theme/theme1.xml
  headers: HeaderPart[]; // word/header*.xml
  footers: FooterPart[]; // word/footer*.xml
  comments?: CommentsPart;
  commentsExtended?: CommentsExtendedPart; // MS 拡張: スレッド/メンション
  footnotes?: FootnotesPart;
  endnotes?: EndnotesPart;
  glossaryDocument?: GlossaryDocument;
  customXml?: CustomXmlPart[];
  embeddings?: EmbeddingPart[]; // OLE 等
  media?: MediaPart[]; // 画像
  vbaProject?: BinaryPart; // macros — pass-through のみ
  extras: PassThrough[]; // 未対応の part
};

type Body = {
  blocks: BlockNode[]; // paragraphs, tables, SDTs in document order
  sectPr?: SectionProperties; // 末尾セクションのプロパティ
};

type BlockNode =
  | Paragraph
  | Table
  | SdtBlock // <w:sdt> block-level
  | RawBlock; // pass-through

type Paragraph = {
  pPr?: ParagraphProperties; // style, alignment, indent, numbering ref, sectPr (inline)
  children: InlineNode[]; // runs, hyperlinks, SDTs, bookmarks, fields, revisions
};

type ParagraphProperties = {
  pStyle?: StyleId;
  numPr?: { numId: number; ilvl: number };
  alignment?: "left" | "center" | "right" | "both" | "distribute";
  indent?: { left?: Twip; right?: Twip; firstLine?: Twip; hanging?: Twip };
  spacing?: { before?: Twip; after?: Twip; line?: Twip; lineRule?: "auto" | "exact" | "atLeast" };
  shading?: Shading;
  borders?: ParagraphBorders;
  sectPr?: SectionProperties; // インライン sectPr (セクション末)
  // ... + tracked-change wrappers (pPrChange)
};

type InlineNode =
  | Run
  | Hyperlink
  | SdtInline
  | BookmarkStart
  | BookmarkEnd
  | FieldStart
  | FieldSeparator
  | FieldEnd // <w:fldChar> 系
  | SimpleField // <w:fldSimple>
  | InsRange
  | DelRange // <w:ins> / <w:del> (tracked changes)
  | RawInline;

type Run = {
  rPr?: RunProperties;
  pieces: RunPiece[];
};

type RunProperties = {
  rStyle?: StyleId;
  rFonts?: {
    ascii?: string;
    eastAsia?: string;
    hAnsi?: string;
    cs?: string;
    hint?: "default" | "eastAsia" | "cs";
  };
  bold?: ToggleProp; // toggle 系は special semantics
  italic?: ToggleProp;
  size?: HalfPoint; // 半ポイント値
  color?: ColorValue;
  highlight?: HighlightColor;
  underline?: { val: UnderlineStyle; color?: ColorValue };
  strike?: ToggleProp;
  vanish?: ToggleProp; // 非表示
  lang?: { val?: string; eastAsia?: string; bidi?: string };
  vertAlign?: "baseline" | "superscript" | "subscript";
  // ... + rPrChange
};

type RunPiece =
  | { kind: "text"; value: string; preserveSpace?: boolean } // <w:t>
  | { kind: "delText"; value: string } // <w:delText> (tracked-changes 削除)
  | {
      kind: "break";
      type?: "page" | "column" | "textWrapping";
      clear?: "none" | "left" | "right" | "all";
    }
  | { kind: "tab" } // <w:tab/> (not the char)
  | { kind: "drawing"; drawing: Drawing } // <w:drawing> — inline/anchor
  | { kind: "pict"; pict: Picture } // <w:pict> 旧 VML
  | { kind: "fieldChar"; type: "begin" | "separate" | "end"; dirty?: boolean }
  | { kind: "instrText"; value: string } // field instruction
  | { kind: "symbol"; font: string; char: string }
  | { kind: "footnoteReference"; id: number }
  | { kind: "endnoteReference"; id: number }
  | { kind: "commentReference"; id: number }
  | { kind: "noBreakHyphen" }
  | { kind: "softHyphen" }
  | { kind: "lastRenderedPageBreak" } // Word キャッシュ
  | RawRunPiece;

type Table = {
  tblPr?: TableProperties;
  tblGrid: { cols: Twip[] };
  rows: TableRow[];
};

type TableRow = {
  trPr?: TableRowProperties; // height, isHeader, cantSplit
  cells: TableCell[];
};

type TableCell = {
  tcPr?: TableCellProperties; // gridSpan, vMerge, borders, shading, width
  blocks: BlockNode[]; // セル内に段落/ネスト表
};

type SectionProperties = {
  type?: "continuous" | "nextPage" | "evenPage" | "oddPage" | "nextColumn";
  pgSz: { w: Twip; h: Twip; orient?: "portrait" | "landscape" };
  pgMar: {
    top: Twip;
    right: Twip;
    bottom: Twip;
    left: Twip;
    header: Twip;
    footer: Twip;
    gutter: Twip;
  };
  cols?: { num: number; sep?: boolean; equalWidth?: boolean; columns?: { w: Twip; space: Twip }[] };
  headerRefs: { type: "default" | "first" | "even"; rId: RelationshipId }[];
  footerRefs: { type: "default" | "first" | "even"; rId: RelationshipId }[];
  pgNumType?: { fmt?: NumberFormat; start?: number };
  titlePg?: boolean;
  lineNumType?: LineNumberType;
};
```

### 単位

- **Twip** (1/20 pt = 1/1440 inch) — 紙寸法・余白・行高さ・インデント
- **HalfPoint** (1/2 pt) — フォントサイズ
- **EMU** (1 inch = 914400 EMU) — DrawingML (画像サイズ等)
- **Eighth-of-a-point** — 罫線太さ

公開 API は `string` リテラル (`"1in"`, `"2cm"`, `"10pt"`, `"720twip"`) を解釈する helper を経由。

### Run-spanning text 問題

Word は `<w:r><w:t>{na</w:t></w:r><w:r><w:t>me}</w:t></w:r>` のように
run が文字列の途中で分割される。検索/置換は **段落単位で run をフラットに連結したビュー** を作って実行し、
位置 → (run index, run 内 offset) のマップを介して書き戻す。
Run 分割を再構成する際は **属性 (rPr) が同一の隣接 run を統合** する正規化を行う。

```ts
doc.findText("{{name}}").replaceWith("山田太郎");
```

挙動:

1. 全 paragraph をスキャン、各 `<w:t>` を連結したフラット文字列を作る
2. マッチ位置を取得 (`{{name}}` は通常 4-5 run に分割されている)
3. 範囲を新文字列で置換 → run 列を再構築 (元の rPr を継承、複数 rPr に跨る場合は **先頭 rPr を引き継ぐ** のがデフォルト)
4. 隣接 run が同一 rPr なら統合

**注意:** `<w:fldChar>`, `<w:tab/>`, `<w:br/>`, ブックマーク, comment range, ins/del で run 列が **テキスト以外の要素** を含むので、フラット化時はそれらを「アンカー」として位置に登録する (置換は文字範囲外を跨がない)。

### Toggle properties (bold/italic/strike etc.)

Word の bold/italic 等は **XOR トグル**。スタイル階層 (docDefaults → linked-paragraph-style → run-style → direct-formatting) を辿るとき、各レイヤの toggle 値を **XOR** する。

- レイヤ 1: bold = on (style)
- レイヤ 2: bold = on (direct)
- 結果: bold = off ! (XOR)

これは Microsoft Word のスタイルとして「強調を更に強調 = 通常に戻す」セマンティクス。`effectiveRunProperties()` の実装で正確に再現する必要がある。

### Style 解決アルゴリズム

```
1) docDefaults (rPrDefault + pPrDefault)
2) Paragraph style chain (basedOn を辿って最上位まで)
   - paragraph 側で linked character style を持つ場合は character 側も同時に解決
3) Numbering の pPr/rPr (numbering.xml の対応 ilvl)
4) Run の rStyle chain
5) 直接書式 (pPr/rPr の direct)
```

各層で:

- Toggle 系は XOR
- 値系 (size, color 等) は **後勝ち**
- 一部 (rFonts) は `ascii`/`eastAsia`/`hAnsi`/`cs` のサブ属性ごとに後勝ち

### Pass-through 仕組み

L1 ノードは `extras: PassThrough[]` を持つ。パーサが知らない子要素は **スロット位置** とともに保管:

```ts
type PassThrough = { slot: number; node: XmlNode };
```

Writer は L1 の既知要素を吐き終わった後、`slot` を見て間に挟む。これで XSD の sequence 制約に違反しない順序で書き戻せる。

---

## 6. OPC (Open Packaging Conventions) 層

`.docx` は ZIP。中身 (代表的):

```
[Content_Types].xml
_rels/.rels
word/document.xml
word/_rels/document.xml.rels
word/styles.xml
word/numbering.xml
word/settings.xml
word/webSettings.xml
word/fontTable.xml
word/theme/theme1.xml
word/header1.xml … header3.xml
word/_rels/header1.xml.rels
word/footer1.xml … footer3.xml
word/_rels/footer1.xml.rels
word/comments.xml
word/commentsExtended.xml          (MS 拡張)
word/commentsIds.xml               (MS 拡張)
word/peopleReferences.xml          (MS 拡張)
word/footnotes.xml
word/_rels/footnotes.xml.rels
word/endnotes.xml
word/glossary/document.xml         (Building Blocks)
word/media/image*.png|jpg|svg
word/embeddings/*.xlsx|docx|bin
word/vbaProject.bin                (マクロ有り docm)
customXml/item1.xml + itemProps1.xml
docProps/core.xml
docProps/app.xml
docProps/custom.xml
```

### API

```ts
const pkg = await OpcPackage.read(bytes);
pkg.parts; // Iterable<Part>
pkg.getPart("/word/document.xml");
pkg.addPart({ name: "/word/media/image1.png", contentType: "image/png", data });
pkg.relationships; // パッケージ rel
pkg.partRelationships("/word/document.xml"); // パート rel
const out = await pkg.write({ keepUntouchedBytes: true });
```

### round-trip の必須事項

- `[Content_Types].xml` の `<Default>` / `<Override>` の順序と内容を保持。
- `.rels` の `Id` (rId1, rId2…) は再採番せず、**削除→追加** 時のみ新規付与 (採番は連番でなく未使用最小値)。
- 触れていないパートは **入力 ZIP のバイト列をそのまま出力 ZIP に書き直す** (CRC・タイムスタンプを再計算しない経路を維持)。
- ZIP は **STORED または DEFLATE** のみ。`extra fields` (Unicode path, NTFS timestamps) は捨ててよい (Word は気にしない)。
- ファイル名は `/` 区切り、先頭の `/` を含めない (`word/document.xml` の形式)。
- 64-bit ZIP (ZIP64) は **read のみ対応**。書き出しは通常 ZIP。100MB 超の docx は珍しいが対応マージン。

---

## 7. WordprocessingML 機能の優先順位

| 優先 | 領域                         | 主要要素                                                                           | 備考                                                                |
| ---- | ---------------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| P0   | OPC + 空 docx                | `[Content_Types].xml`, `.rels`, 最小 `document.xml`                                | 生成と round-trip の基礎                                            |
| P0   | Document/Body 骨格           | `w:document`, `w:body`, `w:p`, `w:r`, `w:t`                                        | 最小フォーマット                                                    |
| P1   | テキスト書式                 | `w:rPr`: bold/italic/size/color/fonts/lang/highlight/underline                     | toggle 系の XOR セマンティクス、CJK の `w:eastAsia`                 |
| P1   | 段落書式                     | `w:pPr`: alignment/indent/spacing/border/shading                                   | スタイル参照と直接書式の優先順位                                    |
| P1   | run-spanning find/replace    | flat view + 位置マップ                                                             | テンプレ docx の基本機能                                            |
| P1   | 改行・タブ・特殊記号         | `w:br`, `w:tab`, `w:noBreakHyphen`, `w:softHyphen`, `w:sym`                        | 文字でなく要素なので扱い要注意                                      |
| P2   | テーブル                     | `w:tbl`, `w:tr`, `w:tc`, `w:tblGrid`, `gridSpan`, `vMerge`                         | セル結合の vMerge ("restart"/"continue") 解決                       |
| P2   | 画像                         | `w:drawing` → `wp:inline`/`wp:anchor` + `a:blip`, media part                       | rId 整合、size の EMU 変換                                          |
| P2   | リスト/箇条書き              | `numbering.xml`, `w:numPr`, `w:abstractNum`, `w:lvl`, `w:lvlText`, `w:lvlOverride` | レベル毎の番号書式・記号、ilvl と numId の関係                      |
| P2   | スタイル                     | `styles.xml`, `w:style`, `w:basedOn`, `w:link`, `w:next`, `w:docDefaults`          | basedOn チェーン解決                                                |
| P2   | ハイパーリンク               | `w:hyperlink` (external rel と internal anchor 両対応)                             | rel 連動                                                            |
| P3   | セクション・余白             | `w:sectPr`, `w:pgSz`, `w:pgMar`, `w:cols`, `w:type`                                | 「inline sectPr」 (段落末尾) の取り回し                             |
| P3   | ヘッダー/フッター            | `w:headerReference`, `w:footerReference`, `header*.xml`, `footer*.xml`             | default / first / even 3 種、画像参照 rel                           |
| P3   | コメント                     | `w:commentRangeStart/End`, `w:commentReference`, `comments.xml`                    | MS 拡張: `commentsExtended.xml` (スレッド), `commentsIds.xml`       |
| P3   | フィールド                   | `w:fldChar`, `w:instrText`, `w:fldSimple`                                          | `MERGEFIELD`, `PAGE`, `NUMPAGES`, `HYPERLINK`, `TOC`, `REF`, `DATE` |
| P3   | 脚注/巻末注                  | `w:footnoteReference`, `footnotes.xml`, endnotes 同等                              | デフォルト sep/contSep の特殊脚注                                   |
| P4   | tracked changes              | `w:ins`, `w:del`, `w:moveTo`, `w:moveFrom`, `w:pPrChange`, `w:rPrChange`, ...      | accept/reject API                                                   |
| P4   | SDT (構造化文書タグ)         | `w:sdt`, `w:sdtPr`, `w:sdtContent`                                                 | bound-data CRUD、Building Blocks ギャラリ                           |
| P4   | ブックマーク                 | `w:bookmarkStart/End`                                                              | overlapping bookmarks の扱い                                        |
| P4   | リビジョン保護               | `w:permStart/End`                                                                  | range protection                                                    |
| P5   | フレーム/テキストボックス    | `w:framePr`, `w:pict` (VML legacy), `mc:AlternateContent`                          | 旧形式の VML フォールバックも保持                                   |
| P5   | 数式 (OMML)                  | `m:oMath`, `m:oMathPara`                                                           | 当面 pass-through、将来編集                                         |
| P5   | カスタム XML パート          | `customXml/item*.xml`, `itemProps*.xml`                                            | pass-through                                                        |
| P5   | グロッサリ (Building Blocks) | `glossary/document.xml`                                                            | サブ docx 構造                                                      |
| P5   | OLE 埋め込み                 | `w:object`, `embeddings/`                                                          | pass-through                                                        |
| P5   | マクロ (vbaProject.bin)      | binary part                                                                        | pass-through のみ                                                   |
| P5   | 検証 (validator)             | XSD + Word 互換ヒューリスティクス                                                  | 別パッケージ                                                        |

---

## 8. Markup Compatibility (mc:AlternateContent)

Word は SDK 拡張用に `mc:AlternateContent` を活用する。

```xml
<mc:AlternateContent>
  <mc:Choice Requires="w14">…Office 2010+ 表現…</mc:Choice>
  <mc:Fallback>…Office 2007 互換表現…</mc:Fallback>
</mc:AlternateContent>
```

方針:

- パース時は **Choice と Fallback の両方** を保持。
- ライブラリは **理解できる Requires をホワイトリスト化** (`w14`, `w15`, `w16`, `w16cid`, `w16se`, `w16sdtdh`, `wp14`)。
- 編集時は **Choice (理解できれば) と Fallback の両方** を一貫して更新するか、Fallback を破棄し再生成するかをポリシー化。
- 書き出し時は両方を保持して出す。

`w:document/@mc:Ignorable` の値 (`"w14 w15 w16se ..."`) は **必ず保持**。これが壊れると Word は古い互換モードで開いてしまう。

---

## 9. 公開 API スケッチ

```ts
import { Docx } from "@word-kit/core";

// 1) 空から生成
const doc = Docx.create({ pageSize: "A4", margins: "normal" });
doc.body.appendParagraph("タイトル", { style: "Heading1" });
doc.body.appendParagraph("本文1段落目");
doc.body.appendTable([
  ["氏名", "スコア"],
  ["Alice", "90"],
]);
doc.body.appendImage(imageBytes, { width: "10cm", anchor: "inline" });

// 2) テンプレ読込
const tpl = await Docx.open(file); // file: Uint8Array | Blob | File
tpl.findText("{{name}}").replaceWith("山田太郎"); // run 跨ぎ対応
tpl.findText("{{date}}").replaceWith(new Date(), { format: "yyyy/MM/dd" });
tpl.tables[0].rows.add(["A", "B", "C"]); // 行追加
tpl.tables[0].row(0).cell(1).text = "X"; // セル直接書換
tpl.pictures[0].replace(imageBytes); // 画像差し替え (アスペクト維持)

// 3) スタイル
doc.styles.add({
  id: "MyAccent",
  type: "character",
  rPr: { bold: true, color: "1f497d" },
});
doc.body.paragraphs[0].runs[0].rPr = { rStyle: "MyAccent" };

// 4) ヘッダー / フッター / セクション
const sec = doc.sections[0];
sec.headers.default.appendParagraph("社内文書");
sec.footers.default.appendPageNumber({ align: "center" });
doc.body.insertSectionBreak({ type: "nextPage", pgSz: { orient: "landscape" } });

// 5) コメント
const c = doc.body.paragraphs[3].addComment({
  author: "山田",
  initials: "Y.Y.",
  text: "確認お願いします",
});
c.reply("確認しました", { author: "鈴木", initials: "S.S." });

// 6) tracked changes
doc.acceptAllRevisions(); // すべて受け入れ
doc.rejectAllRevisions(); // すべて拒否

// 7) 保存
const bytes: Uint8Array = await doc.toUint8Array();
const blob: Blob = await doc.toBlob();
await doc.writeFile("/tmp/out.docx"); // io-node のみ

// 8) 一段下のレイヤを直接触る
const xml = doc.parts.get("/word/document.xml")!.xml; // Raw XML AST
```

### 単位リテラル

公開 API は次を受ける:

- `"1in"`, `"2cm"`, `"5mm"`, `"10pt"`, `"720twip"`, `"914400emu"`
- 数値 (デフォルト解釈は **EMU**、ただし文脈依存; doc strings で明記)
- `u.inches(1.5)`, `u.cm(2)`, `u.pt(10)`, `u.twip(720)` の helper

---

## 10. 実装ロードマップ

### M0 (完了): リポジトリスケルトン

- README / CLAUDE.md / PLAN.md / references/ submodules / docs/specs/ / scripts/fetch-specs.sh
- pnpm workspace / changesets / oxlint / oxfmt / tsup / vitest の scaffold は既存

### M1: OPC 読み書き + 完全 round-trip

- `@word-kit/opc` を実装
- `OpcPackage.read(bytes)` / `OpcPackage.write({ keepUntouchedBytes })`
- ZIP は fflate
- `[Content_Types].xml` / `.rels` のパーサ・シリアライザ
- **アクセプタンス**: Word for Mac / Windows / Online / Google Docs / LibreOffice / Pages 出力の合計 **20 サンプル docx** を読込→書出 → 全てが Word で警告なく開く + LibreOffice headless で convert 成功

### M2: XML AST コア (L0)

- `@word-kit/ooxml-xml`
- パーサ抽象 + fast-xml-parser バックエンド
- writer (順序保存・名前空間プレフィックス保存・`xml:space="preserve"` 保持)
- **アクセプタンス**: 上の 20 サンプルの全 part を parse → serialize → 等価性テスト通過

### M3: WordprocessingML 最小 AST (parser + writer, read-only)

- `Document / Body / Paragraph / Run / Text` の L1 化
- `@word-kit/wml` の P0/P1 範囲: rPr/pPr の主要プロパティ
- pass-through スロット仕組み
- **アクセプタンス**: 20 サンプルの round-trip 後の AST が安定 (再パース→等価)
- 友人案 step 2-3 ("document.xml の読み取り" + "AST 化") を完了

### M4: Run-spanning find/replace

- `doc.findText(string | RegExp)` → match 集合
- `.replaceWith(string)` / `.delete()`
- Run 統合の正規化
- **アクセプタンス**: `{{placeholder}}` 系 100 ケース (分割パターン網羅、CJK 混在、tab/br/symbol 越え) で正解
- 友人案 step 4 を完了

### M5: 空 docx を生成

- `Docx.create()` で最小の 1 段落 docx
- 内蔵 `styles.xml` 雛形 (Normal / Heading1-9 など Word 既定の同名)
- **アクセプタンス**: 生成 docx が Word で警告なく開き、保存しても壊れない

### M6: ブロック挿入 + 段落 API

- `body.appendParagraph/insertParagraph`、`run.append`、`paragraph.runs[]`
- `pPr` の上書き (style, align, indent, spacing, numbering ref)
- `rPr` の上書き (bold/italic/size/color/fonts/lang/highlight)
- **アクセプタンス**: 既存テンプレに段落を入れて保存しても layout 崩れ無し
- 友人案 step 5 を完了

### M7: 画像

- `body.appendImage / paragraph.addImage`、`picture.replace(bytes)`
- 画像 part の追加と rId の自動採番
- 画像メタ (PNG/JPEG/GIF/BMP/TIFF/SVG) の sniff
- inline / anchor 両対応
- **アクセプタンス**: 50MB 級 docx を読み込み、3 枚画像差し替え → Word で表示確認
- 友人案 step 6 を完了

### M8: テーブル

- 行/列/セル CRUD、merge (gridSpan / vMerge)、罫線、shading、tblGrid
- `tables[0].rows.add(["A", "B", "C"])` API
- ネストされた table (cells in cells) を保持
- **アクセプタンス**: 5×5 結合あり罫線ありテーブルを生成し、Word で正しく表示
- 友人案 step 7 を完了

### M9: スタイル (`styles.xml`)

- Style chain (`basedOn`) の解決、`docDefaults`、`linked` style ペア
- Toggle property の XOR セマンティクス
- `styles.add()` / `styles.modify()`
- **アクセプタンス**: 既存 Word テンプレ (社内テンプレ等) のスタイルを正しく再現、style chain を編集後も Word 上で正しく適用
- 友人案 step 8 を完了

### M10: ナンバリング (`numbering.xml`)

- `abstractNum` / `num` / `lvl` / `lvlOverride` / `lvlText`
- 多階層リストの解決 (`%1.%2.` 形式)
- 番号付け再開 (`lvlRestart`)
- **アクセプタンス**: 5 階層多重リスト、bullet と numbered の混在、`startOverride` がすべて Word で正常表示
- 友人案 step 9 を完了

### M11: ヘッダー / フッター / セクション

- `headers/footers` part 追加と rId 連動
- default / first / even 3 タイプ
- セクション切り替え (continuous / nextPage / evenPage / oddPage)
- `pgSz` (横向き含む) / `pgMar` / `cols`
- インライン sectPr の取り回し
- **アクセプタンス**: 縦横混在の複数セクション docx を生成し、Word で正しく表示・ページ番号も連続
- 友人案 step 10 を完了

### M12: コメント / フィールド / 脚注 / 巻末注

- `comments.xml` + `commentsExtended.xml` のスレッド構造、author 一覧
- `fldChar`/`instrText`/`fldSimple` 一式、`MERGEFIELD`/`PAGE`/`NUMPAGES`/`HYPERLINK`/`DATE` の 5 種を編集 API 化
- `footnotes.xml`/`endnotes.xml` 追加/削除
- **アクセプタンス**: コメントスレッド付き docx の往復、MERGEFIELD 差替後 Word で正常表示
- 友人案 step 11 を完了

### M13: tracked changes

- `ins`/`del`/`moveTo`/`moveFrom`/`*Change` の AST 化
- `acceptAllRevisions()` / `rejectAllRevisions()` / 範囲指定 accept/reject
- **アクセプタンス**: 100 件以上の改訂が乗った契約書 docx を accept しても layout 崩れ無し
- 友人案 step 12 を完了

### M14: SDT / ブックマーク / 高度なフィールド

- SDT (Content Controls) の CRUD + 値バインド
- TOC 構築 helper (見出しからの自動生成)
- 数式 (OMML) は pass-through で温存

### M15: 仕上げ

- ブラウザ/Node IO 最適化 (Web Workers アダプタ提供)
- バンドルサイズ目標: core (zip/xml 込み) brotli 後 60KB 未満
- ベンチマーク (parse/edit/serialize の三段)
- ドキュメントサイト (typedoc + 実例)

### M16: 1.0.0 リリース

---

## 11. テスト戦略

### レイヤ

1. **ユニット**: 各パッケージで関数単位。vitest。
2. **AST round-trip プロパティテスト**: fast-check で AST をランダム生成 → write → read → 等価。
3. **ファイル round-trip**: `fixtures/` の実 docx を `parse → serialize → 再 parse` で AST 等価性を assert。
4. **Word 互換テスト**:
   - **LibreOffice headless** で開いて conversion error が無いことを確認 (CI で実行可能)
   - **Microsoft Word** (実物) は手動回帰 (リリース前)
5. **ビジュアル回帰** (M11 以降): LibreOffice で PDF/PNG 化 → snapshot diff (`reg-cli` または `playwright` + image snapshot)
6. **検索置換テスト**: M4 で 100+ パターン (CJK, tab/br 越え, ins/del 越え, comment range 越え)

### Fixtures

- 自作: 30 ファイル (各機能を網羅)
- 外部: Microsoft 公式サンプル、Google Docs export、LibreOffice export、Pages export
- ライセンス問題回避のため、複製コミットせず生成スクリプトで取得 (`scripts/fetch-fixtures.sh`、後で作る)

### CI

- GitHub Actions
- matrix: Node 20 / 22 / 24, macOS / Ubuntu
- ブラウザ: vitest + playwright (Chromium + WebKit + Firefox)
- LibreOffice headless: `soffice --headless --convert-to pdf out.docx` の zero exit を確認

---

## 12. ライセンス

- 本リポジトリ: **MIT** (LICENSE 配置済)
- 依存:
  - fflate (MIT) ✅
  - fast-xml-parser (MIT) ✅
- `references/` 内 OSS は **読むのみ・コピー禁止** (clean-room 実装)
- ECMA-376 仕様は ECMA International から自由配布、引用は OK だが大量複製は避ける
- MS Open Specifications は Microsoft OSP/CP の下で利用可

---

## 13. 開発フロー

1. PLAN.md を都度更新する (生きたドキュメント)
2. 新パッケージは小さく切り、`pnpm -F` で単体テスト可能に
3. PR ベース、changesets で changelog 自動化
4. コミットメッセージは Conventional Commits
5. 設計が分岐する局面では `docs/decisions/NNNN-*.md` に ADR を残す
6. PowerPoint/Excel 用 PR は明示的に拒否 (CLAUDE.md のスコープ規律)

---

## 14. 未解決事項 / 将来検討

- **XML パーサの選択**: fast-xml-parser の `preserveOrder` モードでも属性順や CDATA 保持に限界があれば、SAX ベースの自作パーサに切り替える。判断は M2 のベンチ後。
- **ストリーミング**: 100MB 超 docx (画像多数) を扱うときに all-in-memory が辛い場合、ZIP の central directory を遅延読み + パーツ単位の遅延 parse で対応するか M1 で検証。
- **暗号化 docx**: MS-OFFCRYPTO は仕様が分厚いので別フェーズ。当面はサポート外で例外を投げる。
- **TypeScript の型定義の作り方**: 数千要素の WordprocessingML を手書きは非現実的。XSD → TS の codegen を検討する (M2 後半)。
- **API の Immutable vs Mutable**: 編集 API は性能のために Mutable 寄せ。ただし `clone()` を必ず提供。
- **Web Workers**: ブラウザでメインスレッドを止めないため、ZIP/XML を Worker に逃がす公式アダプタを提供するか検討。
- **TOC 自動更新**: TOC フィールドは Word が再計算するため、generate 時に _cached result_ を正しく書くか、Word に再計算を要求する `<w:updateFields w:val="true"/>` を `settings.xml` に入れる戦略を選ぶ。
- **既存 scaffold の整合**: `package.json` / `pnpm-workspace.yaml` / `biome.json` 相当 (oxlint/oxfmt) は scaffold 済み。`packages/` には `core/` と `opc/` のスケルトンがある (PLAN との整合確認は M1 着手時)。

---

## 15. 参考リソース

すべて `references/` 配下に git submodule として配置済み (詳細は `references/README.md`)。
仕様 PDF/XSD は `scripts/fetch-specs.sh` で `docs/specs/` に取得。
