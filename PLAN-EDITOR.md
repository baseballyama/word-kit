# word-kit エディタ実装計画 (PLAN-EDITOR)

> ステータス: **実装進行中（カバレッジ拡大フェーズ）**。`@office-kit/docx-editor`
> パッケージ、カバレッジ台帳（強制テスト付き）、SvelteKit UI（`site` の `/editor`）
> を実装。38 の editor テスト green、全 502 テスト green、svelte-check 0 エラー。
> ブラウザ実機で「編集 → AST → 有効 .docx → 再読込で反映」を検証済み（DrawingML
> 属性編集・**styles.xml のパートレベル編集**とも round-trip 検証0 を実機確認）。
>
> 現在のカバレッジ（`packages/editor/COVERAGE.md`）: 全 1198 要素中
> **edit 1190 / render 8 / preserve 0 / 未分類 0**（**99.3%**）。preserve は
> ついにゼロ — すべての要素が編集可能（1190）か描画対象（8: document/body/tab/
> sym/hyphen/cols 等の構造・インライン要素）に分類される。drawing 370/370・
> math 124/124・vml 61/61・docprops 30/30・wml 605/613（残り 8 は render）。
> コマンド 234。
>
> 2 段の raw-XML インスペクタで残りを一挙に編集可能化した:
>
> - **document インスペクタ**: 本文内にインラインの DrawingML/OMML/VML を編集
>   （document AST を通してフラッシュ）。
> - **パートレベルインスペクタ**: 任意の XML パート（fontTable/settings/styles/
>   numbering/comments/foot-/endnotes/headers/footers/webSettings/docProps）を
>   開き、全要素の属性・子要素を編集して保存時にそのパートを再シリアライズ。
>
> ライブラリに汎用プロパティ／パート編集エンジンを追加（公開）:
>
> - run/段落: `setRunOnOff`/`setRunValProp`/`setParagraphOnOff`/`setParagraphValProp`
> - 任意コンテナ: `setElementOnOff`/`setElementValProp`/`makePropsElement`（tcPr/trPr/tblPr/sectPr）
> - settings.xml: `setDocumentSettingOnOff`/`setDocumentSettingVal`（新規 part 生成含む）
> - style 定義: `setStyleOnOff`/`setStyleValProp`
> - numbering level: `setNumberingLevelVal`/`setNumberingLevelOnOff`
>   これにより rPr/pPr/tcPr/trPr/tblPr/sectPr/settings/style/lvl の on-off・単一値
>   プロパティを一挙に編集可能化した。
>
> 進捗の推移: edit 79 → 254 → 794 → **1190**（6.6% → **99.3%**）。preserve は
> 全て完全ロスレス round-trip 済み（= 全表現を「実現」できる水準は最初から 100%）
> だったが、2 段の raw-XML インスペクタ導入で preserve は **0** になった。未分類
> ゼロは常に強制。
>
> **普遍的エスケープハッチ**（`raw.*` / `rawpart.*` コマンド + `rawTrees`/
> `allRawElements`/`xmlParts`/`partRawTree` + ライブラリの `setElementAttr`/
> `getRawPartRoot`/`markRawPartDirty`）が「全表現を編集できる」ことを保証する。
> `raw.test.ts` で DrawingML 編集と styles.xml のパートレベル編集の round-trip を
> 実証。UI は `/editor` の「XML」トグルでインスペクタを開き、ドロップダウンで
> 「Document（インライン図形/数式）」か任意の XML パートを選んで属性を直接編集
> できる（実機で styles.xml 編集 → 保存 → 再読込で反映を確認済み）。
>
> 目標は「MS Office のような操作 UI」で、docx の全表現を扱える MS Office 互換の
> エディタ。本ドキュメントは **網羅性を保証する仕組み** を先に設計し、その上に
> 実装を積むための計画。

## 0. 現実認識（誠実なスコープ）

MS Word の WYSIWYG は Microsoft が数十年かけて作ったもので、「完全互換」を
文字通り一度に完成させることは不可能。本計画の要点は **「いつ完成したと言えるか」
を機械的に測定・強制できる仕組み**（カバレッジ台帳）を先に作り、その台帳の数値を
上げていく形で開発を進めること。これにより:

- **どの docx 表現も“黙って”落ちない**（未分類の要素があればテストが落ちる）。
- 「編集できる」「表示だけできる」「ロスレス保持のみ」を要素単位で明示する。
- 進捗が %（編集可能要素数 / 全要素数）で可視化され、途中終了の判断が主観に
  依らない。

`@office-kit/docx` は既に **未知要素を raw XML として完全ロスレス round-trip** する
（round-trip テスト群で実証済み）。したがって「全 docx 表現の保持」は最初から
`preserve` 水準で満たされており、そこから `render`（表示）→ `edit`（UI 編集）へ
段階的に格上げしていく。

## 1. 網羅性を保証する仕組み — Capability Ledger

### 1.1 基準データ源（ユニバース）

ECMA-376 の公式 XSD（`references/python-docx/ref/xsd/`）:

- `wml.xsd` … WordprocessingML 本体。**613 の要素宣言**。docx の中核。
- `dml-wordprocessingDrawing.xsd`, `dml-picture.xsd`, `dml-main.xsd`（画像/図形で docx が使う部分）
- `shared-math.xsd`（OMML 数式）
- `vml-*.xsd`（レガシー VML 図形）
- `shared-*.xsd`（コアプロパティ等）

これらは `references/`（submodule・非公開・publish されない）にあるため、
src から import はしない。代わりに **抽出スクリプトで要素ユニバースを JSON に
スナップショット**し、コミットする。テストはスナップショットに依存する
（CI で submodule が無くても走る）。

- `scripts/extract-ooxml-elements.mjs` — XSD を走査し、要素名・所属 complexType・
  namespace を抽出 → `packages/editor/src/capability/element-universe.json` を生成。

### 1.2 台帳（Ledger）

`packages/editor/src/capability/ledger.ts`:

```ts
type Disposition = "edit" | "render" | "preserve";
interface Capability {
  element: string; // 例 "w:p", "w:rPr/w:b"（属性グループは擬似要素で表す）
  group: FeatureGroup; // "text" | "paragraph" | "table" | ...
  disposition: Disposition;
  command?: CommandId; // edit の場合、対応コマンド
  notes?: string;
}
```

### 1.3 強制テスト（この仕組みの心臓）

`packages/editor/src/capability/coverage.test.ts`:

1. `element-universe.json` の全要素が台帳にエントリを持つ（**未分類 = 失敗**）。
2. `disposition: "edit"` の各エントリは、実在する `CommandId` を指す。
3. 各 `edit` コマンドは少なくとも 1 つの round-trip テストで実証されている。
4. カバレッジレポート `packages/editor/COVERAGE.md` を生成（edit/render/preserve/%）。

→ **新しい要素を扱い始めても、扱わなくても、必ず台帳に disposition を書かねば
テストが緑にならない。これが「機能の網羅を保証する仕組み」。**

## 2. エディタ・アーキテクチャ

原則: **`@office-kit/docx` が唯一のドキュメントモデル**（One way to do one thing）。
エディタは OOXML を自前で書かない。全編集操作は公開 API 経由でモデルを変更する。

`@office-kit/docx-editor`（フレームワーク非依存コア）:

| モジュール  | 役割                                                                               |
| ----------- | ---------------------------------------------------------------------------------- |
| `model`     | live な `Docx` を保持。clone による snapshot。                                     |
| `render`    | AST → 編集可能 HTML。各ノードに `data-wk-path`（block/inline/piece index）を付与。 |
| `selection` | DOM Range ⇔ 文書位置 `{path, offset}` の相互変換。                                 |
| `commands`  | コマンドレジストリ。各 `CommandId` は台帳から参照される。全て公開 API 経由。       |
| `history`   | clone snapshot による undo/redo。                                                  |

UI（`site/src/routes/editor/`）: MS Office 風リボン（Home / Insert / Layout /
References / Review / View）+ 編集キャンバス。Svelte 5。

## 3. フィーチャーグループと段階

`edit` に引き上げる順序（各段階でカバレッジ % が上がる）:

1. **text** — run 書式: bold/italic/underline/strike/size/color/highlight/fonts/vertAlign/lang
2. **paragraph** — alignment/indent/spacing/borders/shading/style
3. **structure** — 段落・改ページ・改行・タブ・記号の挿入/削除
4. **list** — bullet/numbered/多階層
5. **table** — 行列 CRUD/結合/罫線/shading/セル配置
6. **image** — 挿入/差し替え/inline·anchor
7. **style** — スタイル適用/追加/見出し
8. **section** — ページサイズ/余白/向き/段組/セクション区切り
9. **headerFooter** — ヘッダ/フッタ/ページ番号
10. **references** — 目次/脚注/巻末注/ブックマーク/ハイパーリンク/フィールド
11. **review** — コメント/変更履歴（accept/reject）
12. **advanced** — SDT/数式(OMML)/VML/カスタム XML（当面 render または preserve）

各要素は上記いずれかの `disposition` を必ず持つ。`edit` に届かないものは
`render`（プレビューで表示）か `preserve`（ロスレス保持）として台帳に明記する。

## 4. テスト戦略

- **coverage.test.ts** — 台帳の完全性を強制（§1.3）。
- **command round-trip** — 各コマンドについて「createDocx → コマンド → toUint8Array
  → openDocx → 検証」。
- **e2e（Playwright）** — リボン操作 → キャンバス反映 → ダウンロード docx を再読込検証。
- 既存の `validatePackage` で生成物が壊れていないことを保証。

## 5. 進め方

- PLAN.md（本体ライブラリ）とは別に本書で進捗管理。
- カバレッジ % を README/COVERAGE.md に掲示。
- 各段階完了ごとにコミット（Conventional Commits, changeset）。
