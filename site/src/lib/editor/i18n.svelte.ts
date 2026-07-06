/**
 * Lightweight i18n for the editor UI. A single reactive `current` locale drives
 * `t(key)`; because `t` reads the `$state`, any markup calling it re-renders when
 * the locale changes. Falls back to English for missing keys.
 */

export const LOCALES = {
  en: "English",
  ja: "日本語",
  es: "Español",
  fr: "Français",
  de: "Deutsch",
  zh: "中文",
} as const;

export type LocaleId = keyof typeof LOCALES;

export type MessageKey =
  | "tab.home"
  | "tab.insert"
  | "tab.layout"
  | "tab.references"
  | "tab.review"
  | "action.new"
  | "action.open"
  | "action.find"
  | "action.xml"
  | "action.download"
  | "action.undo"
  | "action.redo"
  | "find.find"
  | "find.replaceWith"
  | "find.replaceAll"
  | "find.close"
  | "group.font"
  | "group.text"
  | "group.paragraph"
  | "group.styles"
  | "group.insert"
  | "group.headings"
  | "group.pageSetup"
  | "group.sections"
  | "group.references"
  | "group.tracking"
  | "ins.paragraph"
  | "ins.table"
  | "ins.pageBreak"
  | "ins.link"
  | "ins.picture"
  | "layout.sectionBreak"
  | "layout.pageNumbers"
  | "ref.toc"
  | "ref.footnote"
  | "ref.endnote"
  | "ref.bookmark"
  | "review.newComment"
  | "review.acceptAll"
  | "review.rejectAll"
  | "style.normal"
  | "style.heading1"
  | "style.heading2"
  | "style.heading3"
  | "status.ready"
  | "status.editing"
  | "status.words"
  | "status.chars"
  | "xml.title"
  | "xml.documentSource"
  | "language";

type Dict = Record<MessageKey, string>;

const en: Dict = {
  "tab.home": "Home",
  "tab.insert": "Insert",
  "tab.layout": "Layout",
  "tab.references": "References",
  "tab.review": "Review",
  "action.new": "New",
  "action.open": "Open…",
  "action.find": "Find",
  "action.xml": "XML",
  "action.download": "Download .docx",
  "action.undo": "Undo",
  "action.redo": "Redo",
  "find.find": "Find",
  "find.replaceWith": "Replace with",
  "find.replaceAll": "Replace all",
  "find.close": "Close find",
  "group.font": "Font",
  "group.text": "Text",
  "group.paragraph": "Paragraph",
  "group.styles": "Styles",
  "group.insert": "Insert",
  "group.headings": "Headings",
  "group.pageSetup": "Page setup",
  "group.sections": "Sections",
  "group.references": "References",
  "group.tracking": "Tracking",
  "ins.paragraph": "Paragraph",
  "ins.table": "Table",
  "ins.pageBreak": "Page break",
  "ins.link": "Link",
  "ins.picture": "Picture",
  "layout.sectionBreak": "Section break",
  "layout.pageNumbers": "Page numbers",
  "ref.toc": "Table of contents",
  "ref.footnote": "Footnote",
  "ref.endnote": "Endnote",
  "ref.bookmark": "Bookmark",
  "review.newComment": "New comment",
  "review.acceptAll": "Accept all",
  "review.rejectAll": "Reject all",
  "style.normal": "Normal",
  "style.heading1": "Heading 1",
  "style.heading2": "Heading 2",
  "style.heading3": "Heading 3",
  "status.ready": "Ready.",
  "status.editing": "Editing…",
  "status.words": "words",
  "status.chars": "chars",
  "xml.title": "Raw XML — DrawingML / OMML / VML",
  "xml.documentSource": "Document — inline shapes / math / VML",
  language: "Language",
};

const ja: Dict = {
  "tab.home": "ホーム",
  "tab.insert": "挿入",
  "tab.layout": "レイアウト",
  "tab.references": "参照",
  "tab.review": "校閲",
  "action.new": "新規",
  "action.open": "開く…",
  "action.find": "検索",
  "action.xml": "XML",
  "action.download": ".docx をダウンロード",
  "action.undo": "元に戻す",
  "action.redo": "やり直し",
  "find.find": "検索",
  "find.replaceWith": "置換後",
  "find.replaceAll": "すべて置換",
  "find.close": "検索を閉じる",
  "group.font": "フォント",
  "group.text": "文字",
  "group.paragraph": "段落",
  "group.styles": "スタイル",
  "group.insert": "挿入",
  "group.headings": "見出し",
  "group.pageSetup": "ページ設定",
  "group.sections": "セクション",
  "group.references": "参照",
  "group.tracking": "変更履歴",
  "ins.paragraph": "段落",
  "ins.table": "表",
  "ins.pageBreak": "改ページ",
  "ins.link": "リンク",
  "ins.picture": "画像",
  "layout.sectionBreak": "セクション区切り",
  "layout.pageNumbers": "ページ番号",
  "ref.toc": "目次",
  "ref.footnote": "脚注",
  "ref.endnote": "文末脚注",
  "ref.bookmark": "ブックマーク",
  "review.newComment": "新しいコメント",
  "review.acceptAll": "すべて承認",
  "review.rejectAll": "すべて元に戻す",
  "style.normal": "標準",
  "style.heading1": "見出し 1",
  "style.heading2": "見出し 2",
  "style.heading3": "見出し 3",
  "status.ready": "準備完了。",
  "status.editing": "編集中…",
  "status.words": "単語",
  "status.chars": "文字",
  "xml.title": "生 XML — DrawingML / OMML / VML",
  "xml.documentSource": "ドキュメント — インライン図形 / 数式 / VML",
  language: "言語",
};

const es: Dict = {
  "tab.home": "Inicio",
  "tab.insert": "Insertar",
  "tab.layout": "Diseño",
  "tab.references": "Referencias",
  "tab.review": "Revisar",
  "action.new": "Nuevo",
  "action.open": "Abrir…",
  "action.find": "Buscar",
  "action.xml": "XML",
  "action.download": "Descargar .docx",
  "action.undo": "Deshacer",
  "action.redo": "Rehacer",
  "find.find": "Buscar",
  "find.replaceWith": "Reemplazar con",
  "find.replaceAll": "Reemplazar todo",
  "find.close": "Cerrar búsqueda",
  "group.font": "Fuente",
  "group.text": "Texto",
  "group.paragraph": "Párrafo",
  "group.styles": "Estilos",
  "group.insert": "Insertar",
  "group.headings": "Títulos",
  "group.pageSetup": "Config. de página",
  "group.sections": "Secciones",
  "group.references": "Referencias",
  "group.tracking": "Control de cambios",
  "ins.paragraph": "Párrafo",
  "ins.table": "Tabla",
  "ins.pageBreak": "Salto de página",
  "ins.link": "Enlace",
  "ins.picture": "Imagen",
  "layout.sectionBreak": "Salto de sección",
  "layout.pageNumbers": "Números de página",
  "ref.toc": "Tabla de contenido",
  "ref.footnote": "Nota al pie",
  "ref.endnote": "Nota al final",
  "ref.bookmark": "Marcador",
  "review.newComment": "Nuevo comentario",
  "review.acceptAll": "Aceptar todo",
  "review.rejectAll": "Rechazar todo",
  "style.normal": "Normal",
  "style.heading1": "Título 1",
  "style.heading2": "Título 2",
  "style.heading3": "Título 3",
  "status.ready": "Listo.",
  "status.editing": "Editando…",
  "status.words": "palabras",
  "status.chars": "caracteres",
  "xml.title": "XML sin procesar — DrawingML / OMML / VML",
  "xml.documentSource": "Documento — formas / fórmulas / VML en línea",
  language: "Idioma",
};

const fr: Dict = {
  "tab.home": "Accueil",
  "tab.insert": "Insertion",
  "tab.layout": "Disposition",
  "tab.references": "Références",
  "tab.review": "Révision",
  "action.new": "Nouveau",
  "action.open": "Ouvrir…",
  "action.find": "Rechercher",
  "action.xml": "XML",
  "action.download": "Télécharger .docx",
  "action.undo": "Annuler",
  "action.redo": "Rétablir",
  "find.find": "Rechercher",
  "find.replaceWith": "Remplacer par",
  "find.replaceAll": "Tout remplacer",
  "find.close": "Fermer la recherche",
  "group.font": "Police",
  "group.text": "Texte",
  "group.paragraph": "Paragraphe",
  "group.styles": "Styles",
  "group.insert": "Insertion",
  "group.headings": "Titres",
  "group.pageSetup": "Mise en page",
  "group.sections": "Sections",
  "group.references": "Références",
  "group.tracking": "Suivi",
  "ins.paragraph": "Paragraphe",
  "ins.table": "Tableau",
  "ins.pageBreak": "Saut de page",
  "ins.link": "Lien",
  "ins.picture": "Image",
  "layout.sectionBreak": "Saut de section",
  "layout.pageNumbers": "Numéros de page",
  "ref.toc": "Table des matières",
  "ref.footnote": "Note de bas de page",
  "ref.endnote": "Note de fin",
  "ref.bookmark": "Signet",
  "review.newComment": "Nouveau commentaire",
  "review.acceptAll": "Tout accepter",
  "review.rejectAll": "Tout refuser",
  "style.normal": "Normal",
  "style.heading1": "Titre 1",
  "style.heading2": "Titre 2",
  "style.heading3": "Titre 3",
  "status.ready": "Prêt.",
  "status.editing": "Édition…",
  "status.words": "mots",
  "status.chars": "caractères",
  "xml.title": "XML brut — DrawingML / OMML / VML",
  "xml.documentSource": "Document — formes / maths / VML en ligne",
  language: "Langue",
};

const de: Dict = {
  "tab.home": "Start",
  "tab.insert": "Einfügen",
  "tab.layout": "Layout",
  "tab.references": "Verweise",
  "tab.review": "Überprüfen",
  "action.new": "Neu",
  "action.open": "Öffnen…",
  "action.find": "Suchen",
  "action.xml": "XML",
  "action.download": ".docx herunterladen",
  "action.undo": "Rückgängig",
  "action.redo": "Wiederholen",
  "find.find": "Suchen",
  "find.replaceWith": "Ersetzen durch",
  "find.replaceAll": "Alle ersetzen",
  "find.close": "Suche schließen",
  "group.font": "Schriftart",
  "group.text": "Text",
  "group.paragraph": "Absatz",
  "group.styles": "Formatvorlagen",
  "group.insert": "Einfügen",
  "group.headings": "Überschriften",
  "group.pageSetup": "Seite einrichten",
  "group.sections": "Abschnitte",
  "group.references": "Verweise",
  "group.tracking": "Nachverfolgung",
  "ins.paragraph": "Absatz",
  "ins.table": "Tabelle",
  "ins.pageBreak": "Seitenumbruch",
  "ins.link": "Link",
  "ins.picture": "Bild",
  "layout.sectionBreak": "Abschnittsumbruch",
  "layout.pageNumbers": "Seitenzahlen",
  "ref.toc": "Inhaltsverzeichnis",
  "ref.footnote": "Fußnote",
  "ref.endnote": "Endnote",
  "ref.bookmark": "Textmarke",
  "review.newComment": "Neuer Kommentar",
  "review.acceptAll": "Alle annehmen",
  "review.rejectAll": "Alle ablehnen",
  "style.normal": "Standard",
  "style.heading1": "Überschrift 1",
  "style.heading2": "Überschrift 2",
  "style.heading3": "Überschrift 3",
  "status.ready": "Bereit.",
  "status.editing": "Bearbeiten…",
  "status.words": "Wörter",
  "status.chars": "Zeichen",
  "xml.title": "Roh-XML — DrawingML / OMML / VML",
  "xml.documentSource": "Dokument — Inline-Formen / Formeln / VML",
  language: "Sprache",
};

const zh: Dict = {
  "tab.home": "开始",
  "tab.insert": "插入",
  "tab.layout": "布局",
  "tab.references": "引用",
  "tab.review": "审阅",
  "action.new": "新建",
  "action.open": "打开…",
  "action.find": "查找",
  "action.xml": "XML",
  "action.download": "下载 .docx",
  "action.undo": "撤销",
  "action.redo": "重做",
  "find.find": "查找",
  "find.replaceWith": "替换为",
  "find.replaceAll": "全部替换",
  "find.close": "关闭查找",
  "group.font": "字体",
  "group.text": "文本",
  "group.paragraph": "段落",
  "group.styles": "样式",
  "group.insert": "插入",
  "group.headings": "标题",
  "group.pageSetup": "页面设置",
  "group.sections": "分节",
  "group.references": "引用",
  "group.tracking": "修订",
  "ins.paragraph": "段落",
  "ins.table": "表格",
  "ins.pageBreak": "分页符",
  "ins.link": "链接",
  "ins.picture": "图片",
  "layout.sectionBreak": "分节符",
  "layout.pageNumbers": "页码",
  "ref.toc": "目录",
  "ref.footnote": "脚注",
  "ref.endnote": "尾注",
  "ref.bookmark": "书签",
  "review.newComment": "新建批注",
  "review.acceptAll": "全部接受",
  "review.rejectAll": "全部拒绝",
  "style.normal": "正文",
  "style.heading1": "标题 1",
  "style.heading2": "标题 2",
  "style.heading3": "标题 3",
  "status.ready": "就绪。",
  "status.editing": "编辑中…",
  "status.words": "词",
  "status.chars": "字符",
  "xml.title": "原始 XML — DrawingML / OMML / VML",
  "xml.documentSource": "文档 — 内联图形 / 公式 / VML",
  language: "语言",
};

const DICTS: Record<LocaleId, Dict> = { en, ja, es, fr, de, zh };

let current = $state<LocaleId>("en");

/** The active locale (reactive). */
export function locale(): LocaleId {
  return current;
}

export function setLocale(id: LocaleId): void {
  current = id;
}

/** Translate a message key in the active locale, falling back to English. */
export function t(key: MessageKey): string {
  return DICTS[current][key] ?? en[key] ?? key;
}
