#!/usr/bin/env bash
# fetch-specs.sh
#
# OOXML 関連の公開仕様を docs/specs/ に取得します。
#
# 取得対象:
#   - ECMA-376 5th edition Part 1-4 (PDF) ... ECMA International の自由配布版
#   - ECMA-376 5th edition Schema (XSD)  ... 同上 (Part 1 の正規スキーマ zip)
#   - Microsoft MS-OE376 (HTML)          ... ECMA-376 の Microsoft 実装メモ
#   - Microsoft Open Specifications PPTX / DOCX / XLSX (HTML, 任意)
#
# 取得した PDF/HTML/XSD はリポジトリにコミットしません (.gitignore で除外)。
# 必要なら手元の作業ディレクトリで参照してください。
#
# 使い方:
#   ./scripts/fetch-specs.sh           # すべて取得
#   ./scripts/fetch-specs.sh ecma      # ECMA-376 のみ
#   ./scripts/fetch-specs.sh xsd       # XSD のみ
#   ./scripts/fetch-specs.sh ms        # Microsoft Open Specs のみ

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DST="$ROOT/docs/specs"
mkdir -p "$DST/ecma-376" "$DST/xsd" "$DST/ms-openspecs"

target="${1:-all}"

fetch() {
  local url="$1"
  local out="$2"
  if [ -f "$out" ]; then
    echo "skip (exists): $out"
    return 0
  fi
  echo "fetching: $url"
  curl -L --fail --retry 3 --retry-delay 2 -o "$out.tmp" "$url"
  mv "$out.tmp" "$out"
}

fetch_ecma() {
  # ECMA-376 5th edition (2016). PDF 直リンクは ECMA のサイトから取得可能。
  # 公式トップ: https://ecma-international.org/publications-and-standards/standards/ecma-376/
  fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip" \
        "$DST/ecma-376/ECMA-376-1_5th_edition_december_2016.zip"
  fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip" \
        "$DST/ecma-376/ECMA-376-2_5th_edition_december_2021.zip"
  fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-3_4th_edition_december_2016.pdf" \
        "$DST/ecma-376/ECMA-376-3_4th_edition_december_2016.pdf"
  fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-4_4th_edition_december_2016.pdf" \
        "$DST/ecma-376/ECMA-376-4_4th_edition_december_2016.pdf"
  echo "ECMA-376 fetched. Unzip Part1/Part2 to access PDFs and XSDs."
}

fetch_xsd() {
  # Part 1 zip 内に XSD と RELAX NG が含まれる。Part 2 zip には OPC の XSD が含まれる。
  if [ ! -f "$DST/ecma-376/ECMA-376-1_5th_edition_december_2016.zip" ]; then
    fetch_ecma
  fi
  echo "Extracting XSDs from ECMA-376 Part1 zip ..."
  ( cd "$DST/xsd" && unzip -oq "$DST/ecma-376/ECMA-376-1_5th_edition_december_2016.zip" \
      "*.xsd" "*.rng" "*.rnc" || true )
  echo "Extracting XSDs from ECMA-376 Part2 zip ..."
  ( cd "$DST/xsd" && unzip -oq "$DST/ecma-376/ECMA-376-2_5th_edition_december_2021.zip" \
      "*.xsd" "*.rng" "*.rnc" || true )
}

fetch_ms() {
  # Microsoft Open Specifications は HTML のため、参照 URL のみ記載。
  cat > "$DST/ms-openspecs/INDEX.md" <<'EOF'
# Microsoft Open Specifications (OOXML 拡張)

ECMA-376 で定義されない、または Microsoft 製品が実際に書き出す拡張要素は
以下の Open Specifications を参照する。HTML 公開のためダウンロード対象外。

word-kit のスコープは docx のため、主に MS-OE376 と MS-DOCX を参照する。

- MS-OE376: Office Implementation Information for ECMA-376 (全体補足)
  https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/
- MS-DOCX: Word Extensions to the Office Open XML File Format
  https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/
- MS-ODRAWXML: Office Drawing Extensions
  https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/
- MS-OFFCRYPTO: Office Document Cryptography Structure (現状サポート外)
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/

スコープ外だが参考になり得るもの (現時点では参照不要):
- MS-PPTX: PowerPoint Extensions
- MS-XLSX: Excel Extensions

参考になる横断ドキュメント:
- Office Open XML overview (officeopenxml.com)
  http://officeopenxml.com/
EOF
  echo "Wrote $DST/ms-openspecs/INDEX.md"
}

case "$target" in
  all) fetch_ecma; fetch_xsd; fetch_ms ;;
  ecma) fetch_ecma ;;
  xsd) fetch_xsd ;;
  ms) fetch_ms ;;
  *) echo "unknown target: $target"; exit 1 ;;
esac

echo "Done. See $DST"
