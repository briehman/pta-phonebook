#!/usr/bin/env bash

cd "$(dirname "$0")"

data_pdf="${1?Must specify data PDF as first argument}"
output_pdf="${2?Must specify output PDF as second argument}"

rm -rf tmp
mkdir -p tmp

if [[ ! -e "$data_pdf" ]]; then
  echo "ERROR: $data_pdf is missing" >&2
  exit 1
fi

# Turn the covers to PDFs
for img in \
  covers/front-cover.jpg \
  covers/front-inside-cover.jpg \
  covers/back-inside-cover.jpg \
  covers/back-cover.jpg \
  ; do
  echo "Converting $img to PDF..."
  b="$(basename "$img" .jpg)"
  img2pdf --pagesize Letter "$img" -o tmp/"$b".pdf;
done

# Write the PDF
set -x
pdfunite \
  tmp/front-cover.pdf \
  tmp/front-inside-cover.pdf \
  "$data_pdf" \
  tmp/back-inside-cover.pdf \
  tmp/back-cover.pdf \
  "$output_pdf"
