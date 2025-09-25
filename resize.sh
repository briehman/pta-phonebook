#!/usr/bin/env bash

mkdir -p covers/resized

# Turn the covers to PDFs
for img in \
  covers/modified/front-cover.jpg \
  covers/modified/front-inside-cover.jpg \
  covers/modified/back-inside-cover.jpg \
  covers/modified/back-cover.jpg \
  ; do
  echo "Resizing $img..."
  b="$(basename "$img" .jpg)"
  convert "$img" -resize 2480x3508 -quality 85 covers/"$b".jpg
done
