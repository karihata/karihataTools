# karihataTools

## 概要

- 画像ファイルの検査等の自作マクロをエクセルアドインとしてまとめている

- ImageMagickと連携させたり区切り文字付きで文字結合する関数やシートの一括作成をする

- VBAの勉強用として作成しているため、UIの作り込みは疎らになっている

- Office2016以降で使える「TEXTJOIN」に似たセル同士の文字結合時に区切り文字を追加できる「StrCat」関数が使えます

## 導入
1. Install.vbsを実行
2. KarihataTools.xlamのプロパティを開いて「ブロック解除」をする　※プロパティにブロック解除が無ければ不要


## 撤去
1. Excelを起動
2. [ファイル]-[オプション]-[アドイン]-[設定...]を開く
3. 「Karihatatools」のチェックを外す
4. 「Karihatatools」を選択して[参照(B)...]を開く
5. 「Karihatatools.xlam」ファイルを削除

## 参考サイト
- ImageMagickをVBAで使う
https://sites.google.com/site/torimemoforwork/vba/dll/imagemagick

- indentify_format一覧
https://imagemagick.org/script/escape.php
