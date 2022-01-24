# 郵便番号枠を作成するスクリプトを作成しました

adobe Illustrator で動作する郵便番号枠を作成するスクリプトです。  
郵便料金別納と後納の表示を作成するスクリプトを同梱しました。

動作確認環境  
Windows：2021  
Mac：CS6 CC2019 2021  
古いイラストレーターでは座標系が異なるため描画位置がずれます。

## スクリプトファイルの説明です

- 郵便番号枠を最前面の新規レイヤーに作成します。その際、ドキュメントサイズの矩形（塗り、線ともに色なし）も作成します。

- 最前面に空のレイヤー（非表示ではない、ロックされていない）があれば、そこに作成されますので、**先に開いているドキュメントには十分にご注意ください。**

- アートボードが複数ある場合には、アクティブなアートボードに作成されます。

- ドキュメントを開いていない場合は、定型の最大サイズ（長３）の新規ドキュメントに郵便番号枠を作成します。postcard-の場合は官製はがきサイズの新規ドキュメントに郵便番号枠を作成します。

- 最前面のレイヤーに、オブジェクト（文字や図形等）が存在する場合には作成しません。また、非表示のレイヤー、ロックされたレイヤーにも作成しません。

- 枠は朱色の塗りの矩形の上に紙色（白）の矩形を重ねて表現しています。

- よこ長のサイズの郵便番号枠は、たて長の場合を回転させた状態と同じです。

- 定型または定形外のサイズにあった郵便番号枠を作成します。

- 枠はグループにまとまっています。

- -stroke の付く jsx ファイルは、枠を線で作成しています。ハイフンも線です。

- postcard-の付く jsx ファイルは、はがきの規格サイズのみに対応しており、POSTCARD の文言を所定の位置付近に作成します。

### 定型・定形外サイズに対応します

- zip-code-frame.jsx
- zip-code-frame-stroke.jsx

定型：長３　長４　洋１　洋２　大判ハガキなど  
定形外：角２　角３　角６　角０など

### はがきサイズにのみ対応します

- postcard-zip-code-frame.jsx
- postcard-zip-code-frame-stroke.jsx

## スクリプトの使い方

**Code（緑色）から ZIP ファイルをダウンロードできます。**

圧縮ファイルを展開後、jsx ファイルを、Illustrator の、**ファイルメニュー > スクリプト > その他のスクリプト**より作動させてください。  
通常、Ctrl + F12、command + F12 にショートカットキーが割り当てられています。

ドキュメントが何も開いていない場合は、長３サイズの新規ドキュメントに郵便番号枠を作成します。postcard-の方では、官製はがきサイズで郵便番号枠を作成します。

新規ドキュメントを任意のサイズ（アートボードサイズ）で作成された場合は、そのサイズに合わせて郵便番号枠を作成します。

### 枠の線の太さを変更したい

枠の線の太さを変更される場合は、**-stroke でない**ほうが適しています。-stroke のほうが変更は簡単ですが、枠の内側のサイズが変わってしまうため、規格から外れてしまいます。

### 大きめのアートボードで作業したい

制作物より大きいアートボードサイズで作業される場合は、別ドキュメントに作成しペーストしていただくか、作成後にアートボードサイズを変更して対応してください。

## ご注意

郵便物の種類には、長辺、短辺の長さだけでなく、重量や厚さも含まれていますので、実際の郵便料金等は各自お調べください。

## 郵便料金別納と後納の表示を作成するスクリプトについて

### 郵便料金別納と後納の表示を作成

- postage-payment-method.jsx

郵便料金別納と後納の表示を作成する、postage-payment-method.jsx は、常に新規ドキュメントを作成します。

表示の詳細については、適宜、アレンジする等してご使用ください。

## ライセンス

Illustrator Script  
郵便番号枠を作成するスクリプト  
zip-code-frame.jsx / zip-code-frame-stroke.jsx  
postcard-zip-code-frame.jsx / postcard-zip-code-frame-stroke.jsx  
郵便料金別納と後納の表示を作成するスクリプト  
postage-payment-method.jsx  
リリース：2021.06.27  
アップデート：2021.07.14 postage-payment-method.jsx を同梱  
再リリース：2022.01.25

Copyright (C) 2021 Katsushi Nakamura

This program is free software; you can redistribute it and/or modifyit under the terms of the GNU General Public License as published bythe Free Software Foundation; either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty ofMERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public Licensealong with this program. If not, see <https://www.gnu.org/licenses/>.
