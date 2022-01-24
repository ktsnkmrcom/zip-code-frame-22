/*
Illustrator Script  
郵便番号枠を作成するスクリプト  
zip-code-frame.jsx(1.0.1) / zip-code-frame-stroke.jsx(1.0.1)  
postcard-zip-code-frame.jsx(1.0.1) / postcard-zip-code-frame-stroke.jsx(1.0.1)  
郵便料金別納と後納の表示を作成するスクリプト  
postage-payment-method.jsx(1.0.0)  
リリース：2021.06.27  
アップデート：2021.07.14 postage-payment-method.jsxを同梱

Copyright (C) 2021 Katsushi Nakamura

This program is free software; you can redistribute it and/or modifyit under the terms of the GNU General Public License as published bythe Free Software Foundation; either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty ofMERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public Licensealong with this program. If not, see <https://www.gnu.org/licenses/>.
*/

(function () {
  //即時関数でスコープ
  // アートボード座標を基準にするように設定 cs5以降
  var ver = app.version.split(".");
  if (parseInt(ver[0]) >= 15) {
    app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;
  }

  // CMYKカラーを設定し、CMYKカラーオブジェクトを返す
  function setCMYKColor(c, m, y, k) {
    var CMYK = new CMYKColor();
    CMYK.cyan = c;
    CMYK.magenta = m;
    CMYK.yellow = y;
    CMYK.black = k;
    return CMYK;
  }

  // 塗りのカラーを「カラーなし」に設定
  var noColor = new NoColor();

  // 塗りの色
  var color = new CMYKColor();
  color.cyan = 0;
  color.magenta = 0;
  color.yellow = 0;
  color.black = 0;

  // 線の色
  var borderColor = new CMYKColor();
  borderColor.cyan = 0;
  borderColor.magenta = 0;
  borderColor.yellow = 0;
  borderColor.black = 100;

  // ドキュメントの有無

  if (documents.length >= 0) {
    // 官製はがきサイズ　ミリメートルをポイントに換算
    var mW100 = new UnitValue(100, "mm");
    var s100 = mW100.as("pt");
    var mH148 = new UnitValue(148, "mm");
    var l148 = mH148.as("pt");

    // 新規にファイル作成
    var dp = DocumentPreset;
    dp.colorMode = DocumentColorSpace.CMYK;
    dp.rasterResolution = DocumentRasterResolution.HighResolution;
    dp.units = RulerUnits.Millimeters;
    dp.width = s100;
    dp.height = l148;
    documents.addDocument("Standerd100x148", dp);

    // 丸　後納

    // 料金後納レイヤー作成
    var layObj = activeDocument.layers.add();
    layObj.name = "料金後納　丸";

    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    // 丸枠サイズ　ミリメートルをポイントに換算
    var mW25 = new UnitValue(25, "mm");
    var w25 = mW25.as("pt");
    var x = 20;
    var y = -20;

    // 楕円形を描く
    var ellipseObj = groupObj.pathItems.ellipse(y, x, w25, w25, false, true);
    ellipseObj.fillColor = color; // 塗り
    ellipseObj.stroked = true; // 線を表示する
    ellipseObj.strokeWidth = 0.5; // pointで線の幅を指定する
    ellipseObj.strokeColor = borderColor; // 線

    // 線を描画
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [x + 2, y - w25 / 3],
      [x - 2 + w25, y - w25 / 3],
    ]);
    lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
    lineObj.strokeWidth = 0.5; // pointで線の幅を指定する
    lineObj.strokeColor = borderColor; // 線

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w25 / 2 + x; // X座標
    txtObj.top = y - 8; // Y座標
    txtObj.paragraphs.add("差出事業所名"); // 段落を作成
    txtObj.paragraphs[0].size = 7; // 段落の文字サイズ設定
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w25 / 2 + x; // X座標
    txtObj.top = -(w25 / 2 + 15); // Y座標
    txtObj.paragraphs.add("料金後納"); // 段落を作成
    txtObj.paragraphs.add("郵便");
    txtObj.paragraphs[0].size = 10; // 文字サイズ
    txtObj.paragraphs[1].size = 14;
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央
    txtObj.paragraphs[1].justification = Justification.CENTER;
    txtObj.paragraphs[0].autoLeading = false; // 行間自動を解除
    txtObj.paragraphs[0].leading = 12; // 行間を指定

    // 角　後納

    var layObj = activeDocument.layers.add();
    layObj.name = "料金後納　角";
    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    var mW20 = new UnitValue(20, "mm");
    var w20 = mW20.as("pt");
    var mH20 = new UnitValue(20, "mm");
    var h20 = mH20.as("pt");

    var lineX = 20;
    var lineH = h20;
    var lineY = -20;
    var lineW = w20;

    // 矩形を描く
    var rectObj = groupObj.pathItems.rectangle(lineY, lineX, lineW, lineH);
    rectObj.fillColor = color; //塗り
    rectObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
    rectObj.strokeWidth = 0.5; // pointで線幅を指定する
    rectObj.strokeColor = borderColor; // 線

    // 線を描画
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [lineX, lineY - h20 / 3],
      [lineX + w20, lineY - h20 / 3],
    ]);
    lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
    lineObj.strokeWidth = 0.5; // pointで線幅を指定する
    lineObj.strokeColor = borderColor; // 線

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w20 / 2 + lineX; // X座標
    txtObj.top = -22; // Y座標
    txtObj.paragraphs.add("差出事業所名"); // 段落を作成
    txtObj.paragraphs[0].size = 7; // 文字サイズ
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w20 / 2 + lineX; // X座標
    txtObj.top = -44; // Y座標
    txtObj.paragraphs.add("料金後納"); // 段落を作成
    txtObj.paragraphs.add("郵便");
    txtObj.paragraphs[0].size = 9; // 文字サイズ
    txtObj.paragraphs[1].size = 11;
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央
    txtObj.paragraphs[1].justification = Justification.CENTER;
    txtObj.paragraphs[0].autoLeading = false; // 行間自動を解除
    txtObj.paragraphs[0].leading = 12; // 行間を指定

    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金別納レイヤー作成
    // 料金後納レイヤー作成
    var layObj = activeDocument.layers.add();
    layObj.name = "料金別納　丸";

    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    // 楕円形を描く
    var ellipseObj = groupObj.pathItems.ellipse(y, x, w25, w25, false, true);
    ellipseObj.fillColor = color; // 塗り
    ellipseObj.stroked = true; // 線を表示する
    ellipseObj.strokeWidth = 0.5; // pointで線幅を指定する
    ellipseObj.strokeColor = borderColor; // 線

    // 線を描画
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [x + 2, y - w25 / 3],
      [x - 2 + w25, y - w25 / 3],
    ]);
    lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
    lineObj.strokeWidth = 0.5; // pointで線幅を指定する
    lineObj.strokeColor = borderColor; // 線

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w25 / 2 + x; // X座標
    txtObj.top = y - 8; // Y座標
    txtObj.paragraphs.add("差出事業所名"); // 段落を作成
    txtObj.paragraphs[0].size = 7; // 文字サイズ
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w25 / 2 + 20; // X座標を10ptの位置にする
    txtObj.top = -(w25 / 2 + 15); // Y座標を上から50ptの位置にする
    txtObj.paragraphs.add("料金別納"); // 段落を作成
    txtObj.paragraphs.add("郵便");
    txtObj.paragraphs[0].size = 10; // 文字サイズ
    txtObj.paragraphs[1].size = 14;
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央
    txtObj.paragraphs[1].justification = Justification.CENTER;
    txtObj.paragraphs[0].autoLeading = false; // 行間自動を解除
    txtObj.paragraphs[0].leading = 12; //行間を指定

    // 角　別納

    var layObj = activeDocument.layers.add();
    layObj.name = "料金別納　角";
    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    var mW20 = new UnitValue(20, "mm");
    var w20 = mW20.as("pt");
    var mH20 = new UnitValue(20, "mm");
    var h20 = mH20.as("pt");

    var lineX = 20;
    var lineH = h20;
    var lineY = -20;
    var lineW = w20;

    // 矩形を描く
    var rectObj = groupObj.pathItems.rectangle(lineY, lineX, lineW, lineH);
    rectObj.fillColor = color; // 塗り
    rectObj.stroked = true; // 線を表示する
    rectObj.strokeWidth = 0.5; // 線の幅を指定する
    rectObj.strokeColor = borderColor; // 線

    // 線を描画
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [lineX, lineY - h20 / 3],
      [lineX + w20, lineY - h20 / 3],
    ]);
    lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
    lineObj.strokeWidth = 0.5; // pointで線幅を指定する
    lineObj.strokeColor = borderColor; // 線

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w20 / 2 + lineX; // X座標
    txtObj.top = y - 2; // Y座標
    txtObj.paragraphs.add("差出事業所名"); // 段落を作成
    txtObj.paragraphs[0].size = 7; // 文字サイズ
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央

    // テキストフレームを作成し段落に文字挿入
    var txtObj = groupObj.textFrames.add();
    txtObj.left = w20 / 2 + lineX; // X座標
    txtObj.top = y - 24; // Y座標
    txtObj.paragraphs.add("料金別納"); // 段落を作成
    txtObj.paragraphs.add("郵便");
    txtObj.paragraphs[0].size = 9; // 文字サイズ
    txtObj.paragraphs[1].size = 11;
    txtObj.paragraphs[0].justification = Justification.CENTER; // 行揃え中央
    txtObj.paragraphs[1].justification = Justification.CENTER;
    txtObj.paragraphs[0].autoLeading = false; // 行間自動を解除
    txtObj.paragraphs[0].leading = 12; // 行間を指定
  } else {
  }
})();
