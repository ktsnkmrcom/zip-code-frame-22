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
  // アートボード座標を基準にするように設定 cs5以降
  var ver = app.version.split(".");
  if (parseInt(ver[0]) >= 15) {
    app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;
  }

  // 塗りのカラーを「なし」に指定する設定用
  var noColor = new NoColor();

  // 枠の内側の色　塗りの色
  var color = new CMYKColor();
  color.cyan = 0;
  color.magenta = 0;
  color.yellow = 0;
  color.black = 0;

  // 枠の色　塗りの色　最下面の矩形
  var borderColor = new CMYKColor();
  borderColor.cyan = 0;
  borderColor.magenta = 90;
  borderColor.yellow = 100;
  borderColor.black = 0;

  // ドキュメントの有無　ファイルを開いているか
  var n = documents.length;

  if (n === 0) {
    // ファイル無し
    // 定型最大長３サイズ
    // 定型最大の長３サイズの新規ファイル　ミリメートルをポイントに換算
    var mW120 = new UnitValue(120, "mm");
    var w120 = mW120.as("pt");
    var mH235 = new UnitValue(235, "mm");
    var h235 = mH235.as("pt");

    var dp = DocumentPreset;
    dp.colorMode = DocumentColorSpace.CMYK;
    dp.rasterResolution = DocumentRasterResolution.HighResolution;
    dp.units = RulerUnits.Millimeters;
    dp.width = w120;
    dp.height = h235;

    documents.addDocument("Standerd120x235", dp);

    // 郵便枠の内側サイズ
    var mW5_7 = new UnitValue(5.7, "mm");
    var w5_7 = mW5_7.as("pt");
    var mH8_0 = new UnitValue(8.0, "mm");
    var h8_0 = mH8_0.as("pt");

    // 郵便枠前３の外側サイズ
    var mW6_7 = new UnitValue(6.7, "mm");
    var w6_7 = mW6_7.as("pt");
    var mH9_0 = new UnitValue(9.0, "mm");
    var h9_0 = mH9_0.as("pt");

    // 郵便枠後4の外側サイズ
    var mW6_3 = new UnitValue(6.3, "mm");
    var w6_3 = mW6_3.as("pt");
    var mH8_6 = new UnitValue(8.6, "mm");
    var h8_6 = mH8_6.as("pt");

    // 枠の太さ
    var mB0_5 = new UnitValue(0.5, "mm");
    var b0_5 = mB0_5.as("pt");
    var mB0_3 = new UnitValue(0.3, "mm");
    var b0_3 = mB0_3.as("pt");

    // 右側の指定マージン
    var mRightMargin8_0 = new UnitValue(8.0, "mm");
    var rightMargin = mRightMargin8_0.as("pt");

    // 上側の指定マージン
    var mTopMargin12_0 = new UnitValue(-12.0, "mm");
    var topMargin = mTopMargin12_0.as("pt");

    // 郵便枠の総幅
    var mPostFrameSize47_7 = new UnitValue(47.7, "mm");
    var postFrameSize = mPostFrameSize47_7.as("pt");

    // 新規ドキュメントサイズ時（120x235）の内側枠の開始の起点
    var newStartingPoint = w120 - postFrameSize - rightMargin;

    // 枠の指定配置位置までの間隔
    var mD7_0 = new UnitValue(7.0, "mm");
    var d7_0 = mD7_0.as("pt");

    var mD14_0 = new UnitValue(14.0, "mm");
    var d14_0 = mD14_0.as("pt");

    var mD21_6 = new UnitValue(21.6, "mm");
    var d21_6 = mD21_6.as("pt");

    var mD28_4 = new UnitValue(28.4, "mm");
    var d28_4 = mD28_4.as("pt");

    var mD35_2 = new UnitValue(35.2, "mm");
    var d35_2 = mD35_2.as("pt");

    var mD42_0 = new UnitValue(42.0, "mm");
    var d42_0 = mD42_0.as("pt");

    // ハイフンの矩形
    var lineX = newStartingPoint + d14_0 + w5_7;

    var mLineH = new UnitValue(0.5, "mm");
    var lineH = mLineH.as("pt");

    var lineY = topMargin - (h8_0 / 2 - lineH / 2);

    var lineW = d21_6 - d14_0 - w5_7;

    // トリム用の矩形を描く
    var rectObj = activeDocument.pathItems.rectangle(
      h235 - h235,
      0,
      w120,
      h235
    );
    rectObj.fillColor = noColor;
    rectObj.stroked = true; // 線を表示する
    rectObj.strokeWidth = 0; // 線の幅を指定する
    rectObj.strokeColor = noColor;

    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    // ハイフンを描く
    var rectObj = groupObj.pathItems.rectangle(lineY, lineX, lineW, lineH);
    rectObj.fillColor = borderColor;
    rectObj.stroked = true; // 線を表示する
    rectObj.strokeWidth = 0; // 線の幅を指定する
    rectObj.strokeColor = noColor;

    // 四角形を描く
    var array = [
      [topMargin + b0_5, newStartingPoint - b0_5, w6_7, h9_0, 1],
      [topMargin + b0_5, newStartingPoint + d7_0 - b0_5, w6_7, h9_0, 1],
      [topMargin + b0_5, newStartingPoint + d14_0 - b0_5, w6_7, h9_0, 1],
      [topMargin + b0_3, newStartingPoint + d21_6 - b0_3, w6_3, h8_6, 2],
      [topMargin + b0_3, newStartingPoint + d28_4 - b0_3, w6_3, h8_6, 2],
      [topMargin + b0_3, newStartingPoint + d35_2 - b0_3, w6_3, h8_6, 2],
      [topMargin + b0_3, newStartingPoint + d42_0 - b0_3, w6_3, h8_6, 2],
      [topMargin, newStartingPoint, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d7_0, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d14_0, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d21_6, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d28_4, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d35_2, w5_7, h8_0, 3],
      [topMargin, newStartingPoint + d42_0, w5_7, h8_0, 3],
    ];

    for (var i = 0; i < array.length; i++) {
      for (var j = 0; j < array[i].length; j++) {
        y = array[i][0];
        x = array[i][1];
        w = array[i][2];
        h = array[i][3];
        s = array[i][4];
      }

      if (s === 1) {
        var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
        rectObj.fillColor = borderColor;
        rectObj.stroked = true; // 線を表示する
        rectObj.strokeWidth = 0; // 線の幅を指定する
        rectObj.strokeColor = noColor; //nocolor
      } else if (s === 2) {
        var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
        rectObj.fillColor = borderColor;
        rectObj.stroked = true; // 線を表示する
        rectObj.strokeWidth = 0; // 線の幅を指定する
        rectObj.strokeColor = noColor; //nocolor
      } else {
        var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
        rectObj.fillColor = color;
        rectObj.stroked = true; // 線を表示する
        rectObj.strokeWidth = 0; // 線の幅を指定する
        rectObj.strokeColor = noColor; //nocolor
      }
    }
  } else {
    // ファイル有り
    // 最上面レイヤーのアイテムチェック
    var page = activeDocument.layers[0].pageItems.length;
    if (page === 0) {
      // ドキュメントサイズを取得
      var docWidth = activeDocument.width;
      var docHeight = activeDocument.height;
      // トリム用描画に取得のままが必要

      // 判別用　小数点以下切り捨て
      var mfdocWidth = Math.floor(docWidth);
      var mfdocHeight = Math.floor(docHeight);

      // mmに変換
      var pdocWidth = new UnitValue(mfdocWidth, "pt");
      var mdocWidth = Math.floor(pdocWidth.as("mm"));
      var pdocHeight = new UnitValue(mfdocHeight, "pt");
      var mdocHeight = Math.floor(pdocHeight.as("mm"));

      // 常にmm変換すると1mm短くなるので1を足す
      var mdocWidth = mdocWidth + 1;
      var mdocHeight = mdocHeight + 1;

      // 長辺と短辺　たて長とよこ長判別用
      var longSide = Math.max(mdocWidth, mdocHeight);
      var shortSide = Math.min(mdocWidth, mdocHeight);

      // 定型サイズ最大　ミリメートルをポイントに換算　　定型最大の長３サイズ
      var mL120 = new UnitValue(120, "mm");
      // var l120 = mL120.as("pt");
      var mL235 = new UnitValue(235, "mm");
      // var l235 = mL235.as("pt");

      // 定型最小サイズ　ミリメートルをポイントに換算
      var mL90 = new UnitValue(90, "mm");
      // var l90 = mL90.as("pt");
      var mL140 = new UnitValue(140, "mm");
      // var l140 = mL140.as("pt");

      // 定型外最大サイズ　ミリメートルをポイントに換算
      var mL250 = new UnitValue(250, "mm");
      // var l250 = mL250.as("pt");
      var mL340 = new UnitValue(340, "mm");
      // var l340 = mL340.as("pt");

      // 定型サイズか判定
      if (
        shortSide <= mL120 &&
        shortSide >= mL90 &&
        longSide <= mL235 &&
        longSide >= mL140
      ) {
        // たて長か判定
        if (mdocWidth < mdocHeight) {
          // 定型サイズ内たて
          // 郵便枠の内側サイズ
          var mW5_7 = new UnitValue(5.7, "mm");
          var w5_7 = mW5_7.as("pt");
          var mH8_0 = new UnitValue(8.0, "mm");
          var h8_0 = mH8_0.as("pt");

          // 郵便枠前３の外側サイズ
          var mW6_7 = new UnitValue(6.7, "mm");
          var w6_7 = mW6_7.as("pt");
          var mH9_0 = new UnitValue(9.0, "mm");
          var h9_0 = mH9_0.as("pt");

          // 郵便枠後4の外側サイズ
          var mW6_3 = new UnitValue(6.3, "mm");
          var w6_3 = mW6_3.as("pt");
          var mH8_6 = new UnitValue(8.6, "mm");
          var h8_6 = mH8_6.as("pt");

          // 枠の太さ
          var mB0_5 = new UnitValue(0.5, "mm");
          var b0_5 = mB0_5.as("pt");
          var mB0_3 = new UnitValue(0.3, "mm");
          var b0_3 = mB0_3.as("pt");

          // 右側の指定マージン
          var mRightMargin8_0 = new UnitValue(8.0, "mm");
          var rightMargin = mRightMargin8_0.as("pt");

          // 上側の指定マージン
          var mTopMargin12_0 = new UnitValue(-12.0, "mm");
          var topMargin = mTopMargin12_0.as("pt");

          // 郵便枠の総幅
          var mPostFrameSize47_7 = new UnitValue(47.7, "mm");
          var postFrameSize = mPostFrameSize47_7.as("pt");

          // 定型サイズ　ミリメートルをポイントに換算
          var mW120 = new UnitValue(120, "mm");
          var w120 = mW120.as("pt");
          var mH235 = new UnitValue(235, "mm");
          var h235 = mH235.as("pt");

          // 任意ドキュメントサイズの内側枠の開始の起点
          var newStartingPoint = docWidth - postFrameSize - rightMargin;

          // 枠の指定配置位置までの間隔
          var mD7_0 = new UnitValue(7.0, "mm");
          var d7_0 = mD7_0.as("pt");

          var mD14_0 = new UnitValue(14.0, "mm");
          var d14_0 = mD14_0.as("pt");

          var mD21_6 = new UnitValue(21.6, "mm");
          var d21_6 = mD21_6.as("pt");

          var mD28_4 = new UnitValue(28.4, "mm");
          var d28_4 = mD28_4.as("pt");

          var mD35_2 = new UnitValue(35.2, "mm");
          var d35_2 = mD35_2.as("pt");

          var mD42_0 = new UnitValue(42.0, "mm");
          var d42_0 = mD42_0.as("pt");

          var lineX = newStartingPoint + d14_0 + w5_7;

          var mLineH = new UnitValue(0.5, "mm");
          var lineH = mLineH.as("pt");

          var lineY = topMargin - (h8_0 / 2 - lineH / 2);

          var lineW = d21_6 - d14_0 - w5_7;

          // 非表示レイヤーやロックレイヤーへの対応
          try {
            // トリム用の矩形を描く
            var rectObj = activeDocument.pathItems.rectangle(
              docHeight - docHeight,
              0,
              docWidth,
              docHeight
            );
            rectObj.fillColor = noColor;
            rectObj.stroked = true; // 線を表示する
            rectObj.strokeWidth = 0; // 線の幅を指定する
            rectObj.strokeColor = noColor;

            // グループオブジェクトを生成
            var groupObj = activeDocument.groupItems.add();

            // ハイフンを描く
            var rectObj = groupObj.pathItems.rectangle(
              lineY,
              lineX,
              lineW,
              lineH
            );
            rectObj.fillColor = borderColor;
            rectObj.stroked = true; // 線を表示する
            rectObj.strokeWidth = 0; // 線の幅を指定する
            rectObj.strokeColor = noColor;

            // 四角形を描く
            var array = [
              [topMargin + b0_5, newStartingPoint - b0_5, w6_7, h9_0, 1],
              [topMargin + b0_5, newStartingPoint + d7_0 - b0_5, w6_7, h9_0, 1],
              [
                topMargin + b0_5,
                newStartingPoint + d14_0 - b0_5,
                w6_7,
                h9_0,
                1,
              ],
              [
                topMargin + b0_3,
                newStartingPoint + d21_6 - b0_3,
                w6_3,
                h8_6,
                2,
              ],
              [
                topMargin + b0_3,
                newStartingPoint + d28_4 - b0_3,
                w6_3,
                h8_6,
                2,
              ],
              [
                topMargin + b0_3,
                newStartingPoint + d35_2 - b0_3,
                w6_3,
                h8_6,
                2,
              ],
              [
                topMargin + b0_3,
                newStartingPoint + d42_0 - b0_3,
                w6_3,
                h8_6,
                2,
              ],
              [topMargin, newStartingPoint, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d7_0, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d14_0, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d21_6, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d28_4, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d35_2, w5_7, h8_0, 3],
              [topMargin, newStartingPoint + d42_0, w5_7, h8_0, 3],
            ];

            for (var i = 0; i < array.length; i++) {
              for (var j = 0; j < array[i].length; j++) {
                y = array[i][0];
                x = array[i][1];
                w = array[i][2];
                h = array[i][3];
                s = array[i][4];
              }

              if (s === 1) {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = borderColor;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor;
              } else if (s === 2) {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = borderColor;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor;
              } else {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor;
              }
            }
            // alert("定型サイズ内たて");
          } catch (e) {
            alert(
              "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
            );
          }
        } else {
          // 定型サイズ内よこ
          // 郵便枠の内側サイズ
          var mH5_7 = new UnitValue(5.7, "mm");
          var h5_7 = mH5_7.as("pt");
          var mW8_0 = new UnitValue(8.0, "mm");
          var w8_0 = mW8_0.as("pt");

          // 郵便枠前３の外側サイズ
          var mH6_7 = new UnitValue(6.7, "mm");
          var h6_7 = mH6_7.as("pt");
          var mW9_0 = new UnitValue(9.0, "mm");
          var w9_0 = mW9_0.as("pt");

          // 郵便枠後4の外側サイズ
          var mH6_3 = new UnitValue(6.3, "mm");
          var h6_3 = mH6_3.as("pt");
          var mW8_6 = new UnitValue(8.6, "mm");
          var w8_6 = mW8_6.as("pt");

          // 枠の太さ
          var mB0_5 = new UnitValue(0.5, "mm");
          var b0_5 = mB0_5.as("pt");
          var mB0_3 = new UnitValue(0.3, "mm");
          var b0_3 = mB0_3.as("pt");

          // 下側の指定マージン
          var mBottomMargin = new UnitValue(8.0, "mm");
          var bottomMargin = mBottomMargin.as("pt");

          // 右側の指定マージン
          var mRightMargin = new UnitValue(12.0, "mm");
          var rightMargin = mRightMargin.as("pt");

          // 郵便枠の総幅
          var mPostFrameSize = new UnitValue(47.7, "mm");
          var postFrameSize = mPostFrameSize.as("pt");

          // ドキュメントサイズ時の内側枠の開始の起点
          var newStartingPoint =
            docHeight - docHeight - (docHeight - postFrameSize - bottomMargin);

          // 内側矩形の共通x座標
          var xPoint = docWidth - (w8_0 + rightMargin);

          // 枠の指定配置位置までの間隔
          var mD7_0 = new UnitValue(7.0, "mm");
          var d7_0 = mD7_0.as("pt");

          var mD14_0 = new UnitValue(14.0, "mm");
          var d14_0 = mD14_0.as("pt");

          var mD21_6 = new UnitValue(21.6, "mm");
          var d21_6 = mD21_6.as("pt");

          var mD28_4 = new UnitValue(28.4, "mm");
          var d28_4 = mD28_4.as("pt");

          var mD35_2 = new UnitValue(35.2, "mm");
          var d35_2 = mD35_2.as("pt");

          var mD42_0 = new UnitValue(42.0, "mm");
          var d42_0 = mD42_0.as("pt");

          //ハイフン用
          var mLineW = new UnitValue(0.5, "mm");
          var lineW = mLineW.as("pt");

          var lineX = docWidth - (w8_0 / 2 + lineW / 2) - rightMargin;

          var lineY = newStartingPoint - d14_0 - h5_7;

          var lineH = d21_6 - d14_0 - h5_7;

          // 非表示レイヤーやロックレイヤーへの対応
          try {
            // トリム用の矩形を描く
            var rectObj = activeDocument.pathItems.rectangle(
              docHeight - docHeight,
              0,
              docWidth,
              docHeight
            );
            rectObj.fillColor = noColor;
            rectObj.stroked = true; // 線を表示する
            rectObj.strokeWidth = 0; // 線の幅を指定する
            rectObj.strokeColor = noColor;

            // グループオブジェクトを生成
            var groupObj = activeDocument.groupItems.add();

            // ハイフンを描く
            var rectObj = groupObj.pathItems.rectangle(
              lineY,
              lineX,
              lineW,
              lineH
            );
            rectObj.fillColor = borderColor;
            rectObj.stroked = true; // 線を表示する
            rectObj.strokeWidth = 0; // 線の幅を指定する
            rectObj.strokeColor = noColor;

            // 四角形を描く
            var array = [
              [newStartingPoint + b0_5, xPoint - b0_5, w9_0, h6_7, 1],
              [newStartingPoint - d7_0 + b0_5, xPoint - b0_5, w9_0, h6_7, 1],
              [newStartingPoint - d14_0 + b0_5, xPoint - b0_5, w9_0, h6_7, 1],
              [newStartingPoint - d21_6 + b0_3, xPoint - b0_3, w8_6, h6_3, 2],
              [newStartingPoint - d28_4 + b0_3, xPoint - b0_3, w8_6, h6_3, 2],
              [newStartingPoint - d35_2 + b0_3, xPoint - b0_3, w8_6, h6_3, 2],
              [newStartingPoint - d42_0 + b0_3, xPoint - b0_3, w8_6, h6_3, 2],
              [newStartingPoint, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d7_0, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d14_0, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d21_6, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d28_4, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d35_2, xPoint, w8_0, h5_7, 3],
              [newStartingPoint - d42_0, xPoint, w8_0, h5_7, 3],
            ];

            for (var i = 0; i < array.length; i++) {
              for (var j = 0; j < array[i].length; j++) {
                y = array[i][0];
                x = array[i][1];
                w = array[i][2];
                h = array[i][3];
                s = array[i][4];
              }

              if (s === 1) {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = borderColor;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor; //nocolor
              } else if (s === 2) {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = borderColor;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor; //nocolor
              } else {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = 0; // 線の幅を指定する
                rectObj.strokeColor = noColor; //nocolor
              }
            }
            // alert("定型サイズ内よこ");
          } catch (e) {
            alert(
              "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
            );
          }
        }
      } else {
        //定型サイズではない

        // 定型外サイズか判定
        if (shortSide >= mL90 && longSide >= mL140) {
          // たて長か判定
          if (mdocWidth <= mdocHeight) {
            // 定型外サイズたて長規格内
            // 郵便枠の内側サイズ
            var mW10_0 = new UnitValue(10.0, "mm");
            var w10_0 = mW10_0.as("pt");
            var mH13_5 = new UnitValue(13.5, "mm");
            var h13_5 = mH13_5.as("pt");

            // 郵便枠前３の外側サイズ
            var mW11_4 = new UnitValue(11.4, "mm");
            var w11_4 = mW11_4.as("pt");
            var mH14_9 = new UnitValue(14.9, "mm");
            var h14_9 = mH14_9.as("pt");

            // 郵便枠後4の外側サイズ
            var mW10_8 = new UnitValue(10.8, "mm");
            var w10_8 = mW10_8.as("pt");
            var mH14_3 = new UnitValue(14.3, "mm");
            var h14_3 = mH14_3.as("pt");

            // 枠の太さ
            var mB0_7 = new UnitValue(0.7, "mm");
            var b0_7 = mB0_7.as("pt");
            var mB0_4 = new UnitValue(0.4, "mm");
            var b0_4 = mB0_4.as("pt");

            // 右側の指定マージン
            var mRightMargin18_0 = new UnitValue(18.0, "mm");
            var nsdRightMargin = mRightMargin18_0.as("pt");

            // 上側の指定マージン
            var mTopMargin18_0 = new UnitValue(-18.0, "mm");
            var nsdTopMargin = mTopMargin18_0.as("pt");

            // 郵便枠の総幅
            var mPostFrameSize83_0 = new UnitValue(83.0, "mm");
            var nsdPostFrameSize = mPostFrameSize83_0.as("pt");

            // 任意のドキュメントサイズ時の内側枠の開始の起点
            var startingPoint = docWidth - nsdPostFrameSize - nsdRightMargin;

            // 枠の指定配置位置までの間隔
            var mD12_0 = new UnitValue(12.0, "mm");
            var d12_0 = mD12_0.as("pt");

            var mD24_0 = new UnitValue(24.0, "mm");
            var d24_0 = mD24_0.as("pt");

            var mD37_0 = new UnitValue(37.0, "mm");
            var d37_0 = mD37_0.as("pt");

            var mD49_0 = new UnitValue(49.0, "mm");
            var d49_0 = mD49_0.as("pt");

            var mD61_0 = new UnitValue(61.0, "mm");
            var d61_0 = mD61_0.as("pt");

            var mD73_0 = new UnitValue(73.0, "mm");
            var d73_0 = mD73_0.as("pt");

            var lineX = startingPoint + d24_0 + w10_0;

            var mLineH = new UnitValue(0.7, "mm");
            var lineH = mLineH.as("pt");

            var lineY = nsdTopMargin - (h13_5 / 2 - lineH / 2);

            var lineW = d37_0 - d24_0 - w10_0;

            // 非表示レイヤーやロックレイヤーへの対応
            try {
              // トリム用の矩形を描く
              var rectObj = activeDocument.pathItems.rectangle(
                docHeight - docHeight,
                0,
                docWidth,
                docHeight
              );
              rectObj.fillColor = noColor;
              rectObj.stroked = true; // 線を表示する
              rectObj.strokeWidth = 0; // 線の幅を指定する
              rectObj.strokeColor = noColor;

              // グループオブジェクトを生成
              var groupObj = activeDocument.groupItems.add();

              // ハイフンを描く
              var rectObj = groupObj.pathItems.rectangle(
                lineY,
                lineX,
                lineW,
                lineH
              );
              rectObj.fillColor = borderColor;
              rectObj.stroked = true; // 線を表示する
              rectObj.strokeWidth = 0; // 線の幅を指定する
              rectObj.strokeColor = noColor;

              // 四角形を描く
              var array = [
                [nsdTopMargin + b0_7, startingPoint - b0_7, w11_4, h14_9, 1],
                [
                  nsdTopMargin + b0_7,
                  startingPoint + d12_0 - b0_7,
                  w11_4,
                  h14_9,
                  1,
                ],
                [
                  nsdTopMargin + b0_7,
                  startingPoint + d24_0 - b0_7,
                  w11_4,
                  h14_9,
                  1,
                ],
                [
                  nsdTopMargin + b0_4,
                  startingPoint + d37_0 - b0_4,
                  w10_8,
                  h14_3,
                  2,
                ],
                [
                  nsdTopMargin + b0_4,
                  startingPoint + d49_0 - b0_4,
                  w10_8,
                  h14_3,
                  2,
                ],
                [
                  nsdTopMargin + b0_4,
                  startingPoint + d61_0 - b0_4,
                  w10_8,
                  h14_3,
                  2,
                ],
                [
                  nsdTopMargin + b0_4,
                  startingPoint + d73_0 - b0_4,
                  w10_8,
                  h14_3,
                  2,
                ],
                [nsdTopMargin, startingPoint, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d12_0, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d24_0, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d37_0, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d49_0, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d61_0, w10_0, h13_5, 3],
                [nsdTopMargin, startingPoint + d73_0, w10_0, h13_5, 3],
              ];

              for (var i = 0; i < array.length; i++) {
                for (var j = 0; j < array[i].length; j++) {
                  y = array[i][0];
                  x = array[i][1];
                  w = array[i][2];
                  h = array[i][3];
                  s = array[i][4];
                }

                if (s === 1) {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = borderColor;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                } else if (s === 2) {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = borderColor;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                } else {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                }
              }
              // alert("定型外サイズたて長規格内");
            } catch (e) {
              alert(
                "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
              );
            }
          } else {
            //定形外サイズよこ長規格内
            // 郵便枠の内側サイズ
            var mH10_0 = new UnitValue(10.0, "mm");
            var h10_0 = mH10_0.as("pt");
            var mW13_5 = new UnitValue(13.5, "mm");
            var w13_5 = mW13_5.as("pt");

            // 郵便枠前３の外側サイズ
            var mH11_4 = new UnitValue(11.4, "mm");
            var h11_4 = mH11_4.as("pt");
            var mW14_9 = new UnitValue(14.9, "mm");
            var w14_9 = mW14_9.as("pt");

            // 郵便枠後4の外側サイズ
            var mH10_8 = new UnitValue(10.8, "mm");
            var h10_8 = mH10_8.as("pt");
            var mW14_3 = new UnitValue(14.3, "mm");
            var w14_3 = mW14_3.as("pt");

            // 枠の太さ
            var mB0_7 = new UnitValue(0.7, "mm");
            var b0_7 = mB0_7.as("pt");
            var mB0_4 = new UnitValue(0.4, "mm");
            var b0_4 = mB0_4.as("pt");

            // 下側の指定マージン
            var mBottomMargin = new UnitValue(18.0, "mm");
            var bottomMargin = mBottomMargin.as("pt");

            // 右側の指定マージン
            var mRightMargin = new UnitValue(18.0, "mm");
            var rightMargin = mRightMargin.as("pt");

            // 郵便枠の総幅
            var mPostFrameSize83_0 = new UnitValue(83.0, "mm");
            var nsdPostFrameSize = mPostFrameSize83_0.as("pt");

            // ドキュメントサイズ時の内側枠の開始の起点
            var newStartingPoint =
              docHeight -
              docHeight -
              (docHeight - nsdPostFrameSize - bottomMargin);

            // 内側矩形の共通x座標
            var xPoint = docWidth - (w13_5 + rightMargin);

            // 枠の指定配置位置までの間隔
            var mD12_0 = new UnitValue(12.0, "mm");
            var d12_0 = mD12_0.as("pt");

            var mD24_0 = new UnitValue(24.0, "mm");
            var d24_0 = mD24_0.as("pt");

            var mD37_0 = new UnitValue(37.0, "mm");
            var d37_0 = mD37_0.as("pt");

            var mD49_0 = new UnitValue(49.0, "mm");
            var d49_0 = mD49_0.as("pt");

            var mD61_0 = new UnitValue(61.0, "mm");
            var d61_0 = mD61_0.as("pt");

            var mD73_0 = new UnitValue(73.0, "mm");
            var d73_0 = mD73_0.as("pt");

            // ハイフン用
            var mLineW = new UnitValue(0.7, "mm");
            var lineW = mLineW.as("pt");

            var lineX = docWidth - (w13_5 / 2 + lineW / 2) - rightMargin;

            var lineY = newStartingPoint - d24_0 - h10_0;

            var lineH = d37_0 - d24_0 - h10_0;

            // 非表示レイヤーやロックレイヤーへの対応
            try {
              // トリム用の矩形を描く
              var rectObj = activeDocument.pathItems.rectangle(
                docHeight - docHeight,
                0,
                docWidth,
                docHeight
              );
              rectObj.fillColor = noColor;
              rectObj.stroked = true; // 線を表示する
              rectObj.strokeWidth = 0; // 線の幅を指定する
              rectObj.strokeColor = noColor;

              // グループオブジェクトを生成
              var groupObj = activeDocument.groupItems.add();

              // ハイフンを描く
              var rectObj = groupObj.pathItems.rectangle(
                lineY,
                lineX,
                lineW,
                lineH
              );
              rectObj.fillColor = borderColor;
              rectObj.stroked = true; // 線を表示する
              rectObj.strokeWidth = 0; // 線の幅を指定する
              rectObj.strokeColor = noColor;

              // 四角形を描く
              var array = [
                [newStartingPoint + b0_7, xPoint - b0_7, w14_9, h11_4, 1],
                [
                  newStartingPoint - d12_0 + b0_7,
                  xPoint - b0_7,
                  w14_9,
                  h11_4,
                  1,
                ],
                [
                  newStartingPoint - d24_0 + b0_7,
                  xPoint - b0_7,
                  w14_9,
                  h11_4,
                  1,
                ],
                [
                  newStartingPoint - d37_0 + b0_4,
                  xPoint - b0_4,
                  w14_3,
                  h10_8,
                  2,
                ],
                [
                  newStartingPoint - d49_0 + b0_4,
                  xPoint - b0_4,
                  w14_3,
                  h10_8,
                  2,
                ],
                [
                  newStartingPoint - d61_0 + b0_4,
                  xPoint - b0_4,
                  w14_3,
                  h10_8,
                  2,
                ],
                [
                  newStartingPoint - d73_0 + b0_4,
                  xPoint - b0_4,
                  w14_3,
                  h10_8,
                  2,
                ],
                [newStartingPoint, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d12_0, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d24_0, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d37_0, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d49_0, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d61_0, xPoint, w13_5, h10_0, 3],
                [newStartingPoint - d73_0, xPoint, w13_5, h10_0, 3],
              ];

              for (var i = 0; i < array.length; i++) {
                for (var j = 0; j < array[i].length; j++) {
                  y = array[i][0];
                  x = array[i][1];
                  w = array[i][2];
                  h = array[i][3];
                  s = array[i][4];
                }

                if (s === 1) {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = borderColor;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                } else if (s === 2) {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = borderColor;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                } else {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = 0; // 線の幅を指定する
                  rectObj.strokeColor = noColor;
                }
              }
              // alert("定形外サイズよこ長規格内");
            } catch (e) {
              alert(
                "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
              );
            }
          }
        } else {
          alert("規格外のサイズです");
        }
      }
    } else {
      alert(
        "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：すでにオブジェクトが存在するレイヤー"
      );
    }
  }
})();
