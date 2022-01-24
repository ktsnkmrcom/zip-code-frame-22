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

  // 枠の色　線　ストロークで表現
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

    // 郵便枠前３の内側サイズ
    var mW6_2 = new UnitValue(6.2, "mm");
    var w6_2 = mW6_2.as("pt");
    var mH8_5 = new UnitValue(8.5, "mm");
    var h8_5 = mH8_5.as("pt");

    // 郵便枠後４の内側サイズ
    var mW6_0 = new UnitValue(6.0, "mm");
    var w6_0 = mW6_0.as("pt");
    var mH8_3 = new UnitValue(8.3, "mm");
    var h8_3 = mH8_3.as("pt");

    // 枠　線の太さ
    var mB0_5 = new UnitValue(0.5, "mm");
    var b0_5 = mB0_5.as("pt");
    var mB0_3 = new UnitValue(0.3, "mm");
    var b0_3 = mB0_3.as("pt");

    // 右側の指定マージン
    var mRightMargin = new UnitValue(7.85, "mm");
    var rightMargin = mRightMargin.as("pt");

    // 前３の上側の指定マージン
    var mTopMargin3 = new UnitValue(-11.75, "mm");
    var topMargin3 = mTopMargin3.as("pt");

    // 後４の上側の指定マージン
    var mTopMargin4 = new UnitValue(-11.847, "mm");
    var topMargin4 = mTopMargin4.as("pt");

    // 郵便枠の総幅
    var mPostFrameSize = new UnitValue(48.1, "mm");
    var postFrameSize = mPostFrameSize.as("pt");

    // 新規ドキュメントサイズ時（120x235）の内側枠の開始の起点
    var newStartingPoint = w120 - postFrameSize - rightMargin;

    // 枠の指定配置位置までの間隔
    var mD7_0 = new UnitValue(7.0, "mm");
    var d7_0 = mD7_0.as("pt");

    var mD14_0 = new UnitValue(14.0, "mm");
    var d14_0 = mD14_0.as("pt");

    var mD21_7 = new UnitValue(21.7, "mm");
    var d21_7 = mD21_7.as("pt");

    var mD28_5 = new UnitValue(28.5, "mm");
    var d28_5 = mD28_5.as("pt");

    var mD35_3 = new UnitValue(35.3, "mm");
    var d35_3 = mD35_3.as("pt");

    var mD42_1 = new UnitValue(42.1, "mm");
    var d42_1 = mD42_1.as("pt");

    // ハイフンの矩形
    var lineX = newStartingPoint + d14_0 + w6_2;

    var mLineH = new UnitValue(0.5, "mm");
    var lineH = mLineH.as("pt");

    var lineY = topMargin3 - (h8_5 / 2 - lineH / 2);

    var lineW = d21_7 - d14_0 - w6_2;

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

    // ハイフンを線で描画　はがき定型たて
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [lineX, lineY - lineH / 2],
      [lineX + lineW, lineY - lineH / 2],
    ]);
    lineObj.stroked = true;
    lineObj.strokeWidth = b0_5;
    lineObj.strokeColor = borderColor;

    // 四角形を描く
    var array = [
      [topMargin3, newStartingPoint, w6_2, h8_5, 1],
      [topMargin3, newStartingPoint + d7_0, w6_2, h8_5, 1],
      [topMargin3, newStartingPoint + d14_0, w6_2, h8_5, 1],
      [topMargin4, newStartingPoint + d21_7, w6_0, h8_3, 2],
      [topMargin4, newStartingPoint + d28_5, w6_0, h8_3, 2],
      [topMargin4, newStartingPoint + d35_3, w6_0, h8_3, 2],
      [topMargin4, newStartingPoint + d42_1, w6_0, h8_3, 2],
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
        rectObj.fillColor = color;
        rectObj.stroked = true; // 線を表示する
        rectObj.strokeWidth = b0_5; // 線の幅を指定する
        rectObj.strokeColor = borderColor; //nocolor
      } else {
        var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
        rectObj.fillColor = color;
        rectObj.stroked = true; // 線を表示する
        rectObj.strokeWidth = b0_3; // 線の幅を指定する
        rectObj.strokeColor = borderColor; //nocolor
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
          // 郵便枠前３の内側サイズ
          var mW6_2 = new UnitValue(6.2, "mm");
          var w6_2 = mW6_2.as("pt");
          var mH8_5 = new UnitValue(8.5, "mm");
          var h8_5 = mH8_5.as("pt");

          // 郵便枠後４の内側サイズ
          var mW6_0 = new UnitValue(6.0, "mm");
          var w6_0 = mW6_0.as("pt");
          var mH8_3 = new UnitValue(8.3, "mm");
          var h8_3 = mH8_3.as("pt");

          // 線の太さ
          var mB0_5 = new UnitValue(0.5, "mm");
          var b0_5 = mB0_5.as("pt");
          var mB0_3 = new UnitValue(0.3, "mm");
          var b0_3 = mB0_3.as("pt");

          // 右側の指定マージン
          var mRightMargin = new UnitValue(7.85, "mm");
          var rightMargin = mRightMargin.as("pt");

          // 前３の上側の指定マージン
          var mTopMargin3 = new UnitValue(-11.75, "mm");
          var topMargin3 = mTopMargin3.as("pt");

          // 後４の上側の指定マージン
          var mTopMargin4 = new UnitValue(-11.847, "mm");
          var topMargin4 = mTopMargin4.as("pt");

          // 郵便枠の総幅
          var mPostFrameSize = new UnitValue(48.1, "mm");
          var postFrameSize = mPostFrameSize.as("pt");

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

          var mD21_7 = new UnitValue(21.7, "mm");
          var d21_7 = mD21_7.as("pt");

          var mD28_5 = new UnitValue(28.5, "mm");
          var d28_5 = mD28_5.as("pt");

          var mD35_3 = new UnitValue(35.3, "mm");
          var d35_3 = mD35_3.as("pt");

          var mD42_1 = new UnitValue(42.1, "mm");
          var d42_1 = mD42_1.as("pt");

          var lineX = newStartingPoint + d14_0 + w6_2;

          var mLineH = new UnitValue(0.5, "mm");
          var lineH = mLineH.as("pt");

          var lineY = topMargin3 - (h8_5 / 2 - lineH / 2);

          var lineW = d21_7 - d14_0 - w6_2;

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

            // ハイフンを線で描画　はがき定型たて
            var lineObj = groupObj.pathItems.add();
            lineObj.setEntirePath([
              [lineX, lineY - lineH / 2],
              [lineX + lineW, lineY - lineH / 2],
            ]);
            lineObj.stroked = true;
            lineObj.strokeWidth = b0_5;
            lineObj.strokeColor = borderColor;

            // 四角形を描く
            var array = [
              [topMargin3, newStartingPoint, w6_2, h8_5, 1],
              [topMargin3, newStartingPoint + d7_0, w6_2, h8_5, 1],
              [topMargin3, newStartingPoint + d14_0, w6_2, h8_5, 1],
              [topMargin4, newStartingPoint + d21_7, w6_0, h8_3, 2],
              [topMargin4, newStartingPoint + d28_5, w6_0, h8_3, 2],
              [topMargin4, newStartingPoint + d35_3, w6_0, h8_3, 2],
              [topMargin4, newStartingPoint + d42_1, w6_0, h8_3, 2],
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
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = b0_5; // 線の幅を指定する
                rectObj.strokeColor = borderColor;
              } else {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = b0_3; // 線の幅を指定する
                rectObj.strokeColor = borderColor;
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
          // 郵便枠前３の内側サイズ
          var mH6_2 = new UnitValue(6.2, "mm");
          var h6_2 = mH6_2.as("pt");
          var mW8_5 = new UnitValue(8.5, "mm");
          var w8_5 = mW8_5.as("pt");

          // 郵便枠後４の内側サイズ
          var mH6_0 = new UnitValue(6.0, "mm");
          var h6_0 = mH6_0.as("pt");
          var mW8_3 = new UnitValue(8.3, "mm");
          var w8_3 = mW8_3.as("pt");

          // 線の太さ
          var mB0_5 = new UnitValue(0.5, "mm");
          var b0_5 = mB0_5.as("pt");
          var mB0_3 = new UnitValue(0.3, "mm");
          var b0_3 = mB0_3.as("pt");

          // 下側の指定マージン
          var mBottomMargin = new UnitValue(7.85, "mm");
          var bottomMargin = mBottomMargin.as("pt");

          // 右側の指定マージン
          var mRightMargin3 = new UnitValue(11.75, "mm");
          var rightMargin3 = mRightMargin3.as("pt");

          var mRightMargin4 = new UnitValue(11.847, "mm");
          var rightMargin4 = mRightMargin4.as("pt");

          // 郵便枠の総幅
          var mPostFrameSize = new UnitValue(48.1, "mm");
          var postFrameSize = mPostFrameSize.as("pt");

          // ドキュメントサイズ時の内側枠の開始の起点
          var newStartingPoint =
            docHeight - docHeight - (docHeight - postFrameSize - bottomMargin);

          // 矩形のx座標
          var xPoint3 = docWidth - (w8_5 + rightMargin3);
          var xPoint4 = docWidth - (w8_3 + rightMargin4);

          // 枠の指定配置位置までの間隔
          var mD7_0 = new UnitValue(7.0, "mm");
          var d7_0 = mD7_0.as("pt");

          var mD14_0 = new UnitValue(14.0, "mm");
          var d14_0 = mD14_0.as("pt");

          var mD21_7 = new UnitValue(21.7, "mm");
          var d21_7 = mD21_7.as("pt");

          var mD28_5 = new UnitValue(28.5, "mm");
          var d28_5 = mD28_5.as("pt");

          var mD35_3 = new UnitValue(35.3, "mm");
          var d35_3 = mD35_3.as("pt");

          var mD42_1 = new UnitValue(42.1, "mm");
          var d42_1 = mD42_1.as("pt");

          // ハイフン用
          var mLineW = new UnitValue(0.5, "mm");
          var lineW = mLineW.as("pt");

          var lineX = docWidth - (w8_5 / 2 + lineW / 2) - rightMargin3;

          var lineY = newStartingPoint - d14_0 - h6_2;

          var lineH = d21_7 - d14_0 - h6_2;

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

            // ハイフンを線で描画　定型よこ
            var lineObj = groupObj.pathItems.add();
            lineObj.setEntirePath([
              [lineX + lineW / 2, lineY],
              [lineX + lineW / 2, lineY - lineH],
            ]);
            lineObj.stroked = true;
            lineObj.strokeWidth = b0_5;
            lineObj.strokeColor = borderColor;

            // 四角形を描く
            var array = [
              [newStartingPoint, xPoint3, w8_5, h6_2, 1],
              [newStartingPoint - d7_0, xPoint3, w8_5, h6_2, 1],
              [newStartingPoint - d14_0, xPoint3, w8_5, h6_2, 1],
              [newStartingPoint - d21_7, xPoint4, w8_3, h6_0, 2],
              [newStartingPoint - d28_5, xPoint4, w8_3, h6_0, 2],
              [newStartingPoint - d35_3, xPoint4, w8_3, h6_0, 2],
              [newStartingPoint - d42_1, xPoint4, w8_3, h6_0, 2],
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
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = b0_5; // 線の幅を指定する
                rectObj.strokeColor = borderColor; //nocolor
              } else {
                var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                rectObj.fillColor = color;
                rectObj.stroked = true; // 線を表示する
                rectObj.strokeWidth = b0_3; // 線の幅を指定する
                rectObj.strokeColor = borderColor; //nocolor
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
            // 郵便枠前３の内側サイズ
            var mW10_7 = new UnitValue(10.7, "mm");
            var w10_7 = mW10_7.as("pt");
            var mH14_2 = new UnitValue(14.2, "mm");
            var h14_2 = mH14_2.as("pt");

            // 郵便枠後４の内側サイズ
            var mW10_4 = new UnitValue(10.4, "mm");
            var w10_4 = mW10_4.as("pt");
            var mH13_9 = new UnitValue(13.9, "mm");
            var h13_9 = mH13_9.as("pt");

            // 線の太さ
            var mB0_7 = new UnitValue(0.7, "mm");
            var b0_7 = mB0_7.as("pt");
            var mB0_4 = new UnitValue(0.4, "mm");
            var b0_4 = mB0_4.as("pt");

            // 右側の指定マージン
            var mRightMargin = new UnitValue(17.8, "mm");
            var nsdrightMargin = mRightMargin.as("pt");

            // 前３の上側の指定マージン
            var mTopMargin3 = new UnitValue(-17.648, "mm");
            var nsdtopMargin3 = mTopMargin3.as("pt");

            // 後４の上側の指定マージン
            var mTopMargin4 = new UnitValue(-17.8, "mm");
            var nsdtopMargin4 = mTopMargin4.as("pt");

            // 郵便枠の総幅
            var mPostFrameSize = new UnitValue(83.55, "mm");
            var nsdpostFrameSize = mPostFrameSize.as("pt");

            // 任意のドキュメントサイズ時の内側枠の開始の起点
            var startingPoint = docWidth - nsdpostFrameSize - nsdrightMargin;

            // 枠の指定配置位置までの間隔
            var mD12_0 = new UnitValue(12.0, "mm");
            var mD12_0 = mD12_0.as("pt");

            var mD24_0 = new UnitValue(24.0, "mm");
            var d24_0 = mD24_0.as("pt");

            var mD37_15 = new UnitValue(37.15, "mm");
            var d37_15 = mD37_15.as("pt");

            var mD49_15 = new UnitValue(49.15, "mm");
            var d49_15 = mD49_15.as("pt");

            var mD61_15 = new UnitValue(61.15, "mm");
            var d61_15 = mD61_15.as("pt");

            var mD73_15 = new UnitValue(73.15, "mm");
            var d73_15 = mD73_15.as("pt");

            var lineX = startingPoint + d24_0 + w10_7;

            var mLineH = new UnitValue(0.7, "mm");
            var lineH = mLineH.as("pt");

            var lineY = nsdtopMargin3 - (h14_2 / 2 - lineH / 2);

            var lineW = d37_15 - d24_0 - w10_7;

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

              // ハイフンを線で描画　はがき定型たて
              var lineObj = groupObj.pathItems.add();
              lineObj.setEntirePath([
                [lineX, lineY - lineH / 2],
                [lineX + lineW, lineY - lineH / 2],
              ]);
              lineObj.stroked = true;
              lineObj.strokeWidth = b0_7;
              lineObj.strokeColor = borderColor;

              // 四角形を描く
              var array = [
                [nsdtopMargin3, startingPoint, w10_7, h14_2, 1],
                [nsdtopMargin3, startingPoint + mD12_0, w10_7, h14_2, 1],
                [nsdtopMargin3, startingPoint + d24_0, w10_7, h14_2, 1],
                [nsdtopMargin4, startingPoint + d37_15, w10_4, h13_9, 2],
                [nsdtopMargin4, startingPoint + d49_15, w10_4, h13_9, 2],
                [nsdtopMargin4, startingPoint + d61_15, w10_4, h13_9, 2],
                [nsdtopMargin4, startingPoint + d73_15, w10_4, h13_9, 2],
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
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = b0_7; // 線の幅を指定する
                  rectObj.strokeColor = borderColor;
                } else {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = b0_4; // 線の幅を指定する
                  rectObj.strokeColor = borderColor;
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
            // 郵便枠前３の内側サイズ
            var mH10_7 = new UnitValue(10.7, "mm");
            var h10_7 = mH10_7.as("pt");
            var mW14_2 = new UnitValue(14.2, "mm");
            var w14_2 = mW14_2.as("pt");

            // 郵便枠後４の内側サイズ
            var mH10_4 = new UnitValue(10.4, "mm");
            var h10_4 = mH10_4.as("pt");
            var mW13_9 = new UnitValue(13.9, "mm");
            var w13_9 = mW13_9.as("pt");

            // 線の太さ
            var mB0_7 = new UnitValue(0.7, "mm");
            var b0_7 = mB0_7.as("pt");
            var mB0_4 = new UnitValue(0.4, "mm");
            var b0_4 = mB0_4.as("pt");

            // 右側の指定マージン
            var mBottomMargin = new UnitValue(17.8, "mm");
            var bottomMargin = mBottomMargin.as("pt");

            // 前３の上側の指定マージン
            var mRightMargin3 = new UnitValue(17.648, "mm");
            var nsdRightMargin3 = mRightMargin3.as("pt");

            // 後４の上側の指定マージン
            var mRightMargin4 = new UnitValue(17.8, "mm");
            var nsdRightMargin4 = mRightMargin4.as("pt");

            // 郵便枠の総幅
            var mPostFrameSize = new UnitValue(83.55, "mm");
            var nsdPostFrameSize = mPostFrameSize.as("pt");

            // ドキュメントサイズ時の内側枠の開始の起点
            var newStartingPoint =
              docHeight -
              docHeight -
              (docHeight - nsdPostFrameSize - bottomMargin);

            // 矩形のx座標
            var xPoint3 = docWidth - (w14_2 + nsdRightMargin3);
            var xPoint4 = docWidth - (w13_9 + nsdRightMargin4);

            // 枠の指定配置位置までの間隔
            var mD12_0 = new UnitValue(12.0, "mm");
            var mD12_0 = mD12_0.as("pt");

            var mD24_0 = new UnitValue(24.0, "mm");
            var d24_0 = mD24_0.as("pt");

            var mD37_15 = new UnitValue(37.15, "mm");
            var d37_15 = mD37_15.as("pt");

            var mD49_15 = new UnitValue(49.15, "mm");
            var d49_15 = mD49_15.as("pt");

            var mD61_15 = new UnitValue(61.15, "mm");
            var d61_15 = mD61_15.as("pt");

            var mD73_15 = new UnitValue(73.15, "mm");
            var d73_15 = mD73_15.as("pt");

            // ハイフン用
            var mLineW = new UnitValue(0.7, "mm");
            var lineW = mLineW.as("pt");

            var lineX = docWidth - (w14_2 / 2 + lineW / 2) - nsdRightMargin3;

            var lineY = newStartingPoint - d24_0 - h10_7;

            var lineH = d37_15 - d24_0 - h10_7;

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

              // ハイフンを線で描画　定型外よこ
              var lineObj = groupObj.pathItems.add();
              lineObj.setEntirePath([
                [lineX + lineW / 2, lineY],
                [lineX + lineW / 2, lineY - lineH],
              ]);
              lineObj.stroked = true;
              lineObj.strokeWidth = b0_7;
              lineObj.strokeColor = borderColor;

              // 四角形を描く
              var array = [
                [newStartingPoint, xPoint3, w14_2, h10_7, 1],
                [newStartingPoint - mD12_0, xPoint3, w14_2, h10_7, 1],
                [newStartingPoint - d24_0, xPoint3, w14_2, h10_7, 1],
                [newStartingPoint - d37_15, xPoint4, w13_9, h10_4, 2],
                [newStartingPoint - d49_15, xPoint4, w13_9, h10_4, 2],
                [newStartingPoint - d61_15, xPoint4, w13_9, h10_4, 2],
                [newStartingPoint - d73_15, xPoint4, w13_9, h10_4, 2],
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
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = b0_7; // 線の幅を指定する
                  rectObj.strokeColor = borderColor;
                } else {
                  var rectObj = groupObj.pathItems.rectangle(y, x, w, h);
                  rectObj.fillColor = color;
                  rectObj.stroked = true; // 線を表示する
                  rectObj.strokeWidth = b0_4; // 線の幅を指定する
                  rectObj.strokeColor = borderColor;
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
