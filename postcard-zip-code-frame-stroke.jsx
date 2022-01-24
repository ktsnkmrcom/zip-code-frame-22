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

  // 塗りのカラーを「なし」に指定する設定用
  var noColor = new NoColor();

  // 枠の内側の色　塗り
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

  // ドキュメントの有無
  var n = documents.length;

  if (n === 0) {
    // ファイルが開かれていない
    // 官製はがきサイズ　ミリメートルをポイントに換算
    var mW100 = new UnitValue(100, "mm");
    var w100 = mW100.as("pt");
    var mH148 = new UnitValue(148, "mm");
    var h148 = mH148.as("pt");

    // 新規にファイルを作成
    var dp = DocumentPreset;
    dp.colorMode = DocumentColorSpace.CMYK;
    dp.rasterResolution = DocumentRasterResolution.HighResolution;
    dp.units = RulerUnits.Millimeters;
    dp.width = w100;
    dp.height = h148;
    documents.addDocument("Standerd100x148", dp);

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

    // 線の開始点用
    var mW5_7 = new UnitValue(5.7, "mm");
    var w5_7 = mW5_7.as("pt");

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

    // 新規ドキュメントサイズ時（100x148）の内側枠の開始の起点
    var newStartingPoint = w100 - postFrameSize - rightMargin;

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

    // ハイフンの線
    var lineX = newStartingPoint + d14_0 + w6_2;

    var mLineH = new UnitValue(0.5, "mm");
    var lineH = mLineH.as("pt");

    var lineY = topMargin3 - (h8_5 / 2 - lineH / 2);

    var lineW = d21_7 - d14_0 - w6_2;

    // トリム用の矩形を描く
    var rectObj = activeDocument.pathItems.rectangle(
      h148 - h148,
      0,
      w100,
      h148
    );
    rectObj.fillColor = noColor;
    rectObj.stroked = true; // 線を表示する
    rectObj.strokeWidth = 0; // 線の幅を指定する
    rectObj.strokeColor = noColor;

    // テキストフレームを作成し段落に文字挿入
    var txtObj = activeDocument.textFrames.add();
    txtObj.left = w100 / 2;
    txtObj.top = -10;
    txtObj.contents = "POSTCARD";
    txtObj.paragraphs[0].size = 7; // 段落の文字サイズ設定
    txtObj.paragraphs[0].justification = Justification.CENTER;

    // グループオブジェクトを生成
    var groupObj = activeDocument.groupItems.add();

    // ハイフンを線で描画
    var lineObj = groupObj.pathItems.add();
    lineObj.setEntirePath([
      [lineX, lineY - lineH / 2],
      [lineX + lineW, lineY - lineH / 2],
    ]);
    lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
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
    // ファイルが開かれている
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

      // はがき最大サイズ　ミリメートルをポイントに換算
      var mL107 = new UnitValue(107, "mm");
      var l107 = mL107.as("pt");
      var mL154 = new UnitValue(154, "mm");
      var l154 = mL154.as("pt");

      // はがき最小サイズ　ミリメートルをポイントに換算
      var mL90 = new UnitValue(90, "mm");
      var l90 = mL90.as("pt");
      var mL140 = new UnitValue(140, "mm");
      var l140 = mL140.as("pt");

      // はがきサイズか判定
      if (
        shortSide <= mL107 &&
        shortSide >= mL90 &&
        longSide <= mL154 &&
        longSide >= mL140
      ) {
        // たて長か判定
        if (mdocWidth < mdocHeight) {
          // はがきサイズたて長
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

          // 線の開始点用
          var mW5_7 = new UnitValue(5.7, "mm");
          var w5_7 = mW5_7.as("pt");

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

          // 内側枠の開始の起点
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

            // テキストフレームを作成し段落に文字挿入
            var txtObj = activeDocument.textFrames.add();
            txtObj.left = docWidth / 2;
            txtObj.top = -10;
            txtObj.contents = "POSTCARD";
            txtObj.paragraphs[0].size = 7; // 段落の文字サイズ設定
            txtObj.paragraphs[0].justification = Justification.CENTER;

            // グループオブジェクトを生成
            var groupObj = activeDocument.groupItems.add();

            // ハイフンを線で描画　はがき定型たて
            var lineObj = groupObj.pathItems.add();
            lineObj.setEntirePath([
              [lineX, lineY - lineH / 2],
              [lineX + lineW, lineY - lineH / 2],
            ]);
            lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
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
            // alert("はがきサイズたて長");
          } catch (e) {
            alert(
              "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
            );
          }
        } else {
          // はがきサイズよこ長
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

          // 線の開始点用
          var mW5_7 = new UnitValue(5.7, "mm");
          var w5_7 = mW5_7.as("pt");

          // 線の太さ
          var mB0_5 = new UnitValue(0.5, "mm");
          var b0_5 = mB0_5.as("pt");
          var mB0_3 = new UnitValue(0.3, "mm");
          var b0_3 = mB0_3.as("pt");

          // 下側の指定マージン
          var mBottomMargin = new UnitValue(7.85, "mm");
          var bottomMargin = mBottomMargin.as("pt");

          // 右側の指定マージン　３けた枠
          var mRightMargin3 = new UnitValue(11.75, "mm");
          var rightMargin3 = mRightMargin3.as("pt");

          // 右側の指定マージン　４けた枠
          var mRightMargin4 = new UnitValue(11.847, "mm");
          var rightMargin4 = mRightMargin4.as("pt");

          // 郵便枠の総幅
          var mPostFrameSize = new UnitValue(48.1, "mm");
          var postFrameSize = mPostFrameSize.as("pt");

          // ドキュメントサイズ時の内側枠の開始の起点
          var newStartingPoint =
            docHeight - docHeight - (docHeight - postFrameSize - bottomMargin);

          // よこ長のみ　３けた枠と４けた枠
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

          // ハイフンの矩形
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

            // テキストフレームを作成し段落に文字挿入
            var txtObj = activeDocument.textFrames.add();
            txtObj.left = docWidth / 2;
            txtObj.top = -10;
            txtObj.contents = "POSTCARD";
            txtObj.paragraphs[0].size = 7; // 段落の文字サイズ設定
            txtObj.paragraphs[0].justification = Justification.CENTER;

            // グループオブジェクトを生成
            var groupObj = activeDocument.groupItems.add();

            // ハイフンを線で描画　はがき定型よこ
            var lineObj = groupObj.pathItems.add();
            lineObj.setEntirePath([
              [lineX + lineW / 2, lineY],
              [lineX + lineW / 2, lineY - lineH],
            ]);
            lineObj.stroked = true; // 線を表示する。塗りは指定しないと「なし」に設定される
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
            // alert("はがきサイズよこ長");
          } catch (e) {
            alert(
              "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：非表示レイヤー、ロックされたレイヤー"
            );
          }
        }
      } else {
        alert("はがきの規定サイズ外です");
      }
    } else {
      alert(
        "最前面に新規レイヤー（空のレイヤー）を作成してやりなおす　理由：すでにオブジェクトが存在するレイヤー"
      );
    }
  }
})();
