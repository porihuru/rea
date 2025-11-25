// gyousya.js
// gyousya.txt から宛先・業者情報を読み込み、メイン画面のテキストボックスに反映
// v2025.11.xx-gyousya-02
//
// ■ gyousya.txt の想定フォーマット
//   1行目 : 宛先共通 3項目
//     宛先1行目,宛先2行目,宛先3行目
//       例)
//       滝川駐屯地 会計隊 御中,〒073-0000 北海道滝川市○○,TEL 0124-00-0000
//
//   2行目以降 : 業者ごとの 5項目（1列目がキーとなる「短い業者名」）
//     キー名,代表者名,住所,担当者,連絡先
//       例)
//       トワニ旭川,代表 太郎,〒070-0000 北海道旭川市○○,担当 佐藤,0166-00-0000
//       セイコーフレッシュフーズ,代表 花子,〒003-0000 札幌市白石区○○,担当 高橋,011-000-0000
//
// ■ マッチング方法
//   ・メイン側ヘッダーの「業者名: 株式会社トワニ旭川 …」の文字列から業者名を取り出し、
//     その文字列に gyousya.txt の 1列目（キー）が「含まれていれば」該当業者とみなす
//       例) headerVendor = "株式会社トワニ旭川食品"
//           rec.key      = "トワニ旭川"
//           → "株式会社トワニ旭川食品".indexOf("トワニ旭川") !== -1 なのでマッチ
//
// ■ 公開関数
//   window.applyGyousyaFromHeader(headerText, toTextareaId, vendorTextareaId)
//
//   ・headerText      … メインの headerBox.textContent（納地・業者名が入っている）
//   ・toTextareaId    … 宛先テキストエリアの id（省略時 'toText'）
//   ・vendorTextareaId… 業者テキストエリアの id（省略時 'vendorText'）
//
//   例: ボタン側
//     var headerText = document.getElementById('headerBox').textContent || '';
//     window.applyGyousyaFromHeader(headerText, 'toText', 'vendorText');

(function (global) {
  'use strict';

  // -------------------- 共通ユーティリティ --------------------

  function trim(str) {
    if (str == null) return '';
    return String(str).replace(/^\s+|\s+$/g, '');
  }

  // -------------------- gyousya.txt パース --------------------

  /**
   * gyousya.txt 全文 → レコード配列に変換
   *
   * 戻り値: [{ key, toLines:[3], vendorLines:[5] }, ...]
   *   key         … 1列目（マッチング用の短い業者名）
   *   toLines     … 宛先3行（1行目から取得）
   *   vendorLines … [業者名1行目, 代表者名, 住所, 担当者, 連絡先]
   */
  function parseGyousyaText(text) {
    var normalized = (text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    var lines = normalized.split('\n');

    var defaultTo = null;   // 宛先3行
    var records   = [];
    var i;

    for (i = 0; i < lines.length; i++) {
      var line = trim(lines[i]);
      if (!line) continue;

      var parts = line.split(',');
      var j;
      for (j = 0; j < parts.length; j++) {
        parts[j] = trim(parts[j]);
      }

      if (!defaultTo) {
        // ★最初の非空行 → 宛先
        defaultTo = [
          parts.length > 0 ? parts[0] : '',
          parts.length > 1 ? parts[1] : '',
          parts.length > 2 ? parts[2] : ''
        ];
        continue;
      }

      // 2行目以降 → 業者レコード
      var key = parts.length > 0 ? parts[0] : '';
      if (!key) {
        // キーが空なら無視
        continue;
      }

      var vendorLines = [
        parts.length > 0 ? parts[0] : '',
        parts.length > 1 ? parts[1] : '',
        parts.length > 2 ? parts[2] : '',
        parts.length > 3 ? parts[3] : '',
        parts.length > 4 ? parts[4] : ''
      ];

      records.push({
        key: key,
        toLines: defaultTo || ['', '', ''],
        vendorLines: vendorLines
      });
    }

    return records;
  }

  // -------------------- gyousya.txt ロード処理 --------------------

  var gyousyaRecords = null; // キャッシュ

  function showFileOpenDialog(callback) {
    if (!global.FileReader) {
      alert('このブラウザは FileReader に対応していません。\n別の環境でお試しください。');
      return;
    }

    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.txt';

    input.onchange = function () {
      if (!input.files || !input.files.length) return;
      var file = input.files[0];
      var reader = new FileReader();
      reader.onload = function (ev) {
        var txt = ev.target.result;
        var recs = parseGyousyaText(txt);
        if (!recs || !recs.length) {
          alert('gyousya.txt の内容を正しく読み込めませんでした。\nフォーマットを確認してください。');
          return;
        }
        gyousyaRecords = recs;
        if (callback) callback(gyousyaRecords);
      };
      reader.readAsText(file, 'UTF-8');
    };

    input.click();
  }

  function ensureGyousyaLoaded(callback) {
    if (gyousyaRecords && gyousyaRecords.length) {
      callback(gyousyaRecords);
    } else {
      // ★毎回ダイアログで選ばせる仕様に近い動き：
      //   ページをリロードするたびにキャッシュは消えるので、
      //   その都度このダイアログが出る。
      showFileOpenDialog(callback);
    }
  }

  // -------------------- ヘッダー → 業者名抽出 --------------------

  function extractVendorFromHeader(headerText) {
    var text = headerText || '';
    var lines = text.split(/\r\n|\r|\n/);
    var i;
    for (i = 0; i < lines.length; i++) {
      var t = trim(lines[i]);
      if (!t) continue;
      // 「業者名:」「業者名：」の行を探す
      if (t.indexOf('業者名:') === 0 || t.indexOf('業者名：') === 0) {
        return trim(t.replace(/^業者名[:：]\s*/, ''));
      }
    }
    return '';
  }

  function findGyousyaRecord(headerVendor) {
    if (!gyousyaRecords) return null;
    var hv = trim(headerVendor);
    if (!hv) return null;

    var i, rec;
    for (i = 0; i < gyousyaRecords.length; i++) {
      rec = gyousyaRecords[i];
      if (!rec.key) continue;

      // 「ヘッダーの業者名にキーが含まれる」or「キーに業者名が含まれる」
      if (hv.indexOf(rec.key) !== -1 || rec.key.indexOf(hv) !== -1) {
        return rec;
      }
    }
    return null;
  }

  // -------------------- メイン画面への反映 --------------------

  function applyGyousyaFromHeader(headerText, toTextareaId, vendorTextareaId) {
    ensureGyousyaLoaded(function () {
      var headerVendor = extractVendorFromHeader(headerText || '');
      if (!headerVendor) {
        alert('ヘッダーから業者名を取得できませんでした。\n先に納入台帳テキストを解析してください。');
        return;
      }

      var rec = findGyousyaRecord(headerVendor);
      if (!rec) {
        alert('gyousya.txt に該当する業者が見つかりませんでした。\n1列目のキー名とヘッダーの業者名の対応を確認してください。');
        return;
      }

      var toId     = toTextareaId || 'toText';
      var vendorId = vendorTextareaId || 'vendorText';

      var toTa     = document.getElementById(toId);
      var vendorTa = document.getElementById(vendorId);

      // 宛先 3 行
      if (toTa) {
        toTa.value =
          (rec.toLines[0] || '') + '\n' +
          (rec.toLines[1] || '') + '\n' +
          (rec.toLines[2] || '');
      }

      // 業者 5 行
      if (vendorTa) {
        var vls = rec.vendorLines;
        vendorTa.value =
          (vls[0] || '') + '\n' +
          (vls[1] || '') + '\n' +
          (vls[2] || '') + '\n' +
          (vls[3] || '') + '\n' +
          (vls[4] || '');
      }
    });
  }

  // 公開
  global.applyGyousyaFromHeader = applyGyousyaFromHeader;

})(this);
