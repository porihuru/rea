// print.js
// v2025.11.23-02
// 「全データコピー」で生成したテキストから請求書レイアウトを印刷する
// Edge 95 相当の古めブラウザ＋ローカルサーバ前提

(function () {
  'use strict';

  //==============================
  // 文字列ユーティリティ
  //==============================
  function trim(s) {
    return String(s).replace(/^\s+|\s+$/g, '');
  }

  function splitLines(text) {
    if (!text) return [];
    return String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  }

  // 「1,234」「￥1,234」「¥1,234-」などを数値に
  function parseAmount(s) {
    if (typeof s !== 'string') s = String(s || '');
    s = s.replace(/[¥￥,\s\-]/g, '');
    s = trim(s);
    if (!s) return NaN;
    var num = Number(s);
    return isNaN(num) ? NaN : num;
  }

  // 数値をカンマ付き・少数2桁で整形
  function formatAmountFixed2(num) {
    if (num === null || num === undefined || isNaN(num)) return '';
    var sign = num < 0 ? '-' : '';
    num = Math.abs(num);
    var s = num.toFixed(2); // "1234.50"
    var parts = s.split('.');
    var intPart = parts[0];
    var decPart = parts.length > 1 ? parts[1] : '';
    var res = '';
    while (intPart.length > 3) {
      res = ',' + intPart.slice(-3) + res;
      intPart = intPart.slice(0, -3);
    }
    res = intPart + res;
    if (decPart) res += '.' + decPart;
    return sign + res;
  }

  // 数値をカンマ付き整数に整形
  function formatAmountInt(num) {
    if (num === null || num === undefined || isNaN(num)) return '';
    num = Math.round(num);
    var sign = num < 0 ? '-' : '';
    num = Math.abs(num);
    var s = String(num);
    var res = '';
    while (s.length > 3) {
      res = ',' + s.slice(-3) + res;
      s = s.slice(0, -3);
    }
    res = s + res;
    return sign + res;
  }

  // HTMLエスケープ
  function htmlEscape(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  //==============================
  // パース用ヘルパー
  //==============================

  // 行頭の「日付: 2025/11/23」「宛先1: ○○」のような「キー: 値」形式を解析
  function parseKeyValueLine(line) {
    var m = line.match(/^([^:：]+)[:：]\s*(.*)$/);
    if (!m) return null;
    return {
      key: trim(m[1]),
      value: trim(m[2])
    };
  }

  // [最終ページ集計] セクションをパース
  function parseFooterSection(lines, startIndex) {
    var footer = {
      "課税対象額": '',
      "消費税": '',
      "合計": ''
    };
    for (var i = startIndex + 1; i < lines.length; i++) {
      var line = trim(lines[i]);
      if (!line) continue;
      if (line.indexOf('納入台帳テキスト解析') !== -1) {
        // バージョン行に到達したので終了
        break;
      }
      var kv = parseKeyValueLine(line);
      if (!kv) continue;
      if (kv.key === '課税対象額') {
        footer['課税対象額'] = kv.value;
      } else if (kv.key === '消費税') {
        footer['消費税'] = kv.value;
      } else if (kv.key === '合計') {
        footer['合計'] = kv.value;
      }
    }
    return footer;
  }

  // 明細テーブルの行をパース
  // No, 品名, 規格, 単位, 合計数量, 契約単価, 金額, 備考
  function parseDetailLines(lines, startIndex) {
    var details = [];
    var i;
    // ヘッダー行を探す
    var headerIndex = -1;
    for (i = startIndex; i < lines.length; i++) {
      var t = trim(lines[i]);
      if (!t) continue;
      // ここでは "No\t品名\t規格\t単位\t合計数量\t契約単価\t金額\t備考" 形式を想定
      if (t.indexOf('No') === 0 && t.indexOf('\t') !== -1 && t.indexOf('品名') !== -1) {
        headerIndex = i;
        break;
      }
    }
    if (headerIndex === -1) return {
      details: [],
      endIndex: startIndex
    };

    // 明細行はヘッダーの次の行から開始
    for (i = headerIndex + 1; i < lines.length; i++) {
      var line = lines[i];
      var trimmed = trim(line);
      if (!trimmed) continue;

      // [最終ページ集計] に到達したら終了
      if (trimmed.indexOf('[最終ページ集計]') === 0) {
        break;
      }

      var cols = line.split('\t');
      if (cols.length < 7) {
        // 列数が足りない場合はスキップ
        continue;
      }

      var detail = {
        No: trim(cols[0]),
        品名: trim(cols[1]),
        規格: trim(cols[2]),
        単位: trim(cols[3]),
        合計数量: trim(cols[4]),
        契約単価: trim(cols[5]),
        金額: trim(cols[6]),
        備考: cols.length >= 8 ? trim(cols[7]) : ''
      };

      // No のない行は終了条件とみなす
      if (!detail.No) {
        break;
      }

      details.push(detail);
    }

    return {
      details: details,
      endIndex: i
    };
  }

  // ヘッダー（宛先・業者名など）をパース
  function parseHeader(lines) {
    var header = {
      date: '',
      atesaki1: '',
      atesaki2: '',
      atesaki3: '',
      gyousha: '',
      gyousha_daihyou: '',
      gyousha_addr: '',
      gyousha_tantou: '',
      gyousha_tel: ''
    };

    // 上部情報から日付・宛先・業者を拾う想定
    // 例：
    // 日付: 2025/11/23
    // 宛先1: 滝川駐屯地 会計隊 御中
    // 宛先2: 〒073-0000 北海道滝川市○○
    // 宛先3: TEL 0124-00-0000
    // 業者名: トワニ旭川
    // 代表: 太郎
    // 業者住所: 〒070-0000 北海道旭川市○○
    // 担当: 佐藤
    // 業者TEL: 0166-00-0000

    for (var i = 0; i < lines.length; i++) {
      var line = trim(lines[i]);
      if (!line) continue;

      var kv = parseKeyValueLine(line);
      if (!kv) continue;

      switch (kv.key) {
        case '日付':
          header.date = kv.value;
          break;
        case '宛先1':
          header.atesaki1 = kv.value;
          break;
        case '宛先2':
          header.atesaki2 = kv.value;
          break;
        case '宛先3':
          header.atesaki3 = kv.value;
          break;
        case '業者名':
          header.gyousha = kv.value;
          break;
        case '代表':
          header.gyousha_daihyou = kv.value;
          break;
        case '業者住所':
          header.gyousha_addr = kv.value;
          break;
        case '担当':
          header.gyousha_tantou = kv.value;
          break;
        case '業者TEL':
          header.gyousha_tel = kv.value;
          break;
        default:
          break;
      }
    }

    return header;
  }

  //==============================
  // 印刷用 HTML 生成
  //==============================

  // 1ページ目最大明細数
  var FIRST_PAGE_MAX_ROWS = 15;
  // 2ページ目以降の最大明細数
  var OTHER_PAGE_MAX_ROWS = 25;

  // ページ構造を組み立てる
  function buildPages(header, details, footer) {
    var pages = [];
    var totalRows = details.length;
    var index = 0;

    // 1ページ目
    var firstPageRows = Math.min(totalRows, FIRST_PAGE_MAX_ROWS);
    pages.push({
      pageNo: 1,
      items: details.slice(0, firstPageRows)
    });
    index = firstPageRows;

    // 2ページ目以降
    var pageNo = 2;
    while (index < totalRows) {
      var rows = Math.min(totalRows - index, OTHER_PAGE_MAX_ROWS);
      pages.push({
        pageNo: pageNo,
        items: details.slice(index, index + rows)
      });
      index += rows;
      pageNo++;
    }

    // 合計ページ数をセット
    var totalPages = pages.length;
    for (var i = 0; i < pages.length; i++) {
      pages[i].totalPages = totalPages;
    }

    return pages;
  }

  // ページごとの小計を計算
  function calcPageSubtotal(items) {
    var sum = 0;
    for (var i = 0; i < items.length; i++) {
      var amt = parseAmount(items[i].金額);
      if (!isNaN(amt)) {
        sum += amt;
      }
    }
    return sum;
  }

  // フッター合計（課税対象額・消費税・合計）を解釈
  function parseFooterAmounts(footer) {
    var result = {
      taxable: 0,
      tax: 0,
      total: 0
    };

    if (footer) {
      var v1 = parseAmount(footer['課税対象額']);
      var v2 = parseAmount(footer['消費税']);
      var v3 = parseAmount(footer['合計']);
      if (!isNaN(v1)) result.taxable = v1;
      if (!isNaN(v2)) result.tax = v2;
      if (!isNaN(v3)) result.total = v3;
    }

    return result;
  }

  //==============================
  // 印刷ドキュメント生成
  //==============================

  function buildPrintHtml(header, pages, footer, rawText) {
    var footerAmounts = parseFooterAmounts(footer);

    var html = [];
    html.push('<!DOCTYPE html>');
    html.push('<html lang="ja">');
    html.push('<head>');
    html.push('<meta charset="utf-8">');
    html.push('<title>請求書印刷</title>');
    // 印刷用スタイル
    html.push('<style>');
    html.push('body { margin: 0; padding: 0; font-family: "Yu Gothic", "游ゴシック", sans-serif; font-size: 11pt; }');
    html.push('.page { width: 210mm; height: 297mm; box-sizing: border-box; padding: 20mm 15mm; position: relative; page-break-after: always; }');
    html.push('.title-row { text-align: center; font-size: 18pt; font-weight: bold; margin-bottom: 10mm; }');
    html.push('.title-right { position: absolute; top: 20mm; right: 20mm; text-align: right; font-size: 10pt; }');
    html.push('.atesaki-block { margin-bottom: 6mm; }');
    html.push('.atesaki-label { font-weight: bold; margin-right: 4mm; }');
    html.push('.atesaki-box { display: inline-block; border: 1px solid #000; padding: 3mm 4mm; min-width: 110mm; vertical-align: top; }');
    html.push('.gyousha-row { margin-bottom: 6mm; }');
    html.push('.seikyuu-gaku { display: inline-block; min-width: 60mm; font-weight: bold; }');
    html.push('.gyousha-box { display: inline-block; border: 1px solid #000; padding: 3mm 4mm; min-width: 75mm; max-width: 80mm; vertical-align: top; }');
    html.push('.detail-title { font-weight: bold; margin-top: 4mm; margin-bottom: 1mm; }');
    html.push('.detail-table { width: 100%; border-collapse: collapse; table-layout: fixed; }');
    html.push('.detail-table th, .detail-table td { border: 1px solid #000; padding: 1mm 1.5mm; font-size: 9pt; }');
    html.push('.detail-table th { text-align: center; background-color: #f0f0f0; }');
    html.push('.detail-table td { vertical-align: top; }');
    html.push('.col-no { width: 8mm; text-align: center; }');
    html.push('.col-hinmei { width: 35mm; }');
    html.push('.col-kikaku { width: 25mm; }');
    html.push('.col-tani { width: 10mm; text-align: center; }');
    html.push('.col-qty { width: 20mm; text-align: right; }');
    html.push('.col-tanka { width: 20mm; text-align: right; }');
    html.push('.col-kingaku { width: 25mm; text-align: right; }');
    html.push('.col-bikou { width: auto; }');
    html.push('.page-subtotal { margin-top: 3mm; text-align: right; font-weight: bold; }');
    html.push('.footer-total { position: absolute; bottom: 25mm; right: 20mm; text-align: right; font-weight: bold; }');
    html.push('.footer-total div { margin-bottom: 1.5mm; }');
    html.push('.footer-version { position: absolute; bottom: 10mm; right: 20mm; font-size: 8pt; }');
    html.push('@page { size: A4; margin: 0; }');
    html.push('@media print { body { margin: 0; } .page { box-shadow: none; } }');
    html.push('</style>');
    html.push('</head>');
    html.push('<body>');

    for (var p = 0; p < pages.length; p++) {
      var page = pages[p];
      var isLastPage = (p === pages.length - 1);
      var pageItems = page.items;
      var pageSubtotal = calcPageSubtotal(pageItems);

      html.push('<div class="page">');

      // タイトル行
      html.push('<div class="title-row">請&nbsp;&nbsp;求&nbsp;&nbsp;書</div>');
      html.push('<div class="title-right">');
      html.push('日付: ' + htmlEscape(header.date || '') + '<br>');
      html.push(page.pageNo + ' / ' + page.totalPages);
      html.push('</div>');

      // 宛先ブロック
      html.push('<div class="atesaki-block">');
      html.push('<span class="atesaki-label">宛先</span>');
      html.push('<div class="atesaki-box">');
      if (header.atesaki1) {
        html.push(htmlEscape(header.atesaki1) + '<br>');
      }
      if (header.atesaki2) {
        html.push(htmlEscape(header.atesaki2) + '<br>');
      }
      if (header.atesaki3) {
        html.push(htmlEscape(header.atesaki3));
      }
      html.push('</div>');
      html.push('</div>');

      // 業者・請求額ブロック
      html.push('<div class="gyousha-row">');
      html.push('<div class="seikyuu-gaku">請求額 ￥' + formatAmountInt(footerAmounts.total) + '</div>');
      html.push('<div class="gyousha-box">');
      if (header.gyousha) {
        html.push(htmlEscape(header.gyousha) + '<br>');
      }
      if (header.gyousha_daihyou) {
        html.push('代表 ' + htmlEscape(header.gyousha_daihyou) + '<br>');
      }
      if (header.gyousha_addr) {
        html.push(htmlEscape(header.gyousha_addr) + '<br>');
      }
      if (header.gyousha_tantou || header.gyousha_tel) {
        var line = '';
        if (header.gyousha_tantou) {
          line += '担当 ' + htmlEscape(header.gyousha_tantou) + ' ';
        }
        if (header.gyousha_tel) {
          line += 'TEL ' + htmlEscape(header.gyousha_tel);
        }
        html.push(line);
      }
      html.push('</div>');
      html.push('</div>');

      // 明細タイトル
      html.push('<div class="detail-title">請求明細書</div>');

      // 明細テーブル
      html.push('<table class="detail-table">');
      html.push('<thead>');
      html.push('<tr>');
      html.push('<th class="col-no">No</th>');
      html.push('<th class="col-hinmei">品名</th>');
      html.push('<th class="col-kikaku">規格</th>');
      html.push('<th class="col-tani">単位</th>');
      html.push('<th class="col-qty">合計数量</th>');
      html.push('<th class="col-tanka">契約単価</th>');
      html.push('<th class="col-kingaku">金額</th>');
      html.push('<th class="col-bikou">備考</th>');
      html.push('</tr>');
      html.push('</thead>');
      html.push('<tbody>');

      for (var i = 0; i < pageItems.length; i++) {
        var item = pageItems[i];
        html.push('<tr>');
        html.push('<td class="col-no">' + htmlEscape(item.No || '') + '</td>');
        html.push('<td class="col-hinmei">' + htmlEscape(item.品名 || '') + '</td>');
        html.push('<td class="col-kikaku">' + htmlEscape(item.規格 || '') + '</td>');
        html.push('<td class="col-tani">' + htmlEscape(item.単位 || '') + '</td>');
        html.push('<td class="col-qty">' + htmlEscape(item.合計数量 || '') + '</td>');
        html.push('<td class="col-tanka">' + htmlEscape(item.契約単価 || '') + '</td>');
        html.push('<td class="col-kingaku">' + htmlEscape(item.金額 || '') + '</td>');
        html.push('<td class="col-bikou">' + htmlEscape(item.備考 || '') + '</td>');
        html.push('</tr>');
      }

      html.push('</tbody>');
      html.push('</table>');

      // ページ小計（2ページ以上あるときだけ表示）
      if (pages.length > 1) {
        html.push('<div class="page-subtotal">小計: ' + formatAmountFixed2(pageSubtotal) + '</div>');
      }

      // 最終ページのフッター合計・バージョン
      if (isLastPage) {
        html.push('<div class="footer-total">');
        html.push('<div>合計: ' + formatAmountFixed2(footerAmounts.taxable) + '</div>');
        html.push('<div>消費税: ' + formatAmountInt(footerAmounts.tax) + '</div>');
        html.push('<div>総合計: ' + formatAmountInt(footerAmounts.total) + '</div>');
        html.push('</div>');

        // rawText 末尾のバージョン行を探す
        var versionLine = '';
        var lines = splitLines(rawText || '');
        for (var vi = lines.length - 1; vi >= 0; vi--) {
          var lt = trim(lines[vi]);
          if (!lt) continue;
          if (lt.indexOf('納入台帳テキスト解析') !== -1) {
            versionLine = lt;
            break;
          }
        }
        if (versionLine) {
          html.push('<div class="footer-version">' + htmlEscape(versionLine) + '</div>');
        }
      }

      html.push('</div>'); // .page
    }

    html.push('</body>');
    html.push('</html>');

    return html.join('');
  }

  //==============================
  // メイン関数
  //==============================

  function printLedgerData(allText) {
    if (!allText) {
      alert('印刷データが空です。');
      return;
    }

    var lines = splitLines(allText);
    var header = parseHeader(lines);

    // 明細ブロックを探してパース
    var detailResult = parseDetailLines(lines, 0);
    var details = detailResult.details;

    // フッタ（[最終ページ集計]）を探す
    var footer = null;
    for (var i = 0; i < lines.length; i++) {
      if (trim(lines[i]).indexOf('[最終ページ集計]') === 0) {
        footer = parseFooterSection(lines, i);
        break;
      }
    }

    var pages = buildPages(header, details, footer);
    var html = buildPrintHtml(header, pages, footer, allText);

    var win = window.open('', '_blank');
    if (!win) {
      alert('ポップアップがブロックされました。ブラウザの設定を確認してください。');
      return;
    }
    win.document.open();
    win.document.write(html);
    win.document.close();

    // フォーカスして印刷ダイアログ
    try {
      win.focus();
      win.print();
    } catch (e) {
      // 何もしない
    }
  }

  // グローバル公開
  window.printLedgerData = printLedgerData;

})();
