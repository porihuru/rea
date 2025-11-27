// print.js
// 納入台帳テキスト解析ツール用 印刷プレビュー
// v2025.11.24-03
//
// ●構成（グループ）
//   ・印刷ヘッダー：請求書タイトル／日付／宛先／業者名／請求額
//   ・明細グループ：
//       「請求目明細」の見出し＋明細表＋（必要に応じて）ページ小計
//   ・印刷フッター：合計／消費税／総合計／バージョン文字列
//
// ●ページ分割ルール
//   ・品目数 15 以下：
//       1ページのみ
//       1ページ目：印刷ヘッダー＋明細グループ（小計なし）＋印刷フッター
//   ・品目数 15 を超える：
//       1ページ目：印刷ヘッダー＋明細グループ（15品目＋ページ小計）
//       2ページ目以降：明細グループ（25品目＋ページ小計） … 最終ページまで
//       最終ページ：明細グループ（残り品目＋ページ小計）＋印刷フッター

(function () {
  "use strict";

  // -------------------- ユーティリティ --------------------

  function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // カンマ付き小数2桁
  function formatAmount(val) {
    var num = parseFloat(val);
    if (isNaN(num)) return "";
    var fixed = num.toFixed(2);
    var parts = fixed.split(".");
    var intPart = parts[0];
    var decPart = parts[1];
    var re = /(\d+)(\d{3})/;
    while (re.test(intPart)) {
      intPart = intPart.replace(re, "$1" + "," + "$2");
    }
    return intPart + "." + decPart;
  }

  // カンマ付き整数
  function formatInt(val) {
    var num = parseInt(Math.floor(val), 10);
    if (isNaN(num)) return "";
    var s = String(num);
    var re = /(\d+)(\d{3})/;
    while (re.test(s)) {
      s = s.replace(re, "$1" + "," + "$2");
    }
    return s;
  }

  // 金額文字列を数値に変換（¥, \, , を除去）
  function parseNumber(val) {
    if (val === null || val === undefined) return NaN;
    var s = String(val);
    s = s.replace(/[¥\\,]/g, "");
    s = s.replace(/^\s+|\s+$/g, "");
    if (!s) return NaN;
    var num = parseFloat(s);
    if (isNaN(num)) return NaN;
    return num;
  }

  // 改行を <br> に変換
  function toMultilineHtml(text) {
    if (!text) return "";
    var lines = String(text).split(/\r?\n/);
    var out = [];
    for (var i = 0; i < lines.length; i++) {
      out.push(escapeHtml(lines[i]));
    }
    return out.join("<br>");
  }

  // -------------------- 全データ文字列の解析 --------------------
  // ・行明細と金額合計だけを allText から取り出す
  // ・日付／宛先／業者名は画面の入力欄から取得する

  function parseAllDataText(allText) {
    var result = {
      dateText: "",
      toText: "",
      vendorText: "",
      rows: [],        // {no,name,spec,unit,qty,price,amount,note}
      baseAmount: 0,   // 税抜合計（行金額合計）
      taxAmount: 0,    // 消費税
      totalAmount: 0   // 総合計
    };

    // ---- 日付・宛先・業者名は DOM から取得 ----
    var billDateInput = document.getElementById("billDate");
    var rawDate = billDateInput ? (billDateInput.value || "") : "";
    if (rawDate && window.formatDateForCopy) {
      result.dateText = window.formatDateForCopy(rawDate); // yyyy/mm/dd
    } else {
      result.dateText = rawDate;
    }

    var toTA = document.getElementById("toText");
    result.toText = toTA ? (toTA.value || "") : "";

    var vendorTA = document.getElementById("vendorText");
    result.vendorText = vendorTA ? (vendorTA.value || "") : "";

    // ---- allText から明細行だけを抜き出す ----
    var text = (allText || "")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n");
    var lines = text.split("\n");
    var n = lines.length;

    // "No\t品名" から始まるヘッダー行を探す
    var idxHeader = -1;
    var i;
    for (i = 0; i < n; i++) {
      var line = lines[i];
      if (line.indexOf("No\t") === 0 && line.indexOf("品名") !== -1) {
        idxHeader = i;
        break;
      }
    }
    if (idxHeader === -1) {
      // 明細がなければここで終了（ヘッダー等だけ印刷）
      return result;
    }

    // ヘッダー行の次から [最終ページ集計] か空行までを明細とみなす
    for (i = idxHeader + 1; i < n; i++) {
      var l = lines[i];
      if (!l) break;
      if (l.indexOf("[最終ページ集計]") === 0) break;
      if (l.indexOf("金額の合計:") === 0) break;

      var cols = l.split("\t");
      if (!cols.length || !cols[0]) continue;

      var row = {
        no: cols[0] || "",
        name: cols[1] || "",
        spec: cols[2] || "",
        unit: cols[3] || "",
        qty: cols[4] || "",
        price: cols[5] || "",
        amount: cols[6] || "",
        note: cols[7] || ""
      };
      result.rows.push(row);
    }

    // ---- 金額合計を再計算して税・総合計を算出 ----
    var sum = 0;
    for (i = 0; i < result.rows.length; i++) {
      var amt = parseNumber(result.rows[i].amount);
      if (!isNaN(amt)) {
        sum += amt;
      }
    }
    var base = sum;
    var TAX_RATE = 0.08; // 8%
    var tax = Math.floor(base * TAX_RATE + 1e-6);
    var total = base + tax;

    result.baseAmount = base;
    result.taxAmount = tax;
    result.totalAmount = total;

    return result;
  }

  // -------------------- 請求書 HTML 構築 --------------------

  function buildInvoiceHtml(data) {
    var rows = data.rows || [];

    var FIRST_PAGE_ROWS = 15;
    var OTHER_PAGE_ROWS = 25;

    // ページ構成情報
    // page = { rows:[], showHeader:true/false, showSubtotal:true/false, showFooter:true/false }
    var pages = [];

    if (rows.length === 0) {
      // 明細なし：1ページだけヘッダー＋フッター
      pages.push({
        rows: [],
        showHeader: true,
        showSubtotal: false,
        showFooter: true
      });
    } else if (rows.length <= FIRST_PAGE_ROWS) {
      // 15品目以下
      pages.push({
        rows: rows.slice(0),
        showHeader: true,
        showSubtotal: false,
        showFooter: true
      });
    } else {
      // 15品目超
      var offset = 0;
      // 1ページ目：15件＋小計、ヘッダーあり、フッターなし
      pages.push({
        rows: rows.slice(offset, offset + FIRST_PAGE_ROWS),
        showHeader: true,
        showSubtotal: true,
        showFooter: false
      });
      offset += FIRST_PAGE_ROWS;

      // 2ページ目以降：25件ずつ＋小計、最終ページだけフッターあり
      var totalRows = rows.length;
      while (offset < totalRows) {
        var remaining = totalRows - offset;
        var count = remaining > OTHER_PAGE_ROWS ? OTHER_PAGE_ROWS : remaining;
        var isLast = (offset + count >= totalRows);
        pages.push({
          rows: rows.slice(offset, offset + count),
          showHeader: false,
          showSubtotal: true,
          showFooter: isLast
        });
        offset += count;
      }
    }

    var pageCount = pages.length;

    var baseStr = formatAmount(data.baseAmount || 0);
    var taxStr = formatInt(data.taxAmount || 0);
    var totalStr = formatInt(data.totalAmount || 0);

    var invoiceLine = "";
    if (totalStr) {
      // 「￥1,497,681ー」の形式（ラベル「請求額」はボックス内で表示）
      invoiceLine = "￥" + totalStr + "ー";
    }

    var versionText = window.VERSION_TEXT || "";

    var html = "";
    html += '<!doctype html><html lang="ja"><head>';
    html += '<meta charset="UTF-8">';
    html += "<title>請求書プレビュー</title>";
    html += "<style>";
    html += "@page { margin: 20mm 15mm 20mm 20mm; }"; // 左 20mm
    html += "body { margin: 0; padding: 0; font-family: system-ui, -apple-system, 'Segoe UI', sans-serif; font-size: 12px; }";
    html += ".page { page-break-after: always; }";
    html += ".page:last-child { page-break-after: auto; }";
    html += ".page-inner { width: 100%; box-sizing: border-box; }";

    // ページ番号用
    html += ".page-top { position: relative; margin-bottom: 4px; }";
    html += ".page-no { position: absolute; right: 0; top: 0; font-size: 11px; }";

    // 印刷ヘッダー
    html += ".invoice-header { margin-bottom: 8px; padding: 6px 8px; background: #24527a; color: #fff; }";
    html += ".invoice-title { font-size: 20px; font-weight: bold; text-align: center; letter-spacing: 4px; }";
    html += ".date-line { text-align: right; margin: 4px 0 4px; }";
    html += ".claim-line { margin: 2px 0 4px; font-size: 12px; }";

    html += ".box { border: 1px solid #000; padding: 6px 8px; margin-bottom: 4px; background: #fff; color: #000; }";
    html += ".box-label { font-size: 11px; margin-bottom: 2px; }";
    html += ".box-body { white-space: pre-line; }";
    html += ".box.atena .box-body { text-align: left; }";
    html += ".box.vendor .box-body { text-align: right; }";

    html += ".invoice-amount-wrapper { margin: 6px 0 6px; text-align: right; }";
    html += ".invoice-amount-box { display: inline-block; border: 1px solid #24527a; padding: 4px 10px; min-width: 70mm; }";
    html += ".invoice-amount-label { font-size: 12px; margin-bottom: 2px; text-align: left; }";
    html += ".invoice-amount-value { font-size: 16px; font-weight: bold; text-align: right; }";

    // 明細グループ
    html += ".detail-title { margin-top: 6px; margin-bottom: 2px; font-weight: bold; }";
    html += ".invoice-table { width: 100%; border-collapse: collapse; margin-top: 2px; }";
    html += ".invoice-table th, .invoice-table td { border: 1px solid #000; padding: 2px 4px; vertical-align: top; }";
    html += ".invoice-table th { background: #f0f0f0; text-align: center; }";
    html += ".col-no { width: 10mm; text-align: center; }";
    html += ".col-name { width: auto; }";
    html += ".col-unit { width: 10mm; text-align: center; }";
    html += ".col-qty { width: 18mm; text-align: right; }";
    html += ".col-price { width: 22mm; text-align: right; }";
    html += ".col-amount { width: 24mm; text-align: right; }";
    html += ".col-note { width: 14mm; }";

    html += ".item-name { font-weight: normal; }";
    html += ".item-spec { font-size: 11px; color: #555; }";

    // ページ小計
    html += ".sum-page { margin-top: 4px; text-align: right; }";
    html += ".sum-page .label { display: inline-block; min-width: 70px; }";

    // 印刷フッター
    html += ".footer-area { margin-top: 8px; text-align: right; }";
    html += ".footer-area div { margin-top: 2px; }";
    html += ".footer-area .label { display: inline-block; min-width: 70px; }";

    html += ".version { margin-top: 8px; font-size: 10px; text-align: right; color: #666; }";
    html += "</style>";
    html += "</head><body>";

    for (var p = 0; p < pageCount; p++) {
      var page = pages[p];
      var pageRows = page.rows || [];

      html += '<div class="page"><div class="page-inner">';

      // ページ番号
      html += '<div class="page-top">';
      html += '<div class="page-no">' + (p + 1) + "/" + pageCount + "</div>";
      html += "</div>";

      // ---- 印刷ヘッダー（1ページ目だけ） ----
      if (page.showHeader) {
        html += '<div class="invoice-header">';
        html += '<div class="invoice-title">請　求　書</div>';
        html += "</div>";

        if (data.dateText) {
          html +=
            '<div class="date-line">' +
            escapeHtml(data.dateText) +
            "</div>";
        }

        html += '<div class="claim-line">下記の通りご請求申し上げます。</div>';

        // 宛先枠
        html += '<div class="box atena">';
        html += '<div class="box-label">宛先</div>';
        html +=
          '<div class="box-body">' + toMultilineHtml(data.toText) + "</div>";
        html += "</div>";

        // 業者名枠
        html += '<div class="box vendor">';
        html += '<div class="box-label">業者名</div>';
        html +=
          '<div class="box-body">' +
          toMultilineHtml(data.vendorText) +
          "</div>";
        html += "</div>";

        // 請求額ボックス
        if (invoiceLine) {
          html += '<div class="invoice-amount-wrapper">';
          html += '<div class="invoice-amount-box">';
          html += '<div class="invoice-amount-label">請求額</div>';
          html +=
            '<div class="invoice-amount-value">' +
            escapeHtml(invoiceLine) +
            "</div>";
          html += "</div></div>";
        }
      }

      // ---- 明細グループ ----
      html += '<div class="detail-title">請求目明細</div>';

      if (pageRows.length) {
        html += '<table class="invoice-table">';
        html += "<thead><tr>";
        html += '<th class="col-no">No</th>';
        html += '<th class="col-name">品目・規格</th>';
        html += '<th class="col-unit">単位</th>';
        html += '<th class="col-qty">合計数量</th>';
        html += '<th class="col-price">契約単価</th>';
        html += '<th class="col-amount">金額</th>';
        html += '<th class="col-note">備考</th>';
        html += "</tr></thead>";
        html += "<tbody>";

        var pageBase = 0;

        for (var i = 0; i < pageRows.length; i++) {
          var r = pageRows[i];
          var amtVal = parseNumber(r.amount);
          if (!isNaN(amtVal)) {
            pageBase += amtVal;
          }

          html += "<tr>";
          html +=
            '<td class="col-no">' + escapeHtml(r.no || "") + "</td>";

          // 品目・規格（2行構成）
          html += '<td class="col-name">';
          html +=
            '<div class="item-name">' +
            escapeHtml(r.name || "") +
            "</div>";
          html +=
            '<div class="item-spec">' +
            escapeHtml(r.spec || "") +
            "</div>";
          html += "</td>";

          html +=
            '<td class="col-unit">' +
            escapeHtml(r.unit || "") +
            "</td>";
          html +=
            '<td class="col-qty">' +
            escapeHtml(r.qty || "") +
            "</td>";
          html +=
            '<td class="col-price">' +
            escapeHtml(r.price || "") +
            "</td>";
          html +=
            '<td class="col-amount">' +
            escapeHtml(r.amount || "") +
            "</td>";
          html +=
            '<td class="col-note">' +
            escapeHtml(r.note || "") +
            "</td>";
          html += "</tr>";
        }

        html += "</tbody></table>";

        // ページ小計（必要な場合のみ）
        if (page.showSubtotal) {
          var pageBaseStr = formatAmount(pageBase);
          html += '<div class="sum-page">';
          html +=
            '<span class="label">小計:</span> ' +
            escapeHtml(pageBaseStr) +
            "</div>";
        }
      } else {
        html += "<p>明細がありません。</p>";
      }

      // ---- 印刷フッター（最終ページのみ） ----
      if (page.showFooter) {
        html += '<div class="footer-area">';
        if (baseStr) {
          html +=
            '<div><span class="label">合計:</span> ' +
            escapeHtml(baseStr) +
            "</div>";
        }
        if (taxStr) {
          html +=
            '<div><span class="label">消費税:</span> ' +
            escapeHtml(taxStr) +
            "</div>";
        }
        if (totalStr) {
          html +=
            '<div><span class="label">総合計:</span> ' +
            escapeHtml(totalStr) +
            "</div>";
        }
        html += "</div>";

        if (versionText) {
          html +=
            '<div class="version">' + escapeHtml(versionText) + "</div>";
        }
      }

      html += "</div></div>"; // .page-inner, .page
    }

    html += "</body></html>";
    return html;
  }

  // -------------------- メインエントリ --------------------

  function printLedgerData(allText) {
    var parsed = parseAllDataText(allText || "");

    var win = window.open("", "_blank");
    if (!win) {
      alert("ポップアップがブロックされています。");
      return;
    }
    var doc = win.document;
    doc.open();
    doc.write(buildInvoiceHtml(parsed));
    doc.close();
    win.focus();

    if (win.print) {
      win.print();
    }
  }

  // グローバル公開
  window.printLedgerData = printLedgerData;
})();
