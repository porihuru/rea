// print.js
// 納入台帳テキスト解析ツール用 印刷プレビュー
// v2025.11.24-02
//
// ・index.html 側の「全データコピー用文字列」を引数として受け取るが、
//   日付・宛先・業者名は必ず画面のテキストボックスから取得する。
// ・品名と規格は印刷時に「品名・規格」列にまとめ、
//   上段＝品名、下段＝規格 の 2 行構成で表示する。
// ・印刷ヘッダー（請求書タイトル／日付／宛先／業者名＋請求額）は
//   1ページ目のみに表示。
// ・15品目以下：1ページのみ、ページ小計なし、フッター（合計・消費税額８％・総合計）は
//   最終ページの明細表の中に 3 行として挿入。
// ・15品目超：
//   1ページ目 …… 15 行＋ページ小計行
//   2ページ目以降 …… 25 行＋ページ小計行
//   最終ページ …… ページ小計行 ＋ フッター3行（すべて表の中）
//   （ページ小計は従来どおり残す）

(function () {
  "use strict";

  // -------------------- ユーティリティ --------------------

  function trim(str) {
    return (str || "").replace(/^\s+|\s+$/g, "");
  }

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

  // yyyy-mm-dd → yyyy/mm/dd
  function formatDateYMD(val) {
    if (!val) return "";
    var p = String(val).split("-");
    if (p.length === 3) {
      return p[0] + "/" + p[1] + "/" + p[2];
    }
    return val;
  }

  // 本日の日付（yyyy/mm/dd）
  function getTodayYMD() {
    var d = new Date();
    var y = d.getFullYear();
    var m = d.getMonth() + 1;
    var day = d.getDate();
    var mm = (m < 10 ? "0" : "") + m;
    var dd = (day < 10 ? "0" : "") + day;
    return y + "/" + mm + "/" + dd;
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
    if (rawDate) {
      result.dateText = formatDateYMD(rawDate);
    } else {
      // 空欄なら本日の日付を使用（画面の値は書き換えない）
      result.dateText = getTodayYMD();
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
    var totalRows = rows.length;

    // ページ分割
    var pages = [];
    if (totalRows <= 15) {
      // 15品目以下：1ページのみ・ページ小計なし
      pages.push(rows.slice(0));
    } else {
      // 1ページ目は 15行
      pages.push(rows.slice(0, 15));
      var remain = rows.slice(15);
      while (remain.length > 0) {
        pages.push(remain.slice(0, 25));
        remain = remain.slice(25);
      }
    }
    var pageCount = pages.length || 1;

    // 合計・税・総合計の文字列（.00付き）
    var baseStr = formatAmount(data.baseAmount || 0);
    var taxStr = formatAmount(data.taxAmount || 0);
    var totalStr = formatAmount(data.totalAmount || 0);

    // 請求額（ヘッダー左ボックスに表示）… 「￥1,047,148-」形式
    var invoiceLine = "";
    if (data.totalAmount) {
      var totalIntStr = formatInt(data.totalAmount || 0);
      if (totalIntStr) {
        invoiceLine = "￥" + totalIntStr + "ー";
      }
    }

    var versionText = window.VERSION_TEXT || "";

    var html = "";
    html += '<!doctype html><html lang="ja"><head>';
    html += '<meta charset="UTF-8">';
    html += "<title>請求書プレビュー</title>";
    html += "<style>";
    // 左余白 3cm = 30mm
    html += "@page { margin: 10mm 15mm 10mm 30mm; }";
    html += "body { margin: 0; padding: 0; font-family: system-ui, -apple-system, 'Segoe UI', sans-serif; font-size: 12px; }";
    html += ".page { page-break-after: always; }";
    html += ".page:last-child { page-break-after: auto; }";
    html += ".page-inner { width: 100%; box-sizing: border-box; }";

    html += ".invoice-header { position: relative; margin-bottom: 8px; }";
    html += ".invoice-title { font-size: 20px; font-weight: bold; text-align: center; }";
    html += ".page-no { position: absolute; right: 0; top: 0; font-size: 11px; }";

    html += ".date-line { text-align: right; margin: 4px 0 2px; }";
    html += ".message-line { margin: 2px 0 6px; }";

    html += ".box { border: 1px solid #000; padding: 6px 8px; box-sizing: border-box; }";
    html += ".box-label { font-size: 11px; margin-bottom: 2px; }";
    html += ".box-body { white-space: pre-line; }";
    html += ".box.atena .box-body { text-align: left; }";

    html += ".flex-row { display: flex; gap: 8px; margin-bottom: 4px; }";
    html += ".amount-box { width: 35%; }";
    html += ".amount-box .box-body { text-align: right; font-size: 22px; font-weight: bold; }";
    html += ".vendor-box { flex: 1; }";
    html += ".vendor-box .box-body { text-align: right; }";

    html += ".section-title { margin-top: 8px; margin-bottom: 4px; font-weight: bold; }";

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

    html += ".sum-row-label { text-align: right; font-weight: bold; }";
    html += ".sum-row-amount { text-align: right; font-weight: bold; }";

    html += ".version { margin-top: 8px; font-size: 10px; text-align: right; color: #666; }";
    html += "</style>";
    html += "</head><body>";

    var p, i;

    for (p = 0; p < pageCount; p++) {
      var pageRows = pages[p];
      var isFirstPage = (p === 0);
      var isLastPage = (p === pageCount - 1);

      html += '<div class="page"><div class="page-inner">';

      // ヘッダー
      html += '<div class="invoice-header">';
      html += '<div class="invoice-title">請求書</div>';
      html +=
        '<div class="page-no">' +
        (p + 1) +
        "/" +
        pageCount +
        "</div>";
      html += "</div>";

      if (isFirstPage) {
        // 日付
        if (data.dateText) {
          html +=
            '<div class="date-line">' +
            escapeHtml(data.dateText) +
            "</div>";
        }

        // 「下記のとおりご請求申し上げます。（低減税率対象）」行
        html +=
          '<div class="message-line">下記のとおりご請求申し上げます。（低減税率対象）</div>';

        // 宛先枠
        html += '<div class="box atena" style="margin-bottom:4px;">';
        html += '<div class="box-label">宛先</div>';
        html +=
          '<div class="box-body">' + toMultilineHtml(data.toText) + "</div>";
        html += "</div>";

        // 請求額＋業者名（2分割）
        html += '<div class="flex-row">';
        // 左：請求額
        html += '<div class="box amount-box">';
        html += '<div class="box-label">請求額</div>';
        html +=
          '<div class="box-body">' +
          escapeHtml(invoiceLine || "") +
          "</div>";
        html += "</div>";
        // 右：業者名（タイトル文字は出さない）
        html += '<div class="box vendor-box">';
        html +=
          '<div class="box-body">' +
          toMultilineHtml(data.vendorText) +
          "</div>";
        html += "</div>";
        html += "</div>";
      }

      // 請求明細タイトル
      html += '<div class="section-title">請求明細</div>';

      // 明細表
      if (pageRows && pageRows.length) {
        html += '<table class="invoice-table">';
        html += "<thead><tr>";
        html += '<th class="col-no">No</th>';
        html += '<th class="col-name">品名・規格</th>';
        html += '<th class="col-unit">単位</th>';
        html += '<th class="col-qty">合計数量</th>';
        html += '<th class="col-price">契約単価</th>';
        html += '<th class="col-amount">金額</th>';
        html += '<th class="col-note">備考</th>';
        html += "</tr></thead>";
        html += "<tbody>";

        // 明細行
        var pageSum = 0;
        for (i = 0; i < pageRows.length; i++) {
          var r = pageRows[i];
          var amtNum = parseNumber(r.amount);
          if (!isNaN(amtNum)) {
            pageSum += amtNum;
          }

          html += "<tr>";
          html +=
            '<td class="col-no">' + escapeHtml(r.no || "") + "</td>";

          // 品名・規格（2行構成）
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

        // ページ小計行
        // 15品目以下（1ページのみ）の場合は「小計なし」指定なので追加しない
        var needPageSubtotal =
          (totalRows > 15); // 仕様：15品目超の場合のみページ小計あり
        if (needPageSubtotal) {
          html += "<tr>";
          // No〜契約単価 までを結合
          html +=
            '<td class="sum-row-label" colspan="5">小計</td>';
          html +=
            '<td class="sum-row-amount">' +
            escapeHtml(formatAmount(pageSum)) +
            "</td>";
          html += '<td class="col-note"></td>';
          html += "</tr>";
        }

        // 最終ページのみ：合計／消費税額８％／総合計 をフッターとして
        // 小計行と同じように表の中に 3 行追加（小計は印刷フッターから削除済み）
        if (isLastPage) {
          // 合計
          html += "<tr>";
          html +=
            '<td class="sum-row-label" colspan="5">合計</td>';
          html +=
            '<td class="sum-row-amount">' +
            escapeHtml(baseStr) +
            "</td>";
          html += '<td class="col-note"></td>';
          html += "</tr>";

          // 消費税額８％（.00付き）
          html += "<tr>";
          html +=
            '<td class="sum-row-label" colspan="5">消費税額８％</td>';
          html +=
            '<td class="sum-row-amount">' +
            escapeHtml(taxStr) +
            "</td>";
          html += '<td class="col-note"></td>';
          html += "</tr>";

          // 総合計（.00付き）
          html += "<tr>";
          html +=
            '<td class="sum-row-label" colspan="5">総合計</td>';
          html +=
            '<td class="sum-row-amount">' +
            escapeHtml(totalStr) +
            "</td>";
          html += '<td class="col-note"></td>';
          html += "</tr>";
        }

        html += "</tbody></table>";
      } else {
        html += "<p>明細がありません。</p>";
      }

      // 最終ページのみ：バージョン表示
      if (isLastPage && versionText) {
        html +=
          '<div class="version">' + escapeHtml(versionText) + "</div>";
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
