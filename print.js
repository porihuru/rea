// print.js
// v2025.11.23-02
// 「全データコピー」で生成したテキストから請求書レイアウトを作り、印刷プレビューを開く

(function () {
  // ---------------- ユーティリティ ----------------

  function trim(str) {
    if (str === null || str === undefined) return "";
    return String(str).replace(/^\s+|\s+$/g, "");
  }

  // カンマ・￥・\ を取り除いて数値化
  function parseNumberSimple(val) {
    if (val === null || val === undefined) return NaN;
    var s = String(val);
    s = s.replace(/[¥\\,]/g, "");
    s = trim(s);
    if (!s) return NaN;
    var num = parseFloat(s);
    if (isNaN(num)) return NaN;
    return num;
  }

  // 3桁カンマ＋小数2桁 (例: 1234 → "1,234.00")
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

  // 整数＋3桁カンマ (例: 1497681 → "1,497,681")
  function formatIntAmount(val) {
    var num = parseNumberSimple(val);
    if (isNaN(num)) return "";
    num = Math.round(num);
    var s = String(num);
    var re = /(\d+)(\d{3})/;
    while (re.test(s)) {
      s = s.replace(re, "$1" + "," + "$2");
    }
    return s;
  }

  // HTMLエスケープ
  function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // 「全データコピー」用テキストエリアを探す
  function findAllDataElement() {
    var ids = ["allDataText", "fullCopyText", "copyAllText", "copyAllArea"];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (el) return el;
    }
    return null;
  }

  // ---------------- 全データテキストの解析 ----------------
  // 「納入台帳テキスト解析 v...」〜「金額の合計」「[最終ページ集計]」 の形式から
  // 日付・宛先・業者名・明細行・集計を取り出す
  function parseAllDataText(allText) {
    var normalized = (allText || "")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n");
    var lines = normalized.split("\n");
    var n = lines.length;
    var i;

    var result = {
      versionLine: "",
      date: "",
      toLines: [],
      vendorLines: [],
      rows: [],
      summary: {
        base: "",
        tax: "",
        total: "",
        calcBase: 0,
        calcBaseStr: "",
        rawAmountSumLine: ""
      }
    };

    // バージョン行（最初の非空行）
    for (i = 0; i < n; i++) {
      var t0 = trim(lines[i]);
      if (t0) {
        result.versionLine = t0;
        break;
      }
    }

    var idxDate = -1;
    var idxHeader = -1;

    // 日付行 / 表ヘッダ行 を探す
    for (i = 0; i < n; i++) {
      var line = lines[i];
      var t = trim(line);
      if (!t) continue;

      if (idxDate < 0 && line.indexOf("日付:") === 0) {
        idxDate = i;
      }
      if (
        idxHeader < 0 &&
        t.indexOf("No") === 0 &&
        t.indexOf("品名") >= 0 &&
        t.indexOf("合計数量") >= 0
      ) {
        idxHeader = i;
      }
    }

    if (idxDate >= 0) {
      result.date = trim(lines[idxDate].substring("日付:".length));
    }

    // 宛先と業者名ブロックの検出
    var idxTo = -1;
    var idxVendorBlock = -1;

    // 宛先:
    var startSearch = idxDate >= 0 ? idxDate + 1 : 0;
    for (i = startSearch; i < n; i++) {
      var t1 = trim(lines[i]);
      if (t1.indexOf("宛先:") === 0) {
        idxTo = i;
        break;
      }
    }

    // 宛先行たち
    if (idxTo >= 0) {
      for (i = idxTo + 1; i < n; i++) {
        var rawTo = lines[i];
        var t2 = trim(rawTo);
        if (!t2) {
          if (result.toLines.length) break; // 宛先開始後の空行で終了
          else continue; // 宛先がまだ始まっていない空行はスキップ
        }
        if (t2.indexOf("業者名:") === 0) {
          idxVendorBlock = i;
          break;
        }
        if (t2.indexOf("納地:") === 0 || t2.indexOf("納　地") === 0) {
          break;
        }
        result.toLines.push(t2);
      }
    }

    // 業者名ブロックの開始行（上のループで見つかっていなければ探す）
    if (idxVendorBlock < 0) {
      var startV = idxTo >= 0 ? idxTo + 1 : 0;
      for (i = startV; i < n; i++) {
        var t3 = trim(lines[i]);
        if (t3.indexOf("業者名:") === 0) {
          idxVendorBlock = i;
          break;
        }
        if (t3.indexOf("納地:") === 0 || t3.indexOf("納　地") === 0) {
          break;
        }
      }
    }

    // 業者名ブロックの行
    if (idxVendorBlock >= 0) {
      for (i = idxVendorBlock + 1; i < n; i++) {
        var rawV = lines[i];
        var t4 = trim(rawV);
        if (!t4) {
          if (result.vendorLines.length) break;
          else continue;
        }
        if (t4.indexOf("納地:") === 0 || t4.indexOf("納　地") === 0) {
          break;
        }
        result.vendorLines.push(t4);
      }
    }

    // 明細行（No〜備考）を抽出
    if (idxHeader >= 0) {
      for (i = idxHeader + 1; i < n; i++) {
        var lineRow = lines[i];
        var tr = trim(lineRow);
        if (!tr) continue;

        // サマリ開始
        if (tr.indexOf("金額の合計:") === 0) {
          result.summary.rawAmountSumLine = tr;
          break;
        }
        if (tr.indexOf("[最終ページ集計]") === 0) {
          break;
        }

        // 明細行は "0001" などで始まる
        if (!/^0\d{3}/.test(tr)) {
          continue;
        }

        var cols = lineRow.split("\t");
        var row = {
          no: cols.length > 0 ? trim(cols[0]) : "",
          name: cols.length > 1 ? trim(cols[1]) : "",
          spec: cols.length > 2 ? trim(cols[2]) : "",
          unit: cols.length > 3 ? trim(cols[3]) : "",
          qty: cols.length > 4 ? trim(cols[4]) : "",
          price: cols.length > 5 ? trim(cols[5]) : "",
          amount: cols.length > 6 ? trim(cols[6]) : "",
          note: cols.length > 7 ? trim(cols[7]) : ""
        };
        result.rows.push(row);
      }
    }

    // 集計（課税対象額 / 消費税 / 合計）
    for (i = 0; i < n; i++) {
      var ts = trim(lines[i]);
      if (!ts) continue;
      if (ts.indexOf("課税対象額:") === 0) {
        result.summary.base = trim(ts.substring("課税対象額:".length));
      } else if (ts.indexOf("消費税:") === 0) {
        result.summary.tax = trim(ts.substring("消費税:".length));
      } else if (ts.indexOf("合計:") === 0) {
        result.summary.total = trim(ts.substring("合計:".length));
      }
    }

    // 明細行から金額合計も計算しておく（チェック用）
    var sumBase = 0;
    for (i = 0; i < result.rows.length; i++) {
      var v = parseNumberSimple(result.rows[i].amount);
      if (!isNaN(v)) sumBase += v;
    }
    result.summary.calcBase = sumBase;
    result.summary.calcBaseStr = result.rows.length ? formatAmount(sumBase) : "";

    return result;
  }

  // ---------------- 請求書HTML生成 ----------------

  function buildInvoiceHtml(data) {
    var rows = data.rows || [];
    var title = "請求書";

    // ページ分割設定（必要に応じて調整）
    var rowsPerPage = 25;
    var pageCount = rows.length ? Math.ceil(rows.length / rowsPerPage) : 1;

    var baseText =
      (data.summary && data.summary.base) || data.summary.calcBaseStr || "";
    var taxText = (data.summary && data.summary.tax) || "";
    var totalText = (data.summary && data.summary.total) || "";
    var totalForBill = totalText; // 請求額表示用（￥●●●ー）

    var html = "";
    html += "<!doctype html><html><head><meta charset=\"UTF-8\">";
    html += "<title>" + escapeHtml(title) + "</title>";
    html += "<style>";
    // 左余白2cm
    html += "@page{ margin:20mm 15mm 20mm 20mm; }";
    html += "body{ font-family:'Yu Gothic','Meiryo','MS PGothic',sans-serif; font-size:11pt; }";
    html += ".page{ page-break-after:always; min-height:250mm; }";
    html += ".page:last-child{ page-break-after:auto; }";
    html += ".header{ margin-bottom:8mm; }";
    html += ".title-row{ text-align:center; position:relative; margin-bottom:4mm; }";
    html += ".title-row .title{ font-size:18pt; font-weight:bold; }";
    html += ".title-row .page-no{ position:absolute; right:0; top:0; font-size:10pt; }";
    html += ".date-row{ text-align:right; margin-bottom:3mm; }";
    html += ".row-block{ margin-bottom:3mm; }";
    html += ".label{ display:inline-block; width:48px; font-size:9pt; vertical-align:top; }";
    html += ".box{ display:inline-block; border:1px solid #000; padding:3px 6px; min-height:16pt; font-size:9pt; }";
    html += ".addr-box{ width:70%; text-align:left; }";
    html += ".vendor-box{ width:70%; text-align:right; }";
    html += ".amount-row{ text-align:left; font-size:12pt; font-weight:bold; margin:4mm 0 4mm 0; }";
    html += ".amount-row span.currency{ margin-left:6mm; }";
    html += "table{ border-collapse:collapse; width:100%; font-size:9pt; }";
    html += "th,td{ border:1px solid #000; padding:2px 3px; }";
    html += "th{ text-align:center; }";
    html += "td.num{ text-align:right; white-space:nowrap; }";
    html += "td.text{ text-align:left; }";
    html += ".subtotal{ text-align:right; margin-top:3mm; font-size:10pt; }";
    html += ".summary-block{ margin-top:4mm; text-align:right; font-size:10pt; }";
    html += ".summary-block div{ margin-top:2px; }";
    html += ".version{ margin-top:4mm; text-align:right; font-size:8pt; }";
    html += "</style></head><body>";

    if (!rows.length) {
      // 明細なしの場合でもヘッダは出しておく
      html += '<div class="page">';
      html += '<div class="header">';
      html += '<div class="title-row">';
      html += '<div class="title">' + escapeHtml(title) + "</div>";
      html += '<div class="page-no">1/1</div>';
      html += "</div>";
      html +=
        '<div class="date-row">日付: ' +
        escapeHtml(data.date || "") +
        "</div>";

      html += '<div class="row-block">';
      html += '<span class="label">宛先</span>';
      html += '<span class="box addr-box">';
      if (data.toLines && data.toLines.length) {
        for (var aa = 0; aa < data.toLines.length; aa++) {
          if (aa > 0) html += "<br>";
          html += escapeHtml(data.toLines[aa]);
        }
      }
      html += "</span></div>";

      html += '<div class="row-block">';
      html += '<span class="label">業者名</span>';
      html += '<span class="box vendor-box">';
      if (data.vendorLines && data.vendorLines.length) {
        for (var bb = 0; bb < data.vendorLines.length; bb++) {
          if (bb > 0) html += "<br>";
          html += escapeHtml(data.vendorLines[bb]);
        }
      }
      html += "</span></div>";

      html += '<div class="amount-row">';
      html +=
        "請求額<span class=\"currency\">￥" +
        escapeHtml(formatIntAmount(totalForBill)) +
        "ー</span>";
      html += "</div>";

      html += "<p>明細がありません。</p>";

      if (data.versionLine) {
        html +=
          '<div class="version">' +
          escapeHtml(data.versionLine) +
          "</div>";
      }

      html += "</div></div></body></html>";
      return html;
    }

    for (var p = 0; p < pageCount; p++) {
      var idxStart = p * rowsPerPage;
      var idxEnd = idxStart + rowsPerPage;
      if (idxEnd > rows.length) idxEnd = rows.length;
      var pageRows = rows.slice(idxStart, idxEnd);

      // このページの小計
      var pageSum = 0;
      for (var j = 0; j < pageRows.length; j++) {
        var v = parseNumberSimple(pageRows[j].amount);
        if (!isNaN(v)) pageSum += v;
      }
      var pageSumStr = formatAmount(pageSum);

      html += '<div class="page">';

      // ヘッダ
      html += '<div class="header">';
      html += '<div class="title-row">';
      html += '<div class="title">' + escapeHtml(title) + "</div>";
      html +=
        '<div class="page-no">' + (p + 1) + "/" + pageCount + "</div>";
      html += "</div>";

      html +=
        '<div class="date-row">日付: ' +
        escapeHtml(data.date || "") +
        "</div>";

      // 宛先
      html += '<div class="row-block">';
      html += '<span class="label">宛先</span>';
      html += '<span class="box addr-box">';
      if (data.toLines && data.toLines.length) {
        for (var k = 0; k < data.toLines.length; k++) {
          if (k > 0) html += "<br>";
          html += escapeHtml(data.toLines[k]);
        }
      }
      html += "</span></div>";

      // 業者名
      html += '<div class="row-block">';
      html += '<span class="label">業者名</span>';
      html += '<span class="box vendor-box">';
      if (data.vendorLines && data.vendorLines.length) {
        for (var k2 = 0; k2 < data.vendorLines.length; k2++) {
          if (k2 > 0) html += "<br>";
          html += escapeHtml(data.vendorLines[k2]);
        }
      }
      html += "</span></div>";

      // 請求額
      html += '<div class="amount-row">';
      html +=
        "請求額<span class=\"currency\">￥" +
        escapeHtml(formatIntAmount(totalForBill)) +
        "ー</span>";
      html += "</div>";

      html += "</div>"; // .header

      // 明細表
      html += "<table><thead><tr>";
      html += "<th>No</th>";
      html += "<th>品名</th>";
      html += "<th>規格</th>";
      html += "<th>単位</th>";
      html += "<th>合計数量</th>";
      html += "<th>契約単価</th>";
      html += "<th>金額</th>";
      html += "<th>備考</th>";
      html += "</tr></thead><tbody>";

      for (var r = 0; r < pageRows.length; r++) {
        var row = pageRows[r];
        html += "<tr>";
        html += '<td class="num">' + escapeHtml(row.no) + "</td>";
        html += '<td class="text">' + escapeHtml(row.name) + "</td>";
        html += '<td class="text">' + escapeHtml(row.spec) + "</td>";
        html += '<td class="text">' + escapeHtml(row.unit) + "</td>";
        html += '<td class="num">' + escapeHtml(row.qty) + "</td>";
        html += '<td class="num">' + escapeHtml(row.price) + "</td>";
        html += '<td class="num">' + escapeHtml(row.amount) + "</td>";
        html += '<td class="text">' + escapeHtml(row.note) + "</td>";
        html += "</tr>";
      }

      html += "</tbody></table>";

      // 小計
      html +=
        '<div class="subtotal">小計: ' +
        escapeHtml(pageSumStr) +
        "</div>";

      // 最終ページのみ 合計／消費税／総合計 ＋ バージョン
      if (p === pageCount - 1) {
        html += '<div class="summary-block">';
        html +=
          "<div>合計: " + escapeHtml(baseText || "") + "</div>";
        html +=
          "<div>消費税: " + escapeHtml(taxText || "") + "</div>";
        html +=
          "<div>総合計: " +
          escapeHtml(formatIntAmount(totalText || totalForBill)) +
          "</div>";
        html += "</div>";

        if (data.versionLine) {
          html +=
            '<div class="version">' +
            escapeHtml(data.versionLine) +
            "</div>";
        }
      }

      html += "</div>"; // .page
    }

    html += "</body></html>";
    return html;
  }

  // ---------------- 公開関数：印刷プレビュー ----------------

  function openPrintPreview() {
    var el = findAllDataElement();
    if (!el) {
      alert(
        "全データコピー用のテキストが見つかりません。\n" +
          "print.js 内の findAllDataElement() で id を確認してください。"
      );
      return;
    }
    var txt = el.value || el.textContent || "";
    if (!txt) {
      alert("全データコピー用テキストが空です。\n解析→全データコピーを一度実行してから印刷してください。");
      return;
    }

    var parsed = parseAllDataText(txt);

    var win = window.open("", "_blank");
    if (!win) {
      alert("ポップアップがブロックされました。");
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
  window.openPrintPreview = openPrintPreview;
})();
