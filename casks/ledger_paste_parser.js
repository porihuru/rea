// ledger_paste_parser.js
// 納入台帳貼り付けテキスト解析用パーサ
// v2025.11.23-paste-13 対応版
//
// 出力形式：
// {
//   rows: [
//     {
//       vendor: "滝川駐屯地株式会社トワニ旭川店", // ページヘッダの業者名等(合体文字列)
//       no: "0001",
//       name: "1L豆乳飲料",
//       spec: "",
//       unit: "PC",
//       qty: "11.00",
//       price: "203.00",
//       amount: "2,233.00",
//       note: "<"   // 備考（* や < など）
//     },
//     ...
//   ],
//   summary: {
//     base: "969,582.00",    // 最終ページの「課税対象額」文字列
//     tax:  "77,566",        // 「消費税」文字列
//     total:"1,047,148",     // 「合計」文字列（末尾の「-」は除去）
//     calcBaseStr: "969,582.00", // rows の金額合計（小数2桁）
//     baseCheckMark: "<"     // base == calcBaseStr の場合 "<" それ以外は ""
//   }
// }
//
// ※品名行の先頭に付いている「全角数字（１,２,３…）」は「品名の一部」として扱う。
//   → 数値トークン判定では「半角数字」のみを対象とする。

// ------------------------ 共通ユーティリティ ------------------------

// 空白トリム
function trim(s) {
  if (s == null) return "";
  return String(s).replace(/^\s+|\s+$/g, "");
}

// 全角スペースも含めてトリム
function trimAll(s) {
  if (s == null) return "";
  return String(s)
    .replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
}

// 行が完全に空行かどうか
function isEmptyLine(line) {
  return trimAll(line) === "";
}

// 数字トークン判定（「数量・単価・金額」候補の文字列か）
// → 半角数字・カンマ・小数点・¥・\・マイナスのみ
function isNumberToken(token) {
  if (!token) return false;
  var s = trim(token);
  if (!s) return false;
  // 数字・カンマ・小数点・円記号・マイナスのみで構成される
  return /^[0-9,\.\\¥\-]+$/.test(s);
}

// 行中に含まれる「数字トークン」をすべて抽出
// ここでは「半角数字のみ」を対象とする。
// ※全角数字（１,２,３…）は「品名の一部」と見なすため除外。
function findNumberTokens(line) {
  var tokens = [];
  if (!line) return tokens;
  // ★修正済み：全角数字を除外して、半角の 0-9 のみを対象とする
  var re = /[0-9,\.\\¥\-]+/g;
  var m;
  while ((m = re.exec(line)) !== null) {
    var t = trim(m[0]);
    if (t && isNumberToken(t)) {
      tokens.push(t);
    }
  }
  return tokens;
}

// 「1,234.00」→ "1234.00" などに整形して数値化
function toNumberForCalc(str) {
  if (!str) return NaN;
  var s = String(str);
  s = s.replace(/[\\¥,]/g, "");
  s = trimAll(s);
  if (!s) return NaN;
  var n = parseFloat(s);
  if (isNaN(n)) return NaN;
  return n;
}

// カンマ付き・小数2桁表示
function formatAmount(num) {
  var n = Number(num);
  if (isNaN(n)) return "";
  var fixed = n.toFixed(2);
  var parts = fixed.split(".");
  var intPart = parts[0];
  var decPart = parts[1];
  var re = /(\d+)(\d{3})/;
  while (re.test(intPart)) {
    intPart = intPart.replace(re, "$1,$2");
  }
  return intPart + "." + decPart;
}

// カンマ付き整数
function formatInt(num) {
  var n = Math.floor(Number(num));
  if (isNaN(n)) return "";
  var s = String(n);
  var re = /(\d+)(\d{3})/;
  while (re.test(s)) {
    s = s.replace(re, "$1,$2");
  }
  return s;
}

// ------------------------ ヘッダ（業者名等）抽出 ------------------------

// 最初に出てくる「滝川駐屯地株式会社〜」行などをそのまま vendorFull として返す。
// （納地と業者名を分離しない。合体文字をそのまま保持。）
function extractVendorFull(lines) {
  for (var i = 0; i < lines.length; i++) {
    var line = trimAll(lines[i]);
    if (!line) continue;
    // 例: "滝川駐屯地株式会社トワニ旭川店"
    //     "滝川駐屯地株式会社セイコーフレッシュフーズ"
    if (line.indexOf("滝川駐屯地") >= 0 && line.indexOf("株式会社") >= 0) {
      return line;
    }
  }
  return "";
}

// ------------------------ 明細行抽出 ------------------------

// 「No を含む行」かどうか
function isNoLine(line) {
  if (!line) return false;
  // 「0001」「0100」など4桁連番の単独行、または行内に 0001 があるパターン
  if (/^\s*\d{4}\s*$/.test(line)) return true;
  if (/^\s*\d{4}\s+/.test(line)) return true;
  return false;
}

// 「単位」らしい短いトークンか
function isUnitToken(t) {
  if (!t) return false;
  t = trimAll(t);
  if (!t) return false;
  if (t.length > 3) return false;
  // EA, KG, BG, PC, CN, SH などを想定（英字2〜3文字）
  if (/^[A-Z]{1,3}$/.test(t)) return true;
  return false;
}

// 明細ブロック（1行分）を表す構造
function createEmptyRow(vendorFull) {
  return {
    vendor: vendorFull || "",
    no: "",
    name: "",
    spec: "",
    unit: "",
    qty: "",
    price: "",
    amount: "",
    note: ""
  };
}

// 「*」が含まれる行なら備考マークに反映
function extractNoteMarkFromLine(line) {
  if (!line) return "";
  return line.indexOf("*") >= 0 ? "*" : "";
}

// 本文から「No/品名/規格/単位/数量/単価/金額/備考」を抽出するメインロジック
function parseDetailsFromText(srcText, vendorFull) {
  var lines = String(srcText).split(/\r?\n/);
  var rows = [];

  var i, line;
  var len = lines.length;

  // ページをまたいで現れるので、単純に上から順に走査し、
  // 「No 行」を起点として 1明細ずつ抜き出す。
  var idx = 0;

  while (idx < len) {
    line = lines[idx];

    // No 行を探す
    if (!isNoLine(line)) {
      idx++;
      continue;
    }

    // ----- No 行確定 -----
    var row = createEmptyRow(vendorFull);

    // No 抽出
    var noMatch = line.match(/\d{4}/);
    if (!noMatch) {
      // 想定外だがスキップ
      idx++;
      continue;
    }
    row.no = noMatch[0];

    // 備考マーク（*）はどの行に出てもよいので、この明細ブロック中で検出する。
    var noteMark = extractNoteMarkFromLine(line);

    // 次行以降で「品名」「単位」「数量・単価・金額」のセットを取っていく
    idx++;

    // 1) 品名（複数行ある場合もある）
    var nameLines = [];
    while (idx < len) {
      var l = lines[idx];
      var t = trimAll(l);

      // 空行ならスキップして次へ
      if (!t) {
        idx++;
        continue;
      }

      // 単位行 or 数値だけの行 or 次の No 行が来たら品名ブロック終わり
      //   - 単位候補（EA, KG ...）だけの行
      //   - 完全に数値だけの行（数量ブロック）
      //   - No 行
      var numericTokens = findNumberTokens(t);
      if (isNoLine(t)) break;
      if (isUnitToken(t)) break;
      if (numericTokens.length > 0 && isNumberToken(numericTokens[0]) && !/[*０-９一二三四五六七八九十]/.test(t)) {
        // 「完全に数値のみ」と見なせる行（数量など） → 品名ブロック終端
        break;
      }

      // ★ここで「全角数字で始まる品名」に対応★
      // 先頭が全角数字（１〜９など）の場合も「品名行」として扱う。
      // → findNumberTokens では全角数字を見ていないので numericTokens.length は 0。
      nameLines.push(t);
      idx++;
    }

    row.name = nameLines.join("");

    // 2) 単位行
    var unit = "";
    while (idx < len) {
      var l2 = lines[idx];
      var t2 = trimAll(l2);
      if (!t2) {
        idx++;
        continue;
      }
      if (isUnitToken(t2)) {
        unit = t2;
        idx++;
      }
      break;
    }
    row.unit = unit;

    // 3) 数量・単価・金額ブロック
    //    複数行にまたがっているので、No 〜 次の No の直前までをスキャンして
    //    「数量・単価・金額らしき数値」を3つ拾う。
    var qty = "";
    var price = "";
    var amount = "";

    // 数値候補を集めるバッファ
    var numberBuf = [];

    while (idx < len) {
      var l3 = lines[idx];
      var t3 = trimAll(l3);

      // 次の No 行が来たらブロック終わり
      if (isNoLine(t3)) {
        break;
      }

      // ヘッダー行（（ 分）, 金 額, 合計数量契約単価, 納 地, 等）は無視
      if (
        t3.indexOf("金 額") >= 0 ||
        t3.indexOf("合計数量契約単価") >= 0 ||
        t3.indexOf("納　地：") >= 0 ||
        t3.indexOf("納 入 台 帳") >= 0 ||
        t3.indexOf("納入台帳") >= 0 ||
        t3.indexOf("品　名") >= 0 ||
        t3.indexOf("単位") >= 0 ||
        t3.indexOf("令和") >= 0 ||
        t3.indexOf("課税対象額") >= 0 ||
        t3.indexOf("消費税") >= 0 ||
        t3.indexOf("合　計") >= 0 ||
        t3.indexOf("-　以　下　余　白　-") >= 0 ||
        t3.indexOf("（ 分）") >= 0 ||
        t3.indexOf("（　　　　  　  分）") >= 0 ||
        t3.indexOf("（ 　　　　 　 分）") >= 0
      ) {
        idx++;
        continue;
      }

      // 備考マーク（*）があれば記録だけして数字処理からは除外
      if (t3.indexOf("*") >= 0) {
        noteMark = "*";
        idx++;
        continue;
      }

      // 単純に「数字トークン」だけを拾う（半角数字のみ）
      var nums = findNumberTokens(t3);
      if (nums.length > 0) {
        for (var k = 0; k < nums.length; k++) {
          numberBuf.push(nums[k]);
        }
      }

      // 数字以外のテキストはスキップ
      idx++;
    }

    // numberBuf から「数量・単価・金額」らしき3つを決める
    // 基本方針：
    //   1) 最大値を「金額」とみなす
    //   2) 残りのうち「金額 ÷ 数量 = 単価」に近いペアを探す
    if (numberBuf.length > 0) {
      // 文字列 → 数値へ
      var numericVals = [];
      for (var p = 0; p < numberBuf.length; p++) {
        var nv = toNumberForCalc(numberBuf[p]);
        if (!isNaN(nv)) {
          numericVals.push({ raw: numberBuf[p], val: nv });
        }
      }

      if (numericVals.length > 0) {
        // 金額候補：最大値
        var maxIdx = 0;
        var maxVal = numericVals[0].val;
        for (var q = 1; q < numericVals.length; q++) {
          if (numericVals[q].val > maxVal) {
            maxVal = numericVals[q].val;
            maxIdx = q;
          }
        }
        var amountObj = numericVals[maxIdx];
        amount = amountObj.raw;

        // 残りから数量・単価を推定
        var rest = [];
        for (var r = 0; r < numericVals.length; r++) {
          if (r !== maxIdx) rest.push(numericVals[r]);
        }

        var bestQty = null;
        var bestPrice = null;
        var bestDiff = null;

        for (var a1 = 0; a1 < rest.length; a1++) {
          for (var a2 = 0; a2 < rest.length; a2++) {
            if (a1 === a2) continue;
            var candQty = rest[a1].val;
            var candPrice = rest[a2].val;
            if (candQty === 0) continue;
            var expected = candQty * candPrice;
            var diff = Math.abs(expected - maxVal);
            if (bestDiff === null || diff < bestDiff) {
              bestDiff = diff;
              bestQty = rest[a1];
              bestPrice = rest[a2];
            }
          }
        }

        if (bestQty && bestPrice && bestDiff !== null && bestDiff < 1.0) {
          qty = bestQty.raw;
          price = bestPrice.raw;
        }
      }
    }

    row.qty = qty;
    row.price = price;
    row.amount = amount;
    row.note = noteMark;

    rows.push(row);

    // 次は「次の No 行」から再開
    // （上の while から抜けてきた時点で idx は次の No 行か末尾位置）
  }

  return rows;
}

// ------------------------ 最終ページ集計 ------------------------

// 「課税対象額 / 消費税 / 合計」を抽出
function extractSummary(srcText) {
  var lines = String(srcText).split(/\r?\n/);
  var base = "";
  var tax = "";
  var total = "";

  for (var i = 0; i < lines.length; i++) {
    var line = trimAll(lines[i]);

    // 課税対象額
    if (line.indexOf("課税対象額") >= 0) {
      // 行内に数字があればそれを利用
      var nums = findNumberTokens(line);
      if (nums.length > 0) {
        base = nums[0];
      } else if (i + 1 < lines.length) {
        // 次の行に数値だけ単独であるパターン
        var next = trimAll(lines[i + 1]);
        var nums2 = findNumberTokens(next);
        if (nums2.length > 0) {
          base = nums2[0];
        }
      }
      continue;
    }

    // 消費税
    if (line.indexOf("消費税") >= 0) {
      var nums3 = findNumberTokens(line);
      if (nums3.length > 0) {
        tax = nums3[0];
      } else if (i + 1 < lines.length) {
        var next2 = trimAll(lines[i + 1]);
        var nums4 = findNumberTokens(next2);
        if (nums4.length > 0) {
          tax = nums4[0];
        }
      }
      continue;
    }

    // 合計
    if (line.indexOf("合　計") >= 0 || line.indexOf("合計") >= 0) {
      var nums5 = findNumberTokens(line);
      if (nums5.length > 0) {
        total = nums5[0];
      } else if (i + 1 < lines.length) {
        var next3 = trimAll(lines[i + 1]);
        var nums6 = findNumberTokens(next3);
        if (nums6.length > 0) {
          total = nums6[0];
        }
      }
      continue;
    }
  }

  // 「\1,047,148-」のような表記を想定 -> 記号を取り除く
  base = base.replace(/[\\¥]/g, "").replace(/-/g, "");
  tax  = tax.replace(/[\\¥]/g, "").replace(/-/g, "");
  total = total.replace(/[\\¥]/g, "").replace(/-/g, "");

  return {
    base: base || "",
    tax: tax || "",
    total: total || "",
    calcBaseStr: "",   // 後で rows から計算してセットする
    baseCheckMark: ""  // 同上
  };
}

// ------------------------ メインエクスポート ------------------------

// srcText: 「貼り付けテキスト」全体
function parseLedgerText(srcText) {
  if (srcText == null) {
    return { rows: [], summary: { base: "", tax: "", total: "", calcBaseStr: "", baseCheckMark: "" } };
  }

  var text = String(srcText);

  // ベンダ（納地＋業者名の合体文字列）
  var lines = text.split(/\r?\n/);
  var vendorFull = extractVendorFull(lines);

  // 明細行抽出
  var rows = parseDetailsFromText(text, vendorFull);

  // 最終ページ集計
  var summary = extractSummary(text);

  // rows から金額合計を再計算
  var sum = 0;
  for (var i = 0; i < rows.length; i++) {
    var a = toNumberForCalc(rows[i].amount);
    if (!isNaN(a)) {
      sum += a;
    }
  }
  var calcBaseStr = formatAmount(sum);
  summary.calcBaseStr = calcBaseStr;

  // 原本の「課税対象額」と一致するか
  var baseNum = toNumberForCalc(summary.base);
  if (!isNaN(baseNum) && Math.abs(baseNum - sum) < 0.5) {
    summary.baseCheckMark = "<";
  } else {
    summary.baseCheckMark = "";
  }

  return {
    rows: rows,
    summary: summary
  };
}

// グローバルに公開（ブラウザ用）
if (typeof window !== "undefined") {
  window.parseLedgerText = parseLedgerText;
}
