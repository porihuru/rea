// 2025-12-10 14:00 JST

// ledger_paste_parser.js
// 貼り付けテキスト → 明細 rows ＋ 最終ページ集計 summary への変換専用
// v2025.12.10-01
//
// ●責務
//   - 納入台帳の「貼り付けテキスト」から
//     ・明細行（No / 品名 / 規格(空) / 単位 / 合計数量 / 契約単価 / 金額 / 備考(空)）
//     ・最終ページ集計（課税対象額 / 消費税 / 合計）
//     を抽出して返すだけに特化
//
// ●数量・単価・金額 抽出の最優先ルール（今回修正ポイント）
//   - 各品目ブロック内で、コード行(0001等)以外の行を見て：
//     ① 「数値が2つ以上ある行」を全部集める
//     ② その中で「一番下の行」を「最後の組」とみなす
//     ③ その行の「最後の2つの数値」を [合計数量, 契約単価] とする
//     ④ その行より下で最初に出てくる数値行の「最後の数値」を金額とする
//   - 例）0008 焼き岩のり
//        0.10
//        0.10 19,500.00  ← 最後の組 → [0.10, 19,500.00] = [数量, 単価]
//        1,950.00        ← 次の数値行 → 金額
//
//        0015 ギョーザ
//        340.00
//        340.00 27.00    ← 最後の組 → [340.00, 27.00]
//        9,180.00        ← 金額
//
//        0074 回鍋肉の素
//        1.00
//        1.00 1,200.00   ← 最後の組 → [1.00, 1,200.00]
//        1,200.00        ← 金額
//
//   - 上記ルールで数量・単価・金額がすべて取れるようにしたので、
//     0074 のような最後の品目も正しく数値が入ります。
//
// ●その他
//   - 品名と単位：
//       コード行から始まり、大文字2文字の単位(PC/BG/KG/EA/CA/CN/SH)までを品名として結合。
//       単位が行内に無い場合は、次行以降から単位を探し、出てくるまで品名を連結。
//   - 最終ページ集計：
//       「-　以　下　余　白　-」行より下にある 3つの数値を
//         1行目 → 課税対象額
//         2行目 → 合計（\xxx- を含む）
//         3行目 → 消費税
//       として取得し、
//         summary.base / summary.tax / summary.total
//       に格納する。
//   - rows の金額合計を summary.calcBase / summary.calcBaseStr に入れ、
//     課税対象額と一致すれば summary.baseCheckMark = "<" を付与する。

// -------------------- ユーティリティ --------------------

// 行から数値トークンを抽出
function findNumberTokens(line) {
  var tokens = [];
  var re = /[0-9０-９,\.\\¥-]+/g;
  var m;
  while ((m = re.exec(line)) !== null) {
    tokens.push({ text: m[0], index: m.index });
  }
  return tokens;
}

// 簡易数値パース（カンマ・¥・\ 除去）
function parseNumberSimple(val) {
  if (val === null || val === undefined) return NaN;
  var s = String(val);
  s = s.replace(/[¥\\,]/g, '');
  s = s.replace(/^\s+|\s+$/g, '');
  if (!s) return NaN;
  var num = parseFloat(s);
  if (isNaN(num)) return NaN;
  return num;
}

// 3桁カンマ＋小数2桁
function formatAmount(val) {
  var num = parseFloat(val);
  if (isNaN(num)) return '';
  var fixed = num.toFixed(2);
  var parts = fixed.split('.');
  var intPart = parts[0];
  var decPart = parts[1];
  var re = /(\d+)(\d{3})/;
  while (re.test(intPart)) {
    intPart = intPart.replace(re, '$1' + ',' + '$2');
  }
  return intPart + '.' + decPart;
}

// 合計値の整形（¥, \, 末尾-を除去）
function normalizeTotal(val) {
  if (!val) return '';
  var v = String(val);
  v = v.replace(/[¥\\]/g, '');
  v = v.replace(/-+$/g, '');
  v = v.replace(/^\s+|\s+$/g, '');
  return v;
}

// 品名末尾に単位がくっついている場合の分離
// 例: "冷凍マンゴーチャンクBG" → name="冷凍マンゴーチャンク", unit="BG"
function splitNameAndUnit(fullName) {
  var trimmed = (fullName || '').replace(/\s+$/g, '');
  if (!trimmed) {
    return { name: '', unit: '' };
  }
  var units = ['PC', 'BG', 'KG', 'EA', 'CA', 'CN', 'SH'];
  var re = new RegExp('^(.*)(' + units.join('|') + ')$');
  var m = trimmed.match(re);
  if (m) {
    return {
      name: m[1],
      unit: m[2]
    };
  }
  return { name: trimmed, unit: '' };
}

// -------------------- 明細パーサ --------------------

// 明細パーサ：貼り付けテキスト → 明細配列
function parseDetailsFromText(text) {
  var normalized = (text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  var lines = normalized.split('\n');

  var rows = [];
  var n = lines.length;
  var i, k, p;

  var currentVendor = '';

  // コード検出用（0001, 0002, ...）
  var codeRe = /0\d{3}(?!\d)/g;
  // 単位検出用
  var unitRe = /(PC|BG|KG|EA|CA|CN|SH)(?=\s|$)/;

  for (i = 0; i < n; i++) {
    var line = lines[i];
    var trimmed = line ? line.replace(/^\s+|\s+$/g, '') : '';

    // 令和○年○月 → 次行が「納地＋業者名」の合体行
    if (trimmed.indexOf('令和') !== -1 &&
        trimmed.indexOf('年') !== -1 &&
        trimmed.indexOf('月') !== -1) {
      var vIdx;
      for (vIdx = i + 1; vIdx < n; vIdx++) {
        var vn = lines[vIdx];
        if (!vn) continue;
        var vnTrim = vn.replace(/^\s+|\s+$/g, '');
        if (vnTrim) {
          currentVendor = vnTrim;
          break;
        }
      }
      continue;
    }

    if (!trimmed) continue;

    // 品目開始行（0001, 0002, ... を含む行）
    codeRe.lastIndex = 0;
    var mCode = codeRe.exec(line);
    if (!mCode) continue;

    var code = mCode[0];

    // -------------------- この品目ブロックの終了位置を探す --------------------
    //   - 次の "000x" 行
    //   - ページヘッダー行（「納　地：」「納入台帳」など）
    //   - 「課税対象額」行
    //   - 最終ページの「-　以　下　余　白　-」行（この行まではブロックに含める）
    var blockEnd = n;
    for (k = i + 1; k < n; k++) {
      var l2 = lines[k];
      var t2 = l2 ? l2.replace(/^\s+|\s+$/g, '') : '';
      if (!t2) continue;

      // 次のコード行 → 手前までがこの品目
      codeRe.lastIndex = 0;
      if (codeRe.exec(l2)) {
        blockEnd = k;
        break;
      }

      if (t2.indexOf('納　地') !== -1 && t2.indexOf('業 者 名') !== -1) {
        blockEnd = k;
        break;
      }
      if (t2.indexOf('納入台帳') !== -1) {
        blockEnd = k;
        break;
      }
      if (t2.indexOf('課税対象額') !== -1) {
        blockEnd = k;
        break;
      }
      // 最終ページの「-　以　下　余　白　-」行
      if (t2.indexOf('以') !== -1 && t2.indexOf('余') !== -1) {
        blockEnd = k + 1; // この行までは品目ブロックに含める
        break;
      }
    }

    // -------------------- 品名＆単位の抽出 --------------------
    var nameParts = [];
    var unitFromName = '';

    // コード行の「コード以降」を tail として処理
    var tail = line.substring(mCode.index + 4); // 4桁コードの直後から
    var tailTrim = tail.replace(/^\s+|\s+$/g, '');
    if (tailTrim) {
      var umHead = unitRe.exec(tailTrim);
      if (umHead) {
        // tail 内に単位がある → その手前までが品名
        var unitIndex = umHead.index;
        var nameCandidate = tailTrim.substring(0, unitIndex);
        nameCandidate = nameCandidate.replace(/\s+$/g, '');
        if (nameCandidate) {
          nameParts.push(nameCandidate);
        }
        unitFromName = umHead[1];
      } else {
        // 単位は無いが、コード行に品名の一部がある
        nameParts.push(tailTrim);
      }
    }

    // 2行目以降：品名・単位の続き
    for (p = i + 1; p < blockEnd; p++) {
      var l3 = lines[p];
      if (!l3) continue;
      var t3 = l3.replace(/^\s+|\s+$/g, '');
      if (!t3) continue;

      var um2 = unitRe.exec(t3);
      var tokens3 = findNumberTokens(l3);

      if (!tokens3.length && !um2) {
        // 数字も単位も無い → 完全に品名の続き
        nameParts.push(t3);
        continue;
      }

      if (um2) {
        // この行で単位が出てきた
        var unitIdx2 = um2.index;
        var prefixUnit = t3.substring(0, unitIdx2);
        prefixUnit = prefixUnit.replace(/\s+$/g, '');
        if (prefixUnit) {
          // 例: "（黄） PC 11.00" → "（黄）" を品名に追加
          nameParts.push(prefixUnit);
        }
        if (!unitFromName) {
          unitFromName = um2[1];
        }
        // 単位以降は数量などなので、ここで品名処理は終了
        break;
      }

      // 単位は無いが数字がある行
      // → 行頭の文字列だけを品名の続きにする
      var firstIdx3 = tokens3[0].index;
      if (firstIdx3 > 0) {
        var prefix3 = l3.substring(0, firstIdx3);
        prefix3 = prefix3.replace(/^\s+|\s+$/g, '');
        if (prefix3) {
          nameParts.push(prefix3);
        }
      }
      // 数字が現れた段階で、それ以降は品名ではないとみなして終了
      break;
    }

    var fullName = nameParts.join('');
    var name = '';
    var unit = '';

    if (unitFromName) {
      name = fullName;
      unit = unitFromName;
    } else {
      var nu = splitNameAndUnit(fullName);
      name = nu.name;
      unit = nu.unit;
    }

    // -------------------- 数量・単価・金額を「最後の組ルール」で抽出 --------------------

    var qtyText    = '';
    var priceText  = '';
    var amountText = '';

    // ブロック内の数値行を収集（コードは除外）
    var numericLines = [];     // { idx, tokens: [ {text,index}... ] }
    var pairLines    = [];     // numericLines のうち "数値2個以上" の行

    for (var li = i; li < blockEnd; li++) {
      var ln = lines[li];
      if (!ln) continue;
      var tln = ln.replace(/^\s+|\s+$/g, '');
      if (!tln) continue;

      // 明らかにヘッダー・集計系の行は除外
      if (tln.indexOf('納　地') !== -1 && tln.indexOf('業 者 名') !== -1) continue;
      if (tln.indexOf('納入台帳') !== -1) continue;
      if (tln.indexOf('課税対象額') !== -1) continue;
      if (tln.indexOf('消費税') !== -1 && tln.indexOf('合') !== -1) continue;
      if (tln.indexOf('金　　額') !== -1) continue;
      if (tln.indexOf('合計数量契約単価') !== -1) continue;

      var tokens = findNumberTokens(ln);
      if (!tokens.length) continue;

      // コード(000x)は除外
      var nonCode = [];
      for (var ti = 0; ti < tokens.length; ti++) {
        if (tokens[ti].text === code) continue;
        nonCode.push(tokens[ti]);
      }
      if (!nonCode.length) continue;

      numericLines.push({ idx: li, tokens: nonCode });
      if (nonCode.length >= 2) {
        pairLines.push({ idx: li, tokens: nonCode });
      }
    }

    if (pairLines.length) {
      // 一番下の「数値2個以上の行」が「最後の組」
      var lastPair = pairLines[pairLines.length - 1];
      var tks = lastPair.tokens;
      var len = tks.length;

      var qtyTok   = tks[len - 2];
      var priceTok = tks[len - 1];

      qtyText   = qtyTok.text;
      priceText = priceTok.text;

      // 「最後の組」の行より下で最初に出てくる数値行 → 金額
      var amountFound = false;
      for (var ni = 0; ni < numericLines.length; ni++) {
        var nl = numericLines[ni];
        if (nl.idx <= lastPair.idx) continue;
        // 最後の数値を金額として採用
        var amtTok = nl.tokens[nl.tokens.length - 1];
        amountText = amtTok.text;
        amountFound = true;
        break;
      }

      // 万が一、次の数値行が無い場合は 数量×単価 で金額を算出
      if (!amountFound) {
        var qv = parseNumberSimple(qtyText);
        var pv = parseNumberSimple(priceText);
        if (!isNaN(qv) && !isNaN(pv)) {
          amountText = formatAmount(qv * pv);
        }
      }
    } else if (numericLines.length) {
      // フォールバック：
      // 「数値2個以上の行」が無い＝すべて1個ずつの行だけの場合
      // → 上から 3つの数値を [数量, 単価, 金額] とみなす
      var flat = [];
      for (var fi = 0; fi < numericLines.length; fi++) {
        for (var fj = 0; fj < numericLines[fi].tokens.length; fj++) {
          flat.push(numericLines[fi].tokens[fj]);
        }
      }
      if (flat.length >= 3) {
        qtyText    = flat[0].text;
        priceText  = flat[1].text;
        amountText = flat[2].text;
      } else if (flat.length === 2) {
        qtyText   = flat[0].text;
        priceText = flat[1].text;
        // 金額は数量×単価を計算
        var qv2 = parseNumberSimple(qtyText);
        var pv2 = parseNumberSimple(priceText);
        if (!isNaN(qv2) && !isNaN(pv2)) {
          amountText = formatAmount(qv2 * pv2);
        }
      } else if (flat.length === 1) {
        // 数値1個だけ → 数量だけ分かるとみなす
        qtyText = flat[0].text;
      }
    }
    // numericLines が全く無い場合は qty/price/amount は空のまま

    var row = {
      vendor: currentVendor,
      no:     code,
      name:   name,
      spec:   '',     // 規格はこのツールでは空（UI側で一括入力可）
      unit:   unit,
      qty:    qtyText,
      price:  priceText,
      amount: amountText,
      note:   ''
    };

    rows.push(row);

    // すでに [i..blockEnd-1] を処理したので、次のループは blockEnd-1 の次から
    i = blockEnd - 1;
  }

  return rows;
}

// -------------------- 最終ページ集計 --------------------

// 最終ページ集計（原本値）の抽出
function parseSummaryFromText(text) {
  var normalized = (text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  var lines = normalized.split('\n');
  var n = lines.length;

  var startIdx = -1;
  var i, t;

  // 末尾側から「以　下　余　白」を探す
  for (i = n - 1; i >= 0; i--) {
    t = lines[i] ? lines[i].replace(/^\s+|\s+$/g, '') : '';
    if (!t) continue;
    if (t.indexOf('以') !== -1 && t.indexOf('余') !== -1) {
      startIdx = i + 1;
      break;
    }
  }

  if (startIdx === -1) {
    // 見つからなければ、最後の10行くらいをざっくり見る
    startIdx = n - 10;
    if (startIdx < 0) startIdx = 0;
  }

  var numbers = [];
  for (i = startIdx; i < n; i++) {
    var line = lines[i];
    if (!line) continue;
    var tokens = findNumberTokens(line);
    if (!tokens.length) continue;
    // 各行の最後の数値を採用
    numbers.push(tokens[tokens.length - 1].text);
    if (numbers.length >= 3) break;
  }

  if (numbers.length < 3) {
    return { base: '', tax: '', total: '' };
  }

  // 想定順序：
  //   1行目: 課税対象額
  //   2行目: 合計（\1,047,148- など）
  //   3行目: 消費税
  var baseVal  = numbers[0];
  var totalRaw = numbers[1];
  var taxVal   = numbers[2];
  var totalVal = normalizeTotal(totalRaw);

  return {
    base:  baseVal,
    tax:   taxVal,
    total: totalVal
  };
}

// -------------------- エクスポート関数 --------------------

// 貼り付けテキスト全体 → { rows, summary } にまとめて返す
// summary には以下を追加：
//   summary.calcBase      … 金額合計（数値）
//   summary.calcBaseStr   … 金額合計（"999,999.99"）
//   summary.baseCheckMark … 課税対象額と一致すれば "<"、不一致なら ""
function parseLedgerText(text) {
  var rows = parseDetailsFromText(text || '');
  var summary = parseSummaryFromText(text || '');

  // 金額合計の計算
  var sum = 0;
  var i, v;
  for (i = 0; i < rows.length; i++) {
    v = parseNumberSimple(rows[i].amount);
    if (!isNaN(v)) {
      sum += v;
    }
  }

  summary.calcBase = sum;
  summary.calcBaseStr = rows.length ? formatAmount(sum) : '';

  var baseNum = parseNumberSimple(summary.base);
  var mark = '';
  if (!isNaN(baseNum) && !isNaN(sum)) {
    // 浮動小数の誤差を考慮して ±0.5 以内なら一致とみなす
    if (Math.abs(baseNum - sum) < 0.5) {
      mark = '<';
    }
  }
  summary.baseCheckMark = mark;

  return {
    rows: rows,
    summary: summary
  };
}