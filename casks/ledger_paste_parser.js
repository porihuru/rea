// ledger_paste_parser.js
// 貼り付けテキスト → 明細 rows ＋ 最終ページ集計 summary への変換専用
// v2025.11.24-02
//
// ●責務
//   - 納入台帳の「貼り付けテキスト」から
//     ・明細行（No / 品名 / 規格(空) / 単位 / 合計数量 / 契約単価 / 金額 / 備考(空)）
//     ・最終ページ集計（課税対象額 / 消費税 / 合計）
//     を抽出して返すだけに特化
//
// ●追加仕様
//   - rows の金額合計（amount の合計）を計算し、summary に
//       summary.calcBase      … 数値（合計）
//       summary.calcBaseStr   … "999,999.99" 形式の文字列
//       summary.baseCheckMark … 課税対象額と一致すれば "<"、一致しなければ ""
//     を付加する。
//   - 金額・数量・単価は原本の数字を「そのまま」抽出する方針。
//     ・金額はブロック中の最大値を優先して、そのテキストを使用。
//     ・数量・単価は「数量×単価 ≒ 金額」で最も近いペアを探索して決定する。
//       （数量・単価の値自体は原本の値をそのまま採用）
//
// ●品名＆単位抽出方針（2025.11.24-02）
//   1) コード行（000x を含む行）から次の「単位行」までの間を全部つないで品名とする
//      - 単位行とは、trim後が PC/BG/KG/EA/CA/CN/SH のいずれかだけになっている行
//      - 例:
//          0003
//          苺タルト
//          EA
//        → 品名 = "苺タルト", 単位 = "EA"
//          0004
//          おつまみスティッ
//          ク
//          EA
//        → 品名 = "おつまみスティック", 単位 = "EA"
//      - 数字から始まる品名（1L〜 等）も、コード行〜単位行の間にあればそのまま品名になる
//   2) もしコード行の中にすでに「品名＋単位」が同居していれば、従来通りその行から抽出
//
// ●数量・単価・金額抽出方針
//   - ブロック内の数値からコードを除いたものだけを対象にして：
//     ・3個以上:
//         - 最大値を金額候補とする（amount）
//         - 残りから「数量×単価 ≒ 金額」となるペアを全探索して数量・単価候補を決める
//     ・2個:
//         - 小さい方＝数量、大きい方＝単価、金額＝数量×単価（※金額列が無い場合用のフォールバック）
//     ・1個: 数量のみ分かるとみなす
//     ・0個: すべて空のまま
//
//   ★ページヘッダー行（「納　地：」「納入台帳」など）は
//     ブロック境界として扱い、数量・単価・金額の候補に入らないようにする。
//   ★最終ページの「-　以　下　余　白　-」より下の集計値
//     （課税対象額・消費税・合計）は、明細ブロックとは完全に分離する。

// -------------------- ユーティリティ --------------------

// 数値トークン抽出（数量・単価・金額・集計用）
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

// 品名と単位の分離（末尾単位くっつき用）
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
  var i, k, kk, p;
  var currentVendor = '';

  // コード検出用（0001, 0002, ...）
  var codeRe = /0\d{3}(?!\d)/g;
  // 単位検出用
  var unitRe = /(PC|BG|KG|EA|CA|CN|SH)(?=\s|$)/;

  for (i = 0; i < n; i++) {
    var line = lines[i];
    var trimmed = line ? line.replace(/^\s+|\s+$/g, '') : '';

    // 令和○年○月 → 次行が「納地＋業者名」の行（例：滝川駐屯地株式会社○○）
    if (trimmed.indexOf('令和') !== -1 &&
        trimmed.indexOf('年') !== -1 &&
        trimmed.indexOf('月') !== -1) {
      var vIdx;
      for (vIdx = i + 1; vIdx < n; vIdx++) {
        var vn = lines[vIdx];
        if (!vn) continue;
        var vnTrim = vn.replace(/^\s+|\s+$/g, '');
        if (vnTrim) {
          // ★ここは「納地＋業者名」の合体文字をそのまま保持★
          currentVendor = vnTrim;
          break;
        }
      }
      continue;
    }

    if (!trimmed) {
      continue;
    }

    // 品目開始行（0001, 0002, ... を含む行）
    codeRe.lastIndex = 0;
    var m = codeRe.exec(line);
    if (!m) {
      continue;
    }

    var code = m[0];

    // この品目ブロックの終了位置を探す
    //   - 次の "000x" 行
    //   - ページヘッダー行（「納　地：」「納入台帳」など）
    //   - 最終ページの「以 下 余 白」行（この行までをブロックに含める）
    var blockEnd = n;
    for (k = i + 1; k < n; k++) {
      var l2 = lines[k];
      var t2 = l2 ? l2.replace(/^\s+|\s+$/g, '') : '';
      if (!t2) {
        continue;
      }

      // 次のコード行 → 手前までがこの品目
      codeRe.lastIndex = 0;
      if (codeRe.exec(l2)) {
        blockEnd = k;
        break;
      }

      // ページヘッダー行
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
    // blockEnd が n のままなら、最後までがブロック

    // ブロック内の数値トークン（コード000xも含む）を収集
    // token = { text, value, isCode, globalIndex }
    var tokensBlock = [];
    var globalIndex = 0;
    for (kk = i; kk < blockEnd; kk++) {
      var lNum = lines[kk];
      if (!lNum) continue;
      var tks = findNumberTokens(lNum);
      var ti;
      for (ti = 0; ti < tks.length; ti++) {
        var txt = tks[ti].text;
        var val = parseNumberSimple(txt);
        if (isNaN(val)) continue;
        var token = {
          text: txt,
          value: val,
          isCode: (txt === code),
          globalIndex: globalIndex++
        };
        tokensBlock.push(token);
      }
    }

    // デフォルト値
    var qtyText    = '';
    var priceText  = '';
    var amountText = '';

    // 非コード数値だけを抜き出し
    var nonCodeTokens = [];
    var t;
    for (t = 0; t < tokensBlock.length; t++) {
      if (!tokensBlock[t].isCode) {
        nonCodeTokens.push(tokensBlock[t]);
      }
    }

    if (nonCodeTokens.length >= 3) {
      // -------- 3つ以上ある場合：最大値を金額候補として
      //         「数量×単価 ≒ 金額」を全探索 --------
      var amountToken = null;
      for (t = 0; t < nonCodeTokens.length; t++) {
        var tok = nonCodeTokens[t];
        if (amountToken === null || tok.value > amountToken.value) {
          amountToken = tok;
        }
      }

      var candidates = [];
      for (t = 0; t < nonCodeTokens.length; t++) {
        tok = nonCodeTokens[t];
        if (tok === amountToken) continue;
        if (isNaN(tok.value)) continue;
        if (tok.value <= 0) continue;
        candidates.push(tok);
      }

      var bestQtyTok = null;
      var bestPriceTok = null;
      var bestDiff = Number.POSITIVE_INFINITY;

      if (amountToken && candidates.length >= 2) {
        var qi, pj;
        for (qi = 0; qi < candidates.length; qi++) {
          for (pj = 0; pj < candidates.length; pj++) {
            if (qi === pj) continue;
            var qTok = candidates[qi];
            var pTok = candidates[pj];
            var q = qTok.value;
            var pVal = pTok.value;
            if (q <= 0 || pVal <= 0) continue;

            var prod = q * pVal;
            var diff = Math.abs(prod - amountToken.value);

            if (diff + 1e-6 < bestDiff) {
              bestDiff = diff;
              bestQtyTok = qTok;
              bestPriceTok = pTok;
            } else if (Math.abs(diff - bestDiff) <= 0.5) {
              // 誤差がほぼ同じなら「先に出てきたトークン」を優先
              if (bestQtyTok === null ||
                  qTok.globalIndex < bestQtyTok.globalIndex ||
                  (qTok.globalIndex === bestQtyTok.globalIndex &&
                   pTok.globalIndex < bestPriceTok.globalIndex)) {
                bestQtyTok = qTok;
                bestPriceTok = pTok;
              }
            }
          }
        }
      }

      // 金額は常に「原本の最大値トークン」を採用
      if (amountToken) {
        amountText = amountToken.text;
      }

      if (bestQtyTok && bestPriceTok) {
        // 数量・単価も原本のトークンをそのまま採用
        qtyText   = bestQtyTok.text;
        priceText = bestPriceTok.text;
      } else {
        // 念のためのフォールバック：
        // 非コード数値のうち「小さい順に2つ」を数量・単価とみなす
        if (nonCodeTokens.length >= 2) {
          var arr = [];
          for (t = 0; t < nonCodeTokens.length; t++) {
            arr.push(nonCodeTokens[t]);
          }
          // バブルソート（古いブラウザ用）
          var swapped, tmp;
          do {
            swapped = false;
            for (var s = 0; s < arr.length - 1; s++) {
              if (arr[s].value > arr[s + 1].value) {
                tmp = arr[s];
                arr[s] = arr[s + 1];
                arr[s + 1] = tmp;
                swapped = true;
              }
            }
          } while (swapped);

          var qTok2 = arr[0];
          var pTok2 = arr[1];
          qtyText   = qTok2.text;
          priceText = pTok2.text;

          var qv2 = qTok2.value;
          var pv2 = pTok2.value;
          if (!isNaN(qv2) && !isNaN(pv2)) {
            amountText = formatAmount(qv2 * pv2);
          }
        }
      }

    } else if (nonCodeTokens.length === 2) {
      // -------- 数値が2つだけ：金額列が無いとみなし
      //         小さい方＝数量、大きい方＝単価、金額＝数量×単価 --------
      var tA = nonCodeTokens[0];
      var tB = nonCodeTokens[1];
      var qTok2, pTok2;
      if (tA.value <= tB.value) {
        qTok2 = tA;
        pTok2 = tB;
      } else {
        qTok2 = tB;
        pTok2 = tA;
      }
      qtyText   = qTok2.text;
      priceText = pTok2.text;
      var qv3 = qTok2.value;
      var pv3 = pTok2.value;
      if (!isNaN(qv3) && !isNaN(pv3)) {
        amountText = formatAmount(qv3 * pv3);
      }

    } else if (nonCodeTokens.length === 1) {
      // -------- 数値が1つだけ：数量だけ分かるとみなす --------
      qtyText = nonCodeTokens[0].text;

    } else {
      // 非コード数値が無い場合はそのまま（全て空）
    }

    // -------------------- 品名＆単位の抽出 --------------------

    var nameParts = [];
    var unitFromName = '';

    // ★1) まずコード行の tail に「品名＋単位」が同居していないか見る（従来ロジック）★
    var tail = line.substring(m.index + 4); // 4桁コードの直後から
    var tailTrim = tail.replace(/^\s+|\s+$/g, '');
    if (tailTrim) {
      var umHead = unitRe.exec(tailTrim);
      if (umHead) {
        // tail 内に単位がある → その手前まで全部が品名
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

    // ★2) コード行側で単位が見つからなかった場合：
    //     「コード行の次の行」から「最初に単位行が出てくるところ」までを品名として全部つなぐ★
    if (!unitFromName) {
      for (p = i + 1; p < blockEnd; p++) {
        var l3 = lines[p];
        if (!l3) continue;
        var t3 = l3.replace(/^\s+|\s+$/g, '');
        if (!t3) continue;

        // 行全体が単位だけなら「単位行」とみなす
        var mUnitOnly = t3.match(/^(PC|BG|KG|EA|CA|CN|SH)$/);
        if (mUnitOnly) {
          unitFromName = mUnitOnly[1];
          break; // 単位が確定したので品名抽出はここまで
        }

        // 単位文字列が行の中に紛れているパターン（例: "（黄） PC"）にも一応対応
        var um2 = unitRe.exec(t3);
        if (um2) {
          var unitIdx2 = um2.index;
          var prefixUnit = t3.substring(0, unitIdx2);
          prefixUnit = prefixUnit.replace(/\s+$/g, '');
          if (prefixUnit) {
            nameParts.push(prefixUnit);
          }
          unitFromName = um2[1];
          break;
        }

        // ここまで来たら、この行は「品名の続き」とみなしてまるごと追加
        // → 数字で始まる品名（半角/全角どちらも）もそのまま品名扱いになる
        nameParts.push(t3);
      }
    }

    var fullName = nameParts.join('');
    var name = '';
    var unit = '';

    if (unitFromName) {
      name = fullName;
      unit = unitFromName;
    } else {
      // どうしても単位が取れなかった場合は、末尾くっつき判定でフォロー
      var nu = splitNameAndUnit(fullName);
      name = nu.name;
      unit = nu.unit;
    }

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
    // 完全一致でなくても、浮動小数の誤差を考慮して ±0.5 以内なら一致とみなす
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
