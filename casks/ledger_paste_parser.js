// 2025-12-04 10:00 JST

// ledger_paste_parser.js
// 貼り付けテキスト → 明細 rows ＋ 最終ページ集計 summary への変換専用
// v2025.12.04-01
//
// ●責務
//   - 納入台帳の「貼り付けテキスト」から
//     ・明細行（No / 品名 / 規格(空) / 単位 / 合計数量 / 契約単価 / 金額 / 備考(空)）
//     ・最終ページ集計（課税対象額 / 消費税 / 合計）
//     を抽出して返すだけに特化
//
// ●数量・単価・金額 抽出方針（今回の修正版）
//   1) 各品目ブロック内で「数値トークンが2個以上ある行」をすべて候補にする
//   2) そのうち「一番下（最後）の行」を「合計数量＋契約単価」の行とみなす
//        → その行の「後ろから2番目＝合計数量」「一番後ろ＝契約単価」
//   3) その行より下で「数値を含む最初の行」の数値を「金額」とみなす
//        例: 0.10 19,500.00  ← 最後の組（合計数量・単価）
//            1,950.00        ← 次の行の数値が金額
//   4) このパターンが取れなかった場合だけ、従来の「最大値＝金額＋ペア探索」ロジックでフォールバック
//
//   ※ユーザー指定のルール：
//     「数字がスペースを挟み並んでいる最後の組が合計数量・契約単価、その次の行が金額」
//     を、そのままコード化しています。
//
// ●品名＆単位抽出方針
//   - 000x から始まる行を品目開始として、単位（PC/BG/KG/EA/CA/CN/SH）が出るまでを品名として抽出
//   - 2行以上に分かれている品名も結合して1つの品名とする
//   - 納地＋業者名の合体文字列は vendor としてそのまま保持する（後でユーザが修正）

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

    // 令和○年○月 → 次行が「納地＋業者名」の行
    if (trimmed.indexOf('令和') !== -1 &&
        trimmed.indexOf('年') !== -1 &&
        trimmed.indexOf('月') !== -1) {
      var vIdx;
      for (vIdx = i + 1; vIdx < n; vIdx++) {
        var vn = lines[vIdx];
        if (!vn) continue;
        var vnTrim = vn.replace(/^\s+|\s+$/g, '');
        if (vnTrim) {
          // 納地＋業者名の合体文字列をそのまま vendor として保持
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

    // ブロック内の各行の数値トークンを事前に取得
    var numsByLine = {};
    for (kk = i; kk < blockEnd; kk++) {
      var lTmp = lines[kk];
      numsByLine[kk] = lTmp ? findNumberTokens(lTmp) : [];
    }

    var qtyText    = '';
    var priceText  = '';
    var amountText = '';

    // -------------- 新ルール：最後の「数量＋単価」行と、その直後の金額行 --------------

    // 「数値トークンが2個以上ある行」のうち、最後の行を探す
    var lastPairLine = -1;
    for (p = i; p < blockEnd; p++) {
      var tksP = numsByLine[p];
      if (!tksP || tksP.length < 2) continue;
      // 「000x コード行」が2個以上になるケースは想定しないが、一応除外しておく
      var skipCodeLine = false;
      for (kk = 0; kk < tksP.length; kk++) {
        if (tksP[kk].text === code) {
          skipCodeLine = true;
          break;
        }
      }
      if (skipCodeLine) continue;

      lastPairLine = p; // 毎回更新 → 最後の行が残る
    }

    if (lastPairLine !== -1) {
      var tokensPair = numsByLine[lastPairLine];
      var lenPair = tokensPair.length;
      if (lenPair >= 2) {
        // 後ろから2番目＝合計数量、一番後ろ＝契約単価
        var qTok = tokensPair[lenPair - 2];
        var priceTok = tokensPair[lenPair - 1];
        var qVal = parseNumberSimple(qTok.text);
        var pVal = parseNumberSimple(priceTok.text);
        if (!isNaN(qVal) && !isNaN(pVal)) {
          qtyText   = qTok.text;
          priceText = priceTok.text;

          // 次の「数値を含む行」の数値を金額とみなす
          var amountLine = -1;
          for (kk = lastPairLine + 1; kk < blockEnd; kk++) {
            var tksAmt = numsByLine[kk];
            if (tksAmt && tksAmt.length > 0) {
              amountLine = kk;
              break;
            }
          }
          if (amountLine !== -1) {
            var tokensAmt = numsByLine[amountLine];
            amountText = tokensAmt[tokensAmt.length - 1].text;
          } else {
            // 金額行が無い場合は、数量×単価で計算して補う
            amountText = formatAmount(qVal * pVal);
          }
        }
      }
    }

    // -------------- フォールバック：旧ロジック（最大値＝金額＋ペア探索） --------------
    if (!qtyText && !priceText && !amountText) {
      // ブロック内の数値トークン（コード000xも含む）を収集
      var tokensBlock = [];
      var globalIndex = 0;
      for (kk = i; kk < blockEnd; kk++) {
        var lNum = lines[kk];
        if (!lNum) continue;
        var tks = numsByLine[kk] || findNumberTokens(lNum);
        for (var ti2 = 0; ti2 < tks.length; ti2++) {
          var txt2 = tks[ti2].text;
          var val2 = parseNumberSimple(txt2);
          if (isNaN(val2)) continue;
          tokensBlock.push({
            text: txt2,
            value: val2,
            isCode: (txt2 === code),
            globalIndex: globalIndex++
          });
        }
      }

      var nonCodeTokens = [];
      var t;
      for (t = 0; t < tokensBlock.length; t++) {
        if (!tokensBlock[t].isCode) {
          nonCodeTokens.push(tokensBlock[t]);
        }
      }

      if (nonCodeTokens.length >= 3) {
        // 最大値を金額候補とみなす
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
              var qTok2 = candidates[qi];
              var pTok2 = candidates[pj];
              var qv2 = qTok2.value;
              var pv2 = pTok2.value;
              if (qv2 <= 0 || pv2 <= 0) continue;

              var prod2 = qv2 * pv2;
              var diff2 = Math.abs(prod2 - amountToken.value);

              if (diff2 + 1e-6 < bestDiff) {
                bestDiff = diff2;
                bestQtyTok = qTok2;
                bestPriceTok = pTok2;
              } else if (Math.abs(diff2 - bestDiff) <= 0.5) {
                if (bestQtyTok === null ||
                    qTok2.globalIndex < bestQtyTok.globalIndex ||
                    (qTok2.globalIndex === bestQtyTok.globalIndex &&
                     pTok2.globalIndex < bestPriceTok.globalIndex)) {
                  bestQtyTok = qTok2;
                  bestPriceTok = pTok2;
                }
              }
            }
          }
        }

        if (amountToken) {
          amountText = amountToken.text;
        }

        if (bestQtyTok && bestPriceTok) {
          qtyText   = bestQtyTok.text;
          priceText = bestPriceTok.text;
        } else if (nonCodeTokens.length >= 2) {
          // さらにフォールバック：小さい順に2つを数量・単価とみなす
          var arr = [];
          for (t = 0; t < nonCodeTokens.length; t++) arr.push(nonCodeTokens[t]);
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

          var qTokF = arr[0];
          var pTokF = arr[1];
          qtyText   = qTokF.text;
          priceText = pTokF.text;

          var qv3 = qTokF.value;
          var pv3 = pTokF.value;
          if (!isNaN(qv3) && !isNaN(pv3)) {
            amountText = formatAmount(qv3 * pv3);
          }
        }

      } else if (nonCodeTokens.length === 2) {
        // 数値が2つだけ：小さい方＝数量、大きい方＝単価
        var tA = nonCodeTokens[0];
        var tB = nonCodeTokens[1];
        var qTokF2, pTokF2;
        if (tA.value <= tB.value) {
          qTokF2 = tA;
          pTokF2 = tB;
        } else {
          qTokF2 = tB;
          pTokF2 = tA;
        }
        qtyText   = qTokF2.text;
        priceText = pTokF2.text;
        var qv4 = qTokF2.value;
        var pv4 = pTokF2.value;
        if (!isNaN(qv4) && !isNaN(pv4)) {
          amountText = formatAmount(qv4 * pv4);
        }

      } else if (nonCodeTokens.length === 1) {
        // 数値が1つだけ：数量だけ分かる
        qtyText = nonCodeTokens[0].text;
      }
    }

    // -------------------- 品名＆単位の抽出 --------------------

    var nameParts = [];
    var unitFromName = '';

    // 1行目（コード行）から：コード以降を tail として処理
    var tail = line.substring(m.index + 4); // 4桁コードの直後から
    var tailTrim = tail.replace(/^\s+|\s+$/g, '');
    if (tailTrim) {
      var umHead = unitRe.exec(tailTrim);
      if (umHead) {
        var unitIndex = umHead.index;
        var nameCandidate = tailTrim.substring(0, unitIndex);
        nameCandidate = nameCandidate.replace(/\s+$/g, '');
        if (nameCandidate) {
          nameParts.push(nameCandidate);
        }
        unitFromName = umHead[1];
      } else {
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
      var tokens3 = numsByLine[p] || findNumberTokens(l3);

      if (!tokens3.length && !um2) {
        // 数字も単位も無い → 品名の続き
        nameParts.push(t3);
        continue;
      }

      if (um2) {
        // この行で単位が出てきた
        var unitIdx2 = um2.index;
        var prefixUnit = t3.substring(0, unitIdx2);
        prefixUnit = prefixUnit.replace(/\s+$/g, '');
        if (prefixUnit) {
          nameParts.push(prefixUnit);
        }
        if (!unitFromName) {
          unitFromName = um2[1];
        }
        break;
      }

      // 単位は無いが数字がある行 → 先頭部分だけ品名として追加
      var firstIdx3 = tokens3[0].index;
      if (firstIdx3 > 0) {
        var prefix3 = l3.substring(0, firstIdx3);
        prefix3 = prefix3.replace(/^\s+|\s+$/g, '');
        if (prefix3) {
          nameParts.push(prefix3);
        }
      }
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

    var row = {
      vendor: currentVendor,
      no:     code,
      name:   name,
      spec:   '',
      unit:   unit,
      qty:    qtyText,
      price:  priceText,
      amount: amountText,
      note:   ''
    };

    rows.push(row);

    i = blockEnd - 1;  // このブロックはここまで処理済み
  }

  return rows;
}

// -------------------- 最終ページ集計 --------------------

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
    startIdx = n - 10;
    if (startIdx < 0) startIdx = 0;
  }

  var numbers = [];
  for (i = startIdx; i < n; i++) {
    var line = lines[i];
    if (!line) continue;
    var tokens = findNumberTokens(line);
    if (!tokens.length) continue;
    numbers.push(tokens[tokens.length - 1].text);
    if (numbers.length >= 3) break;
  }

  if (numbers.length < 3) {
    return { base: '', tax: '', total: '' };
  }

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

function parseLedgerText(text) {
  var rows = parseDetailsFromText(text || '');
  var summary = parseSummaryFromText(text || '');

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