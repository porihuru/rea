// ledger_paste_parser.js
// 貼り付けテキスト → 明細 rows ＋ 最終ページ集計 summary への変換専用
// v2025.11.26-01
//
// ●責務
//   - 納入台帳の「貼り付けテキスト」から
//     ・明細行（No / 品名 / 規格(空) / 単位 / 合計数量 / 契約単価 / 金額 / 備考(空)）
//     ・最終ページ集計（課税対象額 / 消費税 / 合計）
//     を抽出して返すだけに特化
//
// ●ポイント
//   - 「合計数量・契約単価・金額」は *計算で作らず*、原本に印字されている数値をそのまま抽出する。
//     （どの数値が数量・単価・金額かを見分けるために、内部で q×p≈amount のチェックは使うが、
//      金額自体は必ず元の文字列を採用する）
//   - rows の金額合計（amount の合計）を計算し、summary に
//       summary.calcBase      … 数値（合計）
//       summary.calcBaseStr   … "999,999.99" 形式の文字列
//       summary.baseCheckMark … 課税対象額と一致すれば "<"、一致しなければ ""
//     を付加する。
//   - 貼り付けデータのレイアウト変化
//       0001（コード単独行）
//       品名（複数行に折り返される場合あり）
//       単位行（PC/BG/KG/EA/CA/CN/SH）
//       日別の数量行 …
//       「合計数量 契約単価」行
//       「金額」行
//       ＊印行 など
//     に対応する。

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
  var codeRe = /0\d{3}(?!\d)/g;
  var unitRe = /(PC|BG|KG|EA|CA|CN|SH)(?=\s|$)/; // 単位検出用

  for (i = 0; i < n; i++) {
    var line = lines[i];
    var trimmed = line ? line.replace(/^\s+|\s+$/g, '') : '';

    // 令和○年○月 → 次行が業者名（納地＋会社名の合体文字列）
    if (trimmed.indexOf('令和') !== -1 &&
        trimmed.indexOf('年') !== -1 &&
        trimmed.indexOf('月') !== -1) {
      var vIdx;
      for (vIdx = i + 1; vIdx < n; vIdx++) {
        var vn = lines[vIdx];
        if (!vn) continue;
        var vnTrim = vn.replace(/^\s+|\s+$/g, '');
        if (vnTrim) {
          // 「滝川駐屯地株式会社トワニ旭川店」等、合体文字をそのまま保持
          currentVendor = vnTrim;
          break;
        }
      }
      continue;
    }

    if (!trimmed) {
      continue;
    }

    // 品目開始行（0001〜0xxx）
    codeRe.lastIndex = 0;
    var m = codeRe.exec(line);
    if (!m) {
      continue;
    }

    var code = m[0];

    // この品目ブロックの終了位置を探す
    k = i + 1;
    for (; k < n; k++) {
      var l2 = lines[k];
      var t2 = l2 ? l2.replace(/^\s+|\s+$/g, '') : '';
      if (!t2) {
        continue;
      }
      codeRe.lastIndex = 0;
      if (codeRe.exec(l2)) {
        // 次の 0xxx が出たらブロック終了
        break;
      }
      // 「-　以　下　余　白　-」など → 集計エリアに入るので終了
      if (t2.indexOf('以') !== -1 && t2.indexOf('余') !== -1) {
        break;
      }
      // 「課税対象額」行に入ったら終了
      if (t2.indexOf('課税対象額') !== -1) {
        break;
      }
      // ページ先頭ヘッダー行で終了
      if (t2.indexOf('納入台帳') !== -1 && t2.indexOf('No') !== -1) {
        break;
      }
      if (t2.indexOf('納') !== -1 && t2.indexOf('業 者 名') !== -1) {
        break;
      }
    }
    var blockEnd = k; // [i .. blockEnd-1] が1品目

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

    // ---------- 数量・単価・金額の決定ロジック ----------
    //
    // ・「数量×単価 ≒ 金額」でどれがどれかを判定するが、
    //   金額そのものは *必ず元の文字列* を使う。
    //
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

      if (amountToken) {
        amountText = amountToken.text;  // ★金額は必ず原本値
      }
      if (bestQtyTok) {
        qtyText = bestQtyTok.text;
      }
      if (bestPriceTok) {
        priceText = bestPriceTok.text;
      }

    } else if (nonCodeTokens.length === 2) {
      // -------- 数値が2つだけ：小さい方を数量・大きい方を金額とみなす --------
      var tA = nonCodeTokens[0];
      var tB = nonCodeTokens[1];
      var smallTok, bigTok;
      if (tA.value <= tB.value) {
        smallTok = tA;
        bigTok = tB;
      } else {
        smallTok = tB;
        bigTok = tA;
      }
      qtyText    = smallTok.text;
      amountText = bigTok.text;
      // 単価は原本に無いと判断し、空のままにしておく（計算で作らない）

    } else if (nonCodeTokens.length === 1) {
      // -------- 数値が1つだけ：数量だけ分かるとみなす --------
      qtyText = nonCodeTokens[0].text;

    } else {
      // 非コード数値が無い場合はそのまま（全て空）
    }

    // -------------------- 品名＆単位の抽出 --------------------

    var nameParts = [];
    var unitFromName = '';

    // 1行目（コード行）は「0001」のみのことが多いので、tail はほぼ空になる。
    // 2行目以降で品名・単位を拾う。
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
      var tokens3 = findNumberTokens(l3);

      if (!tokens3.length && !um2) {
        // 数字も単位も無い → 完全に品名の続きとみなす
        nameParts.push(t3);
        continue;
      }

      if (um2) {
        // この行で単位が出てきた
        var unitIdx2 = um2.index;
        var prefixUnit = t3.substring(0, unitIdx2);
        prefixUnit = prefixUnit.replace(/\s+$/g, '');
        if (prefixUnit) {
          // 例: "手巻おにぎりほぐ し鮭EA" → "手巻おにぎりほぐ し鮭" を品名に追加
          nameParts.push(prefixUnit);
        }
        if (!unitFromName) {
          unitFromName = um2[1];
        }
        // 単位以降は数量などなので、ここで品名処理は終了
        break;
      }

      // 単位は無いが数字がある行
      // → 行頭の文字列だけを品名の続きにするか判断
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

    var row = {
      vendor: currentVendor, // 「滝川駐屯地株式会社トワニ旭川店」など合体文字のまま
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

    // i はこのブロックの末尾まで進める
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

  // 末尾側から「以　下　余　白」付近を探す
  for (i = n - 1; i >= 0; i--) {
    t = lines[i] ? lines[i].replace(/^\s+|\s+$/g, '') : '';
    if (!t) continue;
    if (t.indexOf('以') !== -1 && t.indexOf('余') !== -1) {
      startIdx = i + 1;
      break;
    }
  }

  // 見つからない場合は、末尾10行くらいから拾うフォールバック
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
    // 各行の最後の数値を採用
    numbers.push(tokens[tokens.length - 1].text);
    if (numbers.length >= 3) break;
  }

  if (numbers.length < 3) {
    return { base: '', tax: '', total: '' };
  }

  // 想定順：
  //   1行目 … 課税対象額
  //   2行目 … 合計（\1,497,681- など）
  //   3行目 … 消費税
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

  // 金額合計の計算（amount の合計）
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
    // ±0.5 以内なら一致とみなす
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
