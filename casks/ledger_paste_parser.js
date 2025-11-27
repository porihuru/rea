// ledger_paste_parser.js
// 貼り付けテキスト → 明細 rows ＋ 最終ページ集計 summary への変換専用
// v2025.11.23-08
//
// ●責務
//   - 納入台帳の「貼り付けテキスト」から
//     ・明細行（No / 品名 / 規格(空) / 単位 / 合計数量 / 契約単価 / 金額 / 備考(空)）
//     ・最終ページ集計（課税対象額 / 消費税 / 合計）
//     を抽出して返すだけに特化
//
// ●仕様ポイント
//   - 数量・単価・金額は「元テキストに存在する数値トークンだけ」を使用する。
//     → 新しく計算した金額を作らない（フォールバックの数量×単価=金額は廃止）。
//   - 各品目ブロック内では、数値候補から
//       ・最大値を金額候補
//       ・数量×単価 ≒ 金額候補
//     となるペアを探索し、元のトークンを数量・単価・金額として採用する。
//     （採用する値そのものは、すべてテキスト内のトークン。計算値は使わない。）
//   - 「4,257.00  -　以　下　余　白　」の行は
//       → 最終品目ブロックに含める（＝金額として使えるようにする）
//       → その次の行からが最終ページ集計として扱わ
//         れるように、行分割ルールを調整する。

(function (global) {
  'use strict';

  // -----------------------------
  //  数値ユーティリティ
  // -----------------------------

  // 「969,582.00」「\1,047,148-」「77,566」「18.00」「0.50」などを正規化パース
  function parseNumberToken(raw) {
    if (!raw) return NaN;
    var s = String(raw);
    // 通貨記号・全角の円マークなどを除去
    s = s.replace(/[¥\\]/g, '');
    // カンマ・末尾ハイフンなどを除去
    s = s.replace(/,/g, '');
    s = s.replace(/-+$/g, '');
    s = s.replace(/^\s+|\s+$/g, '');
    if (!s) return NaN;
    var n = parseFloat(s);
    if (isNaN(n)) return NaN;
    return n;
  }

  // 3桁カンマ＋小数2桁でフォーマット
  function formatAmount(num) {
    var n = parseFloat(num);
    if (isNaN(n)) return '';
    var fixed = n.toFixed(2);
    var parts = fixed.split('.');
    var intPart = parts[0];
    var decPart = parts[1];
    var re = /(\d+)(\d{3})/;
    while (re.test(intPart)) {
      intPart = intPart.replace(re, '$1' + ',' + '$2');
    }
    return intPart + '.' + decPart;
  }

  // 整数用（3桁カンマ）
  function formatInt(num) {
    var n = parseInt(Math.floor(num), 10);
    if (isNaN(n)) return '';
    var s = String(n);
    var re = /(\d+)(\d{3})/;
    while (re.test(s)) {
      s = s.replace(re, '$1' + ',' + '$2');
    }
    return s;
  }

  // 「4,257.00  -　以　下　余　白　」のように金額＋余白文字が混在する行から、
  // 金額部分（最後の数値トークン）だけを取り出して返す。
  function extractLastNumberTokenFromLine(line) {
    if (!line) return null;
    // 数値（カンマ・小数点・円記号・末尾ハイフンを含む）っぽいトークンを全部抽出
    var tokens = line.match(/([¥\\]?\d[\d,]*(?:\.\d+)?-?)/g);
    if (!tokens || tokens.length === 0) {
      return null;
    }
    return tokens[tokens.length - 1];
  }

  // 文字列中のすべての数値トークンを抽出して返す
  function extractAllNumberTokens(text) {
    if (!text) return [];
    var tokens = text.match(/([¥\\]?\d[\d,]*(?:\.\d+)?-?)/g);
    return tokens || [];
  }

  // -----------------------------
  //  最終ページ集計の抽出
  // -----------------------------

  function parseSummaryFromText(text) {
    // 「-　以　下　余　白」「以下余白」あたり以降を最終ページ集計として扱う
    var idx = text.indexOf('以　下　余　白');
    if (idx < 0) {
      idx = text.indexOf('以下余白');
    }
    if (idx < 0) {
      // 最終ページ集計らしき部分が見つからない場合は空
      return {
        base: '',
        tax: '',
        total: '',
        calcBaseStr: '',
        baseCheckMark: ''
      };
    }

    var tail = text.substring(idx);
    // 行に分割
    var lines = tail.split(/\r?\n/);

    // 「以　下　余　白」行の1つ下から、「課税対象額」「消費税」「合　計」などが
    // 出てくるまでの範囲を使って、数値トークンを集める。
    var numberTokens = [];
    var started = false;
    var i, line, tokens;

    for (i = 0; i < lines.length; i++) {
      line = lines[i];

      if (!started) {
        // 余白行を含む最初の2〜3行くらいまではスキップしても構わないが、
        // シンプルに「以　下　余　白」を含む行の次から開始する想定。
        if (line.indexOf('余　白') >= 0 || line.indexOf('以下余白') >= 0) {
          started = true;
        }
        continue;
      }

      // 「課税対象額」「消費税」「合　計」などのラベル行を見つけたら終了。
      if (line.indexOf('課税対象額') >= 0 ||
          line.indexOf('消費税') >= 0 ||
          line.indexOf('合　計') >= 0 ||
          line.indexOf('合計') >= 0) {
        break;
      }

      // この範囲にある数値トークンを全部集める
      tokens = extractAllNumberTokens(line);
      if (tokens && tokens.length > 0) {
        numberTokens = numberTokens.concat(tokens);
      }
    }

    if (numberTokens.length < 3) {
      // 想定より少ない場合は、その後ろの行（課税対象額などの右側）を個別に見る
      // ここでは簡易的に fallback→空を返しておく。
      return {
        base: '',
        tax: '',
        total: '',
        calcBaseStr: '',
        baseCheckMark: ''
      };
    }

    // 通常は「969,582.00」「\1,047,148-」「77,566」のように
    // 「課税対象額」「合計」「消費税」など順不同？の3つが並んでいる。
    // 用途上は「課税対象額」「消費税」「合計」の3つを返せればよいので、
    // とりあえず数値の「大きさ」で役割分担する。
    var parsed = numberTokens.map(function (tok) {
      return { raw: tok, value: parseNumberToken(tok) };
    }).filter(function (obj) {
      return !isNaN(obj.value);
    });

    if (parsed.length < 3) {
      return {
        base: '',
        tax: '',
        total: '',
        calcBaseStr: '',
        baseCheckMark: ''
      };
    }

    // valueの降順にソートして「合計」を最大値とみなす
    parsed.sort(function (a, b) {
      return b.value - a.value;
    });

    var totalObj = parsed[0]; // 最大値 → 合計
    var rest = parsed.slice(1);

    // 残りの中で税金っぽい値（最大値からある程度小さい）を消費税、
    // それ以外を課税対象額として扱う。
    // ここでは単純に「一番小さい値＝消費税」「残りの最大値＝課税対象額」ぐらいのルールにする。
    rest.sort(function (a, b) {
      return a.value - b.value;
    });
    var taxObj = rest[0];                 // 最小値 → 消費税
    var baseObj = rest[rest.length - 1];  // 最大値 → 課税対象額

    // 金額の合計（calcBase）としては「課税対象額」を採用
    var calcBase = baseObj.value;

    return {
      base: formatAmount(baseObj.value),
      tax: formatInt(taxObj.value),
      total: formatInt(totalObj.value),
      calcBaseStr: formatAmount(calcBase),
      baseCheckMark: '' // ここでは原本との比較は行わない
    };
  }

  // -----------------------------
  //  明細ブロック抽出の前処理
  // -----------------------------

  // 大きなテキストを、ページヘッダごとにブロックに切る
  // （「No」「納入台帳」あたりを起点にしてもよいが、ここではシンプルにそのまま）
  function splitIntoLines(text) {
    if (!text) return [];
    var lines = text.split(/\r?\n/);
    // 末尾の空行を削る
    while (lines.length > 0 && !lines[lines.length - 1].trim()) {
      lines.pop();
    }
    return lines;
  }

  // 「0001」「0002」…のような No 行か？
  function isNoLine(line) {
    return /^\s*\d{4}\s*$/.test(line);
  }

  // 「BG」「EA」「PC」「KG」「CN」「SH」などの単位行か？
  function isUnitLine(line) {
    // 大文字2〜3文字程度を単位とみなす（例外はあるが一旦これで）
    return /^\s*[A-Z]{1,3}\s*$/.test(line);
  }

  // 「1L豆乳飲料」「三杯酢もずく」「うなぎ長蒲焼」などの品名行か？
  // ここでは「No 行の次」「単位行の1つ上」など位置関係で判断するので、
  // 単独では厳密には判定しない。
  // → isCandidateNameLine として、最低限の記号条件だけにしておく。
  function isCandidateNameLine(line) {
    if (!line) return false;
    var s = line.trim();
    if (!s) return false;
    // 数値だけの行は品名ではない
    if (/^[\d,.\s]+$/.test(s)) return false;
    // 「No」「納入台帳」「金 額」等のキーワードを含む行は除外
    if (s.indexOf('納入台帳') >= 0) return false;
    if (s.indexOf('金') >= 0 && s.indexOf('額') >= 0) return false;
    if (s.indexOf('合計数量') >= 0) return false;
    if (s.indexOf('課税対象額') >= 0) return false;
    if (s.indexOf('消費税') >= 0) return false;
    if (s.indexOf('合　計') >= 0 || s.indexOf('合計') >= 0) return false;
    if (s.indexOf('以　下　余　白') >= 0) return false;
    if (s.indexOf('以下余白') >= 0) return false;

    return true;
  }

  // -----------------------------
  //  明細ブロックの抽出
  // -----------------------------

  function parseDetailsFromText(text) {
    var lines = splitIntoLines(text);
    var i;

    // ページヘッダ相当の行はそのまま残しつつ、No 行を起点に品目ブロックを検出
    // ブロック構造：
    //   [複数行の数量情報など]
    //   0001
    //   品名行
    //   単位行(EA/BG/PC…)
    //   [複数行の数量/単価/金額など]
    // といった塊を1品目とみなす。
    var blocks = []; // { noIndex, nameIndex, unitIndex, lineStart, lineEnd }
    var n = lines.length;

    for (i = 0; i < n; i++) {
      if (!isNoLine(lines[i])) continue;

      var noIndex = i;
      var nameIndex = -1;
      var unitIndex = -1;

      // No 行の次の行〜数行先に「品名候補」がある想定
      var j;
      for (j = noIndex + 1; j < Math.min(noIndex + 4, n); j++) {
        if (isCandidateNameLine(lines[j])) {
          nameIndex = j;
          break;
        }
      }
      if (nameIndex < 0) continue;

      // 品名行の次〜数行先に「単位行」がある想定
      for (j = nameIndex + 1; j < Math.min(nameIndex + 4, n); j++) {
        if (isUnitLine(lines[j])) {
          unitIndex = j;
          break;
        }
      }
      if (unitIndex < 0) continue;

      // ブロックの終了位置は、
      //   - 次の No 行の直前
      //   - または テーブルヘッダ（「No」「納入台帳」など）・「以　下　余　白」の直前
      var k = unitIndex + 1;
      while (k < n) {
        var line = lines[k];
        if (isNoLine(line)) break;
        if (line.indexOf('No') >= 0 && line.indexOf('納入台帳') >= 0) break;
        if (line.indexOf('以　下　余　白') >= 0 || line.indexOf('以下余白') >= 0) break;
        if (line.indexOf('課税対象額') >= 0 ||
            line.indexOf('消費税') >= 0 ||
            line.indexOf('合　計') >= 0 ||
            line.indexOf('合計') >= 0) {
          break;
        }
        k++;
      }

      blocks.push({
        noIndex: noIndex,
        nameIndex: nameIndex,
        unitIndex: unitIndex,
        lineStart: noIndex, // ブロック先頭
        lineEnd: k - 1      // ブロック末尾（次のNo/ヘッダの直前）
      });

      // 次の検索は、このブロック末尾の次の行から再開
      i = k - 1;
    }

    // 1ブロックごとに数量・単価・金額を数値トークンから抽出
    var rows = [];

    blocks.forEach(function (blk) {
      var noLine   = lines[blk.noIndex]   || '';
      var nameLine = lines[blk.nameIndex] || '';
      var unitLine = lines[blk.unitIndex] || '';

      // No はゼロ埋め4桁をそのまま使う
      var noMatch = noLine.match(/(\d{4})/);
      var noStr = noMatch ? noMatch[1] : '';

      // 品名はその行のテキスト全体（トリム）
      var nameStr = nameLine.trim();

      // 単位は EA/BG/PC/KG/CN/SH 等を想定
      var unitStr = unitLine.trim();

      // ブロック全体（lineStart〜lineEnd）を結合して数値トークン抽出
      var blockTextArr = [];
      var idx;
      for (idx = blk.lineStart; idx <= blk.lineEnd; idx++) {
        blockTextArr.push(lines[idx]);
      }
      var blockText = blockTextArr.join('\n');

      var tokens = extractAllNumberTokens(blockText);

      if (!tokens || tokens.length === 0) {
        // 数値が全く見つからない → 空で返す
        rows.push({
          vendor: '',
          no: noStr,
          name: nameStr,
          spec: '',
          unit: unitStr,
          qty: '',
          price: '',
          amount: '',
          note: ''
        });
        return;
      }

      // トークンを value 付きで持つ
      var tokenObjs = tokens.map(function (t) {
        return { raw: t, value: parseNumberToken(t) };
      }).filter(function (obj) {
        return !isNaN(obj.value);
      });

      if (tokenObjs.length === 0) {
        rows.push({
          vendor: '',
          no: noStr,
          name: nameStr,
          spec: '',
          unit: unitStr,
          qty: '',
          price: '',
          amount: '',
          note: ''
        });
        return;
      }

      // 金額候補：最大値（valueが最大のトークン）
      var maxObj = tokenObjs[0];
      var iTok;
      for (iTok = 1; iTok < tokenObjs.length; iTok++) {
        if (tokenObjs[iTok].value > maxObj.value) {
          maxObj = tokenObjs[iTok];
        }
      }

      var amountRaw = maxObj.raw;
      var amountVal = maxObj.value;

      // 数量×単価 ≒ 金額(最大値) となるペアを探す
      var bestPair = null;
      var tolerance = 0.5; // 誤差許容

      for (iTok = 0; iTok < tokenObjs.length; iTok++) {
        var qObj = tokenObjs[iTok];
        if (qObj === maxObj) continue;

        var jTok;
        for (jTok = 0; jTok < tokenObjs.length; jTok++) {
          var pObj = tokenObjs[jTok];
          if (pObj === maxObj) continue;
          if (pObj === qObj) continue;

          var qv = qObj.value;
          var pv = pObj.value;
          var prod = qv * pv;
          if (Math.abs(prod - amountVal) <= tolerance) {
            bestPair = {
              qtyRaw: qObj.raw,
              qtyVal: qv,
              priceRaw: pObj.raw,
              priceVal: pv
            };
            break;
          }
        }
        if (bestPair) break;
      }

      var qtyRaw  = '';
      var priceRaw = '';

      if (bestPair) {
        qtyRaw   = bestPair.qtyRaw;
        priceRaw = bestPair.priceRaw;
      }

      rows.push({
        vendor: '',
        no: noStr,
        name: nameStr,
        spec: '',
        unit: unitStr,
        qty: qtyRaw,
        price: priceRaw,
        amount: amountRaw,
        note: ''
      });
    });

    return rows;
  }

  // -----------------------------
  //  総合パーサ
  // -----------------------------

  function parseLedgerText(text) {
    if (!text) {
      return {
        rows: [],
        summary: {
          base: '',
          tax: '',
          total: '',
          calcBaseStr: '',
          baseCheckMark: ''
        }
      };
    }

    var rows = parseDetailsFromText(text);
    var summary = parseSummaryFromText(text);

    return {
      rows: rows,
      summary: summary
    };
  }

  // グローバル公開
  global.parseLedgerText = parseLedgerText;

})(this);
