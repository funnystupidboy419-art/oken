/* =========================================================================
 * util.js  共通ユーティリティ
 *  - 文字数カウント / エクセル貼り付け解析 / 保存 / 文体変換 など
 *  ※ 外部ライブラリ・外部通信は一切使用しません。
 * ========================================================================= */
(function (global) {
  'use strict';

  var Util = {};

  /* ---------------------------------------------------------------
   * 文字数
   * 通信表・指導要録は「句読点を含む文字数」で数えるのが一般的。
   * 改行は数えない。
   * ------------------------------------------------------------- */
  Util.countChars = function (s) {
    if (!s) return 0;
    return String(s).replace(/\r?\n/g, '').length;
  };

  /* ---------------------------------------------------------------
   * エクセルからの貼り付け(TSV/CSV)を二次元配列に変換
   *  - タブ区切りを優先。タブが無ければカンマ区切りとして扱う。
   *  - セル内改行を含むダブルクォート囲みにも対応。
   * ------------------------------------------------------------- */
  Util.parseTable = function (text) {
    if (!text) return [];
    var src = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    var delim = src.indexOf('\t') >= 0 ? '\t' : ',';
    var rows = [];
    var row = [];
    var cell = '';
    var inQuote = false;
    for (var i = 0; i < src.length; i++) {
      var c = src.charAt(i);
      if (inQuote) {
        if (c === '"') {
          if (src.charAt(i + 1) === '"') { cell += '"'; i++; }
          else { inQuote = false; }
        } else { cell += c; }
        continue;
      }
      if (c === '"' && cell === '') { inQuote = true; continue; }
      if (c === delim) { row.push(cell); cell = ''; continue; }
      if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; continue; }
      cell += c;
    }
    row.push(cell);
    rows.push(row);
    while (rows.length && rows[rows.length - 1].every(function (v) { return String(v).trim() === ''; })) {
      rows.pop();
    }
    return rows.map(function (r) {
      return r.map(function (v) { return String(v).replace(/^\s+|\s+$/g, ''); });
    });
  };

  /* 全角数字を含む文字列から出席番号を取り出す */
  Util.toNumber = function (v) {
    if (v === null || v === undefined) return null;
    var s = String(v).replace(/[０-９]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    }).replace(/[^0-9]/g, '');
    if (s === '') return null;
    return parseInt(s, 10);
  };

  /* 評価記号の正規化 (◎ / ○ / 空) */
  Util.normalizeMark = function (v) {
    var s = String(v == null ? '' : v).replace(/\s/g, '');
    if (s === '') return '';
    if (/^(◎|＠|@|A|a|ａ|Ａ|二重丸)/.test(s)) return '◎';
    if (/^(○|〇|◯|O|o|Ｏ|ｏ|B|b|1|丸)/.test(s)) return '○';
    return '';
  };

  /* ---------------------------------------------------------------
   * テンプレート展開
   *   {{キー}}      … 差し込み
   *   {敬体|常体}   … 文体による出し分け
   * ------------------------------------------------------------- */
  Util.render = function (tpl, vars, style) {
    if (!tpl) return '';
    var out = String(tpl);
    out = out.replace(/\{([^{}|]*)\|([^{}|]*)\}/g, function (_, polite, plain) {
      return style === 'plain' ? plain : polite;
    });
    out = out.replace(/\{\{([^{}]+)\}\}/g, function (_, key) {
      var v = vars ? vars[String(key).replace(/\s/g, '')] : '';
      return v == null ? '' : String(v);
    });
    return Util.tidy(out);
  };

  /* 余分な空白・重複した読点などを整える */
  Util.tidy = function (s) {
    return String(s || '')
      .replace(/[ 　]+/g, '')
      .replace(/、{2,}/g, '、')
      .replace(/。{2,}/g, '。')
      .replace(/、。/g, '。')
      .replace(/^[、。]+/, '')
      .replace(/^\s+|\s+$/g, '');
  };

  /* =================================================================
   * 文体変換（敬体 → 常体）
   *   自由記述欄の補助機能。機械的な変換のため、指導要録に使う前に
   *   必ず内容を目視で確認すること。
   * ================================================================= */

  /* イ段 → ウ段（五段動詞の終止形） */
  var I_TO_U = {
    'い': 'う', 'き': 'く', 'ぎ': 'ぐ', 'し': 'す', 'ち': 'つ',
    'に': 'ぬ', 'ひ': 'ふ', 'び': 'ぶ', 'み': 'む', 'り': 'る'
  };
  /* イ段 → ア段（五段動詞の未然形） */
  var I_TO_A = {
    'い': 'わ', 'き': 'か', 'ぎ': 'が', 'し': 'さ', 'ち': 'た',
    'に': 'な', 'ひ': 'は', 'び': 'ば', 'み': 'ま', 'り': 'ら'
  };
  /* 語幹がイ段でも一段活用になる代表語（例外辞書） */
  var ICHIDAN_TAIL = [
    'でき', '見', '観', '着', '起き', '落ち', '生き', '過ぎ', '感じ', '信じ',
    '応じ', '閉じ', '足り', '借り', '用い', '率い', '報い', '浴び', '延び', '伸び'
  ];

  function isIchidan(stem) {
    if (/[えけげせぜてでねへべぺめれ]$/.test(stem)) return true;   /* 食べ・考え・見られ 等 */
    for (var i = 0; i < ICHIDAN_TAIL.length; i++) {
      if (stem.length >= ICHIDAN_TAIL[i].length &&
          stem.slice(-ICHIDAN_TAIL[i].length) === ICHIDAN_TAIL[i]) return true;
    }
    return false;
  }

  /* サ変（〜する）かどうか：「活動し」「チャレンジし」など */
  function isSuru(stem) {
    if (stem === '' || stem === 'し') return true;
    return /(?:[一-龠々]{2,}|[ァ-ヶー]{2,})し$/.test(stem);
  }

  function suruBase(stem) {
    return (stem === '' || stem === 'し') ? '' : stem.slice(0, -1);
  }

  /* 「〜ます」の語幹 → 終止形 */
  function plainDict(stem) {
    if (isSuru(stem)) return suruBase(stem) + 'する';
    if (stem === '来' || stem === 'き') return '来る';
    if (isIchidan(stem)) return stem + 'る';
    var last = stem.slice(-1);
    return I_TO_U[last] ? stem.slice(0, -1) + I_TO_U[last] : stem + 'る';
  }

  /* 「〜ました」の語幹 → 過去形（音便あり） */
  function plainPast(stem) {
    if (isSuru(stem)) return suruBase(stem) + 'した';
    if (stem === '来' || stem === 'き') return '来た';
    if (isIchidan(stem)) return stem + 'た';
    if (/(行き|いき)$/.test(stem)) return stem.slice(0, -1) + 'った';   /* 行く は促音便 */
    var last = stem.slice(-1);
    var head = stem.slice(0, -1);
    if (last === 'き') return head + 'いた';
    if (last === 'ぎ') return head + 'いだ';
    if (last === 'し') return head + 'した';
    if (last === 'い' || last === 'ち' || last === 'り') return head + 'った';
    if (last === 'に' || last === 'び' || last === 'み') return head + 'んだ';
    return stem + 'た';
  }

  /* 「〜ません」の語幹 → 否定形 */
  function plainNeg(stem) {
    if (isSuru(stem)) return suruBase(stem) + 'しない';
    if (stem === '来' || stem === 'き') return '来ない';
    if (stem === 'あり') return 'ない';
    if (isIchidan(stem)) return stem + 'ない';
    var last = stem.slice(-1);
    return I_TO_A[last] ? stem.slice(0, -1) + I_TO_A[last] + 'ない' : stem + 'ない';
  }

  function plainNegPast(stem) {
    return plainNeg(stem).replace(/ない$/, 'なかった');
  }

  var STEM = '([ぁ-んァ-ヶ一-龠々ー]+?)';
  var MARK = String.fromCharCode(0xE000);

  Util.toPlain = function (text) {
    if (!text) return '';
    var s = String(text);
    s = s.replace(/ますます/g, MARK);

    /* 「〜ています」「〜ておりました」など補助動詞の形 */
    s = s.replace(/(て|で)(?:い|おり)ませんでした/g, '$1いなかった');
    s = s.replace(/(て|で)(?:い|おり)ません/g, '$1いない');
    s = s.replace(/(て|で)(?:い|おり)ました/g, '$1いた');
    s = s.replace(/(て|で)(?:い|おり)ます/g, '$1いる');

    /* 名詞述語・形容動詞 */
    s = s.replace(/ではありませんでした/g, 'ではなかった');
    s = s.replace(/ではありません/g, 'ではない');
    s = s.replace(/ありませんでした/g, 'なかった');
    s = s.replace(/ありません/g, 'ない');
    s = s.replace(/でしょう/g, 'だろう');
    s = s.replace(/でした/g, 'だった');
    s = s.replace(/ですが/g, 'だが');
    s = s.replace(/ですので/g, 'なので');
    s = s.replace(/でして/g, 'で');
    s = s.replace(/です/g, 'だ');
    s = s.replace(/ございます/g, 'ある');

    /* 動詞 */
    s = s.replace(new RegExp(STEM + 'ませんでした', 'g'), function (m, st) { return plainNegPast(st); });
    s = s.replace(new RegExp(STEM + 'ません', 'g'), function (m, st) { return plainNeg(st); });
    s = s.replace(new RegExp(STEM + 'ましょう', 'g'), function (m, st) { return plainDict(st).replace(/る$/, 'よう'); });
    s = s.replace(new RegExp(STEM + 'ました', 'g'), function (m, st) { return plainPast(st); });
    s = s.replace(new RegExp(STEM + 'まして', 'g'), function (m, st) {
      return plainPast(st).replace(/た$/, 'て').replace(/だ$/, 'で');
    });
    s = s.replace(new RegExp(STEM + 'ますので', 'g'), function (m, st) { return plainDict(st) + 'ので'; });
    s = s.replace(new RegExp(STEM + 'ます', 'g'), function (m, st) { return plainDict(st); });

    s = s.replace(new RegExp(MARK, 'g'), 'ますます');
    return s;
  };

  /* 常体 → 敬体（簡易。あくまで補助） */
  Util.toPolite = function (text) {
    if (!text) return '';
    var s = String(text);
    s = s.replace(/ていた([。、」）)]|$)/g, 'ていました$1');
    s = s.replace(/ている([。、」）)]|$)/g, 'ています$1');
    s = s.replace(/だった([。、」）)]|$)/g, 'でした$1');
    s = s.replace(/([^でま])だ([。」）)]|$)/g, '$1です$2');
    return s;
  };

  /* ---------------------------------------------------------------
   * 個人情報チェック（氏名らしき表現の検出）
   * ------------------------------------------------------------- */
  Util.PERSONAL_PATTERNS = [
    { re: '[一-龠々ぁ-んァ-ヶー]{1,4}(さん|くん|君|ちゃん)', label: '氏名らしい表現（〜さん・くん・ちゃん）' },
    { re: '(氏名|名前|なまえ|フリガナ|ふりがな|生年月日|住所|電話番号)', label: '個人情報の項目名' }
  ];

  /* 氏名ではない「〜さん」（家族・職業の呼び方、教材名など）は警告しない */
  Util.PERSON_EXCLUDE = [
    'お母さん', 'おかあさん', '母さん', 'お父さん', 'おとうさん', '父さん',
    'お兄さん', 'おにいさん', '兄さん', 'お姉さん', 'おねえさん', '姉さん',
    'おばあさん', 'おじいさん', 'ばあさん', 'じいさん', 'おばさん', 'おじさん',
    '赤ちゃん', 'みなさん', '皆さん', 'お客さん', 'お客様', '店員さん',
    '農家さん', '職員さん', '用務員さん', '運転手さん', 'おまわりさん', 'お巡りさん',
    '看護師さん', 'お医者さん', '大工さん', '漁師さん', '駅員さん', '住民さん'
  ];

  Util.findPersonalInfo = function (text) {
    if (!text) return [];
    var hits = [];
    Util.PERSONAL_PATTERNS.forEach(function (p) {
      var re = new RegExp(p.re, 'g');
      var m;
      while ((m = re.exec(text)) !== null) {
        var word = m[0];
        if (Util.PERSON_EXCLUDE.indexOf(word) >= 0) continue;
        if (/屋さん$/.test(word)) continue;
        hits.push({ word: word, label: p.label });
        if (hits.length > 20) return;
      }
    });
    return hits;
  };

  /* ---------------------------------------------------------------
   * 保存（この端末のブラウザ内のみ。外部へは送信しません）
   * ------------------------------------------------------------- */
  Util.save = function (key, data) {
    try {
      localStorage.setItem(key, JSON.stringify(data));
      return true;
    } catch (e) {
      return false;
    }
  };

  Util.load = function (key) {
    try {
      var s = localStorage.getItem(key);
      return s ? JSON.parse(s) : null;
    } catch (e) {
      return null;
    }
  };

  Util.clear = function (key) {
    try { localStorage.removeItem(key); } catch (e) { /* noop */ }
  };

  /* ファイルのダウンロード */
  function downloadByAnchor(filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  /* Claude のプレビュー（Artifact）上では通常のダウンロードリンクが無効化されるため、
     利用できるときは downloads 機能（保存確認ダイアログ）を使う。
     それ以外（自分のパソコンでファイルとして開いた場合など）は従来どおりの方法。 */
  Util.download = function (filename, blob) {
    if (window.claude && typeof window.claude.use === 'function') {
      window.claude.use('downloads').then(function (downloads) {
        if (!downloads) { downloadByAnchor(filename, blob); return; }
        return downloads.save({ filename: filename, data: blob }).catch(function (err) {
          if (err && err.code === 'declined') return; /* 保存しない選択。何もしない */
          alert('保存できませんでした：' + (err && err.message ? err.message : String(err)));
        });
      }).catch(function () { downloadByAnchor(filename, blob); });
      return;
    }
    downloadByAnchor(filename, blob);
  };

  Util.escapeHtml = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };

  Util.todayString = function () {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
  };

  global.Util = Util;
})(window);
