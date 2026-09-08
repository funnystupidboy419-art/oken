/* =========================================================================
 * generate.js  所見の組み立て
 *   - 「どの項目についての文か」を保持したまま文を並べます。
 *   - 並び順は設定（学校ごと）に従い、上にあるものほど優先度が高いものとして
 *     文字数超過時は下の項目から短縮・削除します。
 * ========================================================================= */
(function (global) {
  'use strict';

  var T = global.Templates;
  var U = global.Util;

  var Generator = {};

  /* 所見のまとまり（カテゴリ）定義。並び順は設定で変更します。 */
  Generator.CATEGORIES = [
    { key: 'committee', label: '委員会' },
    { key: 'kakari', label: '係' },
    { key: 'event', label: '行事・行事の係' },
    { key: 'behavior', label: '行動の記録' },
    { key: 'observation', label: '日常の見取り' }
  ];

  /* ------------------------------------------------------------------
   * 語句の組み立て補助
   * ------------------------------------------------------------------ */
  function s(v) { return (v == null ? '' : String(v)).replace(/^\s+|\s+$/g, ''); }

  /* 末尾の助詞が重複しないように付ける（「について」など複数文字にも対応） */
  function withParticle(text, particle) {
    var t = s(text);
    if (!t) return '';
    if (t.slice(-particle.length) === particle) return t;
    if (particle.length === 1 && /[をにでへとやのはがも、。]$/.test(t)) return t;
    return t + particle;
  }

  function suffix(text, tail) {
    var t = s(text);
    return t ? t + tail : '';
  }

  /* 文例の選択（出席番号と「別案」の回数で散らして、同じ文の連続を避ける） */
  function pick(list, seed) {
    if (!list || !list.length) return '';
    var i = ((seed % list.length) + list.length) % list.length;
    return list[i];
  }

  /* ------------------------------------------------------------------
   * カテゴリごとの文の候補づくり
   *   返り値: { cat, label, detail, long, short }  （文体展開前のテンプレート）
   * ------------------------------------------------------------------ */

  function partCommittee(st, seed) {
    var name = s(st.committee);
    if (!name) return [];
    var vars = {
      '委員会': name,
      '役割句': suffix(st.committeeRole, 'として、'),
      '活動句': st.committeeWork ? withParticle(st.committeeWork, 'に') : '活動に'
    };
    return [{
      cat: 'committee',
      label: '委員会',
      detail: name + (st.committeeRole ? '／' + st.committeeRole : ''),
      long: T.committee.long,
      short: T.committee.short,
      vars: vars,
      seed: seed
    }];
  }

  function partKakari(st, seed) {
    var name = s(st.kakari);
    if (!name) return [];
    var vars = {
      '係': name,
      '仕事句': st.kakariWork ? withParticle(st.kakariWork, 'を') : '自分の仕事を'
    };
    return [{
      cat: 'kakari',
      label: '係',
      detail: name,
      long: T.kakari.long,
      short: T.kakari.short,
      vars: vars,
      seed: seed
    }];
  }

  function partEvent(st, seed) {
    var list = [];
    (st.events || []).forEach(function (ev, i) {
      var name = s(ev.name);
      if (!name) return;
      list.push({
        cat: 'event',
        label: '行事・行事の係',
        detail: name + (ev.role ? '／' + ev.role : ''),
        long: T.event.long,
        short: T.event.short,
        vars: {
          '行事': name,
          '役割句': suffix(ev.role, 'として、'),
          '頑張り句': suffix(ev.effort, '、')
        },
        seed: seed + i
      });
    });
    return list;
  }

  function partBehavior(st, seed, settings) {
    var marks = st.behavior || {};
    var order = (settings.behaviorOrder && settings.behaviorOrder.length)
      ? settings.behaviorOrder
      : T.BEHAVIOR_ITEMS.map(function (b) { return b.key; });

    var picked = order.filter(function (key) { return marks[key] === '◎' || marks[key] === '○'; });

    if (settings.behaviorPick === 'mark') {
      /* ◎ を優先し、その中では学校の並び順を保つ */
      picked = picked.filter(function (k) { return marks[k] === '◎'; })
        .concat(picked.filter(function (k) { return marks[k] === '○'; }));
    }
    var max = settings.behaviorMax || 2;
    picked = picked.slice(0, max);

    return picked.map(function (key, i) {
      var item = T.BEHAVIOR_ITEMS.filter(function (b) { return b.key === key; })[0];
      var mark = marks[key];
      var tpl = T.behavior[key] || {};
      var longList = tpl[mark] || tpl['○'] || [];
      var shortList = (tpl.short && (tpl.short[mark] || tpl.short['○'])) || longList;
      return {
        cat: 'behavior',
        label: '行動の記録',
        detail: (item ? item.label : key) + '（' + mark + '）',
        long: longList,
        short: shortList,
        vars: {},
        seed: seed + i
      };
    });
  }

  function partObservation(st, seed) {
    var list = [];
    (st.observations || []).forEach(function (ob, i) {
      var text = s(ob.text);
      if (!text) return;
      var pat = null;
      T.observationPatterns.forEach(function (p) { if (p.key === ob.pattern) pat = p; });
      if (!pat) pat = T.observationPatterns[0];
      list.push({
        cat: 'observation',
        label: '日常の見取り',
        detail: '見取り' + (i + 1),
        long: [pat.tpl],
        short: [pat.tpl],
        vars: { '見取り': text },
        raw: pat.key === 'as-is' ? text : null,
        seed: seed + i
      });
    });
    return list;
  }

  /* ------------------------------------------------------------------
   * 文字列化
   *   同じ文末（「〜取り組んでいました。」など）が続くと、まとめて読んだときに
   *   コピー＆ペーストしたような印象になる。保護者が読む文章として自然に
   *   なるよう、直前の文と文末が重なる候補は避けて選び直す。
   * ------------------------------------------------------------------ */
  function endingKey(text) {
    /* 文末の述語部分（おおむね最後の8〜10文字）を比較のキーにする */
    return String(text || '').slice(-9);
  }

  function renderCandidate(part, tpl, style) {
    var text = U.render(tpl, part.vars, style);
    /* 「そのまま使う」の見取りは、指導要録用に常体へ機械変換する */
    if (part.raw && style === 'plain') text = U.toPlain(text);
    if (text && !/[。」）)]$/.test(text)) text += '。';
    return text;
  }

  /* 直前の文と文末が同じにならないよう、候補を順にずらしながら選ぶ */
  function renderPartVaried(part, useShort, style, bump, prevEnding) {
    var list = useShort ? part.short : part.long;
    if (!list || !list.length) return { text: '', ending: prevEnding };
    var n = list.length;
    var start = (((part.seed || 0) + (bump || 0)) % n + n) % n;
    var fallback = null;
    for (var k = 0; k < n; k++) {
      var text = renderCandidate(part, list[(start + k) % n], style);
      if (fallback === null) fallback = text;
      if (!text) continue;
      if (!prevEnding || endingKey(text) !== prevEnding) {
        return { text: text, ending: endingKey(text) };
      }
    }
    return { text: fallback, ending: endingKey(fallback) };
  }

  function assemble(parts, limit, style, bumps) {
    var prevLongEnding = null;
    var prevShortEnding = null;
    var rendered = parts.map(function (p) {
      var bump = bumps && bumps[p.cat];
      var longR = renderPartVaried(p, false, style, bump, prevLongEnding);
      var shortR = renderPartVaried(p, true, style, bump, prevShortEnding);
      if (longR.text) prevLongEnding = longR.ending;
      if (shortR.text) prevShortEnding = shortR.ending;
      return {
        cat: p.cat,
        label: p.label,
        detail: p.detail,
        longText: longR.text,
        shortText: shortR.text,
        useShort: false
      };
    }).filter(function (p) { return p.longText || p.shortText; });

    function total(list) {
      return U.countChars(list.map(function (p) {
        return p.useShort ? p.shortText : p.longText;
      }).join(''));
    }

    /* 1) 後ろ（優先度の低い方）から順に短い文例へ置き換える */
    for (var i = rendered.length - 1; i >= 0 && limit && total(rendered) > limit; i--) {
      if (rendered[i].shortText && rendered[i].shortText !== rendered[i].longText) {
        rendered[i].useShort = true;
      }
    }
    /* 2) それでも収まらなければ、後ろの文から落とす（最低1文は残す） */
    var kept = rendered.slice();
    while (limit && kept.length > 1 && total(kept) > limit) {
      kept.pop();
    }

    var parts2 = kept.map(function (p) {
      return {
        cat: p.cat,
        label: p.label,
        detail: p.detail,
        text: p.useShort ? p.shortText : p.longText,
        shortened: p.useShort
      };
    });
    /* 文字数制限のために使わなかった文（画面で確認できるように返す） */
    var unused = rendered.slice(kept.length).map(function (p) {
      return {
        cat: p.cat,
        label: p.label,
        detail: p.detail,
        text: p.useShort ? p.shortText : p.longText
      };
    });
    var text = parts2.map(function (p) { return p.text; }).join('');
    return { text: text, parts: parts2, unused: unused, dropped: unused.length };
  }

  /* ------------------------------------------------------------------
   * 通信表／指導要録の所見（行動・活動・見取り）
   *   mode: 'report'(通信表・敬体) / 'youroku'(指導要録・常体)
   * ------------------------------------------------------------------ */
  Generator.build = function (student, settings, mode) {
    var style = (mode === 'youroku') ? 'plain' : 'polite';
    var limit = (mode === 'youroku') ? settings.limitYouroku : settings.limitReport;
    var seed = (student.no || 0);
    var bumps = student.bumps || {};

    var byCat = {
      committee: partCommittee(student, seed),
      kakari: partKakari(student, seed),
      event: partEvent(student, seed),
      behavior: partBehavior(student, seed, settings),
      observation: partObservation(student, seed)
    };

    var order = (settings.order && settings.order.length)
      ? settings.order
      : Generator.CATEGORIES.map(function (c) { return c.key; });

    var parts = [];
    order.forEach(function (key) {
      if (settings.enabled && settings.enabled[key] === false) return;
      parts = parts.concat(byCat[key] || []);
    });

    var result = assemble(parts, limit, style, bumps);

    /* 通信表のみ、任意で結びの一文を添える */
    if (mode === 'report' && settings.useClosing && result.text) {
      var closing = U.render(pick(T.closing, seed + (bumps.closing || 0)), {}, style);
      if (!limit || U.countChars(result.text + closing) <= limit) {
        result.text += closing;
        result.parts.push({ cat: 'closing', label: '結び', detail: '', text: closing });
      } else {
        result.unused.push({ cat: 'closing', label: '結び', detail: '文字数が足りないため未使用', text: closing });
      }
    }

    result.length = U.countChars(result.text);
    result.limit = limit;
    result.over = !!(limit && result.length > limit);
    result.mode = mode;
    return result;
  };

  /* ------------------------------------------------------------------
   * 道徳の所見
   *   複数教材をまとめて「大くくりのまとまり」として記述します。
   * ------------------------------------------------------------------ */
  Generator.buildMoral = function (student, settings, mode) {
    var style = (mode === 'youroku') ? 'plain' : 'polite';
    var limit = (mode === 'youroku') ? settings.moralLimitYouroku : settings.moralLimitReport;
    var seed = (student.no || 0);
    var bump = (student.bumps && student.bumps.moral) || 0;

    var parts = [];
    if (settings.moralIntro) {
      parts.push({
        cat: 'moral-intro', label: '導入', detail: '',
        long: T.moral.intro.long, short: T.moral.intro.short,
        vars: { '学期': settings.term || '学期' }, seed: seed + bump
      });
    }
    (student.moral || []).forEach(function (m, i) {
      var material = s(m.material);
      if (!material) return;
      var note = s(m.note);
      var set = note ? T.moral.material : T.moral.material.noNote;
      parts.push({
        cat: 'moral', label: '道徳（教材）',
        detail: material + (m.value ? '／' + m.value : ''),
        long: set.long,
        short: set.short,
        vars: {
          '教材': material,
          '項目句': s(m.value) || '大切なこと',
          '気付き句': note ? '「' + note.replace(/[。]$/, '') + '」' : ''
        },
        seed: seed + i + bump
      });
    });
    if (settings.moralClosing && parts.length) {
      parts.push({
        cat: 'moral-closing', label: '結び', detail: '',
        long: T.moral.closing.long, short: T.moral.closing.short,
        vars: {}, seed: seed + bump
      });
    }

    var result = assemble(parts, limit, style, null);
    result.length = U.countChars(result.text);
    result.limit = limit;
    result.over = !!(limit && result.length > limit);
    result.mode = mode;
    return result;
  };

  /* ------------------------------------------------------------------
   * 総合的な学習の時間の所見
   * ------------------------------------------------------------------ */
  Generator.buildSougou = function (student, settings, mode) {
    var style = (mode === 'youroku') ? 'plain' : 'polite';
    var limit = (mode === 'youroku') ? settings.sougouLimitYouroku : settings.sougouLimitReport;
    var seed = (student.no || 0);
    var bump = (student.bumps && student.bumps.sougou) || 0;

    var parts = [];
    (student.sougou || []).forEach(function (g, i) {
      var unit = s(g.unit);
      if (!unit) return;
      parts.push({
        cat: 'sougou', label: '総合的な学習',
        detail: unit + (g.theme ? '／' + g.theme : ''),
        long: T.sougou.long,
        short: T.sougou.short,
        vars: {
          '単元': unit,
          '課題句': g.theme ? withParticle(g.theme, 'について') : '',
          '活動句': suffix(g.activity, '、'),
          '成果句': suffix(g.result, '、')
        },
        seed: seed + i + bump
      });
    });
    if (settings.sougouClosing && parts.length) {
      parts.push({
        cat: 'sougou-closing', label: '結び', detail: '',
        long: T.sougou.closing.long, short: T.sougou.closing.short,
        vars: {}, seed: seed + bump
      });
    }

    var result = assemble(parts, limit, style, null);
    result.length = U.countChars(result.text);
    result.limit = limit;
    result.over = !!(limit && result.length > limit);
    result.mode = mode;
    return result;
  };

  global.Generator = Generator;
})(window);
