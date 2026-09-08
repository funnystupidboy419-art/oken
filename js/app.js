/* =========================================================================
 * app.js  画面の制御
 *   - 入力（エクセル貼り付け対応）／設定／生成／確認／出力
 *   - 個人情報（氏名など）は扱いません。管理は出席番号のみです。
 * ========================================================================= */
(function () {
  'use strict';

  var U = window.Util;
  var T = window.Templates;
  var G = window.Generator;
  var STORAGE_KEY = 'shoken-app-v1';

  /* =================================================================
   * 状態
   * ================================================================= */
  function defaultSettings() {
    return {
      year: '', grade: '', className: '', term: '2学期', count: 35,
      order: ['observation', 'event', 'committee', 'kakari', 'behavior'],
      enabled: { observation: true, event: true, committee: true, kakari: true, behavior: true },
      behaviorOrder: T.BEHAVIOR_ITEMS.map(function (b) { return b.key; }),
      behaviorMax: 2,
      behaviorPick: 'order',
      limitReport: 120, limitYouroku: 60,
      moralLimitReport: 120, moralLimitYouroku: 60,
      sougouLimitReport: 120, sougouLimitYouroku: 60,
      useClosing: false, moralIntro: true, moralClosing: true, sougouClosing: true,
      warnShort: false, keepEdited: true
    };
  }

  function defaultStudent(no) {
    return {
      no: no,
      committee: '', committeeRole: '', committeeWork: '',
      kakari: '', kakariWork: '',
      events: [emptyEvent(), emptyEvent(), emptyEvent()],
      behavior: {},
      observations: [emptyObs(), emptyObs(), emptyObs()],
      manual: {},
      bumps: {}
    };
  }
  function emptyEvent() { return { name: '', role: '', effort: '' }; }
  function emptyObs() { return { text: '', pattern: 'sugata' }; }

  var state = U.load(STORAGE_KEY) || {};
  state.settings = Object.assign(defaultSettings(), state.settings || {});
  state.students = state.students || {};
  state.moral = state.moral || [];
  state.sougou = state.sougou || [];

  /* 保存し忘れた項目の補完 */
  function student(no) {
    var key = String(no);
    if (!state.students[key]) state.students[key] = defaultStudent(no);
    var st = state.students[key];
    st.no = no;
    if (!st.events) st.events = [emptyEvent(), emptyEvent(), emptyEvent()];
    while (st.events.length < 3) st.events.push(emptyEvent());
    if (!st.observations) st.observations = [emptyObs(), emptyObs(), emptyObs()];
    while (st.observations.length < 3) st.observations.push(emptyObs());
    if (!st.behavior) st.behavior = {};
    if (!st.manual) st.manual = {};
    if (!st.bumps) st.bumps = {};
    return st;
  }

  function numbers() {
    var list = [];
    for (var i = 1; i <= (state.settings.count || 35); i++) list.push(i);
    return list;
  }

  var saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      var ok = U.save(STORAGE_KEY, state);
      var el = document.getElementById('saveState');
      var now = new Date();
      el.textContent = ok
        ? '保存しました ' + now.getHours() + ':' + ('0' + now.getMinutes()).slice(-2)
        : '保存できませんでした（ブラウザの設定をご確認ください）';
    }, 250);
  }

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  /* =================================================================
   * 汎用グリッド（エクセルからの貼り付けに対応）
   * ================================================================= */
  function buildGrid(table, config) {
    function render() {
      var html = '<thead><tr><th>' + (config.cornerLabel || '番号') + '</th>';
      config.columns.forEach(function (col) {
        html += '<th' + (col.width ? ' style="min-width:' + col.width + '"' : '') + '>' +
          U.escapeHtml(col.label) + '</th>';
      });
      if (config.deletable) html += '<th class="col-del">削除</th>';
      html += '</tr></thead><tbody>';

      var rows = config.rowCount();
      for (var r = 0; r < rows; r++) {
        html += '<tr><th>' + U.escapeHtml(config.rowHeader(r)) + '</th>';
        for (var c = 0; c < config.columns.length; c++) {
          var col = config.columns[c];
          var value = config.getCell(r, c);
          html += '<td>';
          if (col.type === 'select') {
            html += '<select data-r="' + r + '" data-c="' + c + '">';
            col.options.forEach(function (op) {
              html += '<option value="' + U.escapeHtml(op.value) + '"' +
                (String(op.value) === String(value) ? ' selected' : '') + '>' +
                U.escapeHtml(op.label) + '</option>';
            });
            html += '</select>';
          } else {
            html += '<input type="text" data-r="' + r + '" data-c="' + c + '"' +
              ' value="' + U.escapeHtml(value) + '"' +
              (col.list ? ' list="' + col.list + '"' : '') +
              (col.placeholder ? ' placeholder="' + U.escapeHtml(col.placeholder) + '"' : '') + '>';
          }
          html += '</td>';
        }
        if (config.deletable) {
          html += '<td class="col-del"><button type="button" class="btn btn--small" data-del="' + r + '">×</button></td>';
        }
        html += '</tr>';
      }
      html += '</tbody>';
      table.innerHTML = html;
    }

    table.addEventListener('input', function (e) {
      var el = e.target;
      if (!el.dataset || el.dataset.r === undefined) return;
      config.setCell(+el.dataset.r, +el.dataset.c, el.value);
      save();
    });
    table.addEventListener('change', function (e) {
      var el = e.target;
      if (el.dataset.r === undefined) return;
      config.setCell(+el.dataset.r, +el.dataset.c, el.value);
      /* 入力した記号を ◎ ○ にそろえて表示に反映する */
      if (config.reflect && el.tagName === 'INPUT') el.value = config.getCell(+el.dataset.r, +el.dataset.c);
      save();
    });
    table.addEventListener('click', function (e) {
      var del = e.target.dataset && e.target.dataset.del;
      if (del === undefined || del === null || del === '') return;
      if (config.onDelete) { config.onDelete(+del); render(); save(); }
    });

    /* エクセルからの貼り付け：選択中のセルを左上として一括入力 */
    table.addEventListener('paste', function (e) {
      var el = e.target;
      if (!el.dataset || el.dataset.r === undefined) return;
      var text = (e.clipboardData || window.clipboardData).getData('text');
      if (!text || (text.indexOf('\t') < 0 && text.indexOf('\n') < 0)) return; /* 単一セルは通常の貼り付け */
      e.preventDefault();
      var rows = U.parseTable(text);
      var r0 = +el.dataset.r, c0 = +el.dataset.c;
      rows.forEach(function (row, i) {
        row.forEach(function (value, j) {
          var r = r0 + i, c = c0 + j;
          if (config.ensureRow) config.ensureRow(r);
          if (r < config.rowCount() && c < config.columns.length) config.setCell(r, c, value);
        });
      });
      render();
      save();
    });

    render();
    return { render: render };
  }

  /* =================================================================
   * 1. 設定
   * ================================================================= */
  function bindSetting(id, key, type) {
    var el = document.getElementById(id);
    if (!el) return;
    if (type === 'check') el.checked = !!state.settings[key];
    else el.value = state.settings[key];
    el.addEventListener('change', function () {
      if (type === 'check') state.settings[key] = el.checked;
      else if (type === 'number') state.settings[key] = parseInt(el.value, 10) || 0;
      else state.settings[key] = el.value;
      save();
      if (key === 'count') renderAllGrids();
    });
    if (type !== 'check') el.addEventListener('input', function () { el.dispatchEvent(new Event('change')); });
  }

  function renderOrderList() {
    var ul = document.getElementById('orderList');
    var labels = {};
    G.CATEGORIES.forEach(function (c) { labels[c.key] = c.label; });
    /* 設定に無いカテゴリを補う */
    G.CATEGORIES.forEach(function (c) {
      if (state.settings.order.indexOf(c.key) < 0) state.settings.order.push(c.key);
    });

    ul.innerHTML = state.settings.order.map(function (key, i) {
      var on = state.settings.enabled[key] !== false;
      return '<li>' +
        '<span class="order-no">' + (i + 1) + '</span>' +
        '<label class="order-name"><input type="checkbox" data-key="' + key + '"' + (on ? ' checked' : '') + '> ' +
        U.escapeHtml(labels[key] || key) + '</label>' +
        '<button type="button" class="btn btn--small" data-up="' + i + '">↑</button>' +
        '<button type="button" class="btn btn--small" data-down="' + i + '">↓</button>' +
        '</li>';
    }).join('');
  }

  function renderBehaviorOrderList() {
    var ul = document.getElementById('behaviorOrderList');
    var labels = {};
    T.BEHAVIOR_ITEMS.forEach(function (b) { labels[b.key] = b.label; });
    T.BEHAVIOR_ITEMS.forEach(function (b) {
      if (state.settings.behaviorOrder.indexOf(b.key) < 0) state.settings.behaviorOrder.push(b.key);
    });
    ul.innerHTML = state.settings.behaviorOrder.map(function (key, i) {
      return '<li>' +
        '<span class="order-no">' + (i + 1) + '</span>' +
        '<span class="order-name">' + U.escapeHtml(labels[key] || key) + '</span>' +
        '<button type="button" class="btn btn--small" data-up="' + i + '">↑</button>' +
        '<button type="button" class="btn btn--small" data-down="' + i + '">↓</button>' +
        '</li>';
    }).join('');
  }

  function moveItem(list, from, to) {
    if (to < 0 || to >= list.length) return;
    var item = list.splice(from, 1)[0];
    list.splice(to, 0, item);
  }

  function setupOrderLists() {
    var ul = document.getElementById('orderList');
    ul.addEventListener('click', function (e) {
      var d = e.target.dataset || {};
      if (d.up !== undefined) moveItem(state.settings.order, +d.up, +d.up - 1);
      else if (d.down !== undefined) moveItem(state.settings.order, +d.down, +d.down + 1);
      else return;
      renderOrderList(); save();
    });
    ul.addEventListener('change', function (e) {
      var key = e.target.dataset && e.target.dataset.key;
      if (!key) return;
      state.settings.enabled[key] = e.target.checked;
      save();
    });

    var bl = document.getElementById('behaviorOrderList');
    bl.addEventListener('click', function (e) {
      var d = e.target.dataset || {};
      if (d.up !== undefined) moveItem(state.settings.behaviorOrder, +d.up, +d.up - 1);
      else if (d.down !== undefined) moveItem(state.settings.behaviorOrder, +d.down, +d.down + 1);
      else return;
      renderBehaviorOrderList();
      grids.behavior.render();
      save();
    });
  }

  /* =================================================================
   * 2〜7. 入力グリッド
   * ================================================================= */
  var grids = {};

  function setupRosterGrid() {
    var cols = [
      { key: 'committee', label: '委員会', width: '9em', placeholder: '例）図書委員会' },
      { key: 'committeeRole', label: '委員会での役割', width: '8em', placeholder: '例）委員長' },
      { key: 'committeeWork', label: '委員会の活動内容', width: '14em', placeholder: '例）本の貸し出し' },
      { key: 'kakari', label: '係', width: '8em', placeholder: '例）黒板係' },
      { key: 'kakariWork', label: '係の仕事', width: '14em', placeholder: '例）授業のたびに黒板を消す仕事' }
    ];
    grids.roster = buildGrid(document.getElementById('gridRoster'), {
      columns: cols,
      rowCount: function () { return numbers().length; },
      rowHeader: function (r) { return String(r + 1); },
      getCell: function (r, c) { return student(r + 1)[cols[c].key] || ''; },
      setCell: function (r, c, v) { student(r + 1)[cols[c].key] = v; }
    });
  }

  function setupEventGrid() {
    var cols = [];
    for (var i = 0; i < 3; i++) {
      cols.push({ ev: i, field: 'name', label: '行事' + (i + 1), width: '9em', placeholder: '例）運動会' });
      cols.push({ ev: i, field: 'role', label: '行事' + (i + 1) + 'の係', width: '8em', placeholder: '例）応援団' });
      cols.push({ ev: i, field: 'effort', label: '行事' + (i + 1) + 'の頑張り', width: '13em', placeholder: '例）大きな声で応援し' });
    }
    grids.event = buildGrid(document.getElementById('gridEvent'), {
      columns: cols,
      rowCount: function () { return numbers().length; },
      rowHeader: function (r) { return String(r + 1); },
      getCell: function (r, c) { return student(r + 1).events[cols[c].ev][cols[c].field] || ''; },
      setCell: function (r, c, v) { student(r + 1).events[cols[c].ev][cols[c].field] = v; }
    });
  }

  function setupBehaviorGrid() {
    /* エクセルからの貼り付けを受け取れるよう、選択式ではなく文字入力にしています。
       ◎ ○ のほか A・B・@・o なども入力できます（確定時に ◎ ○ へそろえます）。 */
    function cols() {
      var labels = {};
      T.BEHAVIOR_ITEMS.forEach(function (b) { labels[b.key] = b.label; });
      return state.settings.behaviorOrder.map(function (key) {
        return { key: key, label: labels[key] || key, width: '5.5em', list: 'markOptions' };
      });
    }
    var current = cols();
    grids.behavior = buildGrid(document.getElementById('gridBehavior'), {
      get columns() { return current; },
      reflect: true,
      rowCount: function () { return numbers().length; },
      rowHeader: function (r) { return String(r + 1); },
      getCell: function (r, c) { return student(r + 1).behavior[current[c].key] || ''; },
      setCell: function (r, c, v) { student(r + 1).behavior[current[c].key] = U.normalizeMark(v); }
    });
    var baseRender = grids.behavior.render;
    grids.behavior.render = function () { current = cols(); baseRender(); };
  }

  function setupObservationGrid() {
    var patternOptions = T.observationPatterns.map(function (p) {
      return { value: p.key, label: p.label };
    });
    var cols = [];
    for (var i = 0; i < 3; i++) {
      cols.push({ ob: i, field: 'text', label: '見取り' + (i + 1), width: '22em', placeholder: T.observationPatterns[0].hint });
      cols.push({ ob: i, field: 'pattern', label: '文型' + (i + 1), type: 'select', options: patternOptions, width: '13em' });
    }
    grids.observation = buildGrid(document.getElementById('gridObservation'), {
      columns: cols,
      rowCount: function () { return numbers().length; },
      rowHeader: function (r) { return String(r + 1); },
      getCell: function (r, c) { return student(r + 1).observations[cols[c].ob][cols[c].field] || ''; },
      setCell: function (r, c, v) {
        var ob = student(r + 1).observations[cols[c].ob];
        ob[cols[c].field] = v;
        if (cols[c].field === 'text' && !ob.pattern) ob.pattern = 'sugata';
      }
    });
  }

  function setupMoralGrid() {
    var cols = [
      { key: 'no', label: '出席番号', width: '5em' },
      { key: 'material', label: '教材名', width: '14em', placeholder: '例）ないた赤おに' },
      { key: 'value', label: '内容項目', width: '13em', list: 'moralValues', placeholder: '例）友情、信頼' },
      { key: 'note', label: '児童の振り返り・気付き', width: '26em', placeholder: '例）本当の友達は相手のことを考えられる人だ' }
    ];
    grids.moral = buildGrid(document.getElementById('gridMoral'), {
      cornerLabel: '行',
      columns: cols,
      deletable: true,
      rowCount: function () { return state.moral.length; },
      rowHeader: function (r) { return String(r + 1); },
      ensureRow: function (r) { while (state.moral.length <= r) state.moral.push({ no: '', material: '', value: '', note: '' }); },
      getCell: function (r, c) { return (state.moral[r] || {})[cols[c].key] || ''; },
      setCell: function (r, c, v) {
        if (!state.moral[r]) state.moral[r] = { no: '', material: '', value: '', note: '' };
        state.moral[r][cols[c].key] = (cols[c].key === 'no') ? (U.toNumber(v) || '') : v;
      },
      onDelete: function (r) { state.moral.splice(r, 1); }
    });
  }

  function setupSougouGrid() {
    var cols = [
      { key: 'no', label: '出席番号', width: '5em' },
      { key: 'unit', label: '単元名', width: '14em', placeholder: '例）大豆はかせになろう' },
      { key: 'theme', label: '探究課題', width: '13em', placeholder: '例）大豆の育ち方' },
      { key: 'activity', label: '活動の様子', width: '20em', placeholder: '例）農家の方に聞き取りを行い' },
      { key: 'result', label: '成果・まとめ', width: '20em', placeholder: '例）分かったことを新聞にまとめ' }
    ];
    grids.sougou = buildGrid(document.getElementById('gridSougou'), {
      cornerLabel: '行',
      columns: cols,
      deletable: true,
      rowCount: function () { return state.sougou.length; },
      rowHeader: function (r) { return String(r + 1); },
      ensureRow: function (r) {
        while (state.sougou.length <= r) state.sougou.push({ no: '', unit: '', theme: '', activity: '', result: '' });
      },
      getCell: function (r, c) { return (state.sougou[r] || {})[cols[c].key] || ''; },
      setCell: function (r, c, v) {
        if (!state.sougou[r]) state.sougou[r] = { no: '', unit: '', theme: '', activity: '', result: '' };
        state.sougou[r][cols[c].key] = (cols[c].key === 'no') ? (U.toNumber(v) || '') : v;
      },
      onDelete: function (r) { state.sougou.splice(r, 1); }
    });
  }

  function renderAllGrids() {
    Object.keys(grids).forEach(function (k) { grids[k].render(); });
  }

  /* =================================================================
   * まとめて貼り付け（見出し行を読み取って対応づけ）
   * ================================================================= */
  var HEADER_MAP = {
    roster: [
      { key: 'committee', words: ['委員会'] },
      { key: 'committeeRole', words: ['役割', '役職', '委員会での役割'] },
      { key: 'committeeWork', words: ['活動', '活動内容'] },
      { key: 'kakari', words: ['係'] },
      { key: 'kakariWork', words: ['仕事', '係の仕事'] }
    ]
  };

  function applyBulkPaste(kind, text) {
    var rows = U.parseTable(text);
    if (!rows.length) return 0;

    /* 1行目が見出しかどうか判定（1列目が数字でなければ見出しとみなす） */
    var hasHeader = U.toNumber(rows[0][0]) === null;
    var header = hasHeader ? rows[0] : null;
    var body = hasHeader ? rows.slice(1) : rows;
    var applied = 0;

    if (kind === 'roster') {
      var mapping = HEADER_MAP.roster.map(function (m, i) { return { key: m.key, col: i + 1 }; });
      if (header) {
        mapping = [];
        HEADER_MAP.roster.forEach(function (m) {
          for (var c = 1; c < header.length; c++) {
            var h = header[c];
            for (var w = 0; w < m.words.length; w++) {
              if (h.indexOf(m.words[w]) >= 0 && !mapping.some(function (x) { return x.col === c; })) {
                mapping.push({ key: m.key, col: c });
                return;
              }
            }
          }
        });
        if (!mapping.length) mapping = HEADER_MAP.roster.map(function (m, i) { return { key: m.key, col: i + 1 }; });
      }
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no) return;
        var st = student(no);
        mapping.forEach(function (m) { if (row[m.col] !== undefined) st[m.key] = row[m.col]; });
        applied++;
      });

    } else if (kind === 'event') {
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no) return;
        var st = student(no);
        for (var i = 0; i < 3; i++) {
          var base = 1 + i * 3;
          if (row[base] !== undefined) st.events[i].name = row[base];
          if (row[base + 1] !== undefined) st.events[i].role = row[base + 1];
          if (row[base + 2] !== undefined) st.events[i].effort = row[base + 2];
        }
        applied++;
      });

    } else if (kind === 'behavior') {
      /* 見出しがあれば項目名で対応づけ、無ければ設定の並び順のとおり */
      var order = state.settings.behaviorOrder.slice();
      var colKey = {};
      if (header) {
        for (var c = 1; c < header.length; c++) {
          var found = null;
          T.BEHAVIOR_ITEMS.forEach(function (b) {
            var short = b.label.split('・')[0];
            if (header[c] && (header[c].indexOf(b.label) >= 0 || header[c].indexOf(short) >= 0)) found = b.key;
          });
          if (found) colKey[c] = found;
        }
      }
      if (!Object.keys(colKey).length) {
        order.forEach(function (key, i) { colKey[i + 1] = key; });
      }
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no) return;
        var st = student(no);
        Object.keys(colKey).forEach(function (c) {
          if (row[c] !== undefined) st.behavior[colKey[c]] = U.normalizeMark(row[c]);
        });
        applied++;
      });

    } else if (kind === 'observation') {
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no) return;
        var st = student(no);
        for (var i = 0; i < 3; i++) {
          if (row[i + 1] !== undefined) {
            st.observations[i].text = row[i + 1];
            if (!st.observations[i].pattern) st.observations[i].pattern = 'sugata';
          }
        }
        applied++;
      });

    } else if (kind === 'moral') {
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no || !row[1]) return;
        state.moral.push({ no: no, material: row[1] || '', value: row[2] || '', note: row[3] || '' });
        applied++;
      });

    } else if (kind === 'sougou') {
      body.forEach(function (row) {
        var no = U.toNumber(row[0]);
        if (!no || !row[1]) return;
        state.sougou.push({ no: no, unit: row[1] || '', theme: row[2] || '', activity: row[3] || '', result: row[4] || '' });
        applied++;
      });
    }
    return applied;
  }

  function setupBulkPaste() {
    $$('[data-paste-apply]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var kind = btn.dataset.pasteApply;
        var area = $('[data-paste="' + kind + '"]');
        var text = area.value;
        if (!text.replace(/\s/g, '')) { alert('貼り付ける内容がありません。'); return; }
        var hits = U.findPersonalInfo(text);
        if (hits.length && !confirm(
          '氏名など個人情報らしい表現が含まれています（例：' + hits[0].word + '）。\n' +
          'このアプリは出席番号だけで管理します。それでも取り込みますか？')) return;
        var n = applyBulkPaste(kind, text);
        renderAllGrids();
        save();
        alert(n + ' 件を取り込みました。');
        area.value = '';
      });
    });
  }

  /* =================================================================
   * 8. 所見の確認・出力
   * ================================================================= */
  var BLOCKS = [
    { id: 'report', title: '通信表の所見（行動・活動・見取り）', style: '敬体（です・ます）', build: 'build', mode: 'report' },
    { id: 'youroku', title: '指導要録の所見（行動・活動・見取り）', style: '常体（だ・である）', build: 'build', mode: 'youroku' },
    { id: 'moralReport', title: '道徳の所見（通信表）', style: '敬体（です・ます）', build: 'buildMoral', mode: 'report' },
    { id: 'moralYouroku', title: '道徳の所見（指導要録）', style: '常体（だ・である）', build: 'buildMoral', mode: 'youroku' },
    { id: 'sougouReport', title: '総合的な学習の所見（通信表）', style: '敬体（です・ます）', build: 'buildSougou', mode: 'report' },
    { id: 'sougouYouroku', title: '総合的な学習の所見（指導要録）', style: '常体（だ・である）', build: 'buildSougou', mode: 'youroku' }
  ];

  /* 生成用に、教材・単元の行をその子のものだけ集めたビューを作る */
  function studentView(no) {
    var st = student(no);
    var view = Object.create(st);
    view.moral = state.moral.filter(function (m) { return U.toNumber(m.no) === no; });
    view.sougou = state.sougou.filter(function (g) { return U.toNumber(g.no) === no; });
    return view;
  }

  function generateFor(no) {
    var view = studentView(no);
    var out = {};
    BLOCKS.forEach(function (b) {
      out[b.id] = G[b.build](view, state.settings, b.mode);
    });
    return out;
  }

  function textOf(st, block, generated) {
    var manual = st.manual[block.id];
    return (manual !== undefined && manual !== null) ? manual : generated[block.id].text;
  }

  function renderReview() {
    var listEl = document.getElementById('reviewList');
    var filter = document.getElementById('reviewFilter').value;
    var html = '';
    var counters = { total: 0, over: 0, warn: 0, empty: 0, edited: 0 };

    numbers().forEach(function (no) {
      var st = student(no);
      var gen = generateFor(no);
      var cardIssues = [];
      var body = '';
      var anyText = false;

      BLOCKS.forEach(function (b) {
        var g = gen[b.id];
        var text = textOf(st, b, gen);
        var len = U.countChars(text);
        var limit = g.limit;
        var edited = st.manual[b.id] !== undefined && st.manual[b.id] !== null;
        if (text) anyText = true;
        if (edited) counters.edited++;

        var cls = '', note = '';
        if (limit && len > limit) { cls = 'is-over'; note = '（上限超過）'; cardIssues.push(b.title + 'が文字数超過'); }
        else if (state.settings.warnShort && limit && text && len < Math.floor(limit * 0.6)) {
          cls = 'is-short'; note = '（短め）';
        }

        var hits = U.findPersonalInfo(text);
        if (hits.length) cardIssues.push(b.title + 'に個人情報らしい表現：' + hits[0].word);

        body += '<div class="shoken-block" data-block="' + b.id + '">' +
          '<div class="shoken-block__title">' + U.escapeHtml(b.title) +
          ' <span class="style-note">' + U.escapeHtml(b.style) + '</span>' +
          (edited ? ' <span class="edited-badge">手直し済み</span>' : '') +
          '<span class="counter ' + cls + '">' + len + ' / ' + (limit || '—') + ' 字' + note + '</span>' +
          '</div>' +
          '<textarea rows="3" data-no="' + no + '" data-field="' + b.id + '">' + U.escapeHtml(text) + '</textarea>' +
          '<div class="field-row">' +
          '<button type="button" class="btn btn--small" data-regen="' + b.id + '" data-no="' + no + '">生成しなおす</button>' +
          '<button type="button" class="btn btn--small" data-alt="' + b.id + '" data-no="' + no + '">別の文例にする</button>' +
          '<button type="button" class="btn btn--small" data-copy="' + b.id + '" data-no="' + no + '">コピー</button>' +
          '</div>' +
          breakdownHtml(g) +
          '</div>';
      });

      if (!anyText) counters.empty++;
      if (cardIssues.length) counters.warn++;
      counters.total++;

      var over = cardIssues.some(function (m) { return m.indexOf('文字数超過') >= 0; });
      if (over) counters.over++;

      if (filter === 'over' && !over) return;
      if (filter === 'warn' && !cardIssues.length) return;
      if (filter === 'empty' && anyText) return;

      html += '<article class="student-card">' +
        '<div class="student-card__head">' +
        '<span class="student-card__no">出席番号 ' + no + '</span>' +
        '<span class="student-card__tags">' + U.escapeHtml(tagsOf(st, no)) + '</span>' +
        '</div>' +
        '<div class="student-card__body">' +
        (cardIssues.length ? '<p class="alert alert--warn">要確認：' + U.escapeHtml(cardIssues.join(' ／ ')) + '</p>' : '') +
        body +
        '</div></article>';
    });

    listEl.innerHTML = html || '<p class="hint">該当する児童がいません。</p>';
    document.getElementById('reviewSummary').innerHTML =
      '対象 ' + counters.total + ' 人／文字数超過 ' + counters.over + ' 件／要確認 ' + counters.warn +
      ' 人／未作成 ' + counters.empty + ' 人／手直し済み ' + counters.edited + ' 箇所';
  }

  function tagsOf(st, no) {
    var tags = [];
    if (st.committee) tags.push('委員会:' + st.committee + (st.committeeRole ? '(' + st.committeeRole + ')' : ''));
    if (st.kakari) tags.push('係:' + st.kakari);
    st.events.forEach(function (ev) { if (ev.name) tags.push('行事:' + ev.name + (ev.role ? '(' + ev.role + ')' : '')); });
    var marks = Object.keys(st.behavior || {}).filter(function (k) { return st.behavior[k]; });
    if (marks.length) tags.push('行動:' + marks.length + '項目');
    var mo = state.moral.filter(function (m) { return U.toNumber(m.no) === no; }).length;
    if (mo) tags.push('道徳:' + mo + '教材');
    var so = state.sougou.filter(function (g) { return U.toNumber(g.no) === no; }).length;
    if (so) tags.push('総合:' + so + '単元');
    return tags.join('　');
  }

  function breakdownHtml(g) {
    var items = (g.parts || []).map(function (p) {
      return '<li><span class="tag">' + U.escapeHtml(p.label + (p.detail ? '：' + p.detail : '')) + '</span>' +
        U.escapeHtml(p.text) + '</li>';
    });
    (g.unused || []).forEach(function (p) {
      items.push('<li><span class="tag tag--unused">文字数の都合で未使用｜' +
        U.escapeHtml(p.label + (p.detail ? '：' + p.detail : '')) + '</span>' + U.escapeHtml(p.text) + '</li>');
    });
    if (!items.length) return '';
    return '<ul class="breakdown">' + items.join('') + '</ul>';
  }

  function setupReview() {
    var listEl = document.getElementById('reviewList');

    listEl.addEventListener('input', function (e) {
      var el = e.target;
      if (el.tagName !== 'TEXTAREA' || !el.dataset.field) return;
      var st = student(+el.dataset.no);
      st.manual[el.dataset.field] = el.value;
      var counter = el.parentNode.querySelector('.counter');
      var block = BLOCKS.filter(function (b) { return b.id === el.dataset.field; })[0];
      if (counter && block) {
        var limit = limitOf(block);
        var len = U.countChars(el.value);
        counter.textContent = len + ' / ' + (limit || '—') + ' 字' + (limit && len > limit ? '（上限超過）' : '');
        counter.className = 'counter' + (limit && len > limit ? ' is-over' : '');
      }
      save();
    });

    listEl.addEventListener('click', function (e) {
      var d = e.target.dataset || {};
      if (d.regen) {
        var st = student(+d.no);
        delete st.manual[d.regen];
        save(); renderReview();
      } else if (d.alt) {
        var st2 = student(+d.no);
        bump(st2, d.alt);
        delete st2.manual[d.alt];
        save(); renderReview();
      } else if (d.copy) {
        var ta = e.target.closest('.shoken-block').querySelector('textarea');
        ta.select();
        try { document.execCommand('copy'); } catch (err) { /* noop */ }
        e.target.textContent = 'コピーしました';
        setTimeout(function () { e.target.textContent = 'コピー'; }, 1200);
      }
    });

    document.getElementById('reviewFilter').addEventListener('change', renderReview);

    document.getElementById('btnGenerateAll').addEventListener('click', function () {
      var keep = document.getElementById('keepEdited').checked;
      if (!keep && !confirm('手直しした所見も含めて、すべて作り直します。よろしいですか？')) return;
      numbers().forEach(function (no) {
        var st = student(no);
        if (!keep) st.manual = {};
      });
      save();
      renderReview();
    });

    document.getElementById('keepEdited').checked = state.settings.keepEdited !== false;
    document.getElementById('keepEdited').addEventListener('change', function () {
      state.settings.keepEdited = this.checked;
      save();
    });

    document.getElementById('btnPrint').addEventListener('click', function () { window.print(); });
    document.getElementById('btnExportXlsx').addEventListener('click', exportXlsx);
    document.getElementById('btnExportCsv').addEventListener('click', exportCsv);
  }

  function limitOf(block) {
    var s = state.settings;
    return {
      report: s.limitReport, youroku: s.limitYouroku,
      moralReport: s.moralLimitReport, moralYouroku: s.moralLimitYouroku,
      sougouReport: s.sougouLimitReport, sougouYouroku: s.sougouLimitYouroku
    }[block.id];
  }

  function bump(st, blockId) {
    if (blockId === 'moralReport' || blockId === 'moralYouroku') {
      st.bumps.moral = (st.bumps.moral || 0) + 1;
    } else if (blockId === 'sougouReport' || blockId === 'sougouYouroku') {
      st.bumps.sougou = (st.bumps.sougou || 0) + 1;
    } else {
      ['committee', 'kakari', 'event', 'behavior', 'observation', 'closing'].forEach(function (cat) {
        st.bumps[cat] = (st.bumps[cat] || 0) + 1;
      });
    }
  }

  /* =================================================================
   * 出力
   * ================================================================= */
  function collectAll() {
    return numbers().map(function (no) {
      var st = student(no);
      var gen = generateFor(no);
      var row = { no: no, gen: gen, st: st, text: {} };
      BLOCKS.forEach(function (b) { row.text[b.id] = textOf(st, b, gen); });
      return row;
    });
  }

  function fileBase() {
    var s = state.settings;
    return ['所見', s.year, s.grade, s.className, s.term].filter(Boolean).join('_') + '_' + U.todayString();
  }

  function buildSheets() {
    var all = collectAll();
    var s = state.settings;

    var reportSheet = {
      name: '通信表所見（敬体）',
      cols: [6, 52, 7, 46, 7, 46, 7],
      wrapFrom: 1,
      rows: [['番号', '所見（行動・活動・見取り）', '字数', '道徳', '字数', '総合的な学習', '字数']]
    };
    var yourokuSheet = {
      name: '指導要録所見（常体）',
      cols: [6, 44, 7, 40, 7, 40, 7],
      wrapFrom: 1,
      rows: [['番号', '所見（行動・活動・見取り）', '字数', '道徳', '字数', '総合的な学習', '字数']]
    };
    var checkSheet = {
      name: 'チェック一覧',
      cols: [6, 26, 7, 7, 10, 34],
      wrapFrom: 5,
      rows: [['番号', '所見の種類', '字数', '上限', '状態', '注意']]
    };
    var inputSheet = {
      name: '入力データ',
      cols: [6, 12, 10, 16, 10, 16, 12, 10, 12, 10, 12, 10, 24, 24, 24],
      wrapFrom: 12,
      rows: [[
        '番号', '委員会', '委員会の役割', '委員会の活動', '係', '係の仕事',
        '行事1', '行事1の係', '行事2', '行事2の係', '行事3', '行事3の係',
        '見取り1', '見取り2', '見取り3'
      ].concat(s.behaviorOrder.map(function (k) {
        var it = T.BEHAVIOR_ITEMS.filter(function (b) { return b.key === k; })[0];
        return it ? it.label : k;
      }))]
    };
    var moralSheet = {
      name: '道徳（教材別）',
      cols: [6, 20, 20, 46],
      wrapFrom: 3,
      rows: [['番号', '教材名', '内容項目', '児童の振り返り']]
    };
    var sougouSheet = {
      name: '総合（単元別）',
      cols: [6, 20, 18, 28, 28],
      wrapFrom: 2,
      rows: [['番号', '単元名', '探究課題', '活動の様子', '成果・まとめ']]
    };

    all.forEach(function (row) {
      reportSheet.rows.push([
        row.no,
        row.text.report, U.countChars(row.text.report),
        row.text.moralReport, U.countChars(row.text.moralReport),
        row.text.sougouReport, U.countChars(row.text.sougouReport)
      ]);
      yourokuSheet.rows.push([
        row.no,
        row.text.youroku, U.countChars(row.text.youroku),
        row.text.moralYouroku, U.countChars(row.text.moralYouroku),
        row.text.sougouYouroku, U.countChars(row.text.sougouYouroku)
      ]);

      BLOCKS.forEach(function (b) {
        var text = row.text[b.id];
        var limit = limitOf(b);
        var len = U.countChars(text);
        var notes = [];
        if (!text) notes.push('未作成');
        if (limit && len > limit) notes.push('文字数超過');
        if (state.settings.warnShort && text && limit && len < Math.floor(limit * 0.6)) notes.push('文字数が少なめ');
        U.findPersonalInfo(text).forEach(function (h) { notes.push('個人情報らしい表現：' + h.word); });
        if (row.st.manual[b.id] !== undefined && row.st.manual[b.id] !== null) notes.push('手直し済み');
        checkSheet.rows.push([
          row.no, b.title, len, limit || '', notes.length ? '要確認' : 'OK', notes.join(' / ')
        ]);
      });

      var st = row.st;
      inputSheet.rows.push([
        row.no, st.committee, st.committeeRole, st.committeeWork, st.kakari, st.kakariWork,
        st.events[0].name, st.events[0].role, st.events[1].name, st.events[1].role,
        st.events[2].name, st.events[2].role,
        st.observations[0].text, st.observations[1].text, st.observations[2].text
      ].concat(s.behaviorOrder.map(function (k) { return st.behavior[k] || ''; })));
    });

    state.moral.forEach(function (m) {
      if (!m.no && !m.material) return;
      moralSheet.rows.push([U.toNumber(m.no) || '', m.material, m.value, m.note]);
    });
    state.sougou.forEach(function (g) {
      if (!g.no && !g.unit) return;
      sougouSheet.rows.push([U.toNumber(g.no) || '', g.unit, g.theme, g.activity, g.result]);
    });

    return [reportSheet, yourokuSheet, checkSheet, inputSheet, moralSheet, sougouSheet];
  }

  function exportXlsx() {
    try {
      var blob = window.Xlsx.build(buildSheets());
      U.download(fileBase() + '.xlsx', blob);
    } catch (e) {
      alert('エクセルの書き出しに失敗しました：' + e.message);
    }
  }

  function exportCsv() {
    var all = collectAll();
    var rows = [['番号', '通信表所見', '字数', '通信表_道徳', '通信表_総合', '指導要録所見', '字数', '要録_道徳', '要録_総合']];
    all.forEach(function (r) {
      rows.push([
        r.no,
        r.text.report, U.countChars(r.text.report), r.text.moralReport, r.text.sougouReport,
        r.text.youroku, U.countChars(r.text.youroku), r.text.moralYouroku, r.text.sougouYouroku
      ]);
    });
    U.download(fileBase() + '.csv', window.Xlsx.toCsv(rows));
  }

  /* =================================================================
   * データの書き出し・読み込み・消去
   * ================================================================= */
  function setupDataButtons() {
    document.getElementById('btnExportJson').addEventListener('click', function () {
      var blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
      U.download(fileBase() + '_入力データ.json', blob);
    });

    document.getElementById('fileImportJson').addEventListener('change', function () {
      var file = this.files && this.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var data = JSON.parse(reader.result);
          if (!data || typeof data !== 'object') throw new Error('形式が違います');
          if (!confirm('読み込むと、今の入力内容は置き換わります。よろしいですか？')) return;
          state = data;
          state.settings = Object.assign(defaultSettings(), state.settings || {});
          state.students = state.students || {};
          state.moral = state.moral || [];
          state.sougou = state.sougou || [];
          save();
          location.reload();
        } catch (e) {
          alert('読み込めませんでした：' + e.message);
        }
      };
      reader.readAsText(file);
      this.value = '';
    });

    document.getElementById('btnClearAll').addEventListener('click', function () {
      if (!confirm('入力したデータをすべて消します。元に戻せません。よろしいですか？')) return;
      if (!confirm('本当に消してよろしいですか？')) return;
      U.clear(STORAGE_KEY);
      location.reload();
    });

    document.getElementById('btnAddMoral').addEventListener('click', function () {
      state.moral.push({ no: '', material: '', value: '', note: '' });
      grids.moral.render(); save();
    });
    document.getElementById('btnAddMoralAll').addEventListener('click', function () {
      var name = prompt('全員分の行を追加します。教材名を入力してください。');
      if (!name) return;
      var value = prompt('内容項目（省略できます）') || '';
      numbers().forEach(function (no) {
        state.moral.push({ no: no, material: name, value: value, note: '' });
      });
      grids.moral.render(); save();
    });
    document.getElementById('btnAddSougou').addEventListener('click', function () {
      state.sougou.push({ no: '', unit: '', theme: '', activity: '', result: '' });
      grids.sougou.render(); save();
    });
    document.getElementById('btnAddSougouAll').addEventListener('click', function () {
      var name = prompt('全員分の行を追加します。単元名を入力してください。');
      if (!name) return;
      numbers().forEach(function (no) {
        state.sougou.push({ no: no, unit: name, theme: '', activity: '', result: '' });
      });
      grids.sougou.render(); save();
    });
  }

  /* =================================================================
   * タブ
   * ================================================================= */
  function setupTabs() {
    document.getElementById('tabs').addEventListener('click', function (e) {
      var tab = e.target.dataset && e.target.dataset.tab;
      if (!tab) return;
      $$('.tab').forEach(function (t) { t.classList.toggle('is-active', t.dataset.tab === tab); });
      $$('.panel').forEach(function (p) { p.classList.toggle('is-active', p.dataset.panel === tab); });
      if (tab === 'review') renderReview();
      window.scrollTo(0, 0);
    });
  }

  /* =================================================================
   * 起動
   * ================================================================= */
  function init() {
    /* 内容項目の入力補助 */
    document.getElementById('moralValues').innerHTML =
      T.MORAL_VALUES.map(function (v) { return '<option value="' + U.escapeHtml(v) + '">'; }).join('');

    bindSetting('setYear', 'year');
    bindSetting('setGrade', 'grade');
    bindSetting('setClass', 'className');
    bindSetting('setTerm', 'term');
    bindSetting('setCount', 'count', 'number');
    bindSetting('limitReport', 'limitReport', 'number');
    bindSetting('limitYouroku', 'limitYouroku', 'number');
    bindSetting('moralLimitReport', 'moralLimitReport', 'number');
    bindSetting('moralLimitYouroku', 'moralLimitYouroku', 'number');
    bindSetting('sougouLimitReport', 'sougouLimitReport', 'number');
    bindSetting('sougouLimitYouroku', 'sougouLimitYouroku', 'number');
    bindSetting('behaviorMax', 'behaviorMax', 'number');
    bindSetting('behaviorPick', 'behaviorPick');
    bindSetting('useClosing', 'useClosing', 'check');
    bindSetting('moralIntro', 'moralIntro', 'check');
    bindSetting('moralClosing', 'moralClosing', 'check');
    bindSetting('sougouClosing', 'sougouClosing', 'check');
    bindSetting('warnShort', 'warnShort', 'check');

    renderOrderList();
    renderBehaviorOrderList();
    setupOrderLists();

    setupRosterGrid();
    setupEventGrid();
    setupBehaviorGrid();
    setupObservationGrid();
    setupMoralGrid();
    setupSougouGrid();

    setupBulkPaste();
    setupReview();
    setupDataButtons();
    setupTabs();

    if (!state.moral.length) state.moral.push({ no: '', material: '', value: '', note: '' });
    if (!state.sougou.length) state.sougou.push({ no: '', unit: '', theme: '', activity: '', result: '' });
    grids.moral.render();
    grids.sougou.render();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
