/*
 * 温度モニター MyGeotab Add-In
 * Didcom TPM+ で設定した BLE 温度センサー（診断項目 "Remote probe 1〜4 temperature"）の
 * StatusData を MyGeotab API から読み、車両一覧・温度推移・逸脱履歴を表示する。
 * MyGeotab 外（ローカルでファイルを開いた場合）はデモデータで動作する。
 */
/* global geotab */
var geotab = window.geotab || { addin: {} };
geotab.addin = geotab.addin || {};

geotab.addin.didcomTemp = function () {
  'use strict';

  var PROFILES = {
    frozen: { min: -25, max: -18 },
    chilled: { min: 0, max: 10 },
    constant: { min: 15, max: 25 }
  };
  var DEFAULTS = { profile: 'frozen', min: -25, max: -18, minExcursionMin: 5, staleMinutes: 30, diagnostics: [], doorDiagnostics: [], doorInvert: false,
    riseRate: 3, riseWindowMin: 60 };   // 急上昇：過去60分の変化が +3℃/時 以上
  var RISE_COLOR = 'rgba(234,88,12,.35)';
  var DOOR_COLOR = 'rgba(234,179,8,.28)', DOOR_LOOKBACK_MS = 15 * 60000; // 逸脱開始の15分前までのドア開を関連ありとみなす
  var COLORS = ['#2563eb', '#ea580c', '#16a34a', '#9333ea'];
  var AUTO_MS = 5 * 60 * 1000;
  var VERSION = '0.10.3';

  var api = null, demo = false, storageKey = 'didcomTemp:demo';
  var settings = Object.assign({}, DEFAULTS);
  var data = null;            // { vehicles, diagnostics, series: {devId: {diagId: [{t, v}]}} }
  var selectedId = null, timer = null, loading = false;
  var root, $ = function (id) { return document.getElementById(id); };

  // ---------- ユーティリティ ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmtT(d) { return d ? d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '-'; }
  function fmtDur(ms) { var m = Math.round(ms / 60000); return m < 60 ? m + '分' : Math.floor(m / 60) + '時間' + (m % 60) + '分'; }
  function fmtV(v) { return v == null ? '-' : v.toFixed(1) + '℃'; }
  function ago(d) { if (!d) return '-'; var m = Math.round((Date.now() - d) / 60000); return m < 1 ? 'たった今' : m < 60 ? m + '分前' : m < 1440 ? Math.floor(m / 60) + '時間前' : Math.floor(m / 1440) + '日前'; }
  // 「Remote probe 1 temperature」/「リモートプローブ1温度」など英語・日本語どちらの名称にも対応
  var PROBE_RE = /(remote\s*probe|リモート\s*プローブ|プローブ)\s*#?\s*([1-4])/i;
  function isProbeTemp(name) { return PROBE_RE.test(name || '') && /(temperature|温度)/i.test(name || '') && !/(error|alarm|open|short|エラー|アラーム|断線|短絡)/i.test(name || ''); }
  function sensorLabel(name) {
    var m = PROBE_RE.exec(name || '') || /RemoteProbe([1-4])TemperatureId/.exec(name || '');
    return m ? '温度' + m[m.length - 1] : name;
  }
  function call(method, params) { return new Promise(function (res, rej) { api.call(method, params, res, rej); }); }
  function multi(calls) { return new Promise(function (res, rej) { api.multiCall(calls, res, rej); }); }
  function showMsg(text, info) { var el = $('dt-message'); el.textContent = text; el.className = 'dt-message' + (info ? ' info' : ''); el.hidden = !text; }

  function loadSettings() {
    try { var s = JSON.parse(localStorage.getItem(storageKey)); if (s) settings = Object.assign({}, DEFAULTS, s); } catch (e) { /* ignore */ }
  }
  function saveSettings() { try { localStorage.setItem(storageKey, JSON.stringify(settings)); } catch (e) { /* ignore */ } }

  // ---------- データ取得（ライブ） ----------
  function findDiagnostics(query) {
    var names = query ? ['%' + query + '%'] : ['%プローブ%', '%probe%'];
    return multi(names.map(function (n) { return ['Get', { typeName: 'Diagnostic', search: { name: n }, resultsLimit: 200 }]; }))
      .then(function (results) {
        var seen = {}, out = [];
        results.forEach(function (list) { list.forEach(function (d) { if (!seen[d.id]) { seen[d.id] = 1; out.push({ id: d.id, name: d.name }); } }); });
        return out.sort(function (a, b) { return a.name.localeCompare(b.name); });
      });
  }

  // GeotabのKnownId（言語に依存しない固定ID）で検出。失敗時は名称検索にフォールバック。
  var PROBE_IDS = ['RemoteProbe1TemperatureId', 'RemoteProbe2TemperatureId', 'RemoteProbe3TemperatureId', 'RemoteProbe4TemperatureId'];
  var detectInfo = '';
  function detectProbes() {
    var byId = Promise.all(PROBE_IDS.map(function (id) {
      return call('Get', { typeName: 'Diagnostic', search: { id: id } })
        .then(function (list) { return list && list.length ? { id: list[0].id || id, name: list[0].name || id } : null; },
              function () { return null; });
    })).then(function (r) { return r.filter(Boolean); });
    return byId.then(function (found) {
      if (found.length) { detectInfo = '固定ID ' + found.length + '件'; return found; }
      return findDiagnostics('プローブ').then(function (list) {
        var hit = list.filter(function (d) { return isProbeTemp(d.name); });
        detectInfo = '固定ID 0件 / 名称検索 ' + list.length + '件中 ' + hit.length + '件該当';
        return hit;
      }, function (e) { detectInfo = '固定ID 0件 / 名称検索エラー：' + (e && e.message || e); return []; });
    });
  }

  // TPM+の接点センサー（診断「Contact sensor 1/2」）を検出。固定IDは推定のため、名称検索でも補完。
  var DOOR_IDS = ['ContactSensor1Id', 'ContactSensor2Id'];
  var DOOR_RE = /(contact\s*sensor|コンタクト\s*センサー?|接触\s*センサー?|接点\s*センサー?)\s*#?\s*([12])/i;
  function isDoorDiag(name) { return DOOR_RE.test(name || '') && !/(error|alarm|fault|battery|障害|エラー|アラーム|切断|電池|バッテリ)/i.test(name || ''); }
  function detectDoors() {
    return Promise.all(DOOR_IDS.map(function (id) {
      return call('Get', { typeName: 'Diagnostic', search: { id: id } })
        .then(function (l) { return l && l.length ? { id: l[0].id || id, name: l[0].name || id } : null; }, function () { return null; });
    })).then(function (found) {
      found = found.filter(Boolean);
      if (found.length) return found;
      return multi(['%Contact sensor%', '%コンタクト%', '%接触センサー%', '%接点センサー%'].map(function (n) {
        return ['Get', { typeName: 'Diagnostic', search: { name: n }, resultsLimit: 200 }];
      })).then(function (results) {
        var seen = {}, out = [];
        results.forEach(function (l) { l.forEach(function (d) { if (!seen[d.id] && isDoorDiag(d.name)) { seen[d.id] = 1; out.push({ id: d.id, name: d.name }); } }); });
        return out;
      }, function () { return []; });
    });
  }

  function loadLive(hours) {
    var ready = settings.diagnostics.length ? Promise.resolve() :
      detectProbes().then(function (list) {
        // 初回: リモートプローブ1〜4温度（KnownId: RemoteProbeNTemperatureId）を自動選択
        settings.diagnostics = list;
        if (list.length) saveSettings();
      });
    // ドアセンサーは未設定かつ一度も自動設定していない場合のみ自動検出（利用者が外した選択は尊重）
    if (!(settings.doorDiagnostics || []).length && !settings.doorAutoDone) {
      ready = ready.then(detectDoors).then(function (list) {
        if (list.length) { settings.doorDiagnostics = list; settings.doorAutoDone = true; saveSettings(); }
      });
    }
    return ready.then(function () {
      if (!settings.diagnostics.length) { var e = new Error('温度センサーの診断項目が自動検出できませんでした（' + detectInfo + '）。「設定」から対象の診断項目を検索して選んでください。'); e.info = true; throw e; }
      var now = new Date(), from = new Date(now - hours * 3600e3);
      var calls = [['Get', { typeName: 'Device', search: { fromDate: now.toISOString() }, resultsLimit: 5000 }]];
      settings.diagnostics.forEach(function (d) {
        calls.push(['Get', { typeName: 'StatusData', search: { diagnosticSearch: { id: d.id }, fromDate: from.toISOString(), toDate: now.toISOString() }, resultsLimit: 50000 }]);
      });
      var doorDiags = settings.doorDiagnostics || [], nT = settings.diagnostics.length;
      doorDiags.forEach(function (d) {
        calls.push(['Get', { typeName: 'StatusData', search: { diagnosticSearch: { id: d.id }, fromDate: from.toISOString(), toDate: now.toISOString() }, resultsLimit: 50000 }]);
      });
      return multi(calls).then(function (r) {
        var devices = r[0], series = {}, door = {};
        function collect(target, list, diagId) {
          list.forEach(function (s) {
            var dev = s.device && s.device.id; if (!dev || typeof s.data !== 'number') return;
            target[dev] = target[dev] || {};
            (target[dev][diagId] = target[dev][diagId] || []).push({ t: new Date(s.dateTime), v: s.data });
          });
        }
        settings.diagnostics.forEach(function (d, i) { collect(series, r[i + 1], d.id); });
        doorDiags.forEach(function (d, j) { collect(door, r[nT + 1 + j], d.id); });
        [series, door].forEach(function (obj) {
          Object.keys(obj).forEach(function (dev) { Object.keys(obj[dev]).forEach(function (k) { obj[dev][k].sort(function (a, b) { return a.t - b.t; }); }); });
        });
        var vehicles = devices.filter(function (dv) { return series[dv.id]; }).map(function (dv) { return { id: dv.id, name: dv.name }; });
        return { vehicles: vehicles, diagnostics: settings.diagnostics, series: series, door: door, doorDiags: doorDiags, from: from, to: now };
      });
    });
  }

  // ---------- デモデータ ----------
  function loadDemo(hours) {
    var now = Date.now(), from = now - hours * 3600e3, step = 5 * 60000;
    var diags = [{ id: 'd1', name: 'Remote probe 1 temperature' }, { id: 'd2', name: 'Remote probe 2 temperature' }];
    var names = ['名古屋 800 い 25-01', '名古屋 800 い 25-02', '品川 800 か 10-11', '品川 800 か 10-12', '大阪 800 さ 31-07', '大阪 800 さ 31-08', '福岡 800 た 44-20', '札幌 800 な 52-03'];
    var seed = 7; function rnd() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
    var vehicles = [], series = {}, door = {};
    names.forEach(function (n, vi) {
      var id = 'v' + vi; vehicles.push({ id: id, name: n }); series[id] = {};
      diags.forEach(function (d, di) {
        var arr = [], end = vi === 7 ? now - 3 * 3600e3 : now;
        for (var t = from; t <= end; t += step) {
          var v = -21 + di * 0.6 + (rnd() - 0.5) * 1.6;
          var h = (now - t) / 3600e3;
          if (vi === 2 && h > 5 && h < 5.6) v += 9;                 // 荷積み時の逸脱（35分）
          if (vi === 4 && Math.abs(h - 9) < 0.06) v += 7;           // ドア開閉の一瞬（除外される）
          if (vi === 5 && h < 1.2) v += (1.2 - h) * 8;               // 現在も上昇中
          if (vi === 1) {                                            // 範囲内だが急上昇中（冷凍機の故障の兆候）
            if (h < 0.9) v += -3 + (0.9 - h) * 4;                    // 直近54分で +4℃/時
            else if (h < 3) v -= 3;
            else if (h < 3.5) v -= 3 * (3.5 - h) / 0.5;
          }
          arr.push({ t: new Date(t), v: Math.round(v * 10) / 10 });
        }
        series[id][d.id] = arr;
      });
      // ドア開閉（[何時間前に開いたか, 開いていた分数(null=開いたまま)]）
      var opens = [[12 + vi * 0.7, 6]];               // 荷下ろし
      if (vi === 2) opens.push([5.62, 25]);           // 荷積み → 逸脱の原因
      if (vi === 4) opens.push([9.05, 4]);
      if (vi === 5) opens.push([1.25, null]);         // 開いたまま → 温度上昇中
      if (vi % 3 === 0) opens.push([3 + vi * 0.1, 8]);
      var dArr = [];
      opens.forEach(function (o) {
        var st = now - o[0] * 3600e3; if (st < from) return;
        dArr.push({ t: new Date(st), v: 1 });
        if (o[1] != null) dArr.push({ t: new Date(st + o[1] * 60000), v: 0 });
      });
      dArr.sort(function (a, b) { return a.t - b.t; });
      door[id] = { door1: dArr };
    });
    return Promise.resolve({ vehicles: vehicles, diagnostics: diags, series: series, door: door, doorDiags: [{ id: 'door1', name: '接点センサー1（ドア）' }], from: new Date(from), to: new Date(now) });
  }

  // ---------- 集計 ----------
  function isOut(v) { return v < settings.min || v > settings.max; }
  function hasDoor() { return !!(data && data.doorDiags && data.doorDiags.length); }

  // ドアが開いていた区間 [{start, end, ongoing}]。値が0以外を「開」とし、doorInvert で反転。
  function doorIntervals(devId) {
    var out = [], s = (data.door || {})[devId] || {};
    Object.keys(s).forEach(function (k) {
      var cur = null;
      s[k].forEach(function (p) {
        var open = (p.v !== 0) !== !!settings.doorInvert;
        if (open && !cur) cur = { start: p.t, end: null };
        else if (!open && cur) { cur.end = p.t; out.push(cur); cur = null; }
      });
      if (cur) { cur.end = data.to; cur.ongoing = true; out.push(cur); }
    });
    return out.sort(function (a, b) { return a.start - b.start; });
  }

  function excursions(devId) {
    var out = [], s = data.series[devId] || {};
    data.diagnostics.forEach(function (d) {
      var arr = s[d.id] || [], cur = null;
      arr.forEach(function (p, i) {
        if (isOut(p.v)) {
          if (!cur) cur = { devId: devId, diag: d, start: p.t, end: null, peak: p.v };
          if (Math.abs(p.v - (settings.min + settings.max) / 2) > Math.abs(cur.peak - (settings.min + settings.max) / 2)) cur.peak = p.v;
        } else if (cur) { cur.end = p.t; out.push(cur); cur = null; }
        if (cur && i === arr.length - 1) { cur.lastT = p.t; out.push(cur); }
      });
    });
    return out.filter(function (e) {
      e.duration = (e.end || e.lastT) - e.start;
      return e.duration >= settings.minExcursionMin * 60000 || !e.end;
    });
  }

  // ---------- 温度の変化の速さ（℃/時） ----------
  // 各計測点に p.r（過去 riseWindowMin 分の変化を1時間あたりに換算）を付与。期間が短すぎる場合は null。
  function computeRates(a) {
    var win = Math.max(10, settings.riseWindowMin || 60) * 60000, j = 0;
    for (var i = 0; i < a.length; i++) {
      var ti = +a[i].t;
      while (j < i && +a[j].t < ti - win) j++;
      var k = j < i ? j : i - 1;                                  // 期間内に前の点が無ければ直前の点（2倍の期間まで）
      if (k < 0 || ti - +a[k].t > 2 * win || ti - +a[k].t < 10 * 60000) { a[i].r = null; continue; }
      a[i].r = (a[i].v - a[k].v) / (ti - +a[k].t) * 3600e3;
    }
  }
  function riseOn() { return (settings.riseRate || 0) > 0; }
  function isRise(r) { return riseOn() && r != null && r >= settings.riseRate; }
  function fmtRate(r) { return r == null ? '-' : (r >= 0 ? '+' : '') + r.toFixed(1) + '℃/時'; }

  function vehicleState(v) {
    var s = data.series[v.id] || {}, latest = [], last = null;
    data.diagnostics.forEach(function (d) {
      var arr = s[d.id]; if (!arr || !arr.length) return;
      var p = arr[arr.length - 1]; latest.push({ diag: d, v: p.v, t: p.t, r: p.r });
      if (!last || p.t > last) last = p.t;
    });
    var status = !latest.length ? 'none'
      : (Date.now() - last > settings.staleMinutes * 60000) ? 'stale'
      : latest.some(function (l) { return isOut(l.v); }) ? 'ng'
      : latest.some(function (l) { return isRise(l.r); }) ? 'rise' : 'ok';
    return { latest: latest, last: last, status: status };
  }
  var STATUS_TEXT = { ok: '正常', ng: '逸脱', rise: '急上昇', stale: 'データ未受信', none: 'データなし' };
  var STATUS_ORDER = { ng: 0, rise: 1, stale: 2, ok: 3, none: 4 };

  // ---------- 描画 ----------
  function render() {
    var filter = ($('dt-filter').value || '').trim();
    Object.keys(data.series).forEach(function (dev) { Object.keys(data.series[dev]).forEach(function (k) { computeRates(data.series[dev][k]); }); });
    var rows = data.vehicles.map(function (v) { return Object.assign({ v: v }, vehicleState(v)); })
      .sort(function (a, b) { return STATUS_ORDER[a.status] - STATUS_ORDER[b.status] || a.v.name.localeCompare(b.v.name); });
    var events = [];
    data.vehicles.forEach(function (v) {
      var doors = doorIntervals(v.id);
      excursions(v.id).forEach(function (e) {
        var eEnd = e.end || e.lastT;
        e.door = doors.some(function (iv) { return iv.start < eEnd && iv.end > e.start - DOOR_LOOKBACK_MS; });
        events.push(e);
      });
    });
    events.sort(function (a, b) { return b.start - a.start; });
    data.events = events;

    $('kpi-total').textContent = rows.length;
    $('kpi-stale').parentNode.title = settings.staleMinutes + '分以上、温度データが届いていない車両';
    ['ok', 'ng', 'rise', 'stale'].forEach(function (k) {
      var n = rows.filter(function (r) { return r.status === k; }).length;
      $('kpi-' + k).textContent = n;
      // 温度逸脱中・急上昇は1件以上でカードを塗りつぶして警告
      if (k === 'ng' || k === 'rise' || k === 'stale') $('kpi-' + k).parentNode.classList.toggle('active', n > 0);
    });
    $('kpi-events').textContent = events.length;

    if (!selectedId && rows.length) selectedId = rows[0].v.id;
    $('dt-vehicles').innerHTML = rows.filter(function (r) { return !filter || r.v.name.indexOf(filter) >= 0; }).map(function (r) {
      var temps = r.latest.map(function (l) {
        var rate = r.status !== 'stale' && l.r != null
          ? ' <span class="dt-rate' + (isRise(l.r) ? ' up' : '') + '">' + (l.r >= 0 ? '↑' : '↓') + fmtRate(l.r) + '</span>' : '';
        return '<span class="dt-temp' + (isOut(l.v) ? ' ng' : '') + '"><small>' + esc(sensorLabel(l.diag.name)) + '</small>' + fmtV(l.v) + rate + '</span>';
      }).join('');
      return '<tr data-id="' + esc(r.v.id) + '"' + (r.v.id === selectedId ? ' class="selected"' : '') + '>' +
        '<td><span class="dt-status ' + r.status + '">' + STATUS_TEXT[r.status] + '</span></td>' +
        '<td>' + esc(r.v.name) + '</td><td>' + (temps || '-') + '</td><td title="' + fmtT(r.last) + '">' + ago(r.last) + '</td></tr>';
    }).join('') || '<tr><td colspan="4" class="dt-muted">温度データのある車両がありません</td></tr>';

    var names = {}; data.vehicles.forEach(function (v) { names[v.id] = v.name; });
    $('dt-events').innerHTML = events.map(function (e) {
      return '<tr><td>' + esc(names[e.devId]) + '</td><td>' + esc(sensorLabel(e.diag.name)) + '</td><td>' + fmtT(e.start) + '</td>' +
        '<td>' + (e.end ? fmtT(e.end) : '<span class="dt-ongoing">継続中</span>') + '</td><td>' + fmtDur(e.duration) + '</td><td>' + fmtV(e.peak) + '</td>' +
        '<td>' + (!hasDoor() ? '<span class="dt-muted">未設定</span>' : e.door ? '<span class="dt-door-tag">あり</span>' : '-') + '</td></tr>';
    }).join('') || '<tr><td colspan="7" class="dt-muted">期間内の逸脱はありません</td></tr>';

    renderDetail();
  }

  function renderDetail(forceW) {
    var v = data.vehicles.filter(function (x) { return x.id === selectedId; })[0];
    $('dt-detail-title').textContent = v ? '温度推移：' + v.name : '温度推移';
    $('dt-range-label').textContent = '基準 ' + settings.min + '〜' + settings.max + '℃';
    var chart = $('dt-chart');
    if (!v) { chart.innerHTML = '<div class="empty">車両を選択してください</div>'; $('dt-legend').innerHTML = ''; $('dt-stats').innerHTML = ''; return; }
    if (!forceW) ensureLogs(v.id);   // 吹き出しの「場所」用に位置記録を先読み

    var s = data.series[v.id] || {}, lines = [];
    data.diagnostics.forEach(function (d, i) { var a = s[d.id]; if (a && a.length) lines.push({ d: d, a: a, c: COLORS[i % COLORS.length] }); });
    if (!lines.length) {
      chartCtx = null; zoom = null; updateZoomUi();
      chart.innerHTML = '<div class="empty">期間内のデータがありません</div>'; $('dt-legend').innerHTML = ''; $('dt-stats').innerHTML = ''; return;
    }

    // 表示範囲（拡大中はその範囲。車両が変わった・期間外になった場合は全体に戻す）
    var full0 = +data.from, full1 = +data.to;
    if (zoom && (zoom.devId !== v.id || zoom.t1 <= full0 || zoom.t0 >= full1)) zoom = null;
    var t0 = zoom ? Math.max(zoom.t0, full0) : full0, t1 = zoom ? Math.min(zoom.t1, full1) : full1;
    updateZoomUi(t0, t1);

    // 範囲内の点（線が途切れないよう前後1点も描画に含める）
    var all = [];
    lines.forEach(function (ln) {
      var a = ln.a, i0 = 0, i1 = a.length - 1;
      while (i0 < a.length && +a[i0].t < t0) i0++;
      while (i1 >= 0 && +a[i1].t > t1) i1--;
      ln.vis = i0 <= i1 ? a.slice(i0, i1 + 1) : [];
      ln.draw = a.slice(Math.max(0, i0 - 1), Math.min(a.length, Math.max(i1, i0 - 1) + 2));
      ln.vis.forEach(function (p) { all.push(p.v); });
    });
    // 範囲内に計測点が無い場合は、範囲の両端の値を前後の計測値から推定
    var estimated = !all.length, est = [], estAvgs = [];
    if (estimated) {
      lines.forEach(function (ln) {
        var a0 = interp(ln.a, t0), a1 = interp(ln.a, t1);
        if (a0 && a1) { est.push(a0.v, a1.v); estAvgs.push((a0.v + a1.v) / 2); }   // 直線区間の平均は両端の平均
      });
    }
    var scaleVals = all.length ? all : est.length ? est : [].concat.apply([], lines.map(function (ln) { return ln.draw.map(function (p) { return p.v; }); }));
    if (!scaleVals.length) scaleVals = [settings.min, settings.max];

    // 表示幅に合わせて描画（縮小で文字が小さくならないように）
    // 「大きく表示」中は画面の高さに合わせて縦も大きくする（レポート出力時は常に300）
    var H = forceW ? 300 : sizeMode === 'full' ? Math.max(300, Math.round(visibleHeight() - fullReserve()))
      : sizeMode === 'small' ? 190 : 300;
    var W = forceW || Math.max(360, Math.round(chart.clientWidth || 800)), bigY = !forceW && sizeMode === 'full', L = bigY ? 58 : 44, R = 12, T = 12, B = 44;
    var dmin = Math.min.apply(null, scaleVals), dmax = Math.max.apply(null, scaleVals), lo, hi;
    if (zoom) { var pad = Math.max(0.5, (dmax - dmin) * 0.1); lo = dmin - pad; hi = dmax + pad; }   // 拡大中はデータに合わせる
    else { lo = Math.min(dmin, settings.min) - 2; hi = Math.max(dmax, settings.max) + 2; }          // 全体表示はしきい値も含める
    var x = function (t) { return L + (t - t0) / (t1 - t0) * (W - L - R); };
    var y = function (val) { return T + (hi - val) / (hi - lo) * (H - T - B); };
    var PW = W - L - R, PH = H - T - B;
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="温度推移グラフ">' +
      '<defs><clipPath id="dt-clip"><rect x="' + L + '" y="' + T + '" width="' + PW + '" height="' + PH + '"/></clipPath>';
    // しきい値の外側（上限より上・下限より下）の領域。線をここで切り抜いて赤く重ねる
    var yTop = Math.max(T, Math.min(H - B, y(settings.max))), yBot = Math.max(T, Math.min(H - B, y(settings.min)));
    var outRects = (yTop > T ? '<rect x="' + L + '" y="' + T + '" width="' + PW + '" height="' + (yTop - T).toFixed(1) + '"/>' : '') +
      (yBot < H - B ? '<rect x="' + L + '" y="' + yBot.toFixed(1) + '" width="' + PW + '" height="' + (H - B - yBot).toFixed(1) + '"/>' : '');
    svg += '<clipPath id="dt-clip-out">' + (outRects || '<rect x="0" y="0" width="0" height="0"/>') + '</clipPath></defs>';
    svg += '<g clip-path="url(#dt-clip)"><rect x="' + L + '" y="' + y(settings.max) + '" width="' + PW + '" height="' + (y(settings.min) - y(settings.max)) + '" fill="rgba(21,128,61,.08)"/></g>';
    svg += '<g fill="rgba(220,38,38,.07)">' + outRects + '</g>';   // 範囲外の領域を薄い赤で塗る
    // ドア開の区間を背景に帯で表示（表示範囲と重なる部分だけ集計）
    var doors = hasDoor() ? doorIntervals(v.id) : [], doorsVis = [], doorMs = 0;
    doors.forEach(function (iv) {
      var s0 = Math.max(+iv.start, t0), s1 = Math.min(+iv.end, t1); if (s1 <= s0) return;
      doorsVis.push(iv); doorMs += s1 - s0;
      var w = Math.max(2, x(s1) - x(s0));
      svg += '<rect class="dt-door-band" x="' + x(s0).toFixed(1) + '" y="' + T + '" width="' + w.toFixed(1) + '" height="' + PH + '" fill="' + DOOR_COLOR + '"/>';
    });
    var stepV = niceStep((hi - lo) / 6), dec = stepV < 1 ? 1 : 0;
    for (var g = Math.ceil(lo / stepV) * stepV; g <= hi + 1e-9; g += stepV) {
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(g) + '" y2="' + y(g) + '" stroke="#e5e7eb"/>' +
        '<text x="' + (L - 6) + '" y="' + (y(g) + 4) + '" font-size="' + (bigY ? 13 : 11) + '" text-anchor="end" fill="' + (bigY ? '#374151' : '#6b7280') + '">' + g.toFixed(dec) + (bigY ? '℃' : '') + '</text>';
    }
    [settings.min, settings.max].forEach(function (th) {
      if (th < lo || th > hi) return;
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(th) + '" y2="' + y(th) + '" stroke="#dc2626" stroke-dasharray="4 3"/>';
    });
    // 時間軸：日本時間のきりのよい時刻に目盛り（間隔は表示範囲から自動選択）。日付は変わり目だけ2段目に表示
    var JST = 9 * 3600e3, MIN = 60000, HR = 3600e3;
    var STEPS = [5 * MIN, 10 * MIN, 15 * MIN, 30 * MIN, HR, 2 * HR, 3 * HR, 4 * HR, 6 * HR, 12 * HR, 24 * HR, 48 * HR];
    var maxTicks = Math.max(2, Math.floor(PW / 80)), step = STEPS[STEPS.length - 1];
    for (var si = 0; si < STEPS.length; si++) { if ((t1 - t0) / STEPS[si] <= maxTicks) { step = STEPS[si]; break; } }
    var axisY = H - B;
    svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + axisY + '" y2="' + axisY + '" stroke="#9ca3af"/>';
    var prevDay = null;
    for (var tt = Math.ceil((t0 + JST) / step) * step - JST; tt <= t1; tt += step) {
      var tx = x(tt); if (tx < L + 14 || tx > W - R - 14) continue;   // 端で文字が切れないよう省略
      var jd = new Date(tt + JST).toISOString();                       // 日本時間の YYYY-MM-DDTHH:MM
      var day = jd.slice(5, 10).replace('-', '/'), hm = jd.slice(11, 16);
      svg += '<line x1="' + tx.toFixed(1) + '" x2="' + tx.toFixed(1) + '" y1="' + T + '" y2="' + axisY + '" stroke="#eef0f3"/>' +
        '<line x1="' + tx.toFixed(1) + '" x2="' + tx.toFixed(1) + '" y1="' + axisY + '" y2="' + (axisY + 5) + '" stroke="#9ca3af"/>' +
        '<text x="' + tx.toFixed(1) + '" y="' + (axisY + 18) + '" font-size="12" font-weight="600" text-anchor="middle" fill="#374151">' + hm + '</text>';
      if (day !== prevDay) {
        svg += '<text x="' + tx.toFixed(1) + '" y="' + (axisY + 33) + '" font-size="11" text-anchor="middle" fill="#6b7280">' + day + '</text>';
        prevDay = day;
      }
    }
    svg += '<g clip-path="url(#dt-clip)">';
    var markersShown = false, riseShown = false, outShown = false;
    lines.forEach(function (ln) {
      var pts = ln.draw.map(function (p) { return x(+p.t).toFixed(1) + ',' + y(p.v).toFixed(1); }).join(' ');
      // 急上昇の区間をオレンジの太い帯で強調（線の下に描く）
      if (!estimated) ln.draw.forEach(function (p, i) {
        if (i === 0 || !isRise(p.r)) return;
        var q = ln.draw[i - 1]; riseShown = true;
        svg += '<line x1="' + x(+q.t).toFixed(1) + '" y1="' + y(q.v).toFixed(1) + '" x2="' + x(+p.t).toFixed(1) + '" y2="' + y(p.v).toFixed(1) +
          '" stroke="' + RISE_COLOR + '" stroke-width="8" stroke-linecap="round"/>';
      });
      var dash = estimated ? ' stroke-dasharray="6 4"' : '';
      svg += '<polyline points="' + pts + '" fill="none" stroke="' + ln.c + '" stroke-width="1.8"' + dash + '/>';
      // しきい値の外に出た部分だけ赤で重ね描き（またいだ位置で正確に切り替わる）
      if (ln.draw.some(function (p) { return isOut(p.v); })) {
        outShown = true;
        svg += '<polyline class="dt-out-line" points="' + pts + '" fill="none" stroke="#dc2626" stroke-width="2.6" stroke-linejoin="round" clip-path="url(#dt-clip-out)"' + dash + '/>';
      }
      // 計測点を常に表示（点が密集して線が見えにくくなる場合のみ省略）
      if (ln.vis.length && ln.vis.length <= PW / 4) {
        markersShown = true;
        ln.vis.forEach(function (p) { svg += '<circle class="dt-mpt" cx="' + x(+p.t).toFixed(1) + '" cy="' + y(p.v).toFixed(1) + '" r="2.6" fill="' + (isOut(p.v) ? '#dc2626' : ln.c) + '"/>'; });
      }
    });
    svg += '</g>';
    if (estimated) {
      svg += '<rect x="' + (L + 6) + '" y="' + (T + 6) + '" width="168" height="20" rx="4" fill="#fef3c7" stroke="#f59e0b"/>' +
        '<text x="' + (L + 14) + '" y="' + (T + 20) + '" font-size="11" fill="#92400e">推定（範囲内に計測点なし）</text>';
    }
    // ドラッグ選択の範囲（初期は非表示）
    svg += '<rect class="dt-sel" x="0" y="' + T + '" width="0" height="' + PH + '" fill="rgba(37,99,235,.12)" stroke="rgba(37,99,235,.5)" style="display:none;pointer-events:none"/>';
    // ホバー用の縦線と点（初期は非表示）
    svg += '<g class="dt-hover" style="display:none;pointer-events:none"><line class="dt-hl" y1="' + T + '" y2="' + (H - B) + '" stroke="#6b7280" stroke-dasharray="3 3"/>' +
      lines.map(function (ln) { return '<circle r="4" fill="' + ln.c + '" stroke="#fff" stroke-width="1.5"/>'; }).join('') + '</g>';
    chart.innerHTML = svg + '</svg><div class="dt-tip" hidden></div>' +
      // 大きく表示中は、見出しが画面外に隠れても操作できるよう、グラフの右上にもボタンを置く
      (!forceW && sizeMode === 'full' ? '<div class="dt-chart-tools">' +
        '<button type="button" data-size="small" title="グラフを小さく表示">' + ICON_SHRINK + '小さく表示</button>' +
        '<button type="button" data-size="normal" title="標準の大きさに戻す（Esc）">' + ICON_SHRINK + '標準に戻す</button></div>' : '');
    chartCtx = { W: W, H: H, L: L, R: R, T: T, B: B, t0: t0, t1: t1, x: x, y: y, lines: lines, doors: doors };
    $('dt-legend').innerHTML = lines.map(function (ln) { return '<span><i style="background:' + ln.c + '"></i>' + esc(sensorLabel(ln.d.name)) + '</span>'; }).join('') +
      '<span><i class="dt-th-swatch"></i>しきい値</span>' +
      (outShown ? '<span><i class="dt-out-swatch" style="background:#dc2626"></i>範囲外（危険）</span>' : '') +
      (hasDoor() ? '<span><i class="dt-door-swatch" style="background:' + DOOR_COLOR + '"></i>ドア開</span>' : '') +
      (markersShown ? '<span><i class="dt-dot"></i>計測点</span>' : '') +
      (riseShown ? '<span><i class="dt-rise-swatch" style="background:' + RISE_COLOR + '"></i>急上昇（+' + settings.riseRate + '℃/時以上）</span>' : '') +
      '<span class="dt-hint">' +
        '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5M11 8v6M8 11h6"/></svg><em>ドラッグで拡大</em>' +
        '<b class="dt-hint-sep"></b>' +
        '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg><em>ダブルクリックで全体表示</em>' +
      '</span>';

    // 統計は表示範囲で計算
    var avg = all.length ? all.reduce(function (a, b) { return a + b; }, 0) / all.length : null;
    var exMs = data.events.filter(function (e) { return e.devId === v.id; }).reduce(function (a, e) {
      return a + Math.max(0, Math.min(+(e.end || e.lastT), t1) - Math.max(+e.start, t0));
    }, 0);
    var tempStats;
    if (all.length) {
      tempStats = [['最低', fmtV(Math.min.apply(null, all))], ['最高', fmtV(Math.max.apply(null, all))], ['平均', fmtV(avg)]];
    } else if (est.length) {
      var eavg = estAvgs.reduce(function (a, b) { return a + b; }, 0) / estAvgs.length;
      tempStats = [['最低（推定）', '≈ ' + fmtV(Math.min.apply(null, est)), 1], ['最高（推定）', '≈ ' + fmtV(Math.max.apply(null, est)), 1], ['平均（推定）', '≈ ' + fmtV(eavg), 1]];
    } else {
      tempStats = [['最低', '-'], ['最高', '-'], ['平均', '-']];
    }
    // 表示範囲内の最大上昇速度
    var maxR = null, maxRT = null;
    lines.forEach(function (ln) { ln.vis.forEach(function (p) { if (p.r != null && (maxR == null || p.r > maxR)) { maxR = p.r; maxRT = p.t; } }); });
    var rateStat = ['最大上昇速度', maxR == null ? '-' : fmtRate(maxR) + ' <small class="dt-muted">' + fmtHM(maxRT) + '</small>', 0, isRise(maxR)];
    $('dt-stats').innerHTML = tempStats.concat([rateStat, ['逸脱時間合計', exMs ? fmtDur(exMs) : 'なし']])
      .concat(hasDoor() ? [['ドア開 回数', doorsVis.length + '回'], ['ドア開 合計', doorMs ? fmtDur(doorMs) : 'なし']] : [])
      .map(function (p) { return '<div' + (p[2] ? ' class="dt-est"' : p[3] ? ' class="dt-rise"' : '') + '><span>' + p[0] + '</span><strong>' + p[1] + '</strong></div>'; }).join('') +
      (estimated && est.length ? '<p class="dt-est-note">※ 表示範囲内に計測点がないため、最低・最高・平均は範囲の前後の計測値を直線で結んで推定した値です。</p>'
        : estimated ? '<p class="dt-est-note">※ 表示範囲内に計測点がなく、前後の計測値も離れすぎているため推定できません。</p>' : '');
  }

  // 時刻 t の温度を前後の計測値から直線補間で推定（前後どちらかが無い、または6時間以上空く場合は null）
  var INTERP_MAX_GAP = 6 * 3600e3;
  function interp(a, t) {
    if (!a.length || t < +a[0].t || t > +a[a.length - 1].t) return null;
    var lo = 0, hi = a.length - 1;
    while (hi - lo > 1) { var m = (lo + hi) >> 1; if (+a[m].t <= t) lo = m; else hi = m; }
    var p = a[lo], q = a[hi];
    if (+q.t - +p.t > INTERP_MAX_GAP) return null;
    var v = +q.t === +p.t ? p.v : p.v + (q.v - p.v) * (t - p.t) / (q.t - p.t);
    return { v: v, prev: p, next: q };
  }

  function niceStep(raw) {
    var steps = [0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50];
    for (var i = 0; i < steps.length; i++) if (steps[i] >= raw) return steps[i];
    return 100;
  }

  // ---------- 拡大表示 ----------
  var zoom = null, drag = null;   // zoom: {devId, t0, t1}

  // ---------- 表示サイズ（小さく表示 / 標準 / 大きく表示） ----------
  var sizeMode = 'normal';   // 'small' | 'normal' | 'full'
  var ICON_SHRINK = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 4v5H4"/><path d="M15 4v5h5"/><path d="M9 20v-5H4"/><path d="M15 20v-5h5"/></svg>';
  // 大きく表示する範囲：ブラウザ全体ではなく、Add-Inが見えている範囲（MyGeotabの左メニュー・上部バーを除く）
  function scrollParent(el) {
    for (var p = el.parentElement; p && p !== document.body && p !== document.documentElement; p = p.parentElement) {
      var oy = getComputedStyle(p).overflowY;
      if ((oy === 'auto' || oy === 'scroll') && p.clientHeight > 0) return p;
    }
    return null;
  }
  function fullBox() {
    var vw = window.innerWidth, vh = window.innerHeight, root = $('didcomTemp');
    var rr = root.getBoundingClientRect(), sp = scrollParent(root), ar = sp ? sp.getBoundingClientRect() : { left: 0, top: 0, right: vw, bottom: vh };
    var left = Math.max(0, ar.left, rr.left), right = Math.min(vw, ar.right, rr.right);
    var top = Math.max(0, ar.top), bottom = Math.min(vh, ar.bottom);
    if (right - left < 300 || bottom - top < 200) return { left: 0, top: 0, width: vw, height: vh };
    return { left: left, top: top, width: right - left, height: bottom - top };
  }
  function placeFull() {
    var card = document.querySelector('#didcomTemp .dt-detail'), b = fullBox();
    card.style.left = b.left + 'px'; card.style.top = b.top + 'px'; card.style.width = b.width + 'px'; card.style.height = b.height + 'px';
    return b;
  }
  function visibleHeight() {
    var h = sizeMode === 'full' ? fullBox().height : window.innerHeight || 800, sh = (window.screen && window.screen.availHeight) || h;
    return Math.min(h, sh - 150);
  }
  // グラフ以外（見出し・凡例・統計）に必要な高さ
  function fullReserve() {
    var card = document.querySelector('#didcomTemp .dt-detail');
    var head = card.querySelector('.dt-card-head'), leg = $('dt-legend'), st = $('dt-stats');
    return (head ? head.offsetHeight : 40) + (leg ? leg.offsetHeight : 24) + (st ? st.offsetHeight : 70) + 80;
  }
  function setSize(mode) {
    sizeMode = mode;
    var card = document.querySelector('#didcomTemp .dt-detail');
    card.classList.toggle('dt-full', mode === 'full');
    card.classList.toggle('dt-small', mode === 'small');
    var bf = $('dt-full'), bs = $('dt-small');
    bf.querySelector('span').textContent = mode === 'full' ? '元に戻す' : '大きく表示';
    bf.title = mode === 'full' ? '通常の表示に戻す（Esc）' : 'グラフを画面いっぱいに大きく表示（Escで戻る）';
    bf.classList.toggle('on', mode === 'full');
    bs.querySelector('span').textContent = mode === 'small' ? '元に戻す' : '小さく表示';
    bs.title = mode === 'small' ? '通常の大きさに戻す' : 'グラフを低く小さく表示';
    bs.classList.toggle('on', mode === 'small');
    if (mode === 'full') placeFull();
    else { card.style.left = card.style.top = card.style.width = card.style.height = ''; }
    hideHover();
    if (data) renderDetail();
    if (mode === 'full') card.scrollTop = 0;
  }
  function updateZoomUi(t0, t1) {
    var btn = $('dt-zoom-reset'); if (btn) btn.hidden = !zoom;
    $('dt-range-label').textContent = (zoom ? '表示範囲 ' + fmtT(new Date(t0)) + '〜' + fmtT(new Date(t1)) + '（' + fmtDur(t1 - t0) + '）・' : '') +
      '基準 ' + settings.min + '〜' + settings.max + '℃';
  }
  function resetZoom() { if (!zoom) return; zoom = null; hideHover(); renderDetail(); }
  function plotPx(ev) {
    var c = chartCtx, svgEl = $('dt-chart').querySelector('svg'); if (!c || !svgEl) return null;
    var r = svgEl.getBoundingClientRect();
    return (ev.clientX - r.left) * c.W / r.width;
  }
  function clampPx(px) { var c = chartCtx; return Math.max(c.L, Math.min(c.W - c.R, px)); }
  function onDown(ev) {
    if (!chartCtx || (ev.pointerType === 'mouse' && ev.button !== 0)) return;
    if (ev.target.closest && ev.target.closest('.dt-chart-tools')) return;   // グラフ上のボタンはドラッグ扱いしない
    var px = plotPx(ev); if (px == null || px < chartCtx.L || px > chartCtx.W - chartCtx.R) return;
    drag = { x0: px, cx0: ev.clientX, moved: false };
    try { $('dt-chart').setPointerCapture(ev.pointerId); } catch (e) { /* ignore */ }
    onHover(ev);
  }
  function onMove(ev) {
    if (drag && chartCtx) {
      if (Math.abs(ev.clientX - drag.cx0) > 5) drag.moved = true;
      if (drag.moved) {
        var px = clampPx(plotPx(ev)), sel = $('dt-chart').querySelector('.dt-sel');
        sel.setAttribute('x', Math.min(drag.x0, px).toFixed(1)); sel.setAttribute('width', Math.abs(px - drag.x0).toFixed(1));
        sel.style.display = ''; hideHover(); ev.preventDefault(); return;
      }
    }
    onHover(ev);
  }
  function onUp(ev) {
    if (!drag) return;
    var d = drag; drag = null;
    var sel = $('dt-chart').querySelector('.dt-sel'); if (sel) sel.style.display = 'none';
    if (!d.moved || !chartCtx) return;
    var c = chartCtx, px = clampPx(plotPx(ev)), a = Math.min(d.x0, px), b = Math.max(d.x0, px);
    if (b - a < 8) return;
    var toT = function (p) { return c.t0 + (p - c.L) / (c.W - c.L - c.R) * (c.t1 - c.t0); };
    var ta = toT(a), tb = toT(b), MIN = 5 * 60000;
    if (tb - ta < MIN) { var mid = (ta + tb) / 2; ta = mid - MIN / 2; tb = mid + MIN / 2; }   // 最小5分
    zoom = { devId: selectedId, t0: ta, t1: tb };
    hideHover(); renderDetail();
  }

  // ---------- グラフのホバー表示 ----------
  var chartCtx = null;
  function nearest(a, t) {
    if (!a.length) return null;
    var lo = 0, hi = a.length - 1;
    while (hi - lo > 1) { var m = (lo + hi) >> 1; if (+a[m].t < t) lo = m; else hi = m; }
    return Math.abs(a[lo].t - t) <= Math.abs(a[hi].t - t) ? a[lo] : a[hi];
  }
  function fmtHM(d) { return d.toLocaleTimeString('ja-JP', { timeZone: 'Asia/Tokyo', hour: '2-digit', minute: '2-digit' }); }
  function hideHover() {
    lastHoverEv = null;
    var chart = $('dt-chart'), g = chart.querySelector('.dt-hover'), tip = chart.querySelector('.dt-tip');
    if (g) g.style.display = 'none'; if (tip) tip.hidden = true;
  }

  // ---------- 測定場所（GPS位置と住所） ----------
  var addrCache = {}, addrPending = {}, logsLoading = {}, lastHoverEv = null;

  // 選択中の車両の位置記録（LogRecord）を期間分まとめて1回だけ読み込む
  function ensureLogs(devId) {
    if (!data || !devId) return;
    data.logs = data.logs || {};
    if (data.logs[devId] || logsLoading[devId]) return;
    if (demo) { data.logs[devId] = demoLogs(devId); return; }
    var d = data; logsLoading[devId] = true;
    call('Get', { typeName: 'LogRecord', search: { deviceSearch: { id: devId }, fromDate: d.from.toISOString(), toDate: d.to.toISOString() }, resultsLimit: 100000 })
      .then(function (list) {
        d.logs[devId] = (list || []).filter(function (r) { return r.latitude || r.longitude; })
          .map(function (r) { return { t: +new Date(r.dateTime), lat: r.latitude, lon: r.longitude, spd: r.speed || 0 }; })
          .sort(function (a, b) { return a.t - b.t; });
      }, function () { d.logs[devId] = []; })
      .then(function () { logsLoading[devId] = false; if (lastHoverEv && d === data) onHover(lastHoverEv); });
  }

  // 時刻 t の位置：直前の記録（停車中は記録が止まるため）。直後10分以内の記録の方が近ければそちら
  function positionAt(devId, t) {
    var a = data && data.logs && data.logs[devId];
    if (!a) return undefined;                  // 読み込み中
    var lo = 0, hi = a.length;
    while (lo < hi) { var m = (lo + hi) >> 1; if (a[m].t <= t) lo = m + 1; else hi = m; }
    var p = lo > 0 ? a[lo - 1] : null, n = lo < a.length ? a[lo] : null;
    if (n && n.t - t <= 10 * 60000 && (!p || n.t - t < t - p.t)) return n;
    return p && t - p.t <= 12 * 3600e3 ? p : null;
  }

  function cleanAddr(s) { return String(s || '').replace(/,?\s*(Japan|日本)\s*$/i, '').replace(/〒?\s*\d{3}-\d{4}\s*/, '').trim(); }
  // 住所を「〜県〜市」程度に短縮（政令市は区まで、郡は町村まで）。日本の住所形式でなければそのまま返す
  var PREF_RE = /(東京都|北海道|京都府|大阪府|[^\s,、]{2,3}県)/;
  function shortAddr(s) {
    var orig = cleanAddr(s);
    s = orig.replace(/\s+/g, '');
    var pm = PREF_RE.exec(s);
    if (!pm) {   // 英語表記など：末尾の「市, 県」だけを残す（例：Toyohashi, Aichi）
      var parts = orig.split(',').map(function (p) { return p.replace(/\d{3}-?\d{4}/, '').trim(); }).filter(Boolean);
      return parts.slice(-2).join(', ');
    }
    var pref = pm[1], rest = s.slice(pm.index + pref.length), m;
    if ((m = /^(.+?市)(?!市)(.+?区)?/.exec(rest))) return pref + m[1] + (m[2] && m[2].length <= 5 ? m[2] : '');  // 〇〇市（〇〇区）
    if ((m = /^(.+?郡.+?[町村])/.exec(rest))) return pref + m[1];                                               // 〇〇郡〇〇町
    if ((m = /^(.+?[区町村])/.exec(rest))) return pref + m[1];                                                  // 東京23区など
    return pref;
  }
  // GetAddresses の結果から短い住所を作る（都道府県・市区町村の項目があれば優先）
  function shortFromResult(a) {
    if (!a) return '';
    var region = a.region || a.state || '', city = a.city || '';
    // 住所全文に都道府県があればそこから切り出す（city 項目は政令市の区だけのことがあるため）
    if (a.formattedAddress && PREF_RE.test(a.formattedAddress.replace(/\s+/g, ''))) return shortAddr(a.formattedAddress);
    if (region && city && PREF_RE.test(region)) return shortAddr(region + city);
    return shortAddr(a.formattedAddress || [region, city, a.street].filter(Boolean).join(' '));
  }
  // 住所（キャッシュ済みなら文字列、取得中なら null）
  function addressFor(pos) {
    var key = pos.lat.toFixed(4) + ',' + pos.lon.toFixed(4);
    if (addrCache[key] !== undefined) return addrCache[key];
    if (demo) return (addrCache[key] = demoAddress(pos));
    if (!addrPending[key]) {
      addrPending[key] = true;
      call('GetAddresses', { coordinates: [{ x: pos.lon, y: pos.lat }] }).then(function (r) {
        var a = r && r[0];
        addrCache[key] = shortFromResult(a);
      }, function () { addrCache[key] = ''; })
        .then(function () { delete addrPending[key]; if (lastHoverEv) onHover(lastHoverEv); });
    }
    return null;
  }

  function placeRow(t) {
    var pos = positionAt(selectedId, t), body;
    if (pos === undefined) body = '<span class="dt-muted">取得中…</span>';
    else if (!pos) body = '<span class="dt-muted">位置データなし</span>';
    else {
      var ad = addressFor(pos);
      body = (ad === null ? '<span class="dt-muted">住所を取得中…</span>' : esc(ad || (pos.lat.toFixed(5) + ', ' + pos.lon.toFixed(5)))) +
        ' <span class="dt-muted">（' + (pos.spd > 3 ? '走行中 ' + Math.round(pos.spd) + 'km/h' : '停車中') + '）</span>' +
        (Math.abs(pos.t - t) > 30 * 60000 ? ' <span class="dt-muted">' + fmtT(new Date(pos.t)) + ' 時点の位置</span>' : '');
    }
    return '<div class="dt-tip-place"><span class="dt-tip-place-label">場所</span>' + body + '</div>';
  }

  // デモ用の位置記録（車両名の地名を中心に、日中は走行・夜間は停車）
  var DEMO_CITIES = { '名古屋': [35.170, 136.882, '愛知県名古屋市', ['中区', '中村区', '港区', '熱田区', '北区']],
    '品川': [35.609, 139.730, '東京都品川区', ['東品川', '大井', '八潮', '南品川', '勝島']],
    '大阪': [34.693, 135.502, '大阪府大阪市', ['北区', '此花区', '住之江区', '西淀川区', '中央区']],
    '福岡': [33.590, 130.402, '福岡県福岡市', ['博多区', '東区', '中央区', '南区', '西区']],
    '札幌': [43.062, 141.354, '北海道札幌市', ['中央区', '白石区', '東区', '北区', '豊平区']] };
  function demoCity(devId) {
    var v = data.vehicles.filter(function (x) { return x.id === devId; })[0], key = v ? v.name.slice(0, 2) : '名古屋';
    return DEMO_CITIES[key] || DEMO_CITIES['名古屋'];
  }
  function demoLogs(devId) {
    var c = demoCity(devId), out = [], seed = devId.length * 31 + devId.charCodeAt(devId.length - 1);
    for (var t = +data.from; t <= +data.to; t += 10 * 60000) {
      var h = (new Date(t + 9 * 3600e3).getUTCHours()), driving = h >= 6 && h < 19;
      if (!driving && out.length && out[out.length - 1].spd === 0) continue;   // 停車中は記録しない
      var k = (t / 3600e3 + seed) * 0.35;
      out.push({ t: t, lat: c[0] + Math.sin(k) * 0.05, lon: c[1] + Math.cos(k * 1.3) * 0.06, spd: driving ? 25 + Math.round(Math.abs(Math.sin(k * 3)) * 35) : 0 });
    }
    return out;
  }
  function demoAddress(pos) {
    var best = null;
    Object.keys(DEMO_CITIES).forEach(function (k) {
      var c = DEMO_CITIES[k], d = Math.abs(c[0] - pos.lat) + Math.abs(c[1] - pos.lon);
      if (!best || d < best.d) best = { c: c, d: d };
    });
    var wards = best.c[3], i = Math.abs(Math.round(pos.lat * 1000 + pos.lon * 1000)) % wards.length;
    return best.c[2] + wards[i] + '（デモ）';
  }

  function onHover(ev) {
    lastHoverEv = ev;
    ensureLogs(selectedId);
    var c = chartCtx, chart = $('dt-chart'), svgEl = chart.querySelector('svg'), g = chart.querySelector('.dt-hover'), tip = chart.querySelector('.dt-tip');
    if (!c || !svgEl || !g || !tip) return;
    var rect = svgEl.getBoundingClientRect(), scale = c.W / rect.width;
    var px = (ev.clientX - rect.left) * scale;
    if (px < c.L || px > c.W - c.R || ev.clientY < rect.top || ev.clientY > rect.bottom) { hideHover(); return; }
    var tc = c.t0 + (px - c.L) / (c.W - c.L - c.R) * (c.t1 - c.t0);
    // マウス位置から40px相当（最低30分）以内の計測値だけを対象にする
    var maxGap = Math.max(30 * 60000, (c.t1 - c.t0) * 40 / (c.W - c.L - c.R));
    // 表示範囲内で一番近い「実際の計測点」に吸着する。無ければ推定表示（範囲内に計測点が無い拡大時など）
    var best = null;
    c.lines.forEach(function (ln) {
      var p = nearest(ln.a, tc);
      if (p && +p.t >= c.t0 && +p.t <= c.t1 && (!best || Math.abs(p.t - tc) < Math.abs(best.t - tc))) best = p;
    });
    var snapped = !!(best && Math.abs(best.t - tc) <= maxGap);
    var t = snapped ? +best.t : tc;
    var sx = c.x(t);
    g.style.display = '';
    var hl = g.querySelector('.dt-hl'); hl.setAttribute('x1', sx); hl.setAttribute('x2', sx);
    var circles = g.querySelectorAll('circle'), rows = [];
    c.lines.forEach(function (ln, i) {
      var p = nearest(ln.a, t), cir = circles[i], label = esc(sensorLabel(ln.d.name));
      var sw = '<i style="background:' + ln.c + '"></i>';
      var ip = snapped ? null : interp(ln.a, t);
      if (ip) {
        // 近くに計測点が無い → 前後の計測値から推定（白抜きの点で表示）
        cir.style.display = ''; cir.setAttribute('cx', c.x(t).toFixed(1)); cir.setAttribute('cy', c.y(ip.v).toFixed(1));
        cir.setAttribute('fill', '#fff'); cir.setAttribute('stroke', ln.c); cir.setAttribute('stroke-width', '2');
        rows.push('<div>' + sw + label + '<b class="dt-est-val' + (isOut(ip.v) ? ' ng' : '') + '">≈ ' + fmtV(ip.v) + '</b><span class="dt-est-tag">推定</span>' +
          ' <span class="dt-muted">(' + fmtHM(ip.prev.t) + ' / ' + fmtHM(ip.next.t) + ' の計測から)</span></div>');
        return;
      }
      cir.setAttribute('fill', ln.c); cir.setAttribute('stroke', '#fff'); cir.setAttribute('stroke-width', '1.5');
      if (!snapped || !p || Math.abs(p.t - t) > maxGap) { cir.style.display = 'none'; rows.push('<div>' + sw + label + '<span class="dt-muted">データなし</span></div>'); return; }
      cir.style.display = ''; cir.setAttribute('cx', c.x(+p.t).toFixed(1)); cir.setAttribute('cy', c.y(p.v).toFixed(1));
      var off = Math.abs(p.t - t) > 2 * 60000 ? ' <span class="dt-muted">(' + fmtHM(p.t) + ' 計測)</span>' : '';
      var rt = p.r != null ? ' <span class="dt-rate' + (isRise(p.r) ? ' up' : '') + '">' + fmtRate(p.r) + '</span>' : '';
      rows.push('<div>' + sw + label + '<b' + (isOut(p.v) ? ' class="ng"' : '') + '>' + fmtV(p.v) + '</b>' + rt + off + '</div>');
    });
    if (hasDoor()) {
      var iv = c.doors.filter(function (d) { return t >= +d.start && t <= +d.end; })[0];
      rows.push('<div><i class="dt-door-swatch" style="background:' + DOOR_COLOR + '"></i>ドア<b>' + (iv ? '開' : '閉') + '</b>' +
        (iv ? ' <span class="dt-muted">(' + fmtHM(iv.start) + '〜' + (iv.ongoing ? '開いたまま' : fmtHM(iv.end) + '・' + fmtDur(iv.end - iv.start)) + ')</span>' : '') + '</div>');
    }
    rows.push(placeRow(t));
    tip.innerHTML = '<div class="dt-tip-time">' + fmtT(new Date(t)) + (snapped ? ' <span class="dt-muted">計測</span>' : ' <span class="dt-est-tag">計測点なし</span>') + '</div>' + rows.join('');
    tip.hidden = false;
    // 吹き出しは縦線の右側、はみ出す場合は左側に表示
    var cx = rect.left + sx / scale - chart.getBoundingClientRect().left, cw = chart.clientWidth, tw = tip.offsetWidth;
    tip.style.left = (cx + 14 + tw > cw ? Math.max(0, cx - 14 - tw) : cx + 14) + 'px';
    tip.style.top = Math.max(0, ev.clientY - chart.getBoundingClientRect().top - tip.offsetHeight / 2) + 'px';
  }

  // ---------- 更新・CSV ----------
  function refresh() {
    if (loading) return; loading = true;
    $('dt-refresh').disabled = true;
    var hours = +$('dt-hours').value;
    (demo ? loadDemo(hours) : loadLive(hours)).then(function (d) {
      data = d;
      showMsg(demo ? 'MyGeotab外で開いているため、デモデータを表示しています。'
        : !d.vehicles.length ? '監視対象：' + d.diagnostics.map(function (x) { return x.name; }).join('、') + '。期間内にデータがある車両はまだありません（センサーの有効化・MAC登録後に表示されます）。'
        : '', true);
      render();
    }).catch(function (e) {
      if (e && e.info) showMsg(e.message, true);
      else showMsg('データ取得に失敗しました：' + (e && e.message ? e.message : e));
    }).then(function () { loading = false; $('dt-refresh').disabled = false; });
  }
  function schedule() {
    clearInterval(timer); timer = null;
    if ($('dt-auto').checked) timer = setInterval(refresh, AUTO_MS);
  }

  function exportCsv() {
    if (!data || !data.events) return;
    var names = {}; data.vehicles.forEach(function (v) { names[v.id] = v.name; });
    var q = function (s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
    var lines = [['車両', 'センサー', '開始', '終了', '継続時間(分)', '最高/最低(℃)', '下限(℃)', '上限(℃)', 'ドア開'].map(q).join(',')];
    data.events.forEach(function (e) {
      lines.push([names[e.devId], sensorLabel(e.diag.name), fmtT(e.start), e.end ? fmtT(e.end) : '継続中', Math.round(e.duration / 60000), e.peak, settings.min, settings.max,
        !hasDoor() ? '未設定' : e.door ? 'あり' : 'なし'].map(q).join(','));
    });
    var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }); // BOM付き: Excelで文字化けしない
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = '温度逸脱履歴_' + new Date().toISOString().slice(0, 10) + '.csv'; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // ---------- レポート出力（PNG / PDF） ----------
  var FONT = '"Hiragino Kaku Gothic ProN","Yu Gothic UI","Meiryo",system-ui,sans-serif';
  function jstStamp() { return new Date(Date.now() + 9 * 3600e3).toISOString().slice(0, 16).replace(/[-:]/g, '').replace('T', '-'); }
  function fmtFull(d) { return d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }); }
  function download(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 3000);
  }

  // 出力用の幅でグラフを描き直してSVG・凡例・統計を取り出し、画面表示は元に戻す
  function captureChart(width) {
    renderDetail(width);
    var svgEl = $('dt-chart').querySelector('svg'), ctx = chartCtx, out = null;
    if (svgEl) {
      var c = svgEl.cloneNode(true);
      Array.prototype.forEach.call(c.querySelectorAll('.dt-hover,.dt-sel'), function (n) { n.parentNode.removeChild(n); });
      c.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      c.setAttribute('width', width); c.setAttribute('height', 300); c.setAttribute('font-family', FONT);
      out = new XMLSerializer().serializeToString(c);
    }
    var legend = Array.prototype.filter.call($('dt-legend').querySelectorAll('span'), function (s) { return !s.classList.contains('dt-hint'); })
      .map(function (s) {
        var i = s.querySelector('i');
        return { color: i ? getComputedStyle(i).backgroundColor : '#999', dot: !!(i && i.classList.contains('dt-dot')),
          dashed: !!(i && i.classList.contains('dt-th-swatch')), text: s.textContent.trim() };
      });
    var stats = Array.prototype.map.call($('dt-stats').querySelectorAll(':scope > div'), function (d) {
      return { label: d.querySelector('span').textContent, value: d.querySelector('strong').textContent.replace(/\s+/g, ' ').trim(), est: d.classList.contains('dt-est'), rise: d.classList.contains('dt-rise') };
    });
    var noteEl = $('dt-stats').querySelector('.dt-est-note');
    var res = { svg: out, legend: legend, stats: stats, note: noteEl ? noteEl.textContent : '', t0: ctx && ctx.t0, t1: ctx && ctx.t1,
      sensors: ctx ? ctx.lines.map(function (ln) { return sensorLabel(ln.d.name); }) : [] };
    renderDetail();
    return res;
  }
  function loadImage(svg) {
    return new Promise(function (res, rej) {
      var img = new Image();
      img.onload = function () { res(img); };
      img.onerror = function () { rej(new Error('グラフ画像の作成に失敗しました。')); };
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
  }

  function buildReportCanvas() {
    var v = data && data.vehicles.filter(function (x) { return x.id === selectedId; })[0];
    if (!v || !chartCtx) return Promise.reject(new Error('出力するグラフがありません。車両を選択してください。'));
    var BW = 1100, PAD = 32, CW = BW - PAD * 2, S = 2;
    var info = captureChart(CW);
    if (!info.svg) return Promise.reject(new Error('出力するグラフがありません。'));
    return loadImage(info.svg).then(function (img) {
      var evs = (data.events || []).filter(function (e) { return e.devId === v.id && +e.start < info.t1 && +(e.end || e.lastT) > info.t0; });
      var shown = evs.slice(0, 12);
      var statCols = Math.max(1, Math.min(info.stats.length, Math.floor((CW + 10) / 130))), statRows = Math.ceil(info.stats.length / statCols);
      var cv = document.createElement('canvas'); cv.width = BW * S; cv.height = 1400 * S;
      var g = cv.getContext('2d'); g.scale(S, S);
      g.fillStyle = '#fff'; g.fillRect(0, 0, BW, 1400);
      function text(s, x, yy, size, color, weight, align) {
        g.font = (weight || 400) + ' ' + size + 'px ' + FONT; g.fillStyle = color || '#1f2937'; g.textAlign = align || 'left'; g.fillText(s, x, yy);
      }
      var y = PAD;
      text('温度記録レポート', PAD, y + 22, 22, '#111827', 700);
      text('出力日時：' + fmtFull(new Date()), BW - PAD, y + 22, 12, '#6b7280', 400, 'right');
      y += 34; g.fillStyle = '#0b5cad'; g.fillRect(PAD, y, CW, 2); y += 24;
      text('車両：' + v.name + '　　センサー：' + info.sensors.join('、'), PAD, y, 14, '#1f2937', 700); y += 22;
      text('期間：' + fmtFull(new Date(info.t0)) + ' 〜 ' + fmtFull(new Date(info.t1)) + '（' + fmtDur(info.t1 - info.t0) + '）　　基準温度：' +
        settings.min + '〜' + settings.max + '℃' + (riseOn() ? '　　急上昇：+' + settings.riseRate + '℃/時以上' : ''), PAD, y, 13, '#374151'); y += 10;
      g.drawImage(img, PAD, y, CW, 300); y += 308;
      // 凡例
      var lx = PAD;
      info.legend.forEach(function (l) {
        g.fillStyle = l.color;
        if (l.dot) { g.beginPath(); g.arc(lx + 4, y + 9, 3.5, 0, Math.PI * 2); g.fill(); lx += 12; }
        else if (l.dashed) {
          g.strokeStyle = '#dc2626'; g.lineWidth = 2; g.setLineDash([4, 3]);
          g.beginPath(); g.moveTo(lx, y + 9); g.lineTo(lx + 16, y + 9); g.stroke(); g.setLineDash([]); g.lineWidth = 1; lx += 22;
        }
        else { g.fillRect(lx, y + 6, 16, 6); lx += 22; }
        text(l.text, lx, y + 14, 12, '#374151'); lx += g.measureText(l.text).width + 18;
      });
      y += 30;
      // 統計
      var bw = (CW - (statCols - 1) * 10) / statCols;
      info.stats.forEach(function (s, i) {
        var cx = PAD + (i % statCols) * (bw + 10), cy = y + Math.floor(i / statCols) * 58;
        g.fillStyle = s.est ? '#fffbeb' : s.rise ? '#fff7ed' : '#f3f4f6'; g.fillRect(cx, cy, bw, 50);
        if (s.est || s.rise) {
          g.strokeStyle = s.est ? '#f59e0b' : '#fdba74'; if (s.est) g.setLineDash([4, 3]);
          g.strokeRect(cx + 0.5, cy + 0.5, bw - 1, 49); g.setLineDash([]);
        }
        text(s.label, cx + 10, cy + 18, 11, s.est ? '#92400e' : '#6b7280');
        text(s.value, cx + 10, cy + 39, 15, s.rise ? '#c2410c' : s.est ? '#78350f' : '#111827', 700);
      });
      y += statRows * 58;
      if (info.note) { text(info.note, PAD, y + 12, 11, '#92400e'); y += 22; }
      y += 22;
      // 逸脱履歴
      text('温度逸脱履歴（この期間）', PAD, y, 14, '#111827', 700); y += 10;
      var cols = [['センサー', 0], ['開始', 110], ['終了', 290], ['継続時間', 470], ['最高/最低', 610], ['ドア開', 750]];
      g.fillStyle = '#f3f4f6'; g.fillRect(PAD, y, CW, 26);
      cols.forEach(function (c) { text(c[0], PAD + 10 + c[1], y + 17, 12, '#6b7280', 700); });
      y += 26;
      if (!shown.length) { text('この期間の逸脱はありません', PAD + 10, y + 17, 12, '#6b7280'); y += 26; }
      shown.forEach(function (e) {
        var vals = [sensorLabel(e.diag.name), fmtFull(e.start), e.end ? fmtFull(e.end) : '継続中', fmtDur(e.duration), fmtV(e.peak), !hasDoor() ? '未設定' : e.door ? 'あり' : '-'];
        vals.forEach(function (s, i) { text(s, PAD + 10 + cols[i][1], y + 17, 12, (i === 2 && !e.end) || i === 4 ? '#dc2626' : '#1f2937'); });
        g.fillStyle = '#e5e7eb'; g.fillRect(PAD, y + 25, CW, 1); y += 26;
      });
      if (evs.length > shown.length) { text('ほか ' + (evs.length - shown.length) + ' 件（全件は「CSV出力」で確認できます）', PAD + 10, y + 16, 11, '#6b7280'); y += 22; }
      y += 24;
      text('MyGeotab 温度モニター v' + VERSION + '　｜　時刻は日本時間　｜　「（推定）」は計測点が無い範囲を前後の計測値から推定した値です', PAD, y, 10, '#9ca3af');
      // 使った高さで切り出し
      var outH = y + PAD - 12, out = document.createElement('canvas');
      out.width = BW * S; out.height = outH * S;
      out.getContext('2d').drawImage(cv, 0, 0);
      return { canvas: out, name: '温度記録_' + v.name.replace(/[\\/:*?"<>|\s]/g, '') + '_' + jstStamp() };
    });
  }

  // JPEG画像1枚をA4横に貼ったPDFを生成（外部ライブラリ不要）
  function makePdf(jpg, imgW, imgH) {
    var PW = 842, PH = 595, M = 24, sc = Math.min((PW - 2 * M) / imgW, (PH - 2 * M) / imgH);
    var dw = imgW * sc, dh = imgH * sc, dx = (PW - dw) / 2, dy = PH - M - dh;   // 上寄せ
    var content = 'q ' + dw.toFixed(2) + ' 0 0 ' + dh.toFixed(2) + ' ' + dx.toFixed(2) + ' ' + dy.toFixed(2) + ' cm /Im0 Do Q';
    var parts = [], len = 0, offs = [];
    function str(s) { var b = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 0xff; parts.push(b); len += b.length; }
    function obj(n, body) { offs[n] = len; str(n + ' 0 obj\n' + body + '\nendobj\n'); }
    str('%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');
    obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
    obj(2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>');
    obj(3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + PW + ' ' + PH + '] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>');
    offs[4] = len;
    str('4 0 obj\n<< /Type /XObject /Subtype /Image /Width ' + imgW + ' /Height ' + imgH +
      ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + jpg.length + ' >>\nstream\n');
    parts.push(jpg); len += jpg.length;
    str('\nendstream\nendobj\n');
    obj(5, '<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream');
    var xref = len, x = 'xref\n0 6\n0000000000 65535 f \n';
    for (var n = 1; n <= 5; n++) x += ('0000000000' + offs[n]).slice(-10) + ' 00000 n \n';
    str(x + 'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n');
    return new Blob(parts, { type: 'application/pdf' });
  }

  function exportReport(kind) {
    var btn = $(kind === 'pdf' ? 'dt-save-pdf' : 'dt-save-png'); btn.disabled = true;
    buildReportCanvas().then(function (r) {
      return new Promise(function (res, rej) {
        if (kind === 'png') {
          r.canvas.toBlob(function (b) { if (!b) return rej(new Error('画像の作成に失敗しました。')); download(b, r.name + '.png'); res(); }, 'image/png');
        } else {
          r.canvas.toBlob(function (b) {
            if (!b) return rej(new Error('PDFの作成に失敗しました。'));
            var fr = new FileReader();
            fr.onload = function () { download(makePdf(new Uint8Array(fr.result), r.canvas.width, r.canvas.height), r.name + '.pdf'); res(); };
            fr.onerror = function () { rej(new Error('PDFの作成に失敗しました。')); };
            fr.readAsArrayBuffer(b);
          }, 'image/jpeg', 0.92);
        }
      });
    }).catch(function (e) { alert(e && e.message ? e.message : e); })
      .then(function () { btn.disabled = false; });
  }

  // ---------- 設定画面 ----------
  // 診断項目の検索・選択リスト（温度センサー用とドアセンサー用で共用）
  function Picker(prefix, defaultQuery) { this.p = prefix; this.q = defaultQuery; this.draft = []; }
  Picker.prototype.render = function (list, fromSearch) {
    var chosen = {}; this.draft.forEach(function (d) { chosen[d.id] = 1; });
    $(this.p + '-list').innerHTML = list.map(function (d) {
      return '<label><input type="checkbox" data-id="' + esc(d.id) + '" data-name="' + esc(d.name) + '"' + (chosen[d.id] ? ' checked' : '') + '> ' +
        esc(d.name) + ' <code>' + esc(d.id) + '</code></label>';
    }).join('') || '<p class="dt-muted" style="padding:8px">' + (fromSearch ? '該当なし（別のキーワードで検索してください）' : '未選択（検索して選んでください）') + '</p>';
  };
  // 表示中のチェック状態を反映した選択結果（表示されていない選択済み項目は保持）
  Picker.prototype.collect = function () {
    var shown = {}, picked = [];
    Array.prototype.forEach.call($(this.p + '-list').querySelectorAll('input[type=checkbox]'), function (cb) {
      shown[cb.dataset.id] = 1; if (cb.checked) picked.push({ id: cb.dataset.id, name: cb.dataset.name });
    });
    return this.draft.filter(function (d) { return !shown[d.id]; }).concat(picked);
  };
  Picker.prototype.search = function () {
    var self = this;
    self.draft = self.collect();
    if (demo) { self.render(self.draft); return; }
    $(self.p + '-list').innerHTML = '<p class="dt-muted" style="padding:8px">検索中…</p>';
    findDiagnostics($(self.p + '-query').value.trim() || self.q).then(function (l) { self.render(l, true); }).catch(function (e) {
      $(self.p + '-list').innerHTML = '<p class="dt-muted" style="padding:8px">検索に失敗しました：' + esc(e && e.message || e) + '</p>';
    });
  };
  var tempPicker = new Picker('set-diag', ''), doorPicker = new Picker('set-door', 'Contact sensor');

  function openSettings() {
    $('set-profile').value = settings.profile; $('set-min').value = settings.min; $('set-max').value = settings.max;
    $('set-minex').value = settings.minExcursionMin; $('set-stale').value = settings.staleMinutes;
    $('set-rise').value = settings.riseRate; $('set-risewin').value = settings.riseWindowMin;
    $('set-door-invert').checked = !!settings.doorInvert;
    tempPicker.draft = (settings.diagnostics.length ? settings.diagnostics : (data ? data.diagnostics : [])).slice();
    doorPicker.draft = ((settings.doorDiagnostics || []).length ? settings.doorDiagnostics : (demo && data ? data.doorDiags : [])).slice();
    tempPicker.render(tempPicker.draft); doorPicker.render(doorPicker.draft);
    $('dt-settings').hidden = false;
  }
  function saveFromForm() {
    settings.profile = $('set-profile').value;
    settings.min = parseFloat($('set-min').value); settings.max = parseFloat($('set-max').value);
    settings.minExcursionMin = Math.max(0, parseInt($('set-minex').value, 10) || 0);
    settings.staleMinutes = Math.max(1, parseInt($('set-stale').value, 10) || 30);
    settings.riseRate = Math.max(0, parseFloat($('set-rise').value) || 0);
    settings.riseWindowMin = Math.max(10, parseInt($('set-risewin').value, 10) || 60);
    if (!(settings.min < settings.max)) { alert('下限は上限より小さい値にしてください。'); return; }
    settings.doorInvert = $('set-door-invert').checked;
    if (!demo) {
      settings.diagnostics = tempPicker.collect();
      settings.doorDiagnostics = doorPicker.collect();
      settings.doorAutoDone = true;
    }
    saveSettings(); $('dt-settings').hidden = true; refresh();
  }

  function bind() {
    $('dt-refresh').onclick = refresh;
    var ch = $('dt-chart');
    ch.addEventListener('pointerdown', onDown);    // ドラッグ開始（動かさなければタップ/クリックで値表示）
    ch.addEventListener('pointermove', onMove);
    ch.addEventListener('pointerup', onUp);
    ch.addEventListener('pointercancel', function () { drag = null; var s = ch.querySelector('.dt-sel'); if (s) s.style.display = 'none'; });
    ch.addEventListener('pointerleave', function () { if (!drag) hideHover(); });
    ch.addEventListener('dblclick', resetZoom);
    $('dt-zoom-reset').onclick = resetZoom;
    $('dt-full').onclick = function () { setSize(sizeMode === 'full' ? 'normal' : 'full'); };
    $('dt-small').onclick = function () { setSize(sizeMode === 'small' ? 'normal' : 'small'); };
    ch.addEventListener('click', function (e) {
      var b = e.target.closest && e.target.closest('.dt-chart-tools button'); if (b) { e.stopPropagation(); setSize(b.dataset.size); }
    });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && sizeMode === 'full') setSize('normal'); });
    $('dt-save-png').onclick = function () { exportReport('png'); };
    $('dt-save-pdf').onclick = function () { exportReport('pdf'); };
    var rz = null;
    window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (sizeMode === 'full') placeFull(); if (data) renderDetail(); }, 150); });
    $('dt-hours').onchange = function () { zoom = null; refresh(); };
    $('dt-auto').onchange = schedule;
    $('dt-filter').oninput = function () { if (data) render(); };
    $('dt-csv').onclick = exportCsv;
    $('dt-vehicles').onclick = function (e) {
      var tr = e.target.closest('tr[data-id]'); if (!tr) return;
      selectedId = tr.dataset.id; render();
    };
    $('dt-open-settings').onclick = openSettings;
    $('set-cancel').onclick = function () { $('dt-settings').hidden = true; };
    $('set-save').onclick = saveFromForm;
    $('set-diag-search').onclick = function () { tempPicker.search(); };
    $('set-door-search').onclick = function () { doorPicker.search(); };
    $('set-profile').onchange = function () {
      var p = PROFILES[this.value]; if (p) { $('set-min').value = p.min; $('set-max').value = p.max; }
    };
  }

  function start(a, isDemo) {
    api = a; demo = isDemo; root = $('didcomTemp');
    var badge = $('dt-source'); badge.textContent = (demo ? 'デモ' : 'ライブ') + ' v' + VERSION; badge.className = 'dt-badge' + (demo ? ' demo' : '');
    var ready = demo ? Promise.resolve() : new Promise(function (res) {
      api.getSession(function (s) { storageKey = 'didcomTemp:' + s.database; res(); });
    });
    return ready.then(function () { loadSettings(); bind(); });
  }

  return {
    initialize: function (a, state, callback) { start(a, false).then(function () { callback(); }, function () { callback(); }); },
    focus: function (a) { api = a; refresh(); schedule(); },
    blur: function () { clearInterval(timer); timer = null; },
    _standalone: function () { start(null, true).then(function () { refresh(); schedule(); }); }
  };
};

// MyGeotab外（ローカルでHTMLを開いた場合）はデモモードで起動。
// MyGeotab内では initialize が呼ばれるので、2秒待っても呼ばれなければデモとみなす。
(function () {
  var orig = geotab.addin.didcomTemp;
  geotab.addin.didcomTemp = function () {
    var inst = orig();
    var init = inst.initialize;
    inst.initialize = function () { window.__didcomTempInit = true; return init.apply(inst, arguments); };
    return inst;
  };
  window.addEventListener('load', function () {
    setTimeout(function () { if (!window.__didcomTempInit) orig()._standalone(); }, 2000);
  });
})();
