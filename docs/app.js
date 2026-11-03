const $ = (s, r = document) => r.querySelector(s);
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const KINDS = ['A/S', 'CLAIM', '미분류'];
const SHOT = location.search.includes('shot');
let ROWS = [], REC = [], DATA, SET, charts = {}, latestY, latestM;
const DEFAULTS = () => ({ 연: new Set([latestY]), 유무상: new Set(['무상']), 구분: new Set(['A/S', 'CLAIM']), 고객: new Set(), 선종: new Set(['상선', '특수선']) });
let SEL = { 연: new Set(), 유무상: new Set(), 구분: new Set(), 고객: new Set(), 선종: new Set() };
const GROUPS = [['연', '연도'], ['유무상', '유/무상'], ['구분', '구분'], ['선종', '선종'], ['고객', '고객']];
const CHIP_GROUPS = GROUPS.filter(([g]) => g !== '고객');

const fmt = v => { const a = Math.abs(v); return a >= 1e8 ? (v / 1e8).toFixed(2) + '억' : a >= 1e4 ? Math.round(v / 1e4).toLocaleString() + '만' : Math.round(v).toLocaleString(); };
const pct = (v, d = 1) => (v * 100).toFixed(d) + '%';
const sum = (a, f) => a.reduce((s, r) => s + f(r), 0);
const colors = () => ({ 'A/S': css('--c1'), 'CLAIM': css('--c2'), '미분류': css('--c3') });

function skeleton() {
  document.body.innerHTML = `
  <div class="app">
    <aside class="side">
      <div class="brand"><span class="logo"></span><div><b>CLAIM·AS</b><small>비용 대시보드</small></div></div>
      <div id="slicers"></div>
      <div class="grp dd" id="custDD"></div>
      <div class="foot" id="foot"></div>
    </aside>
    <main>
      <header><div><h1 id="ttl">CLAIM · AS 비용 현황</h1><p id="sub"></p></div><button class="reset" id="reset">필터 초기화</button></header>
      <section class="kpis" id="kpis"></section>
      <section class="row r3">
        <div class="card"><h2>연도별 비용</h2><div class="box"><canvas id="cYear"></canvas></div></div>
        <div class="card"><h2 id="tMonth">월별 비용</h2><div class="box"><canvas id="cMonth"></canvas></div></div>
        <div class="card"><h2 id="tCum">누적 추이</h2><div class="box"><canvas id="cCum"></canvas></div></div>
      </section>
      <section class="row r4">
        <div class="card"><h2>불량명칭별</h2><div class="box"><canvas id="cDefect"></canvas></div></div>
        <div class="card"><h2>장비별 TOP 10</h2><div class="box"><canvas id="cEquip"></canvas></div></div>
        <div class="card"><h2>거래처별 TOP 10</h2><div class="box"><canvas id="cVendor"></canvas></div></div>
        <div class="card"><h2>호선별 TOP 10</h2><div class="box"><canvas id="cShip"></canvas></div></div>
      </section>
    </main>
  </div>`;
}

async function init() {
  skeleton();
  const [d, s] = await Promise.all([
    fetch('data.json').then(r => r.json()),
    fetch('settings.json').then(r => r.json()).catch(() => ({}))]);
  DATA = d; SET = s;
  ROWS = d.rows.map(r => { const o = {}; d.dims.forEach((k, i) => o[k] = r[i]); o.v = r[d.dims.length]; o.n = r[d.dims.length + 1]; return o; });
  REC = d.receipts.map(r => ({ y: r[0], m: r[1], k: r[2], p: r[3], n: r[4] }));
  [latestY, latestM] = d.latest.split('-').map(Number);
  SEL = DEFAULTS();
  buildCustDD();
  $('#reset').onclick = () => { SEL = DEFAULTS(); render(); };
  $('#foot').innerHTML = `데이터 기준일<br><b>${d.latest}</b><br>갱신 ${d.generated}`;
  Chart.defaults.font.family = 'Pretendard Variable, Pretendard, "Malgun Gothic", sans-serif';
  if (SHOT) Chart.defaults.animation = false;
  render();
}

function options(g) {
  const tot = {};
  ROWS.forEach(r => { const k = r[g]; if (k !== '' && k != null) tot[k] = (tot[k] || 0) + r.v; });
  let keys = Object.keys(tot);
  if (g === '연') keys.sort((a, b) => b - a);
  else if (g === '고객') keys.sort((a, b) => tot[b] - tot[a]);
  else if (g === '유무상') keys = ['무상', '유상'].filter(k => k in tot);
  else if (g === '구분') keys = KINDS.filter(k => k in tot);
  else keys = ['상선', '특수선', '기타'].filter(k => k in tot);
  return keys.map(k => [g === '연' ? +k : k, k]);
}
function renderSlicers() {
  $('#slicers').innerHTML = CHIP_GROUPS.map(([g, label]) => `
    <div class="grp"><div class="gh"><span>${label}</span>
      <a data-clear="${g}" class="${SEL[g].size ? 'on' : ''}" title="선택 해제">✕</a></div>
      <div class="chips">${options(g).map(([v, t]) => `<button class="chip ${SEL[g].has(v) ? 'on' : ''}" data-g="${g}" data-v="${v}">${t}${g === '연' ? '년' : ''}</button>`).join('')}</div></div>`).join('');
  document.querySelectorAll('.chip').forEach(b => b.onclick = () => {
    const g = b.dataset.g, v = g === '연' ? +b.dataset.v : b.dataset.v;
    SEL[g].has(v) ? SEL[g].delete(v) : SEL[g].add(v); render();
  });
  document.querySelectorAll('[data-clear]').forEach(a => a.onclick = () => { SEL[a.dataset.clear].clear(); render(); });
}
function buildCustDD() {
  const box = $('#custDD'), opts = options('고객').map(o => o[0]);
  box.innerHTML = `<div class="gh"><span>고객</span><a id="custClear" title="선택 해제">✕</a></div>
    <button class="ddbtn" id="ddBtn"><span id="ddLabel">전체</span><i>▾</i></button>
    <div class="ddpanel" id="ddPanel" hidden>
      <input id="ddSearch" placeholder="고객 검색">
      <div class="ddlist" id="ddList">${opts.map(v => `<label data-v="${v}"><input type="checkbox" value="${v}"><span>${v}</span></label>`).join('')}</div>
      <div class="ddfoot"><button id="ddAll">전체 선택</button><button id="ddNone">선택 해제</button></div></div>`;
  $('#ddBtn').onclick = e => { e.stopPropagation(); const p = $('#ddPanel'); p.hidden = !p.hidden; if (!p.hidden) $('#ddSearch').focus(); };
  document.addEventListener('click', e => { if (!box.contains(e.target)) $('#ddPanel').hidden = true; });
  $('#ddSearch').oninput = e => { const q = e.target.value.trim().toLowerCase(); document.querySelectorAll('#ddList label').forEach(l => l.style.display = l.dataset.v.toLowerCase().includes(q) ? '' : 'none'); };
  $('#ddList').onchange = e => { const v = e.target.value; e.target.checked ? SEL.고객.add(v) : SEL.고객.delete(v); render(); };
  $('#ddAll').onclick = () => { opts.forEach(v => SEL.고객.add(v)); render(); };
  $('#ddNone').onclick = () => { SEL.고객.clear(); render(); };
  $('#custClear').onclick = () => { SEL.고객.clear(); render(); };
}
function syncCustDD() {
  const sel = [...SEL.고객];
  $('#ddLabel').textContent = !sel.length ? '전체' : sel.length <= 2 ? sel.join(', ') : `${sel[0]} 외 ${sel.length - 1}`;
  $('#ddBtn').classList.toggle('on', sel.length > 0);
  $('#custClear').className = sel.length ? 'on' : '';
  document.querySelectorAll('#ddList input').forEach(i => i.checked = SEL.고객.has(i.value));
}
const pass = (r, skipYear) => (skipYear || !SEL.연.size || SEL.연.has(r.연)) && (!SEL.유무상.size || SEL.유무상.has(r.유무상)) &&
  (!SEL.구분.size || SEL.구분.has(r.구분)) && (!SEL.고객.size || SEL.고객.has(r.고객)) && (!SEL.선종.size || SEL.선종.has(r.선종));
const group = (rows, kf) => { const m = new Map(); rows.forEach(r => { const k = kf(r); m.set(k, (m.get(k) || 0) + r.v); }); return m; };
const avg3 = y => { const rev = SET.revenue || {}; const v = [y - 2, y - 1, y].map(k => rev[k]).filter(x => x != null); return v.length ? sum(v, x => x) / v.length : null; };

function render() {
  renderSlicers();
  syncCustDD();
  const years = SEL.연.size ? [...SEL.연] : [latestY];
  const Y = Math.max(...years);
  const cut = y => y === latestY ? latestM : 12;
  const cur = ROWS.filter(r => pass(r) && years.includes(r.연) && r.월 <= cut(r.연));
  const prv = ROWS.filter(r => pass(r, true) && r.연 === Y - 1 && r.월 <= cut(Y));
  const curY = ROWS.filter(r => pass(r, true) && r.연 === Y && r.월 <= cut(Y));
  const base = ROWS.filter(r => pass(r, true));
  const act = GROUPS.filter(([g]) => g !== '연' && SEL[g].size).map(([g, l]) => `${l}: ${[...SEL[g]].join(', ')}`);
  $('#sub').textContent = `${[...years].sort().join(', ')}년${years.length === 1 && Y === latestY ? ` 1~${latestM}월` : ''}` + (act.length ? ' · ' + act.join(' · ') : ' · 전체 조건');
  kpis(Y, years, cur, prv, curY, cut(Y));
  yearChart(base, years);
  monthChart(Y, base, cut(Y));
  cumChart(Y, base, cut(Y));
  rank('cDefect', cur, r => r.불량명칭, 9);
  rank('cEquip', cur, r => r.장비 || '(없음)', 10);
  rank('cVendor', cur, r => r.거래처명 || '(없음)', 10);
  rank('cShip', cur, r => r.호선 || '(없음)', 10);
}

function kpis(Y, years, cur, prv, curY, cutM) {
  const c = sum(cur, r => r.v), cy = sum(curY, r => r.v), p = sum(prv, r => r.v), claim = sum(cur.filter(r => r.구분 === 'CLAIM'), r => r.v);
  const delta = p ? (cy - p) / p : null, tgt = (SET.kpi || {})[Y] || {};
  const rev = SET.ratio_basis === 'year' ? (SET.revenue || {})[Y] : avg3(Y);
  const rc = REC.filter(r => years.includes(r.y) && r.m <= (r.y === latestY ? latestM : 12) && (!SEL.유무상.size || SEL.유무상.has(r.p)));
  const bar = (v, over) => `<div class="bar"><i class="${over ? 'over' : ''}" style="width:${Math.min(v, 1) * 100}%"></i></div>`;
  const K = (t, v, s, x = '', cls = '') => `<div class="kpi card ${cls}"><span class="t">${t}</span><b class="v">${v}</b><span class="s">${s}</span>${x}</div>`;
  $('#kpis').innerHTML = [
    K('총 비용', fmt(c) + '<em>원</em>', `${sum(cur, r => r.n).toLocaleString()}건 · ${years.length > 1 ? years.length + '개 연도 합산' : Y + '년'}`, '', 'hero'),
    K(`전년(${Y - 1}) 동기 대비`, delta == null ? '-' : `<span class="${delta > 0 ? 'up' : 'down'}">${delta > 0 ? '▲' : '▼'} ${pct(Math.abs(delta))}</span>`, `전년 ${fmt(p)}원`),
    rev ? K('매출 대비 비용', pct(cy / rev, 2), tgt.ratio_target ? `목표 ${pct(tgt.ratio_target, 2)}` : `매출 ${fmt(rev)}원`, tgt.ratio_target ? bar(cy / rev / tgt.ratio_target, cy / rev > tgt.ratio_target) : '') : K('매출 대비 비용', '-', '매출 미입력'),
    K('CLAIM 비중', c ? pct(claim / c) : '-', `CLAIM ${fmt(claim)}원`),
    K('접수 건수', sum(rc, r => r.n).toLocaleString() + '<em>건</em>', `A/S ${sum(rc.filter(r => r.k.includes('A/S')), r => r.n)} · CLAIM ${sum(rc.filter(r => r.k.includes('CLAIM')), r => r.n)}`)
  ].join('');
}

function mk(id, cfg) {
  charts[id]?.destroy();
  Chart.defaults.color = css('--mute'); Chart.defaults.borderColor = css('--grid');
  cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false }, cfg.options);
  cfg.options.plugins = Object.assign({ legend: { position: 'bottom', labels: { boxWidth: 8, boxHeight: 8, usePointStyle: true, font: { size: 11 } } },
    tooltip: { mode: 'index', intersect: false, callbacks: {
      label: c => { const v = c.chart.options.indexAxis === 'y' ? c.parsed.x : c.parsed.y; return v == null ? null : ` ${c.dataset.label}: ${Math.round(v).toLocaleString()}원`; },
      footer: items => { const bars = items.filter(i => i.dataset.stack); if (bars.length < 2) return ''; return '합계: ' + Math.round(bars.reduce((a, i) => a + (i.chart.options.indexAxis === 'y' ? i.parsed.x : i.parsed.y || 0), 0)).toLocaleString() + '원'; } } } }, cfg.options.plugins);
  charts[id] = new Chart(document.getElementById(id), cfg);
}
const yA = () => ({ ticks: { callback: v => fmt(v) }, grid: { color: css('--grid') }, border: { display: false } });
const xA = { grid: { display: false }, border: { display: false } };
const sumBy = rows => { const a = Array(12).fill(0); rows.forEach(r => a[r.월 - 1] += r.v); return a; };

function yearChart(base, years) {
  const ys = [...new Set(base.map(r => r.연))].sort(), C = colors();
  const dim = (y, c) => years.includes(y) ? c : c + '55';
  mk('cYear', { type: 'bar', data: { labels: ys, datasets: KINDS.map(k => ({ label: k, stack: 's', borderRadius: k === 'CLAIM' ? 5 : 0, backgroundColor: ys.map(y => dim(y, C[k])), data: ys.map(y => sum(base.filter(r => r.연 === y && r.구분 === k), r => r.v)) })) },
    options: { scales: { x: { ...xA, stacked: true }, y: { ...yA(), stacked: true } } } });
}
function monthChart(Y, base, cutM) {
  document.getElementById('tMonth').textContent = `${Y}년 월별 비용 vs ${Y - 1}년`;
  const C = colors(), cy = base.filter(r => r.연 === Y), L = [...Array(12)].map((_, i) => (i + 1) + '월');
  mk('cMonth', { data: { labels: L, datasets: [
    ...KINDS.map(k => ({ type: 'bar', label: k, stack: 's', backgroundColor: C[k], borderRadius: k === 'CLAIM' ? 4 : 0, data: sumBy(cy.filter(r => r.구분 === k)).map((v, i) => i < cutM ? v : null) })),
    { type: 'line', label: `${Y - 1}년`, borderColor: css('--ink'), backgroundColor: css('--ink'), borderDash: [4, 4], borderWidth: 1.5, pointRadius: 0, tension: .3, data: sumBy(base.filter(r => r.연 === Y - 1)) }] },
    options: { scales: { x: { ...xA, stacked: true }, y: { ...yA(), stacked: true } } } });
}
const cumul = a => { let s = 0; return a.map(v => s += v); };
function cumChart(Y, base, cutM) {
  const tgt = ((SET.kpi || {})[Y] || {}).cost_target, L = [...Array(12)].map((_, i) => (i + 1) + '월');
  document.getElementById('tCum').textContent = `${Y}년 누적` + (tgt ? ' · 목표' : '');
  const ds = [{ label: `${Y}년`, borderColor: css('--c1'), backgroundColor: css('--c1') + '22', fill: true, borderWidth: 2.5, pointRadius: 3, tension: .3, data: cumul(sumBy(base.filter(r => r.연 === Y))).map((v, i) => i < cutM ? v : null) },
    { label: `${Y - 1}년`, borderColor: css('--prior'), borderDash: [4, 4], borderWidth: 1.5, pointRadius: 0, tension: .3, data: cumul(sumBy(base.filter(r => r.연 === Y - 1))) }];
  if (tgt) ds.push({ label: '목표', borderColor: css('--c2'), borderDash: [2, 3], borderWidth: 1.5, pointRadius: 0, data: L.map((_, i) => tgt * (i + 1) / 12) });
  mk('cCum', { type: 'line', data: { labels: L, datasets: ds }, options: { scales: { x: xA, y: yA() } } });
}
function rank(id, rows, kf, n) {
  const tot = group(rows, kf), keys = [...tot.entries()].sort((a, b) => b[1] - a[1]).slice(0, n).map(e => e[0]), C = colors();
  mk(id, { type: 'bar', data: { labels: keys.map(k => k.length > 14 ? k.slice(0, 13) + '…' : k), datasets: KINDS.map(k => { const g = group(rows.filter(r => r.구분 === k), kf); return { label: k, stack: 's', backgroundColor: C[k], borderRadius: k === 'CLAIM' ? 3 : 0, data: keys.map(x => g.get(x) || 0) }; }) },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { stacked: true, ticks: { callback: v => fmt(v), maxTicksLimit: 5, maxRotation: 0 }, grid: { color: css('--grid') }, border: { display: false } }, y: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } } } } });
}
init();
