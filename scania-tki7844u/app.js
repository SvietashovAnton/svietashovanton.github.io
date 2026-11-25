/* Бортовой журнал Scania TKI7844U (водитель Андрей) — калькулятор ставки, месячный план, журнал рейсов и диаграммы.
   Логика повторяет формулы рабочей книги Excel (листы «Калькулятор», «Месяц», «Рейсы», «Диаграммы», «ТО и ремонты»). */
(function () {
  'use strict';

  // Приложение ведётся для одного автомобиля и одного водителя: данные хранятся под отдельным ключом.
  const VEHICLE = 'Scania TKI7844U', DRIVER = 'Андрей';
  const STORE_KEY = 'logbook-scania-TKI7844U-v1';
  const OLD_STORE_KEY = 'trip-rate-calculator-v1';

  // ---------- Округление как в Excel ----------
  const EPS = 1e-9;
  function round2(x) { if (!isFinite(x)) return x; const s = Math.sign(x); return s * Math.round(Math.abs(x) * 100 + EPS) / 100; }
  function roundup2(x) { if (!isFinite(x)) return x; const s = Math.sign(x); return s * Math.ceil(Math.abs(x) * 100 - 1e-7) / 100; }
  const isNum = v => typeof v === 'number' && isFinite(v);
  const n0 = v => (isNum(v) ? v : 0);
  const sum = arr => arr.reduce((a, b) => a + n0(b), 0);

  // ---------- Даты (строки YYYY-MM-DD) ----------
  function toDate(s) { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); }
  function fromDate(dt) { return dt.toISOString().slice(0, 10); }
  function addDays(s, k) { const d = toDate(s); d.setUTCDate(d.getUTCDate() + k); return fromDate(d); }
  function monthStart(s) { return s.slice(0, 8) + '01'; }
  function monthEnd(s) { const d = toDate(monthStart(s)); d.setUTCMonth(d.getUTCMonth() + 1); d.setUTCDate(0); return fromDate(d); }
  function weekdayMon(s) { const w = toDate(s).getUTCDay(); return w === 0 ? 7 : w; } // WEEKDAY(x;2)
  const minS = (a, b) => (a < b ? a : b), maxS = (a, b) => (a > b ? a : b);
  const validDate = s => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
  function fmtDate(s) { if (!validDate(s)) return ''; const [y, m, d] = s.split('-'); return `${d}.${m}.${y}`; }
  function todayISO() { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
  const MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];

  // ---------- Форматирование ----------
  const nf2 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 1 });
  const nf4 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  const eur = v => (isNum(v) ? nf2.format(v) + ' €' : '—');
  const num2 = v => (isNum(v) ? nf2.format(v) : '');
  const km = v => (isNum(v) ? nf0.format(v) + ' км' : '—');
  const kmN = v => (isNum(v) ? nf1.format(v) : '');
  const pct = v => (isNum(v) ? nf1.format(v * 100) + ' %' : '—');
  const rate4 = v => (isNum(v) ? nf4.format(v) : '');
  const signCls = v => (isNum(v) ? (v < 0 ? 'neg' : v > 0 ? 'pos' : '') : '');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const $ = id => document.getElementById(id);

  // ---------- Состояние ----------
  function uid() { return Math.random().toString(36).slice(2, 10); }

  function defaultState() {
    return {
      month: {
        calcDate: '2026-11-25', days: 30, targetKm: 10000, driverDaily: 107, paidDays: 25, freeDays: 0, plnRate: 4.3925,
        fixed: [
          { name: 'Страховка OC', pln: 1000, eur: null },
          { name: 'Колёса', pln: 1300, eur: null },
          { name: 'ZUS + PIT', pln: 2450, eur: null },
          { name: 'Лизинг', pln: null, eur: 1150 },
          { name: 'Прочие постоянные', pln: null, eur: null }
        ],
        roadAvg: 0.315158, dieselL100: 29, dieselPrice: 1.51, adblueL100: 3, adbluePrice: 1,
        initKmDate: '2026-11-01', initKm: 0, initEmptyKm: null
      },
      calc: {
        title: 'DE Kiel — BE Arlon (пример)', date: '2026-11-26',
        route: [{ code: 'AT', km: 304, tollKm: 300, manual: null }, { code: 'CZ', km: 200, tollKm: 200, manual: null }],
        approachKm: 60, approachTollKm: null, approachRate: 0.572, approachManual: null,
        bridges: [],
        clientRate: 950, skonto: 0.03, premium: 0.05, other: 0, margins: [0.10, 0.15, 0.20]
      },
      trips: exampleTrips(),
      maint: [
        { id: uid(), date: '2026-11-14', odo: 412300, cat: 'ТО', work: 'Замена масла и фильтров (пример)', service: 'Сервис', pln: 1450, rate: 4.3925, eur: null, inTrip: false, tripNo: '', comment: '' }
      ],
      tariffs: JSON.parse(JSON.stringify(window.DEFAULT_TARIFFS || { road: [], bridge: [] })),
      isExample: true
    };
  }

  function exampleTrips() {
    const t = (date, no, route, routeKm, approachKm, freight, skonto, tolls) => ({ id: uid(), date, no, route, routeKm, approachKm, freight, skonto, premium: 0.05, tolls, other: null, comment: 'Пример' });
    return [
      t('2026-11-02', 'EX-101', 'Rotterdam NL — Brno CZ', 1040, 260, 1680, 0.03, 300),
      t('2026-11-04', 'EX-102', 'Brno CZ — Pardubice CZ', 180, 50, 260, 0.03, 32),
      t('2026-11-06', 'EX-103', 'Wrocław PL — Ostrava CZ', 270, 160, 520, 0.03, 66),
      t('2026-11-09', 'EX-104', 'Ostrava CZ — Dresden DE', 470, 30, 800, 0.03, 80),
      t('2026-11-11', 'EX-105', 'Hamburg DE — München DE', 775, 20, 1500, 0.03, 273),
      t('2026-11-13', 'EX-106', 'Nürnberg DE — Konstanz DE', 380, 140, 700, 0.03, 186.6),
      t('2026-11-17', 'EX-107', 'Ulm DE — Lyon FR', 680, 80, 1250, null, 215),
      t('2026-11-19', 'EX-108', 'Clermont-Ferrand FR — Mannheim DE', 710, 140, 840, null, 200),
      t('2026-11-21', 'EX-109', 'Karlsruhe DE — Praha CZ', 590, 50, 870, null, 210),
      t('2026-11-24', 'EX-110', 'Praha CZ — Augsburg DE', 560, 40, 900, 0.03, 190),
      t('2026-11-25', 'EX-111', 'Augsburg DE — Graz AT', 540, 60, 870, 0.03, 279),
      t('2026-11-27', 'EX-112', 'Graz AT — Ostrava CZ (план)', 540, 60, 950, 0.03, 255)
    ];
  }

  let state = load();
  function load() {
    try {
      let raw = localStorage.getItem(STORE_KEY);
      if (!raw) { raw = localStorage.getItem(OLD_STORE_KEY); if (raw) localStorage.setItem(STORE_KEY, raw); }
      if (raw) { const s = JSON.parse(raw); return mergeDefaults(s); }
    } catch (e) { /* хранилище недоступно */ }
    return defaultState();
  }
  function mergeDefaults(s) {
    const d = defaultState();
    s.month = Object.assign({}, d.month, s.month || {});
    s.calc = Object.assign({}, d.calc, s.calc || {});
    if (!Array.isArray(s.calc.margins)) s.calc.margins = d.calc.margins;
    s.trips = Array.isArray(s.trips) ? s.trips : [];
    s.maint = Array.isArray(s.maint) ? s.maint : [];
    s.tariffs = s.tariffs && s.tariffs.road ? s.tariffs : d.tariffs;
    s.trips.forEach(t => { if (!t.id) t.id = uid(); });
    s.maint.forEach(t => { if (!t.id) t.id = uid(); });
    return s;
  }
  let saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch (e) { /* ignore */ } }, 150);
  }

  // ---------- Расчёт месяца (лист «Месяц» + итоги «Рейсы») ----------
  function computeMonth(s) {
    const m = s.month;
    const r = { errors: [] };
    const cd = validDate(m.calcDate) ? m.calcDate : null;
    r.calcDate = cd;
    if (!cd) { r.errors.push('Укажите дату расчёта.'); }
    r.mStart = cd ? monthStart(cd) : null;
    r.mEnd = cd ? monthEnd(cd) : null;
    r.inMonth = d => cd && validDate(d) && d >= r.mStart && d <= r.mEnd;
    r.done = d => r.inMonth(d) && d <= cd;

    // Постоянные расходы
    const rate = isNum(m.plnRate) && m.plnRate > 0 ? m.plnRate : null;
    r.fixedRows = (m.fixed || []).map(f => {
      const e = n0(f.eur) + (isNum(f.pln) && f.pln ? (rate ? f.pln / rate : NaN) : 0);
      return { name: f.name, eurTotal: e };
    });
    r.fixedTotal = round2(sum(r.fixedRows.map(x => x.eurTotal)));
    if (!isFinite(r.fixedTotal)) { r.fixedTotal = NaN; r.errors.push('Укажите курс PLN.'); }

    const target = isNum(m.targetKm) && m.targetKm > 0 ? m.targetKm : null;
    if (!target) r.errors.push('Укажите целевой пробег.');
    r.target = target;
    const days = m.days, daily = n0(m.driverDaily);
    const daysOk = isNum(days) && days >= 28 && days <= 31 && Number.isInteger(days);
    if (!daysOk) r.errors.push('Проверьте дни месяца (28–31).');
    const paid = n0(m.paidDays), free = n0(m.freeDays);

    r.salaryFull = daysOk ? round2(days * daily) : NaN;            // J9
    r.salaryPaid = round2(paid * daily);                            // J10
    r.salaryFree = round2(free * daily);                            // J11
    r.salaryLeft = round2(r.salaryFull - r.salaryPaid - r.salaryFree); // J12
    r.salaryPlan = round2(r.salaryFull - r.salaryFree);              // J13
    r.daysLeft = daysOk && paid + free <= days ? days - paid - free : NaN; // J8
    if (daysOk && paid + free > days) r.errors.push('Оплаченных и бесплатных дней больше, чем дней в месяце.');
    r.fixedPerKm = target ? r.fixedTotal / target : NaN;            // J14
    r.salaryPerKm = target ? r.salaryPlan / target : NaN;           // J15
    r.costPerKm = r.fixedPerKm + r.salaryPerKm;                     // J16

    // Рейсы
    r.trips = s.trips.map(t => computeTrip(t, s, r));
    const done = r.trips.filter(t => r.done(t.date));
    r.doneTrips = done;
    r.tripCount = done.length;
    r.journalKm = round2(sum(done.map(t => t.totalKm)));             // Рейсы!I6
    r.freight = round2(sum(done.map(t => t.freight)));               // I7
    r.receivable = round2(sum(done.map(t => t.receivable)));         // I8
    r.tripProfit = round2(sum(done.map(t => t.profit)));             // Диаграммы!Q52
    r.tripCost = round2(sum(done.map(t => t.costWithPrem)));
    r.salaryInTrips = round2(sum(done.map(t => t.salary)));          // Месяц!J25
    r.badTrips = done.filter(t => !isNum(t.profit)).length;

    // Начальный пробег
    const initKm = n0(m.initKm);
    r.initInMonth = r.inMonth(m.initKmDate) && !(initKm > 0 && m.initKmDate > cd) ? initKm : 0;
    r.doneKm = round2(r.journalKm + r.initInMonth);                  // I9 = Месяц!E8
    r.remainingKm = target ? Math.max(target - r.doneKm, 0) : NaN;   // J6
    r.overKm = target ? Math.max(r.doneKm - target, 0) : NaN;        // J19
    r.fixedLeft = round2(r.fixedPerKm * r.remainingKm);              // J17
    r.fixedAllocated = round2(r.fixedTotal - r.fixedLeft);           // J18
    r.salaryLeftByKm = round2(r.salaryPerKm * r.remainingKm);        // J26
    r.salaryDiff = round2(r.salaryInTrips - r.salaryPaid);           // J32
    r.salaryPaidPerKm = r.doneKm > 0 ? r.salaryPaid / r.doneKm : NaN; // J33

    // ТО и ремонты
    r.maint = s.maint.map(x => computeMaint(x));
    const md = r.maint.filter(x => r.done(x.date) && isNum(x.total));
    r.maintTotal = round2(sum(md.map(x => x.total)));
    r.maintInTrip = round2(sum(md.filter(x => x.inTrip).map(x => x.total)));
    r.maintExtra = round2(r.maintTotal - r.maintInTrip);

    // Итоги месяца (Диаграммы)
    r.monthProfit = round2(r.tripProfit - r.maintExtra);             // I8
    r.monthExpenses = round2(r.freight - r.monthProfit);             // E8
    r.planShare = target ? r.doneKm / target : NaN;                  // K7

    // Статус
    if (!r.errors.length) {
      if (r.salaryLeftByKm < 0) r.status = { cls: 'bad', text: 'На рейсы отнесено больше зарплаты, чем в фонде месяца.' };
      else if (r.remainingKm === 0) r.status = { cls: 'good', text: r.salaryLeftByKm > 0 ? 'Плановый пробег выполнен, а фонд распределён не полностью.' : 'Цель и фонд месяца выполнены.' };
      else r.status = { cls: '', text: 'Данные месяца заполнены.' };
    } else r.status = { cls: 'bad', text: r.errors[0] };
    return r;
  }

  // ---------- Строка журнала (лист «Рейсы», колонки G–Y) ----------
  function computeTrip(t, s, mr) {
    const m = s.month;
    const o = { id: t.id, date: t.date, no: t.no, route: t.route, routeKm: t.routeKm, approachKm: t.approachKm, freight: t.freight, skonto: t.skonto, premium: t.premium, tollsExact: t.tolls, other: t.other, comment: t.comment };
    const kmOk = (isNum(t.routeKm) || isNum(t.approachKm)) && n0(t.routeKm) >= 0 && n0(t.approachKm) >= 0;
    o.totalKm = kmOk ? round2(n0(t.routeKm) + n0(t.approachKm)) : NaN;
    o.receivable = isNum(t.freight) && t.freight >= 0 ? round2(t.freight - (isNum(t.skonto) ? round2(t.freight * t.skonto) : 0)) : NaN;
    o.skontoAmt = isNum(t.freight) ? round2(t.freight - o.receivable) : NaN;
    o.ratePerKm = isNum(t.freight) && o.totalKm > 0 ? t.freight / o.totalKm : NaN;
    o.status = '';
    if (!validDate(t.date)) { o.status = 'Укажите дату'; return o; }
    if (!kmOk) { o.status = 'Укажите пробег'; return o; }
    if (!mr.inMonth(t.date)) { o.status = 'Другой месяц'; o.otherMonth = true; return o; }
    const G = o.totalKm;
    o.dieselFact = isNum(t.diesel);
    o.diesel = o.dieselFact ? round2(t.diesel) : n0(m.dieselL100) === 0 ? 0 : round2(G * m.dieselL100 / 100 * n0(m.dieselPrice));
    o.adblue = n0(m.adblueL100) === 0 ? 0 : round2(G * m.adblueL100 / 100 * n0(m.adbluePrice));
    o.salary = G > 0 && isNum(t.freight) && t.freight > 0 && isNum(mr.salaryPerKm) ? round2(G * mr.salaryPerKm) : NaN;
    o.fixed = isNum(mr.fixedPerKm) ? round2(G * mr.fixedPerKm) : NaN;
    o.tolls = isNum(t.tolls) ? round2(t.tolls) : round2(G * n0(m.roadAvg));
    o.tollsEstimated = !isNum(t.tolls);
    o.costNoPrem = round2(o.diesel + o.adblue + o.salary + o.fixed + o.tolls + n0(t.other)); // N
    const prem = isNum(t.premium) ? t.premium : 0;
    o.premiumAmt = isNum(t.freight) ? round2(t.freight * prem) : NaN;
    o.costWithPrem = round2(o.costNoPrem + o.premiumAmt);             // P
    o.profit = isNum(o.receivable) && isNum(o.costWithPrem) ? round2(o.receivable - o.costWithPrem) : NaN; // M
    o.margin = isNum(o.profit) && o.receivable > 0 ? o.profit / o.receivable : NaN; // L
    const plan = t.date > mr.calcDate;
    o.future = plan;
    if (!isNum(o.salary)) o.status = 'Проверьте фрахт и фонд зарплаты';
    else if (!isNum(o.profit)) o.status = 'Проверьте расходы';
    else o.status = (plan ? 'План: ' : '') + (o.tollsEstimated ? 'оценка дорог' : 'точные сборы');
    return o;
  }

  function computeMaint(x) {
    const o = Object.assign({}, x);
    const pl = n0(x.pln), e = n0(x.eur);
    if (pl > 0 && !(isNum(x.rate) && x.rate > 0)) { o.total = NaN; o.err = 'Укажите курс'; return o; }
    o.total = round2((pl > 0 ? pl / x.rate : 0) + e);
    return o;
  }

  // ---------- Калькулятор ----------
  function findTariff(list, code) {
    const c = String(code || '').trim().toUpperCase();
    if (!c) return null;
    const hits = list.filter(x => String(x.code).toUpperCase() === c);
    return hits.length === 1 ? hits[0] : null;
  }

  function computeCalc(s, mr) {
    const c = s.calc, m = s.month;
    const r = { warnings: [] };
    r.route = c.route.map(row => {
      const t = findTariff(s.tariffs.road, row.code);
      const rate = t && isNum(t.rate) ? t.rate : NaN;
      let fee = NaN, err = '';
      const empty = !row.code && !isNum(row.km) && !isNum(row.tollKm) && !isNum(row.manual);
      if (empty) return { rate, fee: 0, empty: true };
      if (!row.code || !isNum(row.km) || row.km <= 0) err = 'Страна и км';
      else if (isNum(row.manual)) fee = row.manual >= 0 ? round2(row.manual) : NaN;
      else if (!isNum(row.tollKm) || row.tollKm < 0 || row.tollKm > row.km) err = 'Платные км';
      else if (row.tollKm === 0) fee = 0;
      else if (isNum(rate)) fee = round2(row.tollKm * rate);
      else err = 'Тариф или сумма';
      if (err) fee = NaN;
      return { rate, fee, err, tariff: t };
    });
    r.routeKm = sum(c.route.map(x => x.km));
    r.routeFee = round2(sum(r.route.map(x => x.fee)));
    r.routeErr = r.route.some(x => x.err);
    const ak = n0(c.approachKm);
    r.approachKm = ak;
    if (ak === 0) r.approachFee = 0;
    else if (isNum(c.approachManual)) r.approachFee = round2(c.approachManual);
    else if (isNum(c.approachTollKm) && (c.approachTollKm < 0 || c.approachTollKm > ak)) { r.approachFee = NaN; r.warnings.push('Проверьте платные км доезда.'); }
    else r.approachFee = round2((isNum(c.approachTollKm) ? c.approachTollKm : ak) * n0(c.approachRate));
    r.roadFee = round2(r.routeFee + r.approachFee);                    // G45

    r.bridges = c.bridges.map(b => {
      const t = findTariff(s.tariffs.bridge, b.code);
      const rate = t && isNum(t.rate) ? t.rate : NaN;
      let fee = NaN, err = '';
      if (!b.code && !isNum(b.passes) && !isNum(b.manual)) return { rate, fee: 0, empty: true };
      if (!String(b.code || '').trim()) err = 'Укажите код';
      else if (isNum(b.manual)) fee = round2(b.manual);
      else if (!isNum(rate)) err = 'Тариф или сумма';
      else if (!isNum(b.passes)) fee = round2(rate);
      else if (b.passes >= 1 && Number.isInteger(b.passes)) fee = round2(rate * b.passes);
      else err = 'Число проездов';
      return { rate, fee: err ? NaN : fee, err, tariff: t };
    });
    r.bridgeFee = round2(sum(r.bridges.map(x => x.fee)));
    r.bridgeErr = r.bridges.some(x => x.err);

    const total = r.routeKm + ak;
    r.totalKm = total;
    r.diesel = round2(total * n0(m.dieselL100) / 100 * n0(m.dieselPrice));
    r.adblue = round2(total * n0(m.adblueL100) / 100 * n0(m.adbluePrice));
    r.salary = round2(mr.salaryPerKm * total);
    r.roads = round2(r.roadFee + r.bridgeFee);
    // Постоянные расходы по статьям: сумма статьи × (пробег рейса ÷ целевой пробег)
    const share = mr.target ? total / mr.target : NaN;
    r.fixedItems = mr.fixedRows.filter(x => x.eurTotal).map(x => ({ name: x.name, monthly: x.eurTotal, value: round2(x.eurTotal * share) }));
    const fixedRounded = round2(total * mr.fixedPerKm);
    const adj = round2(fixedRounded - sum(r.fixedItems.map(x => x.value)));
    r.fixedTotal = fixedRounded;
    r.other = round2(n0(c.other) + adj);
    r.cost = round2(r.diesel + r.adblue + r.salary + r.roads + r.other + sum(r.fixedItems.map(x => x.value))); // G26

    const sk = n0(c.skonto), pr = n0(c.premium);
    const skOk = sk >= 0 && sk < 1, prOk = pr >= 0 && pr < 1;
    const G26 = r.cost;
    const perKm = v => (isNum(v) && total > 0 ? roundup2(v / total) : NaN);
    r.offers = [];
    const minNoSk = prOk && 1 - pr > 0 ? roundup2(G26 / (1 - pr)) : NaN;
    r.offers.push({ label: 'Минимум без сконто', margin: 0, rate: minNoSk, perKm: perKm(minNoSk) });
    const minSk = skOk && 1 - sk - pr > 0 ? roundup2((G26 + (G26 > 0 && sk > 0 && pr > 0 ? 0.01 : 0)) / (1 - sk - pr)) : NaN;
    r.offers.push({ label: 'Минимум со сконто', margin: 0, rate: minSk, perKm: perKm(minSk) });
    (c.margins || []).forEach(mg => {
      let v = NaN;
      if (isNum(mg) && mg >= 0 && mg < 1 && (1 - sk) * (1 - mg) - pr > 0) {
        v = pr === 0 ? roundup2(roundup2(G26 / (1 - mg)) / (1 - sk)) : roundup2((G26 + (G26 > 0 ? 0.01 : 0)) / ((1 - sk) * (1 - mg) - pr));
      }
      r.offers.push({ label: 'Маржа ' + pct(mg), margin: mg, rate: v, perKm: perKm(v) });
    });

    const cr = c.clientRate;
    r.hasRate = isNum(cr) && cr >= 0;
    r.skontoAmt = r.hasRate ? round2(cr * sk) : NaN;          // G9
    r.receivable = r.hasRate ? round2(cr - r.skontoAmt) : NaN; // G10
    r.premiumAmt = r.hasRate ? round2(cr * pr) : NaN;          // G27
    r.costWithPrem = round2(G26 + r.premiumAmt);               // G28
    r.profit = round2(r.receivable - r.costWithPrem);          // G11
    r.margin = r.receivable > 0 ? r.profit / r.receivable : NaN; // G12
    r.ratePerKm = r.hasRate && total > 0 ? cr / total : NaN;   // G13
    r.costPerKm = total > 0 ? r.costWithPrem / total : NaN;

    // Сообщение (как F7 / I23)
    let msg = '';
    if (!(r.routeKm > 0)) msg = 'Укажите пробег по странам.';
    else if (r.routeErr) msg = 'Укажите сборы по маршруту.';
    else if (!isNum(r.approachFee)) msg = 'Укажите сборы за доезд.';
    else if (r.bridgeErr) msg = 'Укажите сборы за мосты и тоннели.';
    else if (!mr.target) msg = 'Укажите целевой пробег на вкладке «Месяц».';
    else if (mr.remainingKm === 0) msg = 'Плановый пробег месяца выполнен. Увеличьте целевой пробег месяца.';
    else if (total > mr.remainingKm) msg = 'Рейс больше остатка плана (' + km(mr.remainingKm) + '). Обновите цель на вкладке «Месяц».';
    else if (!isNum(mr.fixedTotal)) msg = 'Заполните месячные расходы.';
    else if (!skOk) msg = 'Укажите сконто 0–99,99 %.';
    else if (!prOk) msg = 'Проверьте премию экспедитора.';
    r.message = msg;
    r.ok = !msg || /больше остатка|выполнен/.test(msg);
    return r;
  }

  // ---------- Недели и пробег (лист «Диаграммы») ----------
  function computeCharts(s, mr) {
    const m = s.month, cd = mr.calcDate;
    const res = { weeks: [] };
    if (!cd) return res;
    const tripsIn = (a, b) => mr.trips.filter(t => validDate(t.date) && t.date >= a && t.date <= b && !t.otherMonth);
    const maintExtraIn = (a, b) => round2(sum(mr.maint.filter(x => validDate(x.date) && x.date >= a && x.date <= b && !x.inTrip).map(x => x.total)));
    let start = mr.mStart, end = minS(mr.mEnd, addDays(start, 7 - weekdayMon(start)));
    for (let i = 1; i <= 6 && start <= mr.mEnd; i++) {
      const w = { label: 'Неделя ' + i, start, end };
      if (start <= cd) {
        const to = minS(end, cd);
        const tr = tripsIn(start, to);
        w.income = round2(sum(tr.map(t => t.freight)));
        w.result = round2(sum(tr.map(t => t.profit)) - maintExtraIn(start, to));
        w.expenses = round2(w.income - w.result);
      }
      res.weeks.push(w);
      start = addDays(end, 1); end = minS(mr.mEnd, addDays(start, 6));
    }
    res.month = { label: 'Месяц', start: mr.mStart, end: cd, income: mr.freight, result: mr.monthProfit, expenses: mr.monthExpenses };

    // Текущая неделя
    const ws = maxS(mr.mStart, addDays(cd, 1 - weekdayMon(cd)));
    res.weekStart = ws;
    const wt = tripsIn(ws, cd);
    const initKm = n0(m.initKm);
    const initW = validDate(m.initKmDate) && m.initKmDate >= ws && m.initKmDate <= cd ? initKm : 0;
    const initM = mr.initInMonth;
    const emptyKnown = isNum(m.initEmptyKm);
    const ie = v => (v === 0 ? 0 : emptyKnown ? Math.min(m.initEmptyKm, v) : 0);
    res.weekKm = round2(sum(wt.map(t => t.totalKm)) + initW);
    res.weekIncome = round2(sum(wt.map(t => t.freight)));
    res.weekResult = round2(sum(wt.map(t => t.profit)) - maintExtraIn(ws, cd));
    res.weekLoaded = round2(sum(wt.map(t => n0(t.routeKm))) + (emptyKnown ? initW - ie(initW) : 0));
    res.weekEmpty = round2(sum(wt.map(t => n0(t.approachKm))) + ie(initW));
    res.weekUnalloc = emptyKnown || initW === 0 ? 0 : initW;
    const mt = mr.doneTrips;
    res.monthLoaded = round2(sum(mt.map(t => n0(t.routeKm))) + (emptyKnown ? initM - ie(initM) : 0));
    res.monthEmpty = round2(sum(mt.map(t => n0(t.approachKm))) + ie(initM));
    res.monthUnalloc = emptyKnown || initM === 0 ? 0 : initM;
    const tot = res.monthLoaded + res.monthEmpty + res.monthUnalloc;
    res.emptyShare = tot > 0 && res.monthUnalloc === 0 ? res.monthEmpty / tot : NaN;
    return res;
  }

  // ---------- Привязка полей ввода ----------
  function getPath(path) { return path.split('.').reduce((o, k) => (o == null ? o : o[k]), state); }
  function setPath(path, v) {
    const ks = path.split('.'); const last = ks.pop();
    const o = ks.reduce((a, k) => a[k], state); o[last] = v;
  }
  function parseNum(str) {
    const t = String(str).replace(/\s/g, '').replace(',', '.');
    if (t === '') return null;
    const v = Number(t); return isFinite(v) ? v : NaN;
  }
  function toInput(v, kind) {
    if (v == null || (typeof v === 'number' && !isFinite(v))) return '';
    if (kind === 'pct') return String(+(v * 100).toFixed(6)).replace('.', ',');
    if (kind === 'num') return String(v).replace('.', ',');
    return String(v);
  }
  function readInput(el) {
    const kind = el.dataset.kind || 'text';
    if (kind === 'num') return parseNum(el.value);
    if (kind === 'pct') { const v = parseNum(el.value); return v == null || !isFinite(v) ? v : v / 100; }
    if (kind === 'date') return el.value || null;
    return el.value;
  }
  function fillInputs(root) {
    (root || document).querySelectorAll('[data-bind]').forEach(el => {
      if (el === document.activeElement) return;
      el.value = toInput(getPath(el.dataset.bind), el.dataset.kind);
      if (el.inputMode === '' && (el.dataset.kind === 'num' || el.dataset.kind === 'pct')) el.inputMode = 'decimal';
    });
  }
  document.addEventListener('input', e => {
    const el = e.target;
    if (!el.dataset || !el.dataset.bind) return;
    setPath(el.dataset.bind, readInput(el));
    state.isExample = false;
    save();
    renderAll(true);
  });

  // ---------- Рендер ----------
  let currentTab = 'calc';
  let MR, CR, CH;

  function renderAll(fromInput) {
    MR = computeMonth(state);
    CR = computeCalc(state, MR);
    CH = computeCharts(state, MR);
    const cd = MR.calcDate;
    $('hdr-month').textContent = cd ? MONTHS[+cd.slice(5, 7) - 1] + ' ' + cd.slice(0, 4) + ' · на ' + fmtDate(cd) : '';
    renderCalc(fromInput);
    renderMonth(fromInput);
    renderTrips();
    renderMaint();
    if (currentTab === 'charts') renderCharts();
    if (!fromInput) fillInputs();
  }

  function kvRows(rows) {
    return rows.map(r => {
      if (r.sep) return `<span class="total">${esc(r.l)}</span><span class="v total ${r.cls || ''}">${r.v}</span><span class="s total">${r.s || ''}</span>`;
      return `<span>${esc(r.l)}</span><span class="v ${r.cls || ''}">${r.v}</span><span class="s">${r.s || ''}</span>`;
    }).join('');
  }

  function tile(label, big, sub, cls) {
    return `<div class="tile ${cls || ''}"><span class="l">${esc(label)}</span><span class="b">${big}</span>${sub ? `<span class="sub">${sub}</span>` : ''}</div>`;
  }

  function renderRouteTable() {
    const c = state.calc;
    $('route-body').innerHTML = c.route.map((row, i) => `<tr>
      <td><input id="rt-code-${i}" data-bind="calc.route.${i}.code" list="road-codes" placeholder="DE" style="text-transform:uppercase"></td>
      <td><input id="rt-km-${i}" class="num" data-bind="calc.route.${i}.km" data-kind="num"></td>
      <td><input id="rt-toll-${i}" class="num" data-bind="calc.route.${i}.tollKm" data-kind="num"></td>
      <td class="n" id="rt-rate-${i}"></td>
      <td><input id="rt-man-${i}" class="num" data-bind="calc.route.${i}.manual" data-kind="num"></td>
      <td class="n" id="rt-fee-${i}"></td>
      <td><button class="btn small" data-del-route="${i}" aria-label="Удалить участок">×</button></td></tr>`).join('');
    $('bridge-body').innerHTML = c.bridges.length ? c.bridges.map((row, i) => `<tr>
      <td><input id="br-code-${i}" data-bind="calc.bridges.${i}.code" list="bridge-codes" placeholder="DK-STORE" style="text-transform:uppercase"></td>
      <td><input id="br-pass-${i}" class="num" data-bind="calc.bridges.${i}.passes" data-kind="num" placeholder="1"></td>
      <td class="n" id="br-rate-${i}"></td>
      <td><input id="br-man-${i}" class="num" data-bind="calc.bridges.${i}.manual" data-kind="num"></td>
      <td class="n" id="br-fee-${i}"></td>
      <td><button class="btn small" data-del-bridge="${i}" aria-label="Удалить объект">×</button></td></tr>`).join('')
      : '<tr><td colspan="6" class="muted">Нет платных мостов и тоннелей на маршруте.</td></tr>';
    fillInputs($('route-body')); fillInputs($('bridge-body'));
  }

  function renderTariffTables() {
    const row = (kind, t, i) => `<tr><td><input id="tf-${kind}-code-${i}" data-bind="tariffs.${kind}.${i}.code" style="min-width:90px"></td>
      <td><input id="tf-${kind}-name-${i}" data-bind="tariffs.${kind}.${i}.name" style="min-width:140px"></td>
      <td><input id="tf-${kind}-rate-${i}" class="num" data-bind="tariffs.${kind}.${i}.rate" data-kind="num" style="min-width:80px"></td>
      <td class="muted" style="min-width:200px">${esc(t.note || '')}</td><td class="muted">${esc(t.src || '')}</td></tr>`;
    $('tariff-road').innerHTML = state.tariffs.road.map((t, i) => row('road', t, i)).join('');
    $('tariff-bridge').innerHTML = state.tariffs.bridge.map((t, i) => row('bridge', t, i)).join('');
    let dl = $('road-codes');
    if (!dl) { dl = document.createElement('datalist'); dl.id = 'road-codes'; document.body.appendChild(dl); }
    dl.innerHTML = state.tariffs.road.map(t => `<option value="${esc(t.code)}">${esc(t.name)}</option>`).join('');
    let db = $('bridge-codes');
    if (!db) { db = document.createElement('datalist'); db.id = 'bridge-codes'; document.body.appendChild(db); }
    db.innerHTML = state.tariffs.bridge.map(t => `<option value="${esc(t.code)}">${esc(t.name)}</option>`).join('');
    fillInputs($('tariff-road')); fillInputs($('tariff-bridge'));
  }

  function renderCalc(fromInput) {
    if (!fromInput || !$('route-body').children.length) renderRouteTable();
    if (!fromInput) renderTariffTables();
    CR.route.forEach((r, i) => {
      const rc = $('rt-rate-' + i), fe = $('rt-fee-' + i);
      if (!rc) return;
      rc.textContent = isNum(r.rate) ? rate4(r.rate) : '';
      rc.title = r.tariff ? r.tariff.name + '. ' + (r.tariff.note || '') : '';
      fe.innerHTML = r.err ? `<span class="pill bad">${esc(r.err)}</span>` : r.empty ? '' : num2(r.fee);
    });
    CR.bridges.forEach((r, i) => {
      const rc = $('br-rate-' + i), fe = $('br-fee-' + i);
      if (!rc) return;
      rc.textContent = isNum(r.rate) ? num2(r.rate) : '';
      rc.title = r.tariff ? r.tariff.name + '. ' + (r.tariff.note || '') : '';
      fe.innerHTML = r.err ? `<span class="pill bad">${esc(r.err)}</span>` : r.empty ? '' : num2(r.fee);
    });
    $('route-km').textContent = kmN(CR.routeKm);
    $('route-fee').textContent = num2(CR.routeFee);
    $('app-fee').textContent = eur(CR.approachFee);
    $('route-fee-all').textContent = eur(CR.roadFee);
    $('bridge-fee').textContent = num2(CR.bridgeFee);

    // Итог
    const good = CR.profit >= 0;
    $('calc-tiles').innerHTML = CR.hasRate ? [
      tile('Прибыль / убыток', `<span class="${signCls(CR.profit)}">${eur(CR.profit)}</span>`, 'маржа ' + pct(CR.margin), good ? 'good' : 'bad'),
      tile('Ставка клиента', eur(state.calc.clientRate), rate4(CR.ratePerKm) + ' €/км'),
      tile('К получению', eur(CR.receivable), 'сконто −' + num2(CR.skontoAmt)),
      tile('Себестоимость с премией', eur(CR.costWithPrem), rate4(CR.costPerKm) + ' €/км')
    ].join('') : tile('Себестоимость без премии', eur(CR.cost), km(CR.totalKm));
    const st = $('calc-status');
    st.textContent = CR.message || 'Маржа = прибыль ÷ сумма к получению. Все суммы в евро на одной базе НДС.';
    st.className = 'status' + (CR.message ? ' bad' : '');
    $('calc-status-pill').innerHTML = CR.hasRate && isNum(CR.profit) ? `<span class="pill ${good ? 'good' : 'bad'}">${good ? 'Рейс выгоден' : 'Рейс убыточен'}</span>` : '';

    // Предложения
    $('offer-body').innerHTML = CR.offers.map(o => `<tr><td>${esc(o.label)}</td><td class="n">${num2(o.rate)}</td><td class="n">${num2(o.perKm)}</td></tr>`).join('');

    // Структура расходов
    const parts = [
      { l: 'Дизель', v: CR.diesel, s: `${nf1.format(n0(state.month.dieselL100))} л × ${num2(state.month.dieselPrice)}`, c: 'var(--c1)' },
      { l: 'AdBlue', v: CR.adblue, s: '', c: 'var(--c5)' },
      { l: 'Зарплатные расходы', v: CR.salary, s: rate4(MR.salaryPerKm) + ' €/км', c: 'var(--c3)' },
      { l: 'Дороги, мосты и тоннели', v: CR.roads, s: '', c: 'var(--c4)' },
      { l: 'Постоянные расходы', v: CR.fixedTotal, s: rate4(MR.fixedPerKm) + ' €/км', c: 'var(--c2)' }
    ];
    const tot = sum(parts.map(p => Math.max(0, p.v)));
    $('cost-bar').innerHTML = tot > 0 ? parts.map(p => `<i style="width:${(Math.max(0, p.v) / tot * 100).toFixed(2)}%;background:${p.c}" title="${esc(p.l)}: ${eur(p.v)}"></i>`).join('') : '';
    $('cost-legend').innerHTML = parts.map(p => `<span><b style="background:${p.c}"></b>${esc(p.l)} ${tot > 0 ? nf0.format(p.v / tot * 100) + ' %' : ''}</span>`).join('');
    const rows = [
      { l: 'Дизель', v: num2(CR.diesel), s: parts[0].s },
      { l: 'AdBlue', v: num2(CR.adblue), s: `${nf1.format(n0(state.month.adblueL100))} л × ${num2(state.month.adbluePrice)}` },
      { l: 'Зарплатные расходы', v: num2(CR.salary), s: parts[2].s },
      { l: 'Дороги, мосты и тоннели', v: num2(CR.roads), s: '' }
    ];
    CR.fixedItems.forEach(f => rows.push({ l: f.name, v: num2(f.value), s: 'пост.' }));
    rows.push({ l: 'Прочие / округление', v: num2(CR.other), s: '' });
    rows.push({ l: 'Себестоимость без премии', v: num2(CR.cost), s: '€', sep: true });
    rows.push({ l: 'Премия экспедитора', v: num2(CR.premiumAmt), s: pct(state.calc.premium) });
    rows.push({ l: 'Себестоимость с премией', v: num2(CR.costWithPrem), s: '€', sep: true });
    $('cost-kv').innerHTML = kvRows(rows);
    $('calc-km-note').textContent = `общий пробег ${km(CR.totalKm)} (маршрут ${kmN(CR.routeKm)} + доезд ${kmN(CR.approachKm)})`;

    // Месячный план
    $('calc-month-date').textContent = 'на ' + fmtDate(MR.calcDate);
    $('calc-month-kv').innerHTML = kvRows([
      { l: 'Целевой общий пробег', v: km(MR.target), s: '' },
      { l: 'Фактически пройдено', v: km(MR.doneKm), s: pct(MR.planShare) },
      { l: 'Осталось до цели', v: km(MR.remainingKm), s: MR.target ? pct(MR.remainingKm / MR.target) : '' },
      { l: 'Осталось оплачиваемых дней', v: isNum(MR.daysLeft) ? String(MR.daysLeft) : '—', s: isNum(state.month.days) ? pct(MR.daysLeft / state.month.days) : '' },
      { l: 'Остаток фонда на пробег', v: eur(MR.salaryLeftByKm), s: '' },
      { l: 'Базовый фонд за все дни', v: eur(MR.salaryFull), s: '' },
      { l: 'Начислено за прошедшие дни', v: eur(MR.salaryPaid), s: '' },
      { l: 'Постоянные расходы в месяц', v: eur(MR.fixedTotal), s: '' },
      { l: 'Постоянные расходы на 1 км', v: rate4(MR.fixedPerKm), s: '€' },
      { l: 'Зарплатные расходы на 1 км', v: rate4(MR.salaryPerKm), s: '€' },
      { l: 'Постоянные и зарплата на 1 км', v: rate4(MR.costPerKm), s: '€', sep: true },
      { l: 'Фрахт по журналу', v: eur(MR.freight), s: '' },
      { l: 'К получению по журналу', v: eur(MR.receivable), s: '' }
    ]);
  }

  function renderMonth(fromInput) {
    const m = state.month;
    if (!fromInput || !$('fixed-body').children.length) {
      $('fixed-body').innerHTML = m.fixed.map((f, i) => `<tr>
        <td><input id="fx-name-${i}" data-bind="month.fixed.${i}.name"></td>
        <td><input id="fx-pln-${i}" class="num" data-bind="month.fixed.${i}.pln" data-kind="num"></td>
        <td><input id="fx-eur-${i}" class="num" data-bind="month.fixed.${i}.eur" data-kind="num"></td>
        <td class="n" id="fx-tot-${i}"></td>
        <td><button class="btn small" data-del-fixed="${i}" aria-label="Удалить статью">×</button></td></tr>`).join('');
      fillInputs($('fixed-body'));
    }
    MR.fixedRows.forEach((r, i) => { const el = $('fx-tot-' + i); if (el) el.textContent = num2(r.eurTotal); });
    $('fixed-total').textContent = num2(MR.fixedTotal);
    const st = $('month-status'); st.textContent = MR.status.text; st.className = 'status ' + MR.status.cls;
    $('month-kv').innerHTML = kvRows([
      { l: 'Осталось до целевого пробега', v: km(MR.remainingKm), s: MR.target ? pct(MR.remainingKm / MR.target) : '' },
      { l: 'Фактически пройдено', v: km(MR.doneKm), s: pct(MR.planShare) },
      { l: 'Пробег сверх целевого', v: km(MR.overKm), s: '' },
      { l: 'Осталось оплачиваемых дней', v: isNum(MR.daysLeft) ? String(MR.daysLeft) : '—', s: isNum(m.days) && isNum(MR.daysLeft) ? pct(MR.daysLeft / m.days) : '' },
      { l: 'Полная зарплата за месяц', v: eur(MR.salaryFull), s: '', sep: true },
      { l: 'Фактическая зарплата по дням', v: eur(MR.salaryPaid), s: MR.salaryFull ? pct(MR.salaryPaid / MR.salaryFull) : '' },
      { l: 'Экономия за бесплатные дни', v: eur(MR.salaryFree), s: MR.salaryFull ? pct(MR.salaryFree / MR.salaryFull) : '' },
      { l: 'Осталось зарплаты по дням', v: eur(MR.salaryLeft), s: MR.salaryFull ? pct(MR.salaryLeft / MR.salaryFull) : '' },
      { l: 'Плановая зарплата месяца', v: eur(MR.salaryPlan), s: '' },
      { l: 'Постоянные расходы на 1 км', v: rate4(MR.fixedPerKm), s: '€', sep: true },
      { l: 'Зарплатные расходы на 1 км', v: rate4(MR.salaryPerKm), s: '€' },
      { l: 'Постоянные и зарплатные на 1 км', v: rate4(MR.costPerKm), s: '€' },
      { l: 'Остаток постоянных расходов по пробегу', v: eur(MR.fixedLeft), s: '' },
      { l: 'Уже распределено постоянных расходов', v: eur(MR.fixedAllocated), s: '' }
    ]);
    $('month-kv2').innerHTML = kvRows([
      { l: 'Внесено рейсов за месяц', v: String(MR.tripCount), s: '' },
      { l: 'Фрахт по журналу', v: eur(MR.freight), s: '' },
      { l: 'К получению по журналу', v: eur(MR.receivable), s: '' },
      { l: 'Зарплатные расходы в завершённых рейсах', v: eur(MR.salaryInTrips), s: '' },
      { l: 'Зарплатные расходы на оставшийся пробег', v: eur(MR.salaryLeftByKm), s: '' },
      { l: 'Разница: расходы − зарплата по дням', v: eur(MR.salaryDiff), s: '', cls: signCls(MR.salaryDiff) },
      { l: 'Фактическая зарплата по дням на 1 км', v: rate4(MR.salaryPaidPerKm), s: '€' }
    ]);
  }

  function renderTrips() {
    const list = MR.trips.slice().sort((a, b) => String(a.date).localeCompare(String(b.date)));
    $('trips-tiles').innerHTML = [
      tile('Рейсов за месяц', String(MR.tripCount), 'по ' + fmtDate(MR.calcDate)),
      tile('Пробег по журналу', km(MR.journalKm), MR.initInMonth ? '+ начальный ' + km(MR.initInMonth) : 'всего ' + km(MR.doneKm)),
      tile('Фрахт', eur(MR.freight), 'к получению ' + num2(MR.receivable)),
      tile('Прибыль рейсов', `<span class="${signCls(MR.tripProfit)}">${eur(MR.tripProfit)}</span>`, MR.receivable > 0 ? 'маржа ' + pct(MR.tripProfit / MR.receivable) : '', MR.tripProfit >= 0 ? 'good' : 'bad')
    ].join('');
    const st = $('trips-status');
    if (MR.badTrips) { st.textContent = 'Заполните себестоимость и фрахт у ' + MR.badTrips + ' рейс(ов) выбранного месяца.'; st.className = 'status bad'; }
    else { st.textContent = 'Зарплатные и постоянные расходы рейса считаются автоматически: ставка на 1 км из месячного плана × общий пробег рейса.'; st.className = 'status'; }
    $('trips-body').innerHTML = list.length ? list.map(t => {
      const pill = t.otherMonth ? '<span class="pill">Другой месяц</span>' : t.future ? `<span class="pill info">${esc(t.status)}</span>` : isNum(t.profit) ? `<span class="pill ${t.tollsEstimated ? '' : 'good'}">${esc(t.status)}</span>` : `<span class="pill bad">${esc(t.status)}</span>`;
      return `<tr class="click${t.future || t.otherMonth ? ' future' : ''}" data-trip="${t.id}" tabindex="0">
        <td class="mono">${fmtDate(t.date)}</td><td>${esc(t.no)}</td><td style="min-width:200px">${esc(t.route)}</td>
        <td class="n">${kmN(t.routeKm)}</td><td class="n">${kmN(t.approachKm)}</td><td class="n">${kmN(t.totalKm)}</td>
        <td class="n">${num2(t.freight)}</td><td class="n">${isNum(t.skonto) ? pct(t.skonto) : ''}</td><td class="n">${num2(t.receivable)}</td>
        <td class="n">${num2(t.ratePerKm)}</td><td class="n ${signCls(t.margin)}">${isNum(t.margin) ? pct(t.margin) : ''}</td><td class="n ${signCls(t.profit)}">${num2(t.profit)}</td>
        <td class="n">${num2(t.costNoPrem)}</td><td class="n">${isNum(t.premium) ? pct(t.premium) : ''}</td><td class="n">${num2(t.costWithPrem)}</td>
        <td class="n" title="${t.dieselFact ? 'Фактическая сумма' : 'По норме расхода'}">${num2(t.diesel)}${t.dieselFact ? '*' : ''}</td><td class="n">${num2(t.adblue)}</td><td class="n">${num2(t.salary)}</td><td class="n">${num2(t.fixed)}</td>
        <td class="n" title="${t.tollsEstimated ? 'Оценка по среднему тарифу' : 'Точные сборы'}">${num2(t.tolls)}${t.tollsEstimated ? '~' : ''}</td><td class="n">${num2(t.other)}</td><td>${pill}</td></tr>`;
    }).join('') : '<tr><td colspan="22" class="muted">Журнал пуст. Добавьте рейс кнопкой «+ Рейс» или из калькулятора.</td></tr>';
    const d = MR.doneTrips;
    $('trips-foot').innerHTML = `<tr><td colspan="3">Итого за месяц по ${fmtDate(MR.calcDate)}</td>
      <td class="n">${kmN(sum(d.map(t => t.routeKm)))}</td><td class="n">${kmN(sum(d.map(t => t.approachKm)))}</td><td class="n">${kmN(MR.journalKm)}</td>
      <td class="n">${num2(MR.freight)}</td><td></td><td class="n">${num2(MR.receivable)}</td><td class="n">${MR.journalKm ? num2(MR.freight / MR.journalKm) : ''}</td>
      <td class="n">${MR.receivable ? pct(MR.tripProfit / MR.receivable) : ''}</td><td class="n ${signCls(MR.tripProfit)}">${num2(MR.tripProfit)}</td>
      <td class="n">${num2(sum(d.map(t => t.costNoPrem)))}</td><td></td><td class="n">${num2(MR.tripCost)}</td>
      <td class="n">${num2(sum(d.map(t => t.diesel)))}</td><td class="n">${num2(sum(d.map(t => t.adblue)))}</td><td class="n">${num2(MR.salaryInTrips)}</td>
      <td class="n">${num2(sum(d.map(t => t.fixed)))}</td><td class="n">${num2(sum(d.map(t => t.tolls)))}</td><td class="n">${num2(sum(d.map(t => t.other)))}</td><td></td></tr>`;
  }

  function renderMaint() {
    $('maint-tiles').innerHTML = [
      tile('Всего ТО и ремонтов за месяц', eur(MR.maintTotal), ''),
      tile('Уже учтено в рейсах', eur(MR.maintInTrip), ''),
      tile('Дополнительно к расходам', eur(MR.maintExtra), 'вычитается из прибыли месяца')
    ].join('');
    const list = MR.maint.slice().sort((a, b) => String(a.date).localeCompare(String(b.date)));
    $('maint-body').innerHTML = list.length ? list.map(x => `<tr class="click" data-maint="${x.id}" tabindex="0">
      <td class="mono">${fmtDate(x.date)}</td><td class="n">${isNum(x.odo) ? nf0.format(x.odo) : ''}</td><td>${esc(x.cat)}</td><td style="min-width:200px">${esc(x.work)}</td><td>${esc(x.service)}</td>
      <td class="n">${num2(x.pln)}</td><td class="n">${rate4(x.rate)}</td><td class="n">${num2(x.eur)}</td>
      <td class="n">${isNum(x.total) ? num2(x.total) : `<span class="pill bad">${esc(x.err || 'Проверьте')}</span>`}</td>
      <td>${x.inTrip ? '<span class="pill good">Да</span>' : '<span class="pill">Нет</span>'}</td><td>${esc(x.tripNo)}</td></tr>`).join('')
      : '<tr><td colspan="11" class="muted">Записей нет. Добавьте расход кнопкой «+ Запись».</td></tr>';
  }

  // ---------- Диаграммы ----------
  const charts = {};
  function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
  function makeChart(id, cfg) {
    if (!window.Chart) return;
    if (charts[id]) charts[id].destroy();
    const fg = cssVar('--muted'), grid = cssVar('--line');
    Chart.defaults.font.family = cssVar('--font-body') || 'sans-serif';
    Chart.defaults.color = fg;
    cfg.options = Object.assign({ locale: 'ru-RU', responsive: true, maintainAspectRatio: false, animation: false,
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, boxHeight: 10 } }, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${nf2.format(c.parsed.y ?? c.parsed.x)}` } } } }, cfg.options || {});
    const sc = cfg.options.scales || {};
    Object.values(sc).forEach(a => { a.grid = Object.assign({ color: grid }, a.grid || {}); a.ticks = Object.assign({ color: fg }, a.ticks || {}); });
    charts[id] = new Chart($(id), cfg);
  }

  function renderCharts() {
    const cd = MR.calcDate;
    $('charts-lead').textContent = cd ? `С 1 по ${fmtDate(cd)}. Неделя — с понедельника (не ранее 1-го числа) по выбранную дату.` : '';
    const res = CH.month || {};
    $('charts-tiles').innerHTML = [
      tile('Доход по журналу', eur(res.income), 'фрахт до сконто'),
      tile('Расходы всего', eur(res.expenses), 'вкл. сконто, премию, ТО'),
      tile('Прибыль / убыток', `<span class="${signCls(res.result)}">${eur(res.result)}</span>`, res.income ? 'от дохода ' + pct(res.result / res.income) : '', res.result >= 0 ? 'good' : 'bad'),
      tile('Выполнено плана', pct(MR.planShare), `пройдено ${km(MR.doneKm)}, осталось ${km(MR.remainingKm)}`)
    ].join('');
    const st = $('charts-status');
    st.textContent = MR.badTrips ? 'Проверьте даты, фрахт и расходы рейсов: часть рейсов без расчёта прибыли.' : `Текущая неделя: ${fmtDate(CH.weekStart)}–${fmtDate(cd)}. Доход ${eur(CH.weekIncome)}, итог ${eur(CH.weekResult)}.`;
    st.className = 'status' + (MR.badTrips ? ' bad' : '');

    const ws = CH.weeks;
    $('weeks-body').innerHTML = ws.concat([CH.month]).map(w => `<tr${w.label === 'Месяц' ? ' style="font-weight:600"' : ''}><td>${w.label}</td><td class="mono">${fmtDate(w.start).slice(0, 5)}–${fmtDate(w.end).slice(0, 5)}</td>
      <td class="n">${num2(w.income)}</td><td class="n">${num2(w.expenses)}</td><td class="n ${signCls(w.result)}">${num2(w.result)}</td></tr>`).join('');

    const labels = ws.map(w => `${w.label} (${fmtDate(w.start).slice(0, 5)}–${fmtDate(w.end).slice(0, 5)})`);
    makeChart('ch-weeks', {
      type: 'bar',
      data: { labels: ws.map(w => w.label), datasets: [
        { label: 'Доход', data: ws.map(w => w.income ?? null), backgroundColor: cssVar('--c1') },
        { label: 'Расходы', data: ws.map(w => w.expenses ?? null), backgroundColor: cssVar('--c2') },
        { label: 'Прибыль', data: ws.map(w => (isNum(w.result) ? Math.max(0, w.result) : null)), backgroundColor: cssVar('--c3') },
        { label: 'Убыток', data: ws.map(w => (isNum(w.result) ? Math.max(0, -w.result) : null)), backgroundColor: cssVar('--c4') }
      ] },
      options: { scales: { x: { grid: { display: false } }, y: { beginAtZero: true } }, plugins: { tooltip: { callbacks: { title: it => labels[it[0].dataIndex], label: c => `${c.dataset.label}: ${nf2.format(c.parsed.y)} €` } } } }
    });

    const wkLabel = `Неделя ${fmtDate(CH.weekStart).slice(0, 5)}–${fmtDate(cd).slice(0, 5)}`;
    const mLabel = `Месяц по ${fmtDate(cd).slice(0, 5)}`;
    makeChart('ch-km', {
      type: 'bar',
      data: { labels: [wkLabel, mLabel], datasets: [
        { label: 'Пройдено', data: [CH.weekKm, MR.doneKm], backgroundColor: cssVar('--c1'), stack: 's' },
        { label: 'До цели месяца', data: [null, MR.remainingKm], backgroundColor: cssVar('--c2'), stack: 's' }
      ] },
      options: { scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true } } }
    });

    const ds = [
      { label: 'С грузом', data: [CH.weekLoaded, CH.monthLoaded], backgroundColor: cssVar('--c1'), stack: 's' },
      { label: 'Порожний', data: [CH.weekEmpty, CH.monthEmpty], backgroundColor: cssVar('--c4'), stack: 's' }
    ];
    if (CH.monthUnalloc || CH.weekUnalloc) ds.push({ label: 'Без разбивки', data: [CH.weekUnalloc, CH.monthUnalloc], backgroundColor: cssVar('--c2'), stack: 's' });
    makeChart('ch-load', { type: 'bar', data: { labels: [wkLabel, mLabel], datasets: ds },
      options: { scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true } } } });
    $('load-note').textContent = isNum(CH.emptyShare) ? `Доля порожнего пробега за месяц: ${pct(CH.emptyShare)}. Порожний = доезды; км маршрута считаются с грузом.`
      : CH.monthUnalloc ? 'Укажите порожний пробег начального остатка на вкладке «Месяц», чтобы посчитать долю порожнего.' : 'За выбранный период пробег не внесён.';

    const dt = MR.doneTrips.filter(t => isNum(t.profit));
    makeChart('ch-trips', {
      type: 'bar',
      data: { labels: dt.map(t => fmtDate(t.date).slice(0, 5) + (t.no ? ' · ' + t.no : '')), datasets: [
        { label: 'Прибыль / убыток', data: dt.map(t => t.profit), backgroundColor: dt.map(t => (t.profit >= 0 ? cssVar('--c3') : cssVar('--c4'))) }
      ] },
      options: { indexAxis: 'y', plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => `${nf2.format(c.parsed.x)} €, маржа ${pct(dt[c.dataIndex].margin)}` } } },
        scales: { x: { beginAtZero: true }, y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } } } }
    });
  }

  // ---------- Вкладки ----------
  function showTab(tab) {
    currentTab = tab;
    document.querySelectorAll('nav.tabs button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === tab)));
    document.querySelectorAll('section.view').forEach(s => { s.hidden = s.id !== 'view-' + tab; });
    try { localStorage.setItem(STORE_KEY + ':tab', tab); } catch (e) { /* ignore */ }
    if (location.hash.slice(1) !== tab) history.replaceState(null, '', '#' + tab);
    if (tab === 'charts') renderCharts();
  }
  document.querySelectorAll('nav.tabs button').forEach(b => b.addEventListener('click', () => showTab(b.dataset.tab)));
  window.addEventListener('hashchange', () => { const h = location.hash.slice(1); if ($('view-' + h)) showTab(h); });

  // ---------- Кнопки таблиц ----------
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    const d = b.dataset;
    if (d.delRoute != null) { state.calc.route.splice(+d.delRoute, 1); changed(); }
    else if (d.delBridge != null) { state.calc.bridges.splice(+d.delBridge, 1); changed(); }
    else if (d.delFixed != null) { state.month.fixed.splice(+d.delFixed, 1); changed(); }
  });
  function changed() { state.isExample = false; save(); renderAll(); }
  $('add-route').onclick = () => { state.calc.route.push({ code: '', km: null, tollKm: null, manual: null }); changed(); };
  $('add-bridge').onclick = () => { state.calc.bridges.push({ code: '', passes: null, manual: null }); changed(); };
  $('add-fixed').onclick = () => { state.month.fixed.push({ name: 'Новая статья', pln: null, eur: null }); changed(); };
  $('add-tariff-road').onclick = () => { state.tariffs.road.push({ code: 'НОВЫЙ', name: '', rate: null, note: 'Свой код', src: '' }); changed(); };
  $('add-tariff-bridge').onclick = () => { state.tariffs.bridge.push({ code: 'НОВЫЙ', name: '', rate: null, note: 'Свой объект', src: '' }); changed(); };
  $('reset-tariffs').onclick = () => confirmBox('Вернуть исходные тарифы? Ваши изменения в справочниках будут потеряны.', 'Вернуть', () => { state.tariffs = JSON.parse(JSON.stringify(window.DEFAULT_TARIFFS)); changed(); toast('Тарифы восстановлены'); });

  $('calc-to-log').onclick = () => {
    const c = state.calc;
    if (!(CR.totalKm > 0) || !isNum(c.clientRate)) { toast('Сначала укажите маршрут и ставку клиента'); return; }
    state.trips.push({ id: uid(), date: validDate(c.date) ? c.date : (MR.calcDate || todayISO()), no: '', route: c.title || '',
      routeKm: CR.routeKm, approachKm: CR.approachKm, freight: c.clientRate, skonto: c.skonto, premium: c.premium,
      tolls: isNum(CR.roads) ? CR.roads : null, other: c.other || null, comment: 'Из калькулятора' });
    changed(); toast('Рейс добавлен в журнал'); showTab('trips');
  };

  // ---------- Диалоги ----------
  const dlg = $('dlg'), form = $('dlg-form');
  function openForm(title, fields, values, onSave, onDelete) {
    form.innerHTML = `<h2>${esc(title)}</h2><div class="fields">${fields.map(f => {
      const v = values[f.k];
      if (f.type === 'check') return `<label class="f" style="grid-template-columns:auto 1fr;align-items:center;gap:8px;color:var(--fg)"><input type="checkbox" id="fd-${f.k}" ${v ? 'checked' : ''} style="width:auto">${esc(f.l)}</label>`;
      const kind = f.type || 'text';
      const attrs = kind === 'date' ? 'type="date"' : kind === 'num' || kind === 'pct' ? 'class="num" inputmode="decimal"' : '';
      return `<label class="f" ${f.wide ? 'style="grid-column:1/-1"' : ''}>${esc(f.l)}<input id="fd-${f.k}" ${attrs} data-kind="${kind}" value="${esc(toInput(v, kind))}" placeholder="${esc(f.ph || '')}"></label>`;
    }).join('')}</div><p class="note" id="fd-err"></p>
    <div class="row-actions" style="justify-content:space-between"><div>${onDelete ? '<button type="button" class="btn danger" id="fd-del">Удалить</button>' : ''}</div>
    <div class="row-actions"><button type="button" class="btn" id="fd-cancel">Отмена</button><button type="submit" class="btn primary">Сохранить</button></div></div>`;
    form.onsubmit = e => {
      e.preventDefault();
      const out = {};
      for (const f of fields) {
        const el = $('fd-' + f.k);
        if (f.type === 'check') { out[f.k] = el.checked; continue; }
        const kind = f.type || 'text';
        let v = kind === 'num' || kind === 'pct' ? parseNum(el.value) : el.value.trim();
        if (kind === 'pct' && isNum(v)) v = v / 100;
        if (Number.isNaN(v)) { $('fd-err').textContent = `Поле «${f.l}»: введите число.`; $('fd-err').className = 'note neg'; el.focus(); return; }
        if (kind === 'date') v = v || null;
        out[f.k] = v;
      }
      if (fields.some(f => f.req && (out[f.k] == null || out[f.k] === ''))) { $('fd-err').textContent = 'Заполните поля: ' + fields.filter(f => f.req).map(f => f.l).join(', ') + '.'; $('fd-err').className = 'note neg'; return; }
      onSave(out); dlg.close();
    };
    $('fd-cancel').onclick = () => dlg.close();
    if (onDelete) $('fd-del').onclick = () => { onDelete(); dlg.close(); };
    dlg.showModal();
  }
  function confirmBox(text, okLabel, onOk) {
    form.innerHTML = `<h2>Подтвердите действие</h2><p style="margin:0">${esc(text)}</p><div class="row-actions" style="justify-content:flex-end"><button type="button" class="btn" id="cf-no">Отмена</button><button type="button" class="btn primary" id="cf-ok">${esc(okLabel)}</button></div>`;
    form.onsubmit = e => e.preventDefault();
    $('cf-no').onclick = () => dlg.close();
    $('cf-ok').onclick = () => { dlg.close(); onOk(); };
    dlg.showModal();
  }

  const TRIP_FIELDS = [
    { k: 'date', l: 'Дата завершения', type: 'date', req: true }, { k: 'no', l: '№ рейса' },
    { k: 'route', l: 'Маршрут', wide: true, ph: 'DE 24103 Kiel — BE 6700 Arlon' },
    { k: 'routeKm', l: 'Маршрут, км', type: 'num', req: true }, { k: 'approachKm', l: 'Доезд, км', type: 'num' },
    { k: 'freight', l: 'Фрахт, €', type: 'num', req: true }, { k: 'skonto', l: 'Сконто, % (пусто = 0)', type: 'pct' },
    { k: 'premium', l: 'Премия экспедитора, %', type: 'pct' }, { k: 'tolls', l: 'Сборы точно, € (пусто = средний тариф)', type: 'num' },
    { k: 'diesel', l: 'Дизель факт, € (пусто = по норме)', type: 'num' },
    { k: 'other', l: 'Прочие расходы, €', type: 'num' }, { k: 'comment', l: 'Комментарий', wide: true }
  ];
  function editTrip(id) {
    const t = state.trips.find(x => x.id === id);
    const vals = t || { date: MR.calcDate, premium: 0.05, skonto: 0.03 };
    openForm(t ? 'Рейс' : 'Новый рейс', TRIP_FIELDS, vals, out => {
      if (t) Object.assign(t, out); else state.trips.push(Object.assign({ id: uid() }, out));
      changed(); toast(t ? 'Рейс сохранён' : 'Рейс добавлен');
    }, t ? () => { state.trips = state.trips.filter(x => x.id !== id); changed(); toast('Рейс удалён'); } : null);
  }
  const MAINT_FIELDS = [
    { k: 'date', l: 'Дата', type: 'date', req: true }, { k: 'odo', l: 'Одометр, км', type: 'num' },
    { k: 'cat', l: 'Категория', ph: 'ТО, шины, ремонт…' }, { k: 'service', l: 'Сервис' },
    { k: 'work', l: 'Работы / детали', wide: true },
    { k: 'pln', l: 'Сумма, PLN', type: 'num' }, { k: 'rate', l: 'PLN за 1 €', type: 'num' }, { k: 'eur', l: 'Сумма, €', type: 'num' },
    { k: 'tripNo', l: '№ рейса' }, { k: 'inTrip', l: 'Уже в расходах рейса', type: 'check' }, { k: 'comment', l: 'Комментарий', wide: true }
  ];
  function editMaint(id) {
    const x = state.maint.find(m => m.id === id);
    const vals = x || { date: MR.calcDate, rate: state.month.plnRate, inTrip: false };
    openForm(x ? 'ТО / ремонт' : 'Новая запись ТО', MAINT_FIELDS, vals, out => {
      if (x) Object.assign(x, out); else state.maint.push(Object.assign({ id: uid() }, out));
      changed(); toast('Запись сохранена');
    }, x ? () => { state.maint = state.maint.filter(m => m.id !== id); changed(); toast('Запись удалена'); } : null);
  }
  $('add-trip').onclick = () => editTrip(null);
  $('add-maint').onclick = () => editMaint(null);
  function rowOpen(e) {
    const tr = e.target.closest('tr[data-trip], tr[data-maint]'); if (!tr) return;
    if (e.type === 'keydown' && e.key !== 'Enter') return;
    if (tr.dataset.trip) editTrip(tr.dataset.trip); else editMaint(tr.dataset.maint);
  }
  ['click', 'keydown'].forEach(ev => { $('trips-body').addEventListener(ev, rowOpen); $('maint-body').addEventListener(ev, rowOpen); });

  // ---------- Экспорт / импорт ----------
  function download(name, text, type) {
    try {
      const blob = new Blob([text], { type });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    } catch (e) { /* ignore */ }
  }
  $('trips-csv').onclick = () => {
    const head = ['Дата', '№ рейса', 'Маршрут', 'Маршрут км', 'Доезд км', 'Общий км', 'Фрахт €', 'Сконто %', 'К получению €', '€/км', 'Маржа %', 'Прибыль €', 'Себест. без премии €', 'Премия %', 'Себест. с премией €', 'Дизель €', 'AdBlue €', 'Зарплата €', 'Постоянные €', 'Сборы €', 'Прочие €', 'Статус', 'Комментарий'];
    const f = v => (isNum(v) ? String(Math.round(v * 10000) / 10000).replace('.', ',') : '');
    const q = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const rows = MR.trips.map(t => [fmtDate(t.date), q(t.no), q(t.route), f(t.routeKm), f(t.approachKm), f(t.totalKm), f(t.freight), f(isNum(t.skonto) ? t.skonto * 100 : null), f(t.receivable), f(t.ratePerKm), f(isNum(t.margin) ? t.margin * 100 : null), f(t.profit), f(t.costNoPrem), f(isNum(t.premium) ? t.premium * 100 : null), f(t.costWithPrem), f(t.diesel), f(t.adblue), f(t.salary), f(t.fixed), f(t.tolls), f(t.other), q(t.status), q(t.comment)].join(';'));
    const csv = '﻿' + head.join(';') + '\n' + rows.join('\n');
    download('reisy-scania-TKI7844U.csv', csv, 'text/csv;charset=utf-8');
    showIO('CSV журнала рейсов', csv.slice(1), false);
  };
  function showIO(title, text, importable) {
    form.innerHTML = `<h2>${esc(title)}</h2>
      <p class="note">${importable ? 'Скопируйте текст, чтобы сохранить резервную копию, или вставьте сюда сохранённую копию и нажмите «Загрузить из текста».' : 'Если файл не скачался, скопируйте текст.'}</p>
      <textarea class="io" id="io-text" spellcheck="false">${esc(text)}</textarea>
      <div class="row-actions" style="justify-content:space-between">
        <div class="row-actions">${importable ? '<button type="button" class="btn" id="io-load">Загрузить из текста</button><button type="button" class="btn" id="io-dl">Скачать .json</button><button type="button" class="btn danger" id="io-reset">Сбросить к примеру</button><button type="button" class="btn danger" id="io-clear">Очистить журнал</button>' : ''}</div>
        <div class="row-actions"><button type="button" class="btn" id="io-copy">Копировать</button><button type="button" class="btn primary" id="io-close">Готово</button></div>
      </div>`;
    form.onsubmit = e => e.preventDefault();
    $('io-close').onclick = () => dlg.close();
    $('io-copy').onclick = () => {
      const ta = $('io-text');
      const p = navigator.clipboard && navigator.clipboard.writeText(ta.value);
      if (p) p.then(() => toast('Скопировано'), () => { ta.select(); toast('Выделено — нажмите Ctrl+C'); }); else { ta.select(); }
    };
    if (importable) {
      $('io-dl').onclick = () => download('bortovoj-zhurnal-scania-TKI7844U-' + (MR.calcDate || 'data') + '.json', JSON.stringify(Object.assign({ vehicle: VEHICLE, driver: DRIVER }, state), null, 1), 'application/json');
      $('io-load').onclick = () => {
        try { const s = JSON.parse($('io-text').value); if (!s.month || !s.trips) throw new Error('fmt'); if (s.vehicle && s.vehicle !== VEHICLE) { toast('Это копия для другого автомобиля (' + s.vehicle + '). Загрузите копию ' + VEHICLE + '.'); return; } delete s.vehicle; delete s.driver; state = mergeDefaults(s); save(); renderAll(); dlg.close(); toast('Данные загружены'); }
        catch (e) { toast('Не удалось прочитать текст: нужен JSON, сохранённый этим приложением'); }
      };
      $('io-reset').onclick = () => { dlg.close(); confirmBox('Заменить все данные примером? Текущие рейсы и настройки будут удалены.', 'Сбросить', () => { state = defaultState(); save(); renderAll(); toast('Загружен пример'); }); };
      $('io-clear').onclick = () => { dlg.close(); confirmBox('Удалить все рейсы и записи ТО? Настройки месяца и калькулятора останутся.', 'Очистить', () => { state.trips = []; state.maint = []; changed(); toast('Журнал очищен'); }); };
    }
    dlg.showModal();
  }
  $('btn-data').onclick = () => showIO('Данные: ' + VEHICLE + ', водитель ' + DRIVER, JSON.stringify(Object.assign({ vehicle: VEHICLE, driver: DRIVER }, state), null, 1), true);

  // Импорт рабочей книги Excel (структура исходного файла)
  $('xlsx-file').addEventListener('change', async e => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    if (!window.XLSX) { toast('Библиотека Excel не загрузилась. Проверьте интернет.'); return; }
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const res = importWorkbook(wb);
      state = mergeDefaults(res); state.isExample = false; save(); renderAll();
      toast(`Импортировано: рейсов ${state.trips.length}, записей ТО ${state.maint.length}`);
    } catch (err) { console.error(err); toast('Не удалось прочитать файл: ' + err.message); }
  });

  function importWorkbook(wb) {
    const sh = name => wb.Sheets[name] || null;
    const v = (ws, a) => { if (!ws) return null; const c = ws[a]; return c ? c.v : null; };
    const num = (ws, a) => { const x = v(ws, a); return typeof x === 'number' ? x : (typeof x === 'string' && x.trim() !== '' && isFinite(+x.replace(',', '.')) ? +x.replace(',', '.') : null); };
    const str = (ws, a) => { const x = v(ws, a); return x == null ? '' : String(x).trim(); };
    const date = (ws, a) => { const x = v(ws, a); if (typeof x !== 'number' || x <= 0) return null; return fromDate(new Date(Math.round((x - 25569) * 86400000))); };
    const M = sh('Месяц'), K = sh('Калькулятор'), R = sh('Рейсы'), T = sh('ТО и ремонты');
    if (!M && !R) throw new Error('нет листов «Месяц» и «Рейсы»');
    const s = defaultState();
    s.trips = []; s.maint = [];
    if (M) {
      const m = s.month;
      m.calcDate = date(M, 'E5') || m.calcDate; m.days = num(M, 'E6'); m.targetKm = num(M, 'E7'); m.driverDaily = num(M, 'E9');
      m.paidDays = num(M, 'E10'); m.freeDays = num(M, 'E11');
      m.fixed = [15, 16, 17, 18, 19].map(r => ({ name: str(M, 'B' + r) || 'Статья', pln: num(M, 'D' + r), eur: num(M, 'E' + r) }));
      m.roadAvg = num(M, 'E36'); m.dieselL100 = num(M, 'E38'); m.dieselPrice = num(M, 'E39'); m.adblueL100 = num(M, 'E40'); m.adbluePrice = num(M, 'E41');
    }
    if (K) {
      const c = s.calc;
      const pr = num(K, 'G15'); if (pr) s.month.plnRate = pr;
      c.title = str(K, 'B3'); c.skonto = num(K, 'G6') ?? 0; c.clientRate = num(K, 'G8');
      c.margins = [num(K, 'B8'), num(K, 'B9'), num(K, 'B10')];
      c.approachKm = num(K, 'C12'); c.approachTollKm = num(K, 'D44'); c.approachRate = num(K, 'E44'); c.approachManual = num(K, 'F44');
      c.other = num(K, 'E21') ?? 0; c.premium = num(K, 'E27') ?? 0;
      c.route = [];
      for (let r = 31; r <= 42; r++) { const code = str(K, 'B' + r); if (code || num(K, 'C' + r) != null) c.route.push({ code, km: num(K, 'C' + r), tollKm: num(K, 'D' + r), manual: num(K, 'F' + r) }); }
      if (!c.route.length) c.route.push({ code: '', km: null, tollKm: null, manual: null });
      c.bridges = [];
      for (let r = 70; r <= 77; r++) { const code = str(K, 'B' + r); if (code || num(K, 'E' + r) != null) c.bridges.push({ code, passes: num(K, 'C' + r), manual: num(K, 'E' + r) }); }
      const road = [], bridge = [];
      for (let r = 31; r <= 62; r++) { const code = str(K, 'I' + r); if (code) road.push({ code, name: str(K, 'J' + r), rate: num(K, 'K' + r), note: str(K, 'L' + r), src: str(K, 'M' + r) }); }
      for (let r = 70; r <= 97; r++) { const code = str(K, 'I' + r); if (code) bridge.push({ code, name: str(K, 'J' + r), rate: num(K, 'K' + r), note: str(K, 'L' + r), src: str(K, 'M' + r) }); }
      if (road.length) s.tariffs = { road, bridge };
      c.date = s.month.calcDate;
    }
    if (R) {
      s.month.initKmDate = date(R, 'D6') || s.month.initKmDate; s.month.initKm = num(R, 'D7') ?? 0; s.month.initEmptyKm = num(R, 'D13');
      const range = XLSX.utils.decode_range(R['!ref']);
      for (let r = 15; r <= range.e.r + 1; r++) {
        const d = date(R, 'B' + r);
        const any = d || str(R, 'C' + r) || str(R, 'D' + r) || num(R, 'E' + r) != null || num(R, 'H' + r) != null;
        if (!any) continue;
        s.trips.push({ id: uid(), date: d, no: str(R, 'C' + r), route: str(R, 'D' + r), routeKm: num(R, 'E' + r), approachKm: num(R, 'F' + r),
          freight: num(R, 'H' + r), skonto: num(R, 'I' + r), premium: num(R, 'O' + r), comment: str(R, 'Q' + r), tolls: num(R, 'R' + r), other: num(R, 'S' + r),
          diesel: R['T' + r] && !R['T' + r].f && typeof R['T' + r].v === 'number' ? R['T' + r].v : null });
      }
    }
    if (T) {
      const range = XLSX.utils.decode_range(T['!ref']);
      for (let r = 15; r <= range.e.r + 1; r++) {
        const d = date(T, 'B' + r);
        if (!d && num(T, 'G' + r) == null && num(T, 'I' + r) == null && !str(T, 'E' + r)) continue;
        s.maint.push({ id: uid(), date: d, odo: num(T, 'C' + r), cat: str(T, 'D' + r), work: str(T, 'E' + r), service: str(T, 'F' + r),
          pln: num(T, 'G' + r), rate: num(T, 'H' + r), eur: num(T, 'I' + r), inTrip: str(T, 'K' + r) === 'Да', tripNo: str(T, 'L' + r), comment: str(T, 'M' + r) });
      }
    }
    return s;
  }

  // ---------- Тост ----------
  let toastTimer;
  function toast(msg) { const t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 2600); }

  // Перерисовка графиков при смене темы
  try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (currentTab === 'charts') renderCharts(); }); } catch (e) { /* ignore */ }
  new MutationObserver(() => { if (currentTab === 'charts') renderCharts(); }).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

  // Для проверки расчётов из консоли / тестов
  window.TripCalc = { computeMonth, computeCalc, computeCharts, importWorkbook, getState: () => state, setState: s => { state = mergeDefaults(s); renderAll(); } };

  // ---------- Старт ----------
  renderAll();
  let start = location.hash.slice(1);
  if (!$('view-' + start)) { try { start = localStorage.getItem(STORE_KEY + ':tab') || 'calc'; } catch (e) { start = 'calc'; } }
  showTab($('view-' + start) ? start : 'calc');
  if (state.isExample) toast('Открыт пример. Загрузите свой файл .xlsx или вводите данные вручную.');
})();
