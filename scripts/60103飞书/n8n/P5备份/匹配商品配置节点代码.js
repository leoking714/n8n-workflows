// ============================================================================
// 60103 匹配 + 数据完整性检查（商品配置 / 运价 / 澳洲分区），只输出 1 条消息
// ============================================================================
const num = (v) => {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
};
const blank = (v) => v === null || v === undefined || String(v).trim() === '';

// ---- 1. 商品配置（字段名自动识别，遍历全部记录取并集）----
const cfgItems = $('解析整个商品配置表生成对象数组').all().map(i => i.json);
if (!cfgItems.length) throw new Error('商品配置表没有读到数据');

// 字段名要同时看「表头定义」（获取所有字段名）和「记录里实际有的 key」
// —— 否则某些字段整列没填时，记录里根本没有这个 key，就会漏检
let fieldNames = [];
try {
  const fl = $('获取所有字段名').first().json;
  fieldNames = (fl.items || []).map(x => x.field_name).filter(Boolean);
} catch (e) { /* 没有这个节点就只用记录 key */ }

const allKeys = [...new Set([...fieldNames, ...cfgItems.flatMap(r => Object.keys(r))])]
  .filter(k => k !== 'record_id');
const findKey = (re, notRe) => allKeys.find(k => re.test(k) && (!notRe || !notRe.test(k))) || null;

const skuKey = findKey(/sku/i, null) && (findKey(/前缀|prefix/i) || findKey(/sku/i));
if (!skuKey) throw new Error('商品配置表找不到 SKU 前缀字段，实际字段：' + allKeys.join(' / '));

// 必填：没有就算不出成本；选填：有默认值
const REQ = [
  { label: '采购价', key: findKey(/采购价|采购单价/, /偏远/) },
  { label: '国内运费', key: findKey(/国内运费/, /偏远/) },
  { label: '重量', key: findKey(/重量|weight/i, null) },
].filter(x => x.key);

if (REQ.length < 2) {
  throw new Error('商品配置表里没识别到价格/重量字段（采购价、国内运费、重量），'
    + '实际表头：' + allKeys.join(' / '));
}

const OPT = [
  { label: '服务费', key: findKey(/毛利|服务费/, /美国/) },
  { label: '每计费单位件数', key: findKey(/每计费单位件数|计费件数/, null) },
].filter(x => x.key);
const weightKey = (REQ.find(x => x.label === '重量') || {}).key || null;

const K = {
  buy: findKey(/^采购价|采购价$/, /偏远/),
  buyR: findKey(/偏远.*采购价/, null),
  dom: findKey(/^国内运费|国内运费$/, /偏远/),
  domR: findKey(/偏远.*国内运费/, null),
  svc: findKey(/服务费$|毛利/, /美国/),
  svcUS: findKey(/服务费_美国|美国.*服务费/, null),
  ovC: findKey(/运价覆盖_国家/, null),
  ovP: findKey(/运价覆盖_公斤价/, null),
  ovF: findKey(/运价覆盖_挂号费/, null),
  fixed: findKey(/账单固定金额/, null),
  pack: findKey(/每计费单位件数|计费件数/, null),
};

const cfgMap = new Map();      // 前缀 → 配置 + ok
const cfgIncomplete = [];      // [{前缀, 缺:[...]}]
for (const r of cfgItems) {
  const cell = String(r[skuKey] ?? '').trim();
  if (!cell) continue;
  const isNonGoodsPrefix = /^SHIP/i.test(cell);          // 运费类/保护费这类非实物，不要求价格
  const missReq = isNonGoodsPrefix ? [] : REQ.filter(f => blank(r[f.key])).map(f => f.label);
  const missOpt = OPT.filter(f => blank(r[f.key])).map(f => f.label);
  const get = (k) => (k ? num(r[k]) : null);
  const conf = {
    wt: weightKey ? num(r[weightKey]) : null,
    buy: get(K.buy),
    buyR: get(K.buyR),
    dom: get(K.dom),
    domR: get(K.domR),
    svc: get(K.svc),
    svcUS: get(K.svcUS),
    ovC: K.ovC ? String(r[K.ovC] ?? '').trim() || null : null,
    ovP: get(K.ovP),
    ovF: get(K.ovF),
    fixed: get(K.fixed),
    pack: get(K.pack) || 1,
    ok: missReq.length === 0,
  };
  for (const a of cell.split(/[,，、]/).map(s => s.trim()).filter(Boolean)) {
    cfgMap.set(a, conf);
    if (missReq.length) cfgIncomplete.push({ 前缀: a, 缺: missReq, 缺选填: missOpt });
  }
}
const prefixList = [...cfgMap.keys()].sort((a, b) => b.length - a.length);

// ---- 2. 运价表（来自「运价表索引」节点，已做过重叠/缺口/空价校验）----
const rateIdx = $('运价表索引').first().json;
const rates = new Map();                 // 国家 → [{lo,hi,p,fee,线路}]
for (const s of (rateIdx.index || [])) {
  if (!rates.has(s.国家)) rates.set(s.国家, []);
  rates.get(s.国家).push(s);
}
for (const arr of rates.values()) arr.sort((a, b) => a.lo - b.lo);

// 国家 → 毛利（服务费兜底）
const margins = new Map();
for (const s of (rateIdx.index || [])) {
  if (s.毛利 !== null && s.毛利 !== undefined && !margins.has(s.国家)) margins.set(s.国家, s.毛利);
}

const rateProblems = rateIdx.problems || {};
const rateBadRows = (rateProblems.badPrice || []).map(m => ({ 国家: m.国家, 段: m.段, 缺: m.缺 }));
const noMargin = (rateIdx.index || []).filter(x => x.毛利 === null || x.毛利 === undefined).length;

// ---- 2.5 国家规则（欧盟税 / 附加费 / 服务费覆盖）----
const countryItems = $('解析国家配置表').all().map(i => i.json);
const rules = new Map();
for (const r of countryItems) {
  const c = String(r['国家'] ?? '').trim();
  if (!c) continue;
  const st = String(r['状态'] ?? '').trim();
  if (st && st !== '生效') continue;                 // 有「状态」就按生效过滤，没这列就不管
  rules.set(c, {
    isEU: String(r['是否欧盟'] ?? '').trim() === '是',
    surcharge: num(r['订单附加费']) ?? 0,
    service: num(r['服务费覆盖']),
    taxUsd: num(r['欧盟税额$']),
  });
}

// ---- 2.6 全局参数（汇率 / 欧盟税）----
const paramItems = $('解析全局参数').all().map(i => i.json);
const P = {};
for (const r of paramItems) {
  const k = String(r['参数键'] ?? '').trim();
  if (k) P[k] = num(r['参数值']);
}
const FX = P['USD_CNY'] ?? 6.65;
const EUR = P['EU_TAX_EUR'] ?? 3;
const EUR2USD = P['EUR_USD'] ?? 0.9;
const EU_TAX = Math.round((EUR / EUR2USD) * 100) / 100;   // 3 ÷ 0.9 = 3.33
const MULT = P['DEFAULT_SERVICE_FEE'] ?? 1.15;
const ROUND_N = P['BILL_ROUND'] ?? 2;

// ---- 3. 澳洲分区：本轮暂停（分区表 9091 行，性能原因先跳过）----
// ---- 4. 订单 ----
const orders = $input.all().map(i => i.json).filter(r => r && r.SKU);
if (!orders.length) throw new Error('上游没有订单明细');

const matchPrefix = (sku) => {
  const s = String(sku).trim();
  for (const p of prefixList) {
    if (s === p || s.startsWith(p + '-') || s.startsWith(p + '_')) return p;
  }
  return null;
};
const guessPrefix = (sku) => {
  const head = String(sku).split('_')[0];
  const parts = head.split('-');
  return parts.length >= 2 ? parts.slice(0, 2).join('-') : head;
};

// ---- 5. 逐行匹配 ----
const missSku = new Map();      // ① 前缀不存在
const missCfg = new Set();      // ② 前缀存在但必填为空
const missRate = new Map();     // ③ 运价缺失
const auMissing = new Map();    // ⑤ 澳洲邮编查不到分区
const orderWeight = new Map();
const blockedBySku = new Set();
const blockedByCfg = new Set();
const orderInfo = new Map();
const orderRows = new Map();    // 交易编号 → 行
let matchedRows = 0;
const allOrders = new Set();

const isNonGoodsOf = (r) =>
  String(r.SKU).toUpperCase().startsWith('SHIP') && !String(r['原始图片'] || '').trim();

for (const r of orders) {
  const txn = r['交易编号'];
  allOrders.add(txn);
  if (!orderRows.has(txn)) orderRows.set(txn, []);
  orderRows.get(txn).push(r);

  let country = String(r['国家中文'] || '').trim();
  if (country === '美国' && String(r['物流渠道'] || '').includes('岛屿')) country = '美国（偏远）';

  // 澳大利亚：分区由「合并澳洲分区」节点按邮编单条查到，拼进 r['澳洲分区']
  let auPc = null;
  if (country === '澳大利亚') {
    auPc = String(r['澳洲邮编'] || r['邮政编码'] || '').replace(/\D/g, '').padStart(4, '0').slice(-4);
    const zone = r['澳洲分区'];
    if (zone !== null && zone !== undefined && zone !== '') {
      country = `澳大利亚（${Number(zone)}区）`;
    } else {
      const m = auMissing.get(auPc || '(空)') || { 邮编: auPc || '(空)', 订单: new Set() };
      m.订单.add(txn);
      auMissing.set(auPc || '(空)', m);
    }
  }

  if (!orderInfo.has(txn)) orderInfo.set(txn, { 国家: country });
  else orderInfo.get(txn).国家 = country;

  const hit = matchPrefix(r.SKU);
  const qty = num(r['商品数量']) ?? 1;

  if (!hit) {
    blockedBySku.add(txn);
    const g = guessPrefix(r.SKU);
    const m = missSku.get(g) || { 标识: g, 订单: new Set(), 行: 0, SKU: new Set(), 样例SKU: r.SKU };
    m.订单.add(txn); m.行 += 1; m.SKU.add(r.SKU);
    missSku.set(g, m);
  } else {
    const cfg = cfgMap.get(hit);
    if (!cfg.ok) {
      blockedByCfg.add(txn);
      missCfg.add(hit);
    } else {
      matchedRows++;
      const rowWt = isNonGoodsOf(r) ? 0 : (cfg.wt ?? 0.17) * qty;
      orderWeight.set(txn, (orderWeight.get(txn) || 0) + rowWt);
    }
  }


}

// ---- 6. 运价段匹配（整单重量 → 段价）----
let checkedRate = 0;
const segByOrder = new Map();
for (const [txn, info] of orderInfo) {
  if (blockedBySku.has(txn) || blockedByCfg.has(txn)) continue;
  // 澳洲邮编查不到分区的单，国家名还是「澳大利亚」，运价表里没有这个国家 → 会被判缺运价，
  // 但真正原因是分区缺失，所以先跳过，避免重复报
  if (auMissing.size && [...auMissing.values()].some(m => m.订单.has(txn))) continue;
  checkedRate++;
  const w = Math.round((orderWeight.get(txn) || 0) * 1000) / 1000;
  const segs = rates.get(info.国家) || [];
  const seg = segs.find(x => x.lo < w && w <= x.hi) || null;
  segByOrder.set(txn, { w, seg, segs });

  if (!seg) {
    blocked.add(txn);
    const k = `${info.国家}|${w}kg`;
    const m = missRate.get(k) || {
      国家: info.国家, 重量: w, 订单: new Set(),
      现有: segs.length ? segs.map(x => `${x.lo}-${x.hi}`).join('、') : '运价表没有该国家',
    };
    m.订单.add(txn);
    missRate.set(k, m);
  }
}

// 阻断单集合（SKU 缺 / 配置不全 / 澳洲缺分区 / 缺运价）
const blocked = new Set([...blockedBySku, ...blockedByCfg]);
for (const m of auMissing.values()) for (const t of m.订单) blocked.add(t);

// ---- 6.5 逐行计价（公斤价 / 挂号费 / 运费 / 成本 / 履行价）----
const detail = [];
const totals = { 运费: 0, 成本: 0, 国内运费: 0, 履行价: 0, 挂号费: 0 };

for (const [txn, rows] of orderRows) {
  const info = orderInfo.get(txn) || {};
  const country = info.国家 || '';
  const priced = segByOrder.get(txn) || null;

  // 运费行 = 同单第一个非固定金额商品
  let freightRow = null;
  const skuQty = new Map();
  for (const r of rows) {
    skuQty.set(r.SKU, (skuQty.get(r.SKU) || 0) + (num(r['商品数量']) ?? 1));
  }
  for (const r of rows) {
    const cfg = cfgMap.get(matchPrefix(r.SKU) || '');
    if (cfg && cfg.fixed !== null) continue;
    freightRow = r;
    break;
  }
  const seenSku = new Set();

  for (const r of rows) {
    const hit = matchPrefix(r.SKU);
    const cfg = hit ? cfgMap.get(hit) : null;
    const qty = num(r['商品数量']) ?? 1;
    const nonGoods = isNonGoodsOf(r);
    const isFixed = cfg ? cfg.fixed !== null : false;

    const wtUnit = cfg && cfg.wt !== null ? cfg.wt : null;
    const rowWt = nonGoods ? 0 : Math.round((wtUnit ?? 0.17) * qty * 1000) / 1000;

    // 运价覆盖优先
    const useOv = cfg && cfg.ovC && cfg.ovC === country && cfg.ovP !== null;
    const p = isFixed ? null
      : (useOv ? cfg.ovP : (priced && priced.seg ? priced.seg.p : null));
    const fee = isFixed ? null
      : (useOv
          ? (cfg.ovF !== null ? cfg.ovF : (priced && priced.seg ? priced.seg.fee : null))
          : (priced && priced.seg ? priced.seg.fee : null));

    const isFreightRow = freightRow === r;
    const 行运费 = (p === null ? 0 : Math.round(p * rowWt * 100) / 100);
    const 挂号费 = (isFreightRow && fee !== null) ? fee : 0;
    const 运费 = Math.round((行运费 + 挂号费) * 100) / 100;

    const 采购价 = nonGoods ? 0 : (cfg ? ((info.国家 === '美国（偏远）' && cfg.buyR !== null) ? cfg.buyR : cfg.buy) ?? 0 : 0);
    const 成本 = Math.round(qty * 采购价 * 100) / 100;

    const 国内运费单价 = nonGoods ? 0 : (cfg ? ((info.国家 === '美国（偏远）' && cfg.domR !== null) ? cfg.domR : cfg.dom) ?? 0 : 0);
    const packN = Math.max(1, (cfg && cfg.pack) || 1);
    const firstOfSku = !seenSku.has(r.SKU);
    seenSku.add(r.SKU);
    const 份数 = firstOfSku ? Math.ceil((skuQty.get(r.SKU) || 0) / packN) : 0;
    const 国内运费合计 = Math.round(国内运费单价 * 份数 * 100) / 100;

    const 固定金额合计 = isFixed ? Math.round(cfg.fixed * 份数 * 100) / 100 : 0;

    // 服务费：美国列 > 普通列 > 运费表毛利 > 1.15
    const isUS = country.startsWith('美国');
    let 服务费 = null;
    if (cfg) {
      if (isUS && cfg.svcUS !== null) 服务费 = cfg.svcUS;
      else if (cfg.svc !== null) 服务费 = cfg.svc;
    }
    const rule = rules.get(country) || {};
    if (服务费 === null && rule.service !== null && rule.service !== undefined) 服务费 = rule.service;
    if (服务费 === null) 服务费 = margins.get(country) ?? MULT;

    const 履行价 = Math.round((运费 + 成本 + 国内运费合计) * 100) / 100;

    // 附加费只落订单首行
    const isFirstRowOfOrder = rows[0] === r;
    const 附加费 = isFirstRowOfOrder ? (rule.surcharge || 0) : 0;

    // 人民币 = (运费 + 件数×采购价 + 国内运费 + 附加费) × 服务费
    let 人民币 = Math.round((运费 + 成本 + 国内运费合计 + 附加费) * 服务费 * 100) / 100;
    // 固定金额行：总价锁死 = 固定金额 × 份数（美元）→ 折人民币
    if (isFixed) 人民币 = Math.round(固定金额合计 * FX * 100) / 100;
    const 美元 = Math.round((人民币 / FX) * 10000) / 10000;

    // 欧盟税（整单一次，只落首行）
    const 欧盟税 = isFirstRowOfOrder && rule.isEU ? EU_TAX : 0;

    if (priced && priced.seg) {
      totals.运费 += 运费;
      totals.挂号费 += 挂号费;
      totals.成本 += 成本;
      totals.国内运费 += 国内运费合计;
      totals.履行价 += 履行价;
    }

    detail.push({
      交易编号: txn,
      订单编号: r['订单编号'] || '',
      国家: country,
      物流公司: r['物流公司'] || '',
      物流渠道: r['物流渠道'] || '',
      内部单号: r['内部单号'] || '',
      图片: r['原始图片'] || '',
      邮政编码: r['邮政编码'] || '',
      SKU: r.SKU,
      前缀: hit || '',
      件数: qty,
      单件重量: wtUnit,
      行重量: rowWt,
      整单重量: priced ? priced.w : null,
      线路: priced && priced.seg ? priced.seg.线路 : '',
      公斤价: p,
      挂号费: fee,
      运费行: isFreightRow,
      行运费,
      挂号费合计: 挂号费,
      运费,
      采购价,
      成本,
      国内运费单价,
      国内运费份数: 份数,
      国内运费合计,
      履行价,
      服务费,
      附加费,
      人民币,
      美元,
      欧盟税,
      固定金额合计,
      缺运价: Boolean(priced && !priced.seg),
      非实物: nonGoods,
      缺配置: !hit || (cfg && !cfg.ok) || false,
    });
  }
}

// 订单级：账单 = ROUND(Σ美元 + 欧盟税, 2)；毛利 = 账单×汇率 − 履行价
const perOrderUsd = new Map();
const perOrderCost = new Map();
const perOrderTax = new Map();
for (const x of detail) {
  perOrderUsd.set(x.交易编号, Math.round(((perOrderUsd.get(x.交易编号) || 0) + x.美元) * 1e6) / 1e6);
  perOrderCost.set(x.交易编号, Math.round(((perOrderCost.get(x.交易编号) || 0) + x.履行价) * 100) / 100);
  if (x.欧盟税) perOrderTax.set(x.交易编号, x.欧盟税);
}
const billOf = new Map();
for (const [txn, usd] of perOrderUsd) {
  const tax = perOrderTax.get(txn) || 0;
  billOf.set(txn, Math.round((usd + tax) * Math.pow(10, ROUND_N)) / Math.pow(10, ROUND_N));
}
for (const x of detail) {
  x.订单美元合计 = perOrderUsd.get(x.交易编号);
  x.账单金额 = billOf.get(x.交易编号);
  x.订单履行价 = perOrderCost.get(x.交易编号);
  x.毛利 = Math.round((x.账单金额 * FX - x.订单履行价) * 100) / 100;
}

for (const k of Object.keys(totals)) totals[k] = Math.round(totals[k] * 100) / 100;
totals.人民币 = Math.round(detail.reduce((s2, x) => s2 + x.人民币, 0) * 100) / 100;
totals.账单 = Math.round([...billOf.entries()]
  .filter(([txn]) => !blocked.has(txn))
  .reduce((s2, [, v]) => s2 + v, 0) * 100) / 100;
// 毛利按订单逐单算再累加（避免整体相减的尾差）
let grossTotal = 0;
for (const [txn, bill] of billOf) {
  if (blocked.has(txn)) continue;                 // 阻断单不计毛利
  const cost = perOrderCost.get(txn) || 0;
  grossTotal += Math.round((bill * FX - cost) * 100) / 100;
}
totals.毛利 = Math.round(grossTotal * 100) / 100;

// ---- 7. 组一条消息 ----
const skuList = [...missSku.values()].sort((a, b) => b.订单.size - a.订单.size);
const cfgList = cfgIncomplete.filter(x => missCfg.has(x.前缀)).sort((a, b) => a.前缀.localeCompare(b.前缀));
const rateList = [...missRate.values()].sort((a, b) => b.订单.size - a.订单.size);


for (const m of rateList) for (const t of m.订单) blocked.add(t);


const TOP = 20;
const L = [];
L.push('【60103 匹配异常汇总】');
L.push(`订单 ${allOrders.size} 单 / ${orders.length} 行；可继续 ${allOrders.size - blocked.size} 单，需补 ${blocked.size} 单。`);

L.push('');
L.push(`━━ ① 商品配置：没有该 SKU 前缀（${skuList.length} 个）━━`);
if (!skuList.length) L.push('无');
else {
  skuList.slice(0, TOP).forEach((m, i) => {
    L.push(`${i + 1}. ${m.标识}（${m.订单.size} 单 / ${m.行} 行 / ${m.SKU.size} 个SKU）`);
    L.push(`   例：${m.样例SKU}`);
  });
  if (skuList.length > TOP) L.push(`… 还有 ${skuList.length - TOP} 个未列出`);
}

L.push('');
L.push(`━━ ② 商品配置：有前缀但价格没填（${cfgList.length} 个）━━`);
if (!cfgList.length) L.push('无');
else {
  const groups = new Map();
  for (const m of cfgList) {
    const k = m.缺.join('、');
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(m.前缀);
  }
  let n = 0;
  for (const [lack, names] of groups) {
    n += 1;
    L.push(`${n}. 缺「${lack}」共 ${names.length} 个：`);
    L.push('   ' + names.slice(0, 40).join('、') + (names.length > 40 ? ` …等 ${names.length} 个` : ''));
  }
  const optMissing = [...new Set(cfgList.flatMap(m => m.缺选填))];
  if (optMissing.length) L.push(`（另有选填项未填：${optMissing.join('、')}，会用默认值）`);
}

L.push('');
L.push(`━━ ③ 运价缺失：该重量没有段（${rateList.length} 项）━━`);
const skippedRate = new Set([...blockedBySku, ...blockedByCfg]).size;
if (!rateList.length) {
  if (skippedRate && checkedRate) {
    L.push(`已校验 ${checkedRate} 单，均能查到运价；${skippedRate} 单因 ①/② 未通过，跳过运价校验`);
  } else if (skippedRate) {
    L.push(`未校验：${skippedRate} 单因 ①/② 未通过，本轮跳过运价匹配`);
  } else L.push('无');
}
else {
  rateList.slice(0, TOP).forEach((m, i) => {
    L.push(`${i + 1}. ${m.国家}｜${m.重量}kg（${m.订单.size} 单）`);
    L.push(`   现有段：${m.现有}`);
  });
  if (rateList.length > TOP) L.push(`… 还有 ${rateList.length - TOP} 项未列出`);
}

L.push('');
L.push(`━━ ④ 运价表：区间/价格不完整的行（${rateBadRows.length} 行）━━`);
if (!rateBadRows.length) L.push('无');
else {
  rateBadRows.slice(0, TOP).forEach((m, i) => {
    L.push(`${i + 1}. ${m.国家}｜${m.段} —— 缺 ${m.缺.join('、')}`);
  });
  if (rateBadRows.length > TOP) L.push(`… 还有 ${rateBadRows.length - TOP} 行未列出`);
}

L.push('');
const auList = [...auMissing.values()].sort((a, b) => b.订单.size - a.订单.size);
L.push(`━━ ⑤ 澳洲分区缺失（${auList.length} 个邮编）━━`);
if (!auList.length) L.push('无');
else {
  auList.slice(0, TOP).forEach((m, i) => {
    L.push(`${i + 1}. 邮编 ${m.邮编}（${m.订单.size} 单）—— 分区表里查不到，需补`);
  });
  if (auList.length > TOP) L.push(`… 还有 ${auList.length - TOP} 个未列出`);
  for (const m of auList) for (const t of m.订单) blocked.add(t);
}

// ⑥ 运价表结构问题（来自「运价表索引」的校验结果）
const rp = rateProblems;
const gCount = (rp.gaps || []).length;
const oCount = (rp.overlap || []).length;
const bCount = (rp.badPrice || []).length;
const zCount = (rp.notFromZero || []).length;
L.push('');
L.push(`━━ ⑥ 运价表结构（重叠 ${oCount} / 缺口 ${gCount} / 空价 ${bCount} / 首段不从0 ${zCount}）━━`);
if (!oCount && !gCount && !bCount) L.push('无影响匹配的问题');
(rp.overlap || []).slice(0, 5).forEach(m => L.push(`重叠：${m.国家} ${m.前段} 与 ${m.后段}（同一个重量会命中两条，只取第一条）`));
(rp.gaps || []).slice(0, 5).forEach(m => L.push(`缺口：${m.国家} ${m.缺口}kg 没有运价（${m.大小}kg）`));
(rp.badPrice || []).slice(0, 5).forEach(m => L.push(`空价：${m.国家} ${m.段} 缺 ${m.缺.join('、')}`));
if (zCount) {
  L.push(`首段不从 0 开始的国家 ${zCount} 个（低于首段下界会被判缺运价）：`
    + (rp.notFromZero || []).slice(0, 8).map(m => `${m.国家}(${m.首段})`).join('、')
    + (zCount > 8 ? ' …' : ''));
}

if (noMargin) {
  L.push('');
  L.push(`提示：运价表有 ${noMargin} 行没填「毛利」，这些国家按默认服务费计算。`);
}

const hasBlock = skuList.length || cfgList.length || rateList.length || rateBadRows.length || auList.length;
if (hasBlock) {
  L.push('');
  L.push('补齐后重跑本流程即可。');
}

return [{
  json: {
    totalOrders: allOrders.size,
    totalRows: orders.length,
    matchedRows,
    blockedOrders: blocked.size,
    missingSku: skuList.length,
    missingCfg: cfgList.length,
    missingRate: rateList.length,
    badRateRows: rateBadRows.length,
    rateProblems: rateProblems.summary || {},
    missingAu: auList.length,
    missingCount: skuList.length + cfgList.length + rateList.length + rateBadRows.length + auList.length,
    message: L.join('\n'),
    missingList: [
      ...skuList.map(m => ({ 级别: '阻断', 类型: 'SKU前缀缺失', 标识: m.标识,
        说明: `商品配置里没有前缀「${m.标识}」`, 订单数: m.订单.size, 行数: m.行, 样例: m.样例SKU })),
      ...cfgList.map(m => ({ 级别: '阻断', 类型: '商品配置不全', 标识: m.前缀,
        说明: `缺 ${m.缺.join('、')}`, 订单数: 0, 行数: 0, 样例: '' })),
      ...rateList.map(m => ({ 级别: '阻断', 类型: '运价缺失', 标识: `${m.国家}|${m.重量}kg`,
        说明: `现有段：${m.现有}`, 订单数: m.订单.size, 行数: 0, 样例: '' })),
      ...rateBadRows.map(m => ({ 级别: '阻断', 类型: '运价表空价', 标识: `${m.国家}|${m.段}`,
        说明: `缺 ${m.缺.join('、')}`, 订单数: 0, 行数: 0, 样例: '' })),
      ...auList.map(m => ({ 级别: '阻断', 类型: '澳洲分区缺失', 标识: `邮编${m.邮编}`,
        说明: '分区表查不到该邮编', 订单数: m.订单.size, 行数: 0, 样例: '' })),
    ],
    totals,
    params: { FX, EUR, EUR2USD, EU_TAX, MULT, ROUND_N },
    detailCount: detail.length,
    detail: detail,
  },
}];
