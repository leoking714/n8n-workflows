// ============================================================================
// 物流运费表 → 索引 + 结构校验
//   输入：「解析物流运费表」的每一行（国家/线路/下界/上界/公斤价/挂号费/毛利）
//   输出：1 个 item { index, problems, countries, segments }
// ============================================================================
const num = (v) => {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
};

const rows = $input.all().map(i => i.json).filter(r => r && r['国家']);
if (!rows.length) throw new Error('物流运费表没有读到数据');

const EPS = 1e-9;
const GAP_TOL = 0.0011;        // 0.001 的缝是「下界开上界闭」写法造成的，忽略

const byCountry = new Map();   // 国家 → [段]
const badPrice = [];           // 价格/区间为空

for (const r of rows) {
  const c = String(r['国家']).trim();
  const line = String(r['线路'] ?? '').trim() || '';
  const lo = num(r['下界']), hi = num(r['上界']);
  const p = num(r['公斤价']), fee = num(r['挂号费']);
  const lack = [];
  if (lo === null) lack.push('下界');
  if (hi === null) lack.push('上界');
  if (p === null) lack.push('公斤价');
  if (fee === null) lack.push('挂号费');
  if (lack.length) {
    badPrice.push({ 国家: c, 线路: line, 段: `${r['下界'] ?? '?'}~${r['上界'] ?? '?'}`, 缺: lack });
    continue;
  }
  if (hi <= lo) {
    badPrice.push({ 国家: c, 线路: line, 段: `${lo}~${hi}`, 缺: ['上界≤下界'] });
    continue;
  }
  if (!byCountry.has(c)) byCountry.set(c, []);
  byCountry.get(c).push({ 国家: c, 线路: line, lo, hi, p, fee, 毛利: num(r['毛利']) });
}

const index = [];
const overlap = [];
const gaps = [];
const notFromZero = [];

for (const [c, segs] of byCountry) {
  segs.sort((a, b) => a.lo - b.lo);
  index.push(...segs);

  if (segs[0].lo > GAP_TOL) {
    notFromZero.push({ 国家: c, 首段: `${segs[0].lo}~${segs[0].hi}` });
  }
  for (let i = 1; i < segs.length; i++) {
    const prev = segs[i - 1], cur = segs[i];
    const diff = cur.lo - prev.hi;
    if (diff < -EPS) {
      overlap.push({ 国家: c, 前段: `${prev.lo}~${prev.hi}`, 后段: `${cur.lo}~${cur.hi}` });
    } else if (diff > GAP_TOL) {
      gaps.push({ 国家: c, 缺口: `${+prev.hi.toFixed(3)}~${+cur.lo.toFixed(3)}`, 大小: +diff.toFixed(3) });
    }
  }
}

const problems = { overlap, gaps, badPrice, notFromZero };

const countries = [...byCountry.keys()].sort();
const lines = [...new Set(rows.map(r => String(r['线路'] ?? '').trim()).filter(Boolean))];

return [{
  json: {
    ok: overlap.length === 0 && badPrice.length === 0,
    countries: countries.length,
    segments: index.length,
    lines,
    index,
    problems,
    summary: {
      重叠: overlap.length,
      缺口: gaps.length,
      价格为空: badPrice.length,
      首段不从0: notFromZero.length,
    },
  },
}];
