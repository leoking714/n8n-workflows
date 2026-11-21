// ============================================================================
// 提取澳大利亚邮编（去重，只查这些邮编，避免全表 9091 行）
//   输入：「解析出需要的数据」的订单明细
//   输出：每个唯一邮编一个 item，下游按邮编单条查询分区
// ============================================================================
const rows = $input.all().map(i => i.json).filter(r => r && r.SKU);

const set = new Set();
for (const r of rows) {
  if (String(r['国家中文'] || '').trim() !== '澳大利亚') continue;
  const pc = String(r['邮政编码'] || '').replace(/\D/g, '').padStart(4, '0').slice(-4);
  if (pc.length === 4) set.add(pc);
}

// 没有澳洲单也要输出 1 条，否则下游节点不会执行、整条链断掉
if (!set.size) set.add('__NONE__');

return [...set].map(pc => ({ json: { 邮编: pc } }));
