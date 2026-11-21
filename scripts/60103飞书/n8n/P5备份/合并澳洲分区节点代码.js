// ============================================================================
// 合并澳洲分区：把「查询澳大利亚分区表」的结果拼回订单
//   输出：订单明细（原样）+ 澳洲邮编 / 澳洲分区
// ============================================================================
const map = new Map();

for (const it of $('查询澳大利亚分区表').all()) {
  for (const rec of (it.json.items || [])) {
    const f = rec.fields || {};
    let pc = f['邮编'];
    if (Array.isArray(pc)) pc = (pc[0] && (pc[0].text ?? pc[0])) ?? '';
    pc = String(pc ?? '').trim().padStart(4, '0');

    let zone = f['分区'];
    if (zone && typeof zone === 'object') zone = zone.value ?? zone;

    if (pc.length === 4 && zone !== null && zone !== undefined && zone !== '' && !map.has(pc)) {
      map.set(pc, Number(zone));
    }
  }
}

const rows = $('解析出需要的数据').all().map(i => i.json).filter(r => r && r.SKU);

return rows.map(r => {
  const o = { ...r };
  if (String(o['国家中文'] || '').trim() === '澳大利亚') {
    const pc = String(o['邮政编码'] || '').replace(/\D/g, '').padStart(4, '0').slice(-4);
    o['澳洲邮编'] = pc;
    o['澳洲分区'] = map.has(pc) ? map.get(pc) : null;
  }
  return { json: o };
});
