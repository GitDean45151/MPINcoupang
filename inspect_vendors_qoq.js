import XLSX from 'xlsx';
import path from 'path';

try {
  const wb = XLSX.readFile(path.resolve('260515_MP.xlsx'));
  const sheet = wb.Sheets['MP'];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  const header = rows[8];
  const vendorRows = [];
  for (let i = 9; i < rows.length; i++) {
    const row = rows[i];
    const obj = {};
    header.forEach((h, idx) => {
      if (h) obj[h] = row[idx];
    });
    if (obj.vendorid) {
      vendorRows.push(obj);
    }
  }

  // Print summary of QoQ columns for group 2 and 3
  const groupStats = {};
  ['group 1', 'group 2', 'group 3', 'group 4'].forEach(g => {
    groupStats[g] = {
      count: 0,
      q2Sales: 0,
      q1Sales: 0,
      qoqSales: 0,
      comm90: 0
    };
  });

  vendorRows.forEach(r => {
    const g = r['커미션 그룹'];
    if (groupStats[g]) {
      groupStats[g].count++;
      groupStats[g].q2Sales += Number(r['2026_Q2_광고비(~260513)']) || 0;
      groupStats[g].q1Sales += Number(r['직전분기 광고비(D열의 동기간)']) || 0;
      groupStats[g].qoqSales += Number(r['2026_Q2_QoQ 대상 권한설정 광고비']) || 0;
      groupStats[g].comm90 += Number(r['2026_Q2_90일 커미션대상 광고비']) || 0;
    }
  });

  console.log('--- Aggregated Stats from Vendor Rows ---');
  Object.keys(groupStats).forEach(g => {
    console.log(`\nGroup: ${g}`);
    console.log(`  Count: ${groupStats[g].count}`);
    console.log(`  Sum 2026_Q2_광고비: ${groupStats[g].q2Sales}`);
    console.log(`  Sum 직전분기 광고비(D열의 동기간): ${groupStats[g].q1Sales}`);
    console.log(`  Sum 2026_Q2_QoQ 대상 권한설정 광고비: ${groupStats[g].qoqSales}`);
    console.log(`  Sum 2026_Q2_90일 커미션대상 광고비: ${groupStats[g].comm90}`);
  });
  
} catch (err) {
  console.error(err);
}
