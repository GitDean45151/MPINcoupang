import XLSX from 'xlsx';
import path from 'path';

try {
  const wb = XLSX.readFile(path.resolve('260515_MP.xlsx'));
  const sheet = wb.Sheets['MP'];
  
  console.log('--- Checking formulas for vendor rows 10 to 15 ---');
  for (let r = 10; r <= 15; r++) {
    console.log(`\nRow ${r} (${sheet[XLSX.utils.encode_col(2 - 1) + r]?.v || 'no-id'}):`);
    // Cols E (5), G (7), H (8), I (9)
    [5, 7, 8, 9].forEach(c => {
      const cellRef = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      const cell = sheet[cellRef];
      if (cell) {
        let display = `  Col ${XLSX.utils.encode_col(c)}: val=${cell.v !== undefined ? cell.v : ''}`;
        if (cell.f) {
          display += ` | formula===${cell.f}`;
        }
        console.log(display);
      }
    });
  }
} catch (err) {
  console.error(err);
}
