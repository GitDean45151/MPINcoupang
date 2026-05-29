import XLSX from 'xlsx';
import path from 'path';

try {
  const wb = XLSX.readFile(path.resolve('260515_MP.xlsx'));
  const sheet = wb.Sheets['MP'];
  
  console.log('--- Printing cell formulas and values in summary area (Rows 1 to 8) ---');
  for (let r = 1; r <= 8; r++) {
    console.log(`\nRow ${r}:`);
    for (let c = 2; c <= 20; c++) { // Columns B (2) to T (20)
      const cellRef = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      const cell = sheet[cellRef];
      if (cell) {
        let display = `  Cell ${cellRef}: val=${cell.v !== undefined ? cell.v : ''}`;
        if (cell.f) {
          display += ` | formula===${cell.f}`;
        }
        console.log(display);
      }
    }
  }
} catch (err) {
  console.error(err);
}
