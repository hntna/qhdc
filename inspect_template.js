const fs = require('fs');
const XLSX = require('./xlsx.full.min.js');
const wb = XLSX.read(fs.readFileSync('Masterlist 2027-2028_Mau_moi.xlsx'), {type:'buffer'});
const ws = wb.Sheets['PL1.1 ML2027-2028'];

for (let r=22; r<=95; r++) {
  const a = ws['A'+r]?.v;
  const b = ws['B'+r]?.v;
  const k = ws['K'+r]?.v;
  console.log(`R${r.toString().padStart(2,'0')} | A=${(a||'').toString().padEnd(4,' ')} | K=${(k||'').toString().padEnd(6,' ')} | B=${(b||'').substring(0,50)}`);
}
