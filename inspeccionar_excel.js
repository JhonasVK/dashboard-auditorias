const XLSX = require('./node_modules/xlsx');
const path = require('path');

const archivo = path.join(__dirname, 'data', 'BBDD_Supervisores.xlsx');

try {
  const wb    = XLSX.readFile(archivo);
  const hoja  = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(hoja, { defval: '' });

  console.log('\n📋 NOMBRE DE HOJAS:', wb.SheetNames);
  console.log('\n📊 TOTAL DE FILAS:', filas.length);
  console.log('\n🏷️  COLUMNAS ENCONTRADAS:');
  if (filas.length > 0) {
    Object.keys(filas[0]).forEach(col => console.log(`   - "${col}"`));
  }
  console.log('\n📄 PRIMERAS 3 FILAS:');
  filas.slice(0, 3).forEach((f, i) => {
    console.log(`\n  Fila ${i+1}:`);
    Object.entries(f).forEach(([k,v]) => console.log(`    ${k}: ${v}`));
  });
} catch(err) {
  console.error('❌ Error:', err.message);
}
