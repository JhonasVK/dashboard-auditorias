const XLSX = require('./node_modules/xlsx');
const path = require('path');

const archivo = path.join(__dirname, 'data', 'BBDD_Supervisores.xlsx');
try {
  const wb = XLSX.readFile(archivo);
  const hoja = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });
  
  if (filas.length > 0) {
    const encabezados = filas[0];
    encabezados.forEach((nombre, index) => {
      // Calcular la letra de la columna
      let colName = '';
      let i = index;
      while (i >= 0) {
        colName = String.fromCharCode((i % 26) + 65) + colName;
        i = Math.floor(i / 26) - 1;
      }
      if (['AC', 'AD', 'AE'].includes(colName)) {
        console.log(`Columna ${colName}: "${nombre}"`);
      }
    });
  }
} catch(err) {
  console.error('Error:', err.message);
}
