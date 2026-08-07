// Generador de BBDD_Supervisores.xlsx con datos realistas
const XLSX = require('./node_modules/xlsx');
const path = require('path');

const supervisores = [
  'Carlos Muñoz', 'Ana Pérez', 'Luis Torres', 'María Rivas',
  'Pedro Soto', 'Jorge Fuentes', 'Carmen López', 'Ricardo Vera',
  'Valentina Castro', 'Felipe Morales'
];

const zonas = ['Zona Norte', 'Zona Sur', 'Zona Centro', 'Zona Oriente', 'Zona Poniente'];

const lugares = {
  'Zona Norte':   ['Sitio Colina', 'Planta Lampa', 'Torre Huechuraba', 'Bodega Quilicura'],
  'Zona Sur':     ['Planta Maipú', 'Centro San Bernardo', 'Bodega Lo Espejo', 'Sitio El Bosque'],
  'Zona Centro':  ['Oficina Central', 'Bodega Central', 'Torre Santiago Centro', 'Planta Estación'],
  'Zona Oriente': ['Centro Peñalolén', 'Sitio La Florida', 'Planta Puente Alto', 'Bodega Macul'],
  'Zona Poniente':['Terminal Pudahuel', 'Planta Cerrillos', 'Bodega Cerro Navia', 'Centro Maipú'],
};

const tipos = ['Seguridad', 'Ambiental', 'Calidad', 'Operacional', 'Infraestructura'];

const estados = {
  alto:  'Aprobado',
  medio: 'Parcial',
  bajo:  'Rechazado',
};

function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function randomItem(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

function formatFecha(date) {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

// Generar 80 registros entre enero 2024 y agosto 2026
const filas = [];
const inicio = new Date('2024-01-01');
const fin    = new Date('2026-08-07');
const rango  = fin - inicio;

for (let i = 0; i < 80; i++) {
  const fecha = new Date(inicio.getTime() + Math.random() * rango);
  const zona  = randomItem(zonas);
  const lugar = randomItem(lugares[zona]);
  const cumpl = randomInt(30, 100);
  const estado = cumpl >= 85 ? estados.alto : cumpl >= 60 ? estados.medio : estados.bajo;
  const hallazgos = cumpl >= 85 ? randomInt(0, 3) : cumpl >= 60 ? randomInt(3, 9) : randomInt(8, 18);

  const observaciones = {
    'Aprobado':  ['Sin observaciones relevantes', 'Cumplimiento satisfactorio', 'Todas las medidas en orden'],
    'Parcial':   ['Requiere mejoras en EPP', 'Procedimientos incompletos', 'Señalización deficiente', 'Documentación pendiente'],
    'Rechazado': ['Riesgo crítico detectado', 'Incumplimiento grave protocolo', 'Acceso no autorizado a zona', 'Equipos sin mantención'],
  };

  filas.push({
    'Fecha':          formatFecha(fecha),
    'Zona':           zona,
    'Lugar':          lugar,
    'Supervisor':     randomItem(supervisores),
    'Tipo Auditoría': randomItem(tipos),
    'Cumplimiento %': cumpl,
    'Estado':         estado,
    'N° Hallazgos':   hallazgos,
    'Observaciones':  randomItem(observaciones[estado]),
  });
}

// Ordenar por fecha descendente
filas.sort((a, b) => {
  const pa = a.Fecha.split('/').reverse().join('');
  const pb = b.Fecha.split('/').reverse().join('');
  return pb.localeCompare(pa);
});

// Crear workbook
const wb  = XLSX.utils.book_new();
const ws  = XLSX.utils.json_to_sheet(filas);

// Ancho de columnas
ws['!cols'] = [
  { wch: 12 }, // Fecha
  { wch: 16 }, // Zona
  { wch: 22 }, // Lugar
  { wch: 20 }, // Supervisor
  { wch: 18 }, // Tipo Auditoría
  { wch: 15 }, // Cumplimiento %
  { wch: 12 }, // Estado
  { wch: 13 }, // N° Hallazgos
  { wch: 40 }, // Observaciones
];

XLSX.utils.book_append_sheet(wb, ws, 'Auditorías');
const outputPath = path.join(__dirname, 'data', 'BBDD_Supervisores_demo.xlsx');
XLSX.writeFile(wb, outputPath);

console.log(`✅ Archivo creado: ${outputPath}`);
console.log(`   Registros generados: ${filas.length}`);
