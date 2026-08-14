/* ============================================================
   DASHBOARD AUDITORÍAS DE TERRENO — Lógica principal
   ============================================================ */

// ─── DATOS DE EJEMPLO ────────────────────────────────────────
// Estos datos se usan solo si no existe data/auditorias.xlsx
const DATOS_EJEMPLO = [
  { fecha: "2024-01-08", zona: "Zona Norte",    lugar: "Sitio Colina",       auditor: "Carlos Muñoz",   cumplimiento: 92, estado: "Aprobado",  hallazgos: 2,  tipo: "Seguridad" },
  { fecha: "2024-01-15", zona: "Zona Sur",      lugar: "Planta Maipú",       auditor: "Ana Pérez",      cumplimiento: 68, estado: "Parcial",   hallazgos: 7,  tipo: "Ambiental" },
  { fecha: "2024-01-22", zona: "Zona Centro",   lugar: "Bodega Central",     auditor: "Luis Torres",    cumplimiento: 45, estado: "Rechazado", hallazgos: 12, tipo: "Calidad"   },
  { fecha: "2024-02-05", zona: "Zona Norte",    lugar: "Torre Santiago",     auditor: "María Rivas",    cumplimiento: 87, estado: "Aprobado",  hallazgos: 3,  tipo: "Seguridad" },
  { fecha: "2024-02-12", zona: "Zona Oriente",  lugar: "Centro Peñalolén",   auditor: "Pedro Soto",     cumplimiento: 78, estado: "Aprobado",  hallazgos: 5,  tipo: "Calidad"   },
  { fecha: "2024-02-20", zona: "Zona Sur",      lugar: "Planta San Bernardo",auditor: "Carlos Muñoz",   cumplimiento: 55, estado: "Parcial",   hallazgos: 9,  tipo: "Ambiental" },
  { fecha: "2024-03-03", zona: "Zona Centro",   lugar: "Oficina Central",    auditor: "Ana Pérez",      cumplimiento: 95, estado: "Aprobado",  hallazgos: 1,  tipo: "Seguridad" },
  { fecha: "2024-03-10", zona: "Zona Norte",    lugar: "Sitio Lampa",        auditor: "Jorge Fuentes",  cumplimiento: 38, estado: "Rechazado", hallazgos: 15, tipo: "Calidad"   },
  { fecha: "2024-03-18", zona: "Zona Poniente", lugar: "Terminal Pudahuel",  auditor: "María Rivas",    cumplimiento: 81, estado: "Aprobado",  hallazgos: 4,  tipo: "Ambiental" },
  { fecha: "2024-04-02", zona: "Zona Sur",      lugar: "Planta Maipú",       auditor: "Luis Torres",    cumplimiento: 74, estado: "Aprobado",  hallazgos: 6,  tipo: "Seguridad" },
  { fecha: "2024-04-09", zona: "Zona Oriente",  lugar: "Centro La Florida",  auditor: "Pedro Soto",     cumplimiento: 62, estado: "Parcial",   hallazgos: 8,  tipo: "Calidad"   },
  { fecha: "2024-04-16", zona: "Zona Norte",    lugar: "Sitio Colina",       auditor: "Jorge Fuentes",  cumplimiento: 89, estado: "Aprobado",  hallazgos: 3,  tipo: "Ambiental" },
  { fecha: "2024-05-07", zona: "Zona Centro",   lugar: "Bodega Central",     auditor: "Carlos Muñoz",   cumplimiento: 58, estado: "Parcial",   hallazgos: 10, tipo: "Seguridad" },
  { fecha: "2024-05-14", zona: "Zona Poniente", lugar: "Terminal Pudahuel",  auditor: "Ana Pérez",      cumplimiento: 91, estado: "Aprobado",  hallazgos: 2,  tipo: "Calidad"   },
  { fecha: "2024-05-21", zona: "Zona Sur",      lugar: "Planta San Bernardo",auditor: "María Rivas",    cumplimiento: 42, estado: "Rechazado", hallazgos: 13, tipo: "Ambiental" },
  { fecha: "2024-06-04", zona: "Zona Norte",    lugar: "Torre Santiago",     auditor: "Luis Torres",    cumplimiento: 96, estado: "Aprobado",  hallazgos: 1,  tipo: "Seguridad" },
  { fecha: "2024-06-11", zona: "Zona Oriente",  lugar: "Centro Peñalolén",   auditor: "Jorge Fuentes",  cumplimiento: 73, estado: "Aprobado",  hallazgos: 5,  tipo: "Calidad"   },
  { fecha: "2024-06-18", zona: "Zona Centro",   lugar: "Oficina Central",    auditor: "Carlos Muñoz",   cumplimiento: 66, estado: "Parcial",   hallazgos: 7,  tipo: "Ambiental" },
];

// ─── ESTADO GLOBAL ────────────────────────────────────────────
let datosOriginales = [...DATOS_EJEMPLO];
let datosFiltrados  = [...DATOS_EJEMPLO];
let paginaActual    = 1;
const FILAS_POR_PAGINA = 15;
// ─── INICIALIZACIÓN ───────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  actualizarFecha();
  cargarExcelAutomatico();
});

// ─── CARGA AUTOMÁTICA desde data/ ────────────────────────────
async function cargarExcelAutomatico() {
  if (window.location.protocol === 'file:') {
    alert("⚠️ ¡Atención! Estás abriendo el archivo directamente.\nPara que el Dashboard pueda leer tu Excel (BBDD_Supervisores.xlsx), debes abrirlo a través del servidor local.\n\n👉 Entra a: http://localhost:3000 en tu navegador.");
  }

  const indicador = document.getElementById('indicador-carga');
  if (indicador) indicador.textContent = '⏳ Cargando datos...';

  // Intenta cargar el archivo real primero, luego el demo
  const archivos = [
    'data/BBDD_Supervisores.xlsx',
    'data/BBDD_Supervisores_demo.xlsx'
  ];

  for (const archivo of archivos) {
    try {
      const resp = await fetch(archivo, { cache: 'no-store' });
      if (!resp.ok) continue;

      const buffer = await resp.arrayBuffer();
      const data   = new Uint8Array(buffer);
      const wb     = XLSX.read(data, { type: 'array', cellDates: true });
      const hoja   = wb.Sheets[wb.SheetNames[0]];
      const filas  = XLSX.utils.sheet_to_json(hoja, { defval: '' });

      datosOriginales = mapearColumnas(filas);
      datosFiltrados  = [...datosOriginales];

      const nombre = archivo.split('/').pop();
      if (indicador) indicador.textContent = `✅ ${datosOriginales.length} registros — ${nombre}`;
      mostrarToast(`✅ Cargado: ${nombre} (${datosOriginales.length} registros)`);
      poblarFiltros();
      renderizar();
      return; // éxito, salir
    } catch (err) {
      console.warn(`No se pudo cargar ${archivo}:`, err.message);
      alert(`Error al cargar el Excel (${archivo}):\n${err.message}`);
    }
  }

  // Si ningún archivo cargó, usar datos de ejemplo
  console.warn('Sin archivo Excel, usando datos de ejemplo');
  datosOriginales = [...DATOS_EJEMPLO];
  datosFiltrados  = [...DATOS_EJEMPLO];
  if (indicador) indicador.textContent = '📊 Datos de ejemplo';
  poblarFiltros();
  renderizar();
}

function actualizarFecha() {
  const el = document.getElementById('fecha-actual');
  if (!el) return;
  const hoy = new Date();
  el.textContent = hoy.toLocaleDateString('es-CL', {
    day: '2-digit', month: 'long', year: 'numeric'
  });
}




// Mapeo de columnas — adaptado exactamente a BBDD_Supervisores (Google Forms)
function mapearColumnas(filas) {
  return filas.map(f => {

    // ── Fecha desde "Marca temporal" (número serial de Excel) ──
    const rawFecha = f['Marca temporal'] || '';
    let fecha = '';
    if (typeof rawFecha === 'number' && rawFecha > 0) {
      const d = new Date(Math.round((rawFecha - 25569) * 86400 * 1000));
      fecha = d.toISOString().split('T')[0];
    } else if (rawFecha instanceof Date) {
      fecha = rawFecha.toISOString().split('T')[0];
    } else if (typeof rawFecha === 'string' && rawFecha.includes('/')) {
      const [d, m, y] = rawFecha.split(' ')[0].split('/');
      const esValida = d && m && y && /^\d{1,2}$/.test(d) && /^\d{1,2}$/.test(m) && /^\d{4}$/.test(y);
      fecha = esValida ? `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}` : '';
    } else {
      fecha = '';
    }

    // ── Nota del técnico → cumplimiento % ──
    const nota = parseFloat(f['Con que nota se evalúa al técnico '] || 0);
    const cumplimiento = nota > 0 ? Math.round((nota / 10) * 100) : 0;

    // ── Puntuación (columna B) ──
    const puntuacion = parseFloat(f['Puntuación'] || 0);

    // ── Nota de Estética del cableado ──
    const estetica = parseFloat(f['Nota de Estética del cableado '] || 0);

    // ── Estado de la auditoría ──
    const estadoRaw = String(f['Estado Auditoria '] || f['Técnico cumple correctamente con su función?  '] || '').trim();
    let estado = 'Pendiente';
    if (estadoRaw.toLowerCase().includes('cumple') || estadoRaw.toLowerCase() === 'si') {
      estado = 'Aprobado';
    } else if (estadoRaw.toLowerCase().includes('no cumple') || estadoRaw.toLowerCase() === 'no') {
      estado = 'Rechazado';
    } else if (estadoRaw !== '') {
      estado = estadoRaw;
    }

    return {
      fecha,
      peticion:      String(f['N° PETICIÓN ']                        || ''),
      tecnico:       String(f['Nombre Técnico']                      || 'Sin nombre'),
      auditor:       String(f['SUPERVISOR']                          || 'Sin supervisor'),
      actividad:     String(f['ACTIVIDAD']                           || 'Sin actividad'),
      tipo:          String(f['AUDITORIA']                           || 'General'),
      direccion:     String(f['DIRECCIÓN (Calle - Numero) ']         || ''),
      cumplimiento,
      nota,
      puntuacion,
      estetica,
      estado,
      observaciones: String(f['Observaciones (detallar observaciones)'] || f['Observaciones '] || ''),
      // Campos adicionales de checklist
      pasamuros:     String(f['Instala Pasa-muros ']                 || ''),
      roseta:        String(f['Roseta Óptica en norma y atornillada ']|| ''),
      limpio:        String(f['Deja Área de Trabajo Limpio ']        || ''),
      capacitacion:  String(f['Capacitación  al Cliente']            || ''),
      funcion:       String(f['Técnico cumple correctamente con su función?  '] || ''),
      alcohol:       String(f['Utiliza Alcohol isopropilico ']       || ''),
      oneclick:      String(f['Utiliza OneClick para realizar limpieza a conectores '] || ''),
      dropSoportes:  String(f['Instala drop con soportes']           || ''),
      dropNorma:     String(f['Drop se encuentra encuentra dentro de norma de instalaciones '] || ''),
      reutiliza:     String(f['Reutiliza Instalación ']               || ''),
      reutilizaOtra: String(f['Reutiliza Instalación de otra compañía '] || ''),

      // Columnas AC, AD, AE
      mesaTrabajo:   String(f['Se llevó a cabo la mesa de trabajo.'] || ''),
      cierreProceso: String(f['Se cerró el proceso con el técnico?'] || ''),
      obsCierre:     String(f['Observaciones ']                      || ''),
    };
  }).filter(f => f.fecha && f.fecha.length > 0);
}

function clasificarEstado(pct) {
  const p = pct > 1 ? pct : pct * 100;
  if (p >= 85) return 'Aprobado';
  if (p >= 60) return 'Parcial';
  return 'Rechazado';
}


// ─── FILTROS ─────────────────────────────────────────────────
function poblarFiltros() {
  const supervisores = [...new Set(datosOriginales.map(d => d.auditor))].sort();
  const tecnicos     = [...new Set(datosOriginales.map(d => d.tecnico))].sort();
  const tipos        = [...new Set(datosOriginales.map(d => d.tipo))].sort();

  const mesesMap = {};
  datosOriginales.forEach(d => {
    if (!d.fecha) return;
    const mesStr = d.fecha.substring(0, 7); // YYYY-MM
    mesesMap[mesStr] = true;
  });
  const mesesOrdenados = Object.keys(mesesMap).sort((a,b) => b.localeCompare(a)); // Descendente
  const mesesDisplay = mesesOrdenados.map(m => {
    const [y, mo] = m.split('-');
    const nombreMes = new Date(+y, +mo - 1).toLocaleDateString('es-CL', { month: 'long', year: 'numeric' });
    return { val: m, display: nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1) };
  });

  const selMes = document.getElementById('filtro-mes');
  if (selMes) {
    const valObj = selMes.value;
    selMes.innerHTML = '<option value="">Todos los meses</option>';
    mesesDisplay.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.val; opt.textContent = m.display;
      if (m.val === valObj) opt.selected = true;
      selMes.appendChild(opt);
    });
  }

  llenarSelect('filtro-zona',     supervisores, 'Todos los supervisores');
  llenarSelect('filtro-auditor',  tecnicos,     'Todos los técnicos');
  llenarSelect('filtro-tipo',     tipos,        'Todos los tipos');
}

function llenarSelect(id, opciones, placeholder) {
  const sel = document.getElementById(id);
  if (!sel) return;
  const val = sel.value;
  sel.innerHTML = `<option value="">${placeholder}</option>`;
  opciones.forEach(o => {
    const opt = document.createElement('option');
    opt.value = o; opt.textContent = o;
    if (o === val) opt.selected = true;
    sel.appendChild(opt);
  });
}

function aplicarFiltros() {
  const mes        = document.getElementById('filtro-mes')?.value       || '';
  const supervisor = document.getElementById('filtro-zona')?.value    || '';
  const tecnico    = document.getElementById('filtro-auditor')?.value || '';
  const tipo       = document.getElementById('filtro-tipo')?.value    || '';
  const texto      = (document.getElementById('filtro-texto')?.value || '').toLowerCase();

  datosFiltrados = datosOriginales.filter(d => {
    if (mes        && (!d.fecha || !d.fecha.startsWith(mes))) return false;
    if (supervisor && d.auditor  !== supervisor) return false;
    if (tecnico    && d.tecnico  !== tecnico)    return false;
    if (tipo       && d.tipo     !== tipo)       return false;
    if (texto      && !`${d.tecnico} ${d.auditor} ${d.direccion} ${d.peticion}`.toLowerCase().includes(texto)) return false;
    return true;
  });

  paginaActual = 1;
  renderizar();
}
function limpiarFiltros() {
  ['filtro-mes','filtro-zona','filtro-auditor','filtro-tipo'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const txt = document.getElementById('filtro-texto');
  if (txt) txt.value = '';
  datosFiltrados = [...datosOriginales];
  paginaActual = 1;
  renderizar();
}

// ─── RENDERIZAR TODO ─────────────────────────────────────────
function renderizar() {
  renderizarKPIs();
  renderizarGraficas();
  renderizarTabla();
  renderizarRanking();
  renderizarHallazgos();
}

// ─── KPI CARDS ───────────────────────────────────────────────
function renderizarKPIs() {
  const total      = datosFiltrados.length;
  const cierres    = datosFiltrados.filter(d => d.cierreProceso && d.cierreProceso.toLowerCase().includes('si')).length;
  const notaProm   = total ? (datosFiltrados.reduce((s,d) => s + (d.nota||0), 0) / total) : 0;
  const puntuacionProm = total ? (datosFiltrados.reduce((s,d) => s + (d.puntuacion||0), 0) / total) : 0;
  const esteticaProm   = total ? (datosFiltrados.reduce((s,d) => s + (d.estetica||0), 0) / total) : 0;

  setKPI('kpi-total',       total,                        'auditorías registradas');
  setKPI('kpi-cumplimiento', notaProm.toFixed(1) + ' / 10', 'nota promedio técnicos');
  setKPI('kpi-cierre',      cierres,                       'cierres con técnico');
  setKPI('kpi-puntuacion',  puntuacionProm.toFixed(1),      'normas y procedimientos');
  setKPI('kpi-estetica',    esteticaProm.toFixed(1),        'estética del cableado');

  renderizarTendenciaMensual(total, notaProm);

  const danilo = datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes('DANILO')).length;
  const rolando = datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes('ROLANDO')).length;
  const julio = datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes('JULIO')).length;

  const superEl = document.getElementById('val-supervisores');
  if (superEl) {
    superEl.innerHTML = `
      <div style="display:flex; justify-content:space-between; color: #3B5069; margin-bottom: 2px;"><span>Danilo:</span> <span>${danilo}</span></div>
      <div style="display:flex; justify-content:space-between; color: #96702F; margin-bottom: 2px;"><span>Rolando:</span> <span>${rolando}</span></div>
      <div style="display:flex; justify-content:space-between; color: #A34A3F;"><span>Julio:</span> <span>${julio}</span></div>
    `;
  }

  renderizarMetaDiaria();
  renderizarPromedioMensual();
  renderizarMesasTrabajo();
}

// ─── META DIARIA (2 auditorías por supervisor) ────────────────
const META_AUDITORIAS_POR_SUPERVISOR = 2;
const SUPERVISORES_META = [
  { key: 'DANILO',  label: 'Danilo',  color: '#3B5069' },
  { key: 'ROLANDO', label: 'Rolando', color: '#96702F' },
  { key: 'JULIO',   label: 'Julio',   color: '#A34A3F' },
];

function fechaLocalHoy() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function fechaLocalAyer() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// Se evalúa el día anterior porque este reporte se lee en la mañana,
// cuando las auditorías del día actual todavía no existen.
function renderizarMetaDiaria() {
  const ayer = fechaLocalAyer();
  const metaTotal = META_AUDITORIAS_POR_SUPERVISOR * SUPERVISORES_META.length;

  const conteos = SUPERVISORES_META.map(s => ({
    ...s,
    count: datosOriginales.filter(d => d.fecha === ayer && d.auditor && d.auditor.toUpperCase().includes(s.key)).length,
  }));
  const totalAyer = conteos.reduce((s, c) => s + c.count, 0);
  const pct = metaTotal ? Math.round((totalAyer / metaTotal) * 100) : 0;

  setKPI('kpi-meta-diaria', pct + '%', `${totalAyer} de ${metaTotal} auditorías de ayer`);

  const detalleEl = document.getElementById('val-meta-diaria-detalle');
  if (detalleEl) {
    detalleEl.innerHTML = conteos.map(c => {
      const cumple = c.count >= META_AUDITORIAS_POR_SUPERVISOR;
      const color = cumple ? '#4F7D64' : c.color;
      const icono = cumple ? '✓' : '';
      return `<div style="display:flex; justify-content:space-between; color:${color}; margin-bottom:2px;"><span>${c.label}:</span> <span>${c.count} / ${META_AUDITORIAS_POR_SUPERVISOR} ${icono}</span></div>`;
    }).join('');
  }
}

// ─── PROMEDIO MENSUAL (auditorías/día vs. meta 2/día) ─────────
// Cuenta días hábiles (lunes a viernes) desde el día 1 del mes actual hasta hoy, inclusive.
function diasHabilesTranscurridos() {
  const hoy = new Date();
  const anio = hoy.getFullYear();
  const mes = hoy.getMonth();
  let habiles = 0;
  for (let dia = 1; dia <= hoy.getDate(); dia++) {
    const dow = new Date(anio, mes, dia).getDay(); // 0=domingo, 6=sábado
    if (dow !== 0 && dow !== 6) habiles++;
  }
  return habiles;
}

function renderizarPromedioMensual() {
  const hoy = new Date();
  const mesActual = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
  const diasTranscurridos = diasHabilesTranscurridos();
  const metaMensual = META_AUDITORIAS_POR_SUPERVISOR * SUPERVISORES_META.length * diasTranscurridos;

  const conteos = SUPERVISORES_META.map(s => {
    const count = datosOriginales.filter(d => d.fecha && d.fecha.startsWith(mesActual) && d.auditor && d.auditor.toUpperCase().includes(s.key)).length;
    const promedioDiario = diasTranscurridos ? count / diasTranscurridos : 0;
    return { ...s, count, promedioDiario };
  });

  const totalMes = conteos.reduce((s, c) => s + c.count, 0);
  const pct = metaMensual ? Math.round((totalMes / metaMensual) * 100) : 0;

  setKPI('kpi-promedio-mensual', pct + '%', `${totalMes} de ${metaMensual} esperadas (${diasTranscurridos} días hábiles)`);

  const detalleEl = document.getElementById('val-promedio-mensual-detalle');
  if (detalleEl) {
    detalleEl.innerHTML = conteos.map(c => {
      const cumple = c.promedioDiario >= META_AUDITORIAS_POR_SUPERVISOR;
      const color = cumple ? '#4F7D64' : c.color;
      return `<div style="display:flex; justify-content:space-between; color:${color}; margin-bottom:2px;"><span>${c.label}:</span> <span>${c.promedioDiario.toFixed(1)} / día</span></div>`;
    }).join('');
  }
}

// ─── MESAS DE TRABAJO POR SUPERVISOR ───────────────────────────
function renderizarMesasTrabajo() {
  const el = document.getElementById('val-mesas-trabajo');
  if (!el) return;

  const conteos = SUPERVISORES_META.map(s => {
    const count = datosFiltrados.filter(d =>
      d.auditor && d.auditor.toUpperCase().includes(s.key) &&
      d.mesaTrabajo && d.mesaTrabajo.trim().toUpperCase() === 'SI'
    ).length;
    return { ...s, count };
  });

  el.innerHTML = conteos.map(c =>
    `<div style="display:flex; justify-content:space-between; color:${c.color}; margin-bottom:2px;"><span>${c.label}:</span> <span>${c.count}</span></div>`
  ).join('');
}

// ─── TENDENCIA vs. MES ANTERIOR ───────────────────────────────
function mesAnteriorStr(mesStr) {
  const [y, m] = mesStr.split('-').map(Number);
  const d = new Date(y, m - 2, 1); // m es 1-indexado; -2 retrocede un mes en base 0
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function filtrarPorMes(mesStr) {
  const supervisor = document.getElementById('filtro-zona')?.value    || '';
  const tecnico    = document.getElementById('filtro-auditor')?.value || '';
  const tipo       = document.getElementById('filtro-tipo')?.value    || '';
  const texto      = (document.getElementById('filtro-texto')?.value  || '').toLowerCase();

  return datosOriginales.filter(d => {
    if (!d.fecha || !d.fecha.startsWith(mesStr)) return false;
    if (supervisor && d.auditor  !== supervisor) return false;
    if (tecnico    && d.tecnico  !== tecnico)    return false;
    if (tipo       && d.tipo     !== tipo)       return false;
    if (texto      && !`${d.tecnico} ${d.auditor} ${d.direccion} ${d.peticion}`.toLowerCase().includes(texto)) return false;
    return true;
  });
}

function renderizarTendenciaMensual(totalActual, notaPromActual) {
  const trendTotalEl = document.getElementById('kpi-total-trend');
  const trendNotaEl  = document.getElementById('kpi-cumplimiento-trend');
  if (!trendTotalEl && !trendNotaEl) return;

  const mesFiltro = document.getElementById('filtro-mes')?.value;
  const mesesConDatos = [...new Set(datosOriginales.map(d => d.fecha).filter(Boolean).map(f => f.substring(0, 7)))].sort();
  const mesRef = mesFiltro || mesesConDatos[mesesConDatos.length - 1];

  if (!mesRef) {
    if (trendTotalEl) trendTotalEl.innerHTML = '';
    if (trendNotaEl) trendNotaEl.innerHTML = '';
    return;
  }

  const filasMesAnterior = filtrarPorMes(mesAnteriorStr(mesRef));

  if (filasMesAnterior.length === 0) {
    if (trendTotalEl) trendTotalEl.innerHTML = `<span class="neutro">Sin datos del mes anterior</span>`;
    if (trendNotaEl) trendNotaEl.innerHTML = `<span class="neutro">Sin datos del mes anterior</span>`;
    return;
  }

  if (trendTotalEl) {
    const deltaTotal = totalActual - filasMesAnterior.length;
    trendTotalEl.innerHTML = formatearTendencia(deltaTotal, `${Math.abs(deltaTotal)} vs. mes anterior`);
  }

  if (trendNotaEl) {
    const notaMesAnterior = filasMesAnterior.reduce((s, d) => s + (d.nota || 0), 0) / filasMesAnterior.length;
    const deltaNota = notaPromActual - notaMesAnterior;
    trendNotaEl.innerHTML = formatearTendencia(deltaNota, `${Math.abs(deltaNota).toFixed(1)} vs. mes anterior`);
  }
}

function formatearTendencia(delta, texto) {
  if (Math.abs(delta) < 0.05) return `<span class="neutro">→ Igual que el mes anterior</span>`;
  const cls = delta > 0 ? 'subida' : 'bajada';
  const flecha = delta > 0 ? '▲' : '▼';
  return `<span class="${cls}">${flecha} ${texto}</span>`;
}

function setKPI(id, valor, sub) {
  const el = document.getElementById(id);
  if (!el) return;
  const valEl = el.querySelector('.kpi-value');
  const subEl = el.querySelector('.kpi-sub');
  if (valEl) valEl.textContent = valor;
  if (subEl) subEl.textContent = sub;
}

// ─── GRÁFICAS ────────────────────────────────────────────────
function renderizarGraficas() {
  renderListaNotas();
}

const CELESTE_PALETTE = ['#7C93AB','#5D7A9B','#4A6483','#8FA6BC','#A9BDCE','#C6D3E0','#3B5069'];
const ESTADO_COLORS   = { aprobado:'#6B9C82', parcial:'#B08A4E', rechazado:'#B25950' };

function renderListaNotas() {
  const cont = document.getElementById('lista-notas');
  if (!cont) return;

  const superFiltro = document.getElementById('filtro-chart-tecnicos')?.value || '';
  const dataForList = superFiltro
    ? datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes(superFiltro))
    : datosFiltrados;

  const techMap = {};
  dataForList.forEach(d => {
    if (!d.tecnico) return;
    if (!techMap[d.tecnico]) techMap[d.tecnico] = { sum: 0, count: 0, auditor: d.auditor || 'Sin Supervisor' };
    techMap[d.tecnico].sum += (d.nota || 0);
    techMap[d.tecnico].count++;
  });

  const list = Object.keys(techMap).map(t => ({
    tecnico: t,
    auditor: techMap[t].auditor,
    pct: Math.round((techMap[t].sum / techMap[t].count / 10) * 100)
  }));

  // Ordenar por supervisor, luego por técnico
  list.sort((a, b) => {
    if (a.auditor === b.auditor) return a.tecnico.localeCompare(b.tecnico);
    return a.auditor.localeCompare(b.auditor);
  });

  if (list.length === 0) {
    cont.innerHTML = `<div style="color:var(--gris-300);font-size:0.9rem;font-style:italic;">Sin datos suficientes</div>`;
    return;
  }

  cont.innerHTML = list.map(item => {
    const color = item.pct >= 85 ? 'var(--verde)' : item.pct >= 60 ? 'var(--ambar)' : 'var(--rojo)';
    const audPrefix = superFiltro ? '' : `<span class="supervisor-link" data-supervisor="${escHtml(item.auditor)}">[${escHtml(item.auditor.split(' ')[0])}]</span> `;
    return `
      <div class="fila">
        <span class="nombre">${audPrefix}<span class="tecnico-link" data-tecnico="${escHtml(item.tecnico)}">${escHtml(item.tecnico)}</span></span>
        <span class="valor" style="color:${color};">${item.pct}%</span>
      </div>`;
  }).join('');
}

// ─── TABLA ───────────────────────────────────────────────────
function renderizarTabla() {
  const tbody = document.getElementById('tabla-body');
  const countEl = document.getElementById('tabla-count');
  const paginacionEl = document.getElementById('tabla-paginacion');
  if (!tbody) return;

  if (countEl) countEl.textContent = `${datosFiltrados.length} registros`;

  if (datosFiltrados.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="7">
        <div class="empty-state">
          <svg class="empty-icon"><use href="#icon-clipboard"/></svg>
          <h3>Sin resultados</h3>
          <p>No hay auditorías que coincidan con los filtros seleccionados.</p>
        </div>
      </td></tr>`;
    if (paginacionEl) paginacionEl.innerHTML = '';
    return;
  }

  const ordenados = [...datosFiltrados].sort((a,b) => (b.fecha||'').localeCompare(a.fecha||''));
  const checklistPorTecnico = analizarChecklistPorTecnico();

  const totalPaginas = Math.max(1, Math.ceil(ordenados.length / FILAS_POR_PAGINA));
  if (paginaActual > totalPaginas) paginaActual = totalPaginas;
  if (paginaActual < 1) paginaActual = 1;
  const inicio = (paginaActual - 1) * FILAS_POR_PAGINA;
  const pagina = ordenados.slice(inicio, inicio + FILAS_POR_PAGINA);

  tbody.innerHTML = pagina.map(d => {
    const nota   = d.nota || 0;
    const pct    = Math.min(Math.max(d.cumplimiento, 0), 100);
    const cls    = pct >= 85 ? '' : pct >= 60 ? 'medio' : 'bajo';
    const estado = d.estado.toLowerCase();
    const badgeCls = estado.includes('aprobado') || estado.includes('cumple') ? 'aprobado' :
                     estado.includes('rechazado') || estado.includes('no cumple') ? 'rechazado' :
                     estado.includes('parcial')   ? 'parcial' : 'pendiente';

    const fechaFormateada = d.fecha
      ? new Date(d.fecha + 'T12:00:00').toLocaleDateString('es-CL', { day:'2-digit', month:'2-digit', year:'numeric' })
      : '—';

    const reincidencias = (checklistPorTecnico[d.tecnico] || []).filter(i => i.count >= 2);
    const badgeReincidente = reincidencias.length
      ? `<span class="badge-reincidente" title="Reincidente en: ${escHtml(reincidencias.map(i => i.label).join(', '))}">⟳ Reincidente</span>`
      : '';

    return `<tr>
      <td>${fechaFormateada}</td>
      <td><code style="font-size:0.75rem;background:#f0f4f8;padding:2px 6px;border-radius:4px">${escHtml(d.peticion)}</code></td>
      <td><strong class="tecnico-link" data-tecnico="${escHtml(d.tecnico)}">${escHtml(d.tecnico)}</strong> ${badgeReincidente}</td>
      <td><span class="supervisor-link" data-supervisor="${escHtml(d.auditor)}">${escHtml(d.auditor)}</span></td>
      <td style="font-size:0.78rem;color:#5A6678">${escHtml(d.direccion)}</td>
      <td>
        <div class="cumplimiento-cell">
          <div class="cumplimiento-bar-bg">
            <div class="cumplimiento-bar ${cls}" style="width:${pct}%"></div>
          </div>
          <span class="cumplimiento-pct">${nota} / 10</span>
        </div>
      </td>
      <td style="font-size:0.75rem;color:#5A6678;max-width:200px">${escHtml(d.observaciones || '—')}</td>
      <td style="font-size:0.75rem;text-align:center">${escHtml(d.mesaTrabajo || '—')}</td>
      <td style="font-size:0.75rem;text-align:center">${escHtml(d.cierreProceso || '—')}</td>
      <td style="font-size:0.75rem;color:#5A6678;max-width:200px">${escHtml(d.obsCierre || '—')}</td>
    </tr>`;
  }).join('');

  if (paginacionEl) {
    if (totalPaginas <= 1) {
      paginacionEl.innerHTML = '';
    } else {
      paginacionEl.innerHTML = `
        <button class="pag-btn" onclick="cambiarPagina(${paginaActual - 1})" ${paginaActual === 1 ? 'disabled' : ''}>‹ Anterior</button>
        <span class="pag-info">Página ${paginaActual} de ${totalPaginas}</span>
        <button class="pag-btn" onclick="cambiarPagina(${paginaActual + 1})" ${paginaActual === totalPaginas ? 'disabled' : ''}>Siguiente ›</button>
      `;
    }
  }
}

function cambiarPagina(n) {
  paginaActual = n;
  renderizarTabla();
}

// ─── UTILIDADES ──────────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ─── FILTRO RÁPIDO (clic en un técnico o supervisor) ───────────
function filtrarPorTecnico(nombre) {
  const sel = document.getElementById('filtro-auditor');
  if (sel) sel.value = nombre;
  aplicarFiltros();
  document.getElementById('tabla-datos')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Filtra por supervisor: muestra en la tabla todos los técnicos que ese supervisor va auditando.
function filtrarPorSupervisor(nombre) {
  const selSup = document.getElementById('filtro-zona');
  const selTec = document.getElementById('filtro-auditor');
  if (selSup) selSup.value = nombre;
  if (selTec) selTec.value = '';
  aplicarFiltros();
  document.getElementById('tabla-datos')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

document.addEventListener('click', (e) => {
  const tec = e.target.closest('.tecnico-link');
  if (tec) { filtrarPorTecnico(tec.dataset.tecnico); return; }
  const sup = e.target.closest('.supervisor-link');
  if (sup) { filtrarPorSupervisor(sup.dataset.supervisor); }
});

function mostrarToast(msg, error = false) {
  const t = document.createElement('div');
  t.className = 'toast' + (error ? ' toast-error' : '');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.classList.add('toast-visible'), 10);
  setTimeout(() => { t.classList.remove('toast-visible'); setTimeout(() => t.remove(), 400); }, 3500);
}

// ─── RANKING TÉCNICOS ───────────────────────────────────────
function renderizarRanking() {
  const containerTop = document.getElementById('ranking-top3');
  const containerBottom = document.getElementById('ranking-bottom3');
  if (!containerTop || !containerBottom) return;

  // 1. Agrupar notas por técnico
  const techMap = {};
  datosFiltrados.forEach(d => {
    if (!d.tecnico) return;
    if (!techMap[d.tecnico]) techMap[d.tecnico] = { sum: 0, count: 0, auditor: d.auditor || 'Sin Supervisor' };
    techMap[d.tecnico].sum += (d.nota || 0);
    techMap[d.tecnico].count++;
  });

  // 2. Convertir a array y calcular promedio
  const list = Object.keys(techMap).map(t => ({
    tecnico: t,
    auditor: techMap[t].auditor,
    promedio: +(techMap[t].sum / techMap[t].count).toFixed(1),
    count: techMap[t].count
  }));

  // Filtrar técnicos sin auditorías válidas (con 0 auditorías, aunque por la lógica arriba todos tendrán al menos 1)
  const listValida = list.filter(l => l.count > 0);

  // 3. Ordenar por promedio descendente
  listValida.sort((a, b) => b.promedio - a.promedio);

  // 4. Obtener Top 3 y Bottom 3
  const top3 = listValida.slice(0, 3);
  let bottom3 = [];
  if (listValida.length > 3) {
    bottom3 = listValida.slice(Math.max(listValida.length - 3, 3)).reverse();
  }

  // 5. Generar HTML interno
  const checklistPorTecnico = analizarChecklistPorTecnico();
  function generarHtml(tecnicos, type) {
    if (tecnicos.length === 0) return `<div style="color:var(--gris-300);font-size:0.9rem;font-style:italic;">Sin datos suficientes</div>`;
    const color = type === 'top' ? 'var(--verde)' : 'var(--rojo)';
    const bg = type === 'top' ? 'var(--verde-bg)' : 'var(--rojo-bg)';
    return tecnicos.map((t, idx) => {
      const fallas = checklistPorTecnico[t.tecnico] || [];
      const fallaTag = (type === 'bottom' && fallas.length)
        ? `<span class="falla-tag">Falla principal: ${escHtml(fallas[0].label)}</span>`
        : '';
      return `
        <div class="ranking-fila" style="background:${bg};">
          <div class="persona">
            <div class="puesto" style="color:${color};">#${idx + 1}</div>
            <div class="info">
              <span class="tecnico tecnico-link" data-tecnico="${escHtml(t.tecnico)}">${escHtml(t.tecnico)}</span>
              <span class="supervisor supervisor-link" data-supervisor="${escHtml(t.auditor)}">${escHtml(t.auditor)}</span>
              ${fallaTag}
            </div>
          </div>
          <div class="medida">
            <span class="nota" style="color:${color};">${t.promedio} / 10</span>
            <span class="cuenta">${t.count} auditoría(s)</span>
          </div>
        </div>
      `;
    }).join('');
  }

  containerTop.innerHTML = generarHtml(top3, 'top');
  containerBottom.innerHTML = generarHtml(bottom3, 'bottom');
}

// ─── TOP HALLAZGOS (CHECKLIST) ─────────────────────────────────
// "malo" indica qué valor cuenta como incumplimiento para cada campo.
// La mayoría es 'NO' (no cumple), pero algunos campos tienen la polaridad
// invertida (ej. reutilizar instalación de otra compañía es lo indeseado).
const CHECKLIST_ITEMS = [
  { key: 'pasamuros',     label: 'No instala pasamuros',                              malo: 'NO' },
  { key: 'roseta',        label: 'Roseta óptica fuera de norma',                      malo: 'NO' },
  { key: 'limpio',        label: 'No deja área de trabajo limpia',                    malo: 'NO' },
  { key: 'capacitacion',  label: 'No capacita al cliente',                            malo: 'NO' },
  { key: 'funcion',       label: 'No cumple correctamente con su función',            malo: 'NO' },
  { key: 'alcohol',       label: 'No utiliza alcohol isopropílico',                   malo: 'NO' },
  { key: 'oneclick',      label: 'No utiliza OneClick en limpieza de conectores',     malo: 'NO' },
  { key: 'dropSoportes',  label: 'No instala drop con soportes',                      malo: 'NO' },
  { key: 'dropNorma',     label: 'Drop fuera de norma de instalaciones',              malo: 'NO' },
  { key: 'reutiliza',     label: 'Reutiliza instalación',                             malo: 'SI' },
  { key: 'reutilizaOtra', label: 'Reutiliza instalación de otra compañía',            malo: 'SI' },
];

// Para cada técnico (sobre datosFiltrados), calcula qué ítems del checklist
// incumple y cuántas veces, ordenado de más a menos frecuente.
// Se usa para la "falla principal" (Bottom 3) y la reincidencia en la tabla.
function analizarChecklistPorTecnico() {
  const porTecnico = {};
  datosFiltrados.forEach(d => {
    if (!d.tecnico) return;
    if (!porTecnico[d.tecnico]) porTecnico[d.tecnico] = [];
    porTecnico[d.tecnico].push(d);
  });

  const resultado = {};
  Object.entries(porTecnico).forEach(([tecnico, registros]) => {
    const items = CHECKLIST_ITEMS.map(item => {
      const count = registros.filter(d =>
        d[item.key] && d[item.key].trim() !== '' &&
        d[item.key].trim().toUpperCase().startsWith(item.malo)
      ).length;
      return { label: item.label, count };
    }).filter(i => i.count > 0).sort((a, b) => b.count - a.count);
    if (items.length) resultado[tecnico] = items;
  });
  return resultado;
}

function renderizarHallazgos() {
  const cont = document.getElementById('hallazgos-list');
  if (!cont) return;

  const resultados = CHECKLIST_ITEMS.map(item => {
    const respuestas = datosFiltrados.filter(d => d[item.key] && d[item.key].trim() !== '');
    const noCumple = respuestas.filter(d => d[item.key].trim().toUpperCase().startsWith(item.malo)).length;
    const pct = respuestas.length ? Math.round((noCumple / respuestas.length) * 100) : 0;
    return { ...item, noCumple, total: respuestas.length, pct };
  }).filter(r => r.total > 0);

  if (resultados.length === 0) {
    cont.innerHTML = `<div style="color:var(--gris-300);font-size:0.9rem;font-style:italic;">Sin datos de checklist disponibles</div>`;
    return;
  }

  resultados.sort((a, b) => b.pct - a.pct);

  cont.innerHTML = resultados.map(r => {
    const color = r.pct >= 50 ? 'var(--rojo)' : r.pct >= 25 ? 'var(--ambar)' : 'var(--verde)';
    return `
      <div class="hallazgo-fila">
        <div class="hallazgo-head">
          <span class="label">${r.label}</span>
          <span class="pct" style="color:${color};">${r.pct}% <span class="conteo">(${r.noCumple}/${r.total})</span></span>
        </div>
        <div class="hallazgo-barra-bg">
          <div class="hallazgo-barra" style="width:${r.pct}%; background:${color};"></div>
        </div>
      </div>`;
  }).join('');
}

// ─── EXPORTACIÓN ─────────────────────────────────────────────
function exportarExcel() {
  if (datosFiltrados.length === 0) {
    mostrarToast('No hay datos para exportar', true);
    return;
  }
  const datosAExportar = datosFiltrados.map(d => ({
    'Fecha': d.fecha,
    'N° Petición': d.peticion,
    'Técnico': d.tecnico,
    'Supervisor': d.auditor,
    'Dirección': d.direccion,
    'Nota': d.nota,
    'Observaciones': d.observaciones,
    'Mesa de Trabajo': d.mesaTrabajo,
    'Cierre Proceso': d.cierreProceso,
    'Obs. Cierre': d.obsCierre
  }));
  const ws = XLSX.utils.json_to_sheet(datosAExportar);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Auditorías');
  const nombreArchivo = `Reporte_Auditorias_${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, nombreArchivo);
  mostrarToast('Exportación completada exitosamente');
}
