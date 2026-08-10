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
let chartDanilo     = null;
let chartRolando    = null;
let chartJulio      = null;
let chartBarras     = null;
let chartLineas     = null;
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
      estado,
      observaciones: String(f['Observaciones (detallar observaciones)'] || f['Observaciones '] || ''),
      // Campos adicionales de checklist
      pasamuros:     String(f['Instala Pasa-muros ']                 || ''),
      roseta:        String(f['Roseta Óptica en norma y atornillada ']|| ''),
      limpio:        String(f['Deja Área de Trabajo Limpio ']        || ''),
      capacitacion:  String(f['Capacitación  al Cliente']            || ''),
      
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
  renderizar();
}

// ─── RENDERIZAR TODO ─────────────────────────────────────────
function renderizar() {
  renderizarKPIs();
  renderizarGraficas();
  renderizarTabla();
  renderizarRanking();
}

// ─── KPI CARDS ───────────────────────────────────────────────
function renderizarKPIs() {
  const total      = datosFiltrados.length;
  const cierres    = datosFiltrados.filter(d => d.cierreProceso && d.cierreProceso.toLowerCase().includes('si')).length;
  const notaProm   = total ? (datosFiltrados.reduce((s,d) => s + (d.nota||0), 0) / total).toFixed(1) : 0;

  setKPI('kpi-total',       total,              'auditorías registradas');
  setKPI('kpi-cumplimiento', notaProm + ' / 10', 'nota promedio técnicos');
  setKPI('kpi-cierre',      cierres,             'cierres con técnico');

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
  renderChartBarras();
  renderChartLineas();
  renderChartsSupervisores();
}

const CELESTE_PALETTE = ['#7C93AB','#5D7A9B','#4A6483','#8FA6BC','#A9BDCE','#C6D3E0','#3B5069'];
const ESTADO_COLORS   = { aprobado:'#6B9C82', parcial:'#B08A4E', rechazado:'#B25950' };

function renderChartBarras() {
  const ctx = document.getElementById('chart-barras')?.getContext('2d');
  if (!ctx) return;

  const superFiltro = document.getElementById('filtro-chart-tecnicos')?.value || '';
  const dataForChart = superFiltro 
    ? datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes(superFiltro))
    : datosFiltrados;

  const techMap = {};
  dataForChart.forEach(d => {
    if (!d.tecnico) return;
    if (!techMap[d.tecnico]) techMap[d.tecnico] = { sum: 0, count: 0, auditor: d.auditor || 'Sin Supervisor' };
    techMap[d.tecnico].sum += (d.nota || 0);
    techMap[d.tecnico].count++;
  });

  const list = Object.keys(techMap).map(t => ({
    tecnico: t,
    auditor: techMap[t].auditor,
    promedio: +(techMap[t].sum / techMap[t].count).toFixed(1)
  }));

  // Ordenar por supervisor, luego por técnico
  list.sort((a,b) => {
    if (a.auditor === b.auditor) return a.tecnico.localeCompare(b.tecnico);
    return a.auditor.localeCompare(b.auditor);
  });

  const labels = list.map(item => {
    const partes = item.tecnico.split(' ');
    const nombreStr = partes.length >= 2 ? `${partes[0]} ${partes[1]}` : item.tecnico;
    const audPrefix = superFiltro ? '' : `[${item.auditor.split(' ')[0]}] `;
    return `${audPrefix}${nombreStr}`;
  });

  const valores = list.map(item => item.promedio);
  
  const backgroundColors = list.map(item => {
    const aud = item.auditor.toUpperCase();
    if (aud.includes('DANILO')) return '#3B5069'; // Azul
    if (aud.includes('ROLANDO')) return '#96702F'; // Naranja
    if (aud.includes('JULIO')) return '#A34A3F'; // Rojo
    return '#B0BBC9'; // Gris default
  });

  // Ajustar altura del contenedor dependiendo de la cantidad de técnicos
  const numTecnicos = labels.length;
  const minHeight = 300;
  const dynamicHeight = Math.max(minHeight, numTecnicos * 25 + 50);
  ctx.canvas.parentNode.style.height = `${dynamicHeight}px`;

  if (chartBarras) chartBarras.destroy();
  chartBarras = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Nota promedio (0-10)',
        data: valores,
        backgroundColor: backgroundColors,
        borderRadius: 4,
        borderSkipped: false,
        barThickness: 8, // Barras muy delgadas
      }]
    },
    options: {
      indexAxis: 'y', // Convertir a gráfico de barras horizontal
      responsive: true,
      maintainAspectRatio: false, // Permitir que la altura sea dinámica
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` Nota: ${ctx.raw} / 10` } }
      },
      scales: {
        x: {
          min: 0, max: 10,
          grid: { color: '#ECEFF4' },
          ticks: { font: { size: 11 } }
        },
        y: {
          grid: { display: false },
          ticks: { font: { size: 10 }, autoSkip: false } // No saltar nombres
        }
      }
    }
  });
}

function renderChartLineas() {
  const ctx = document.getElementById('chart-lineas')?.getContext('2d');
  if (!ctx) return;

  // Obtener meses únicos para el eje X
  const mesesSet = new Set();
  datosFiltrados.forEach(d => { if (d.fecha) mesesSet.add(d.fecha.substring(0,7)); });
  const labels = [...mesesSet].sort();
  
  // Configuración de los supervisores
  const superConfigs = [
    { key: 'DANILO', label: 'Danilo Ojeda', color: '#3B5069' },
    { key: 'ROLANDO', label: 'Rolando Montoya', color: '#96702F' },
    { key: 'JULIO', label: 'Julio Cabrera', color: '#A34A3F' }
  ];

  const datasets = superConfigs.map(conf => {
    const data = labels.map(m => {
      const grupo = datosFiltrados.filter(d => d.fecha && d.fecha.startsWith(m) && d.auditor && d.auditor.toUpperCase().includes(conf.key));
      return grupo.length ? +(grupo.reduce((s,d) => s+(d.nota||0), 0) / grupo.length).toFixed(1) : null;
    });
    
    return {
      label: conf.label,
      data,
      borderColor: conf.color,
      backgroundColor: conf.color,
      borderWidth: 2.5,
      pointBackgroundColor: conf.color,
      pointRadius: 4,
      pointHoverRadius: 6,
      fill: false,
      tension: 0.35,
      spanGaps: true
    };
  });

  const labelsFormateados = labels.map(m => {
    const [y, mo] = m.split('-');
    return new Date(+y, +mo-1).toLocaleDateString('es-CL', { month:'short', year:'2-digit' });
  });

  if (chartLineas) chartLineas.destroy();
  chartLineas = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labelsFormateados,
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { display: true, position: 'top', labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.raw} / 10` } }
      },
      scales: {
        y: {
          min: 0, max: 10,
          grid: { color: '#ECEFF4' },
          ticks: { font: { size: 11 } }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11 } }
        }
      }
    }
  });
}

// ─── GRÁFICAS DE SUPERVISORES ────────────────────────────────
function getWeekNumber(dStr) {
  const d = new Date(dStr);
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + String(weekNo).padStart(2, '0');
}

function renderChartsSupervisores() {
  chartDanilo = renderChartSupervisor('DANILO', chartDanilo, 'chart-danilo', '#3B5069');
  chartRolando = renderChartSupervisor('ROLANDO', chartRolando, 'chart-rolando', '#96702F');
  chartJulio = renderChartSupervisor('JULIO', chartJulio, 'chart-julio', '#A34A3F');
}

function renderChartSupervisor(nombre, chartRef, canvasId, color) {
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return chartRef;

  const tipoVista = document.getElementById('filtro-tiempo')?.value || 'mensual';
  const supervisorData = datosFiltrados.filter(d => d.auditor && d.auditor.toUpperCase().includes(nombre));
  
  const agrupado = {};
  supervisorData.forEach(d => {
    if (!d.fecha) return;
    const clave = tipoVista === 'mensual' ? d.fecha.substring(0,7) : getWeekNumber(d.fecha);
    if (!agrupado[clave]) agrupado[clave] = { count: 0, sum: 0 };
    agrupado[clave].count++;
    agrupado[clave].sum += (d.nota || 0);
  });

  const labels = Object.keys(agrupado).sort();
  const dataPromedio = labels.map(k => +(agrupado[k].sum / agrupado[k].count).toFixed(1));
  const dataCount = labels.map(k => agrupado[k].count);
  
  const labelsFormateados = labels.map(k => {
    if (tipoVista === 'mensual') {
      const [y, mo] = k.split('-');
      return new Date(+y, +mo-1).toLocaleDateString('es-CL', { month:'short', year:'2-digit' });
    }
    return k.replace('-W', ' Sem ');
  });

  if (chartRef) chartRef.destroy();
  return new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labelsFormateados,
      datasets: [
        {
          type: 'line',
          label: 'Nota Promedio',
          data: dataPromedio,
          borderColor: color,
          backgroundColor: color,
          borderWidth: 2,
          yAxisID: 'y',
          tension: 0.3
        },
        {
          type: 'bar',
          label: 'Cantidad Auditorías',
          data: dataCount,
          backgroundColor: color + '40', // opacidad aprox 25%
          borderRadius: 4,
          yAxisID: 'y1'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { type: 'linear', display: true, position: 'left', min: 0, max: 10, title: { display: true, text: 'Nota', font: { size: 10 } }, ticks: { font: { size: 10 } } },
        y1: { type: 'linear', display: true, position: 'right', min: 0, grid: { drawOnChartArea: false }, title: { display: true, text: 'Volumen', font: { size: 10 } }, ticks: { font: { size: 10 }, precision: 0 } }
      }
    }
  });
}

// ─── TABLA ───────────────────────────────────────────────────
function renderizarTabla() {
  const tbody = document.getElementById('tabla-body');
  const countEl = document.getElementById('tabla-count');
  if (!tbody) return;

  if (countEl) countEl.textContent = `${datosFiltrados.length} registros`;

  if (datosFiltrados.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="7">
        <div class="empty-state">
          <div class="empty-icon">📋</div>
          <h3>Sin resultados</h3>
          <p>No hay auditorías que coincidan con los filtros seleccionados.</p>
        </div>
      </td></tr>`;
    return;
  }

  const ordenados = [...datosFiltrados].sort((a,b) => (b.fecha||'').localeCompare(a.fecha||''));

  tbody.innerHTML = ordenados.map(d => {
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

    return `<tr>
      <td>${fechaFormateada}</td>
      <td><code style="font-size:0.75rem;background:#f0f4f8;padding:2px 6px;border-radius:4px">${escHtml(d.peticion)}</code></td>
      <td><strong>${escHtml(d.tecnico)}</strong></td>
      <td>${escHtml(d.auditor)}</td>
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
}

// ─── UTILIDADES ──────────────────────────────────────────────
function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

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
  function generarHtml(tecnicos, type) {
    if (tecnicos.length === 0) return `<div style="color:var(--gris-400);font-size:0.9rem;font-style:italic;">Sin datos suficientes</div>`;
    return tecnicos.map((t, idx) => {
      const colorCls = type === 'top' ? 'color:#4F7D64;' : 'color:#A34A3F;';
      const bgCls = type === 'top' ? 'background:#E8F1EC;' : 'background:#F5E8E6;';
      return `
        <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px; border: 1px solid var(--gris-200); border-radius: 8px; ${bgCls}">
          <div style="display: flex; align-items: center; gap: 12px;">
            <div style="font-weight: 700; font-size: 1.1rem; ${colorCls} min-width: 24px;">#${idx + 1}</div>
            <div style="display: flex; flex-direction: column;">
              <span style="font-weight: 600; color: var(--gris-800); font-size: 0.95rem;">${escHtml(t.tecnico)}</span>
              <span style="font-size: 0.75rem; color: var(--gris-400);">${escHtml(t.auditor)}</span>
            </div>
          </div>
          <div style="display: flex; flex-direction: column; align-items: flex-end;">
            <span style="font-weight: 700; font-size: 1.1rem; ${colorCls}">${t.promedio} / 10</span>
            <span style="font-size: 0.7rem; color: var(--gris-400);">${t.count} auditoría(s)</span>
          </div>
        </div>
      `;
    }).join('');
  }

  containerTop.innerHTML = generarHtml(top3, 'top');
  containerBottom.innerHTML = generarHtml(bottom3, 'bottom');
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
