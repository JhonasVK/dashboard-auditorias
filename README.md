# 📋 Dashboard Auditorías de Terreno

Dashboard web interactivo para visualizar indicadores de auditorías de terreno, 
publicado gratuitamente en GitHub Pages.

---

## 📁 Estructura de la carpeta

```
dashboard-auditorias\
│
├── auditorias.xlsx          ← TU ARCHIVO EXCEL (actualiza este diariamente)
│
├── actualizar.bat           ← DOBLE CLIC para publicar los cambios
├── setup_primera_vez.bat    ← Ejecutar SOLO la primera vez
│
├── index.html               ← Página web del dashboard
├── css\style.css            ← Estilos
├── js\dashboard.js          ← Lógica del dashboard
│
└── data\
    └── auditorias.xlsx      ← Copia automática (no tocar)
```

---

## 🚀 Uso diario (después del setup inicial)

### Flujo de trabajo:

```
1. Abre "auditorias.xlsx" → actualiza tus datos → Guarda y cierra
2. Haz doble clic en "actualizar.bat"
3. Espera ~1 minuto → ¡tu página web ya tiene los datos nuevos!
```

---

## ⚙️ Configuración inicial (solo una vez)

### Requisitos previos

1. **Git para Windows** — Descárgalo en: https://git-scm.com/download/win
2. **Cuenta de GitHub** — Regístrate gratis en: https://github.com

### Pasos

1. Crea un repositorio en GitHub:
   - Ve a https://github.com/new
   - Nombre: `dashboard-auditorias`
   - Tipo: **Público**
   - Sin README ni .gitignore
   - Haz clic en **Create repository**

2. Ejecuta `setup_primera_vez.bat` en esta carpeta
   - Ingresa tu usuario de GitHub cuando lo pida
   - Ingresa tu email y nombre
   - Si pide contraseña, usa un **token de GitHub** (ver abajo)

3. Activa GitHub Pages:
   - Ve a tu repositorio → **Settings** → **Pages**
   - Branch: `main`, carpeta: `/ (root)`
   - Clic en **Save**

4. En ~2 minutos tendrás tu URL:
   ```
   https://TU-USUARIO.github.io/dashboard-auditorias/
   ```

---

## 🔑 Cómo crear un Token de GitHub

Si Git te pide contraseña, usa un token personal:

1. Ve a: https://github.com/settings/tokens
2. Haz clic en **"Generate new token (classic)"**
3. Escribe una descripción (ej: "Dashboard auditorías")
4. Selecciona el permiso **`repo`**
5. Haz clic en **Generate token**
6. **Copia el token** (solo se muestra una vez)
7. Úsalo como contraseña cuando Git lo solicite

---

## 📊 Columnas del Excel (formato recomendado)

Tu Excel puede tener cualquier nombre de columnas — el sistema las detecta automáticamente.  
Para mejores resultados, usa columnas con estos nombres (o similares):

| Columna | Descripción | Ejemplo |
|---|---|---|
| `Fecha` | Fecha de la auditoría | 15/01/2024 |
| `Zona` | Zona o región | Zona Norte |
| `Lugar` | Sitio o establecimiento | Planta Maipú |
| `Auditor` | Nombre del auditor | Ana Pérez |
| `Cumplimiento` | Porcentaje (0–100 o 0–1) | 85 |
| `Estado` | Aprobado / Parcial / Rechazado | Aprobado |
| `Hallazgos` | Cantidad de observaciones | 3 |
| `Tipo` | Categoría de auditoría | Seguridad |

---

## ❓ Problemas comunes

| Problema | Solución |
|---|---|
| "Git no está instalado" | Descarga Git desde https://git-scm.com |
| "No se pudo subir a GitHub" | Verifica tu token de GitHub o conexión a internet |
| "No se pudo copiar el Excel" | Cierra el archivo en Excel antes de ejecutar el .bat |
| La página no se actualiza | Espera 2-3 minutos o refresca con Ctrl+F5 |
| Los datos no se cargan | Verifica que el Excel esté en la carpeta raíz |
