# GestiónApp — Cobranzas & Visitas

App de gestión con Next.js + Google Sheets como base de datos.

---

## 🚀 Deploy en Vercel (3 pasos)

### 1. Subir a GitHub
```bash
# En la carpeta del proyecto
git init
git add .
git commit -m "Initial commit"
gh repo create cobranzas-app --public --push
```

### 2. Conectar en Vercel
1. Ir a [vercel.com](https://vercel.com) → **Add New Project**
2. Importar el repo de GitHub
3. En **Environment Variables** agregar:
   ```
   GOOGLE_SHEET_URL = https://script.google.com/macros/s/...tu-url.../exec
   ```
4. Click **Deploy** ✅

### 3. Alternativa: Vercel CLI
```bash
npm i -g vercel
vercel --prod
# Cuando pida env vars, ingresar GOOGLE_SHEET_URL
```

---

## 📊 Configurar Google Sheets

### Si tu script YA está desplegado:
Verificar que el Apps Script acepta GET con `?action=getData&sheet=Cobranzas`
y POST con `{action, sheet, data}`. Ver `google-apps-script.js` para la estructura.

### Si necesitás actualizar el script:
1. Abrir Google Sheets → **Extensiones → Apps Script**
2. Reemplazar todo el contenido con `google-apps-script.js`
3. Guardar (Ctrl+S)
4. **Ejecutar → initSheets** (crea las hojas automáticamente)
5. **Implementar → Administrar implementaciones → editar** la existente
6. Cambiar versión a "Nueva versión" → **Implementar**

---

## 📁 Estructura del proyecto
```
├── pages/
│   ├── _app.js          # App wrapper
│   ├── index.js         # App principal (UI completa)
│   └── api/
│       └── sheets.js    # Proxy a Google Sheets (evita CORS)
├── styles/
│   └── globals.css      # Tailwind + estilos globales
├── google-apps-script.js # Pegar en Apps Script de tu hoja
├── next.config.js
├── tailwind.config.js
└── package.json
```

---

## 🔌 API del Apps Script

| Acción | Tipo | Params |
|--------|------|--------|
| Leer datos | GET | `?action=getData&sheet=Cobranzas` |
| Agregar fila | POST | `{action:"addRow", sheet, data:{...}}` |
| Actualizar fila | POST | `{action:"updateRow", sheet, id, data:{...}}` |
| Eliminar fila | POST | `{action:"deleteRow", sheet, id}` |

### Campos Cobranzas
`id · cliente · monto · fechaVencimiento · estado · notas`

### Campos Visitas
`id · cliente · direccion · fecha · hora · estado · notas`

---

## 🛠 Desarrollo local
```bash
npm install
npm run dev
# → http://localhost:3000
```

Crear `.env.local`:
```
GOOGLE_SHEET_URL=https://script.google.com/macros/s/.../exec
```
