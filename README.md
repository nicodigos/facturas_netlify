# Netlify invoice app

Proyecto paralelo y autocontenido para migrar la app de Streamlit a Netlify con:

- frontend estatico en `html/css/js`,
- funciones serverless en `netlify/functions`,
- configuracion propia en `.env`.

## Estructura

- `index.html`: shell principal con tabs para base de datos y procesamiento
- `assets/css/styles.css`: estilos
- `assets/js/*.js`: estado, UI, SharePoint, PDF y procesamiento
- `netlify/functions/*.js`: secretos y llamadas server-side

## Variables

Llena `netlify_invoice_app/.env` o configura las mismas variables en Netlify:

- `OPENAI_API_KEY`
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `REDIRECT_URI`
- `SP_HOSTNAME`
- `SP_SITE_PATH`
- `SP_DRIVE_NAME`
- `RECEIPTS_DATABASE_DIR`

## Desarrollo local

```powershell
cd C:\golang\netlify_invoice_app
npm install
npx netlify dev
```

Si `netlify dev` falla con tu version de Node, usa el servidor local del proyecto:

```powershell
cd C:\golang\netlify_invoice_app
npm install
npm run dev
```

Este servidor expone el frontend en `http://localhost:8888` y monta las functions en `/.netlify/functions/*` y `/api/*`.

## Notas

- Las credenciales de OpenAI, Google Vision y el secreto de Microsoft no se exponen al navegador; solo las usan las functions.
- El token de Microsoft si termina en el navegador para operar contra Graph, porque esta app necesita leer/escribir SharePoint desde cliente.
- El frontend descarga librerias desde CDN: `pdf.js`, `pdf-lib` y `SheetJS`.
- La exportacion del resumen procesado queda en CSV. La exportacion filtrada de base usa XLSX en cliente.
- No se reutilizo Python ni Streamlit; esta carpeta puede desplegarse aparte.
