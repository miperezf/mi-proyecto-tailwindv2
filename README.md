# Pedidos Comercial Frutam

Aplicación web para armar pedidos comerciales, agruparlos por **Mail ID** y generar el
cuerpo del correo (HTML) que se pega en Outlook. Publicada en GitHub Pages:
https://miperezf.github.io/mi-proyecto-tailwindv2/

## Stack
- React 19 + Vite 6 + Tailwind CSS v4
- Firebase: Authentication (anónima) y Cloud Firestore
- Despliegue: GitHub Actions → rama `gh-pages` en cada push a `main`

## Desarrollo local
```bash
npm install
cp .env.example .env.local   # completar con la configuración web de Firebase
npm run dev
```

Variables requeridas (en local `.env.local`; en CI como *secrets* del repositorio):
`VITE_FIREBASE_API_KEY`, `VITE_FIREBASE_AUTH_DOMAIN`, `VITE_FIREBASE_PROJECT_ID`,
`VITE_FIREBASE_STORAGE_BUCKET`, `VITE_FIREBASE_MESSAGING_SENDER_ID`,
`VITE_FIREBASE_APP_ID`, `VITE_FIREBASE_MEASUREMENT_ID`.

> La configuración web de Firebase no es secreta (queda en el bundle). La protección real
> de los datos está en las **reglas de Firestore**.

## Datos (Firestore)
- `artifacts/{projectId}/public/data/pedidos/{id}` → `{ header, items }`
  - `items` se guarda como string JSON.
  - `header.status`: `draft` | `sent` | `deleted` (borrado lógico; nunca se borra físicamente).
- `artifacts/{projectId}/public/data/presence/{mailId}/users/{uid}` → aviso de edición simultánea.

## Versionado
- `APP_VERSION` / `APP_VERSION_DATE` en `src/App.tsx`: subir en cada release.
- `WHATSNEW_VERSION`: cambiar solo cuando cambie el contenido del modal de novedades.
- Cada build publica `version.json` (SHA del commit). La app lo consulta cada 5 min y
  avisa cuando hay una versión nueva.

## Seguridad
- Nunca subir credenciales de servicio (`*adminsdk*.json`); están en `.gitignore`.
- Los scripts con Firebase Admin SDK deben ejecutarse fuera de este repositorio.
