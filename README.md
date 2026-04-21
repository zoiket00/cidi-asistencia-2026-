# CIDI Asistencia — Fundación Juanfe

## Estructura

```
cidi-asistencia/
├── server.js          ← servidor Express (proxy + estáticos)
├── package.json
├── .gitignore
└── public/
    ├── index.html
    ├── app.js
    ├── style.css
    └── img/
        └── Logo-Juanfe-verde-1-1.png
```

## Deploy en Railway

1. Subir este proyecto a un repo de GitHub
2. Entrar a railway.app → New Project → Deploy from GitHub
3. Seleccionar el repo
4. Railway detecta el package.json y hace el deploy automático
5. En Settings → Networking → Generate Domain → copiar la URL

## Correr localmente

```bash
npm install
npm start
# Abre http://localhost:3000
```

## Cómo agregar un bebé nuevo

1. La coordinadora abre el Google Sheet maestro
2. Agrega la fila en la hoja del día correspondiente
3. Esperar ~2 min (caché de Google Sheets)
4. La profesora recarga la app → aparece automáticamente
