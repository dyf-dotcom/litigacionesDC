# DC-Litigaciones

Sistema web para gestión de causas judiciales de la Defensoría.

## Stack

- **Frontend:** HTML + CSS + Vanilla JS (sin frameworks)
- **Backend / API:** Google Apps Script (deployado como Web App)
- **Base de datos:** Google Sheets — hoja `Causas`

## Estructura

```
litigacionesDC/
├── app/
│   ├── etapa1_lista.html      ← Lista de todas las causas
│   ├── etapa1_detalle.html    ← Alta y edición — Etapa 1: Instancia
│   ├── etapa2_lista.html      ← Lista filtrada — Etapa 2
│   └── etapa2_detalle.html    ← Edición — Etapa 2: Casación
├── script/
│   └── codigo_apps_script.js  ← API (Google Apps Script)
├── docs/
│   └── Guia_DCLitigaciones.docx
├── .gitignore
└── README.md
```

## Flujo de una causa

```
Alta (Instancia)  →  Casación (TCP)  →  Corte / Federal
  crearEtapa1         actualizarEtapa2    actualizarEtapa3
```

## Acciones de la API

| Acción              | Tipo       | Descripción                        |
|---------------------|------------|------------------------------------|
| `listar`            | Lectura    | Devuelve todas las causas          |
| `buscar`            | Lectura    | Busca por NroCausa o NroCasación   |
| `crearEtapa1`       | Escritura  | Nueva causa (Instancia)            |
| `actualizarEtapa1`  | Update     | Edita datos de Instancia           |
| `actualizarEtapa2`  | Update     | Avanza a Casación (TCP)            |
| `actualizarEtapa3`  | Update     | Cierra en Corte / Federal          |

## Configuración

En cada HTML, reemplazar la constante `API` con la URL del Web App de Google Apps Script:

```js
const API = 'https://script.google.com/macros/s/TU_ID/exec';
```
