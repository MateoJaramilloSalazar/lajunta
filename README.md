# Priorización de temas para nuestras tertulias

Encuesta web estática que usa el método **MaxDiff** (máxima diferencia) para priorizar los temas de nuestras tertulias: en cada pantalla, quien responde elige el tema que más le interesa y el que menos, y con eso se construye un ranking de preferencias.

**Encuesta en línea:** https://mateojaramillosalazar.github.io/lajunta/

## Cómo funciona

- `index.html` es la encuesta completa (HTML, CSS y JavaScript en un solo archivo) y se publica con GitHub Pages.
- Al enviarla, las respuestas se mandan a una **hoja de Google** mediante un **Google Apps Script** publicado como aplicación web (la constante `ENDPOINT` dentro de `index.html`).
- `apps-script/Codigo.gs` es el código del script vinculado a esa hoja. Aquí está solo como respaldo y documentación; el que se ejecuta de verdad es el que está en el editor de Apps Script de la hoja.

## Actualizar la encuesta

```bash
# editar index.html y luego:
git add index.html
git commit -m "Actualiza la encuesta"
git push
```
