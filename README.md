# Carpeta — Suite ofimática ligera

Esqueleto base de una suite tipo Word / Excel / PowerPoint como PWA, pensada
para equipos viejos o con pocos recursos (poca RAM, disco lento, a veces sin
internet).

## Cómo probarla

No se puede abrir con doble clic (`file://`) porque usa módulos ES e
`import()` dinámico, que los navegadores bloquean por CORS en local. Necesita
un servidor estático mínimo:

```bash
cd office-suite
python3 -m http.server 8080
# abre http://localhost:8080
```

Para publicarla igual que tus otros proyectos, súbela a GitHub Pages tal cual
— no necesita build ni Node.

## Decisiones de arquitectura (y por qué, pensando en PCs viejas)

- **Sin framework** (nada de React/Vue): solo JS vanilla con módulos ES. Cero
  bundler, cero build, arranque instantáneo.
- **Carga perezosa por módulo** (`js/core/app-shell.js`): al abrir la app solo
  se descarga la "concha" (launcher + shell). Word/Excel/PowerPoint se
  cargan con `import()` dinámico *solo* cuando el usuario entra a esa app.
- **Librerías pesadas (docx, xlsx, pptxgenjs, mammoth) por CDN, bajo demanda**:
  no se tocan hasta que el usuario abre o exporta ese tipo de archivo
  específico. El navegador las cachea después de la primera vez.
- **Archivos reales, no "en la nube"**: usa la File System Access API
  (`js/core/storage.js`) para abrir/guardar el archivo directo del usuario en
  Chrome/Edge. En navegadores que no la soportan (o Firefox), cae
  automáticamente al flujo clásico de input/descarga — sigue funcionando,
  solo pierde el "guardar rápido sin diálogo".
- **Service worker mínimo** (`sw.js`): solo cachea la concha para que abra
  offline; no pre-descarga las librerías pesadas para no gastar disco de
  entrada.

## Qué ya funciona

| Módulo | Abrir | Guardar/Exportar | Notas |
|---|---|---|---|
| Documentos | `.docx` → editor | editor → `.docx` | Texto con **negrita/cursiva/subrayado conservados de verdad** (incluso mezclados en un mismo párrafo), títulos, listas, **imágenes** incrustadas y **tablas**. Autoguardado local cada ~800ms (red de seguridad, no sustituye a "Guardar"). |
| Hojas de cálculo | `.xlsx` → grid | grid → `.xlsx` | Una sola hoja. **Fórmulas reales**: `=SUMA(A1:A5)`, `=PROMEDIO(B1:B5)`, referencias sueltas (`=A1+B2*3`), rangos. Al guardar, exporta tanto el valor calculado como la fórmula (Excel la reconoce y la recalcula al abrir). |
| Presentaciones | `.pptx` → diapositivas (solo texto) | editor → `.pptx` | Ahora sí lee `.pptx` existentes extrayendo el texto de cada diapositiva (título + cuerpo); no reconstruye diseño, imágenes ni animaciones — es lectura de contenido, no una reconstrucción visual completa. |

## Siguientes pasos sugeridos (por prioridad)

1. **Presentaciones: reconstrucción visual real del `.pptx`** (posiciones, imágenes, estilos) — hoy solo se recupera el texto plano de cada diapositiva.
2. **Hoja de cálculo: más funciones** (SI/IF, CONTAR, BUSCARV) y varias hojas por libro.
3. **Documentos: estilos de párrafo adicionales** (alineación, tamaño y color de fuente) al exportar — hoy el tamaño/color visual del editor no viaja al `.docx`.
4. **Sincronizar autoguardado también en Hojas de cálculo y Presentaciones** (hoy solo Documentos lo tiene).
5. Icono real de la app (los `icons/icon-*.png` actuales son placeholder).

Dime con cuál de los tres módulos seguimos afinando primero y le metemos
profundidad ahí.
