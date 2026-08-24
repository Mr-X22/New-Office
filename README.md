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
| Documentos | `.docx` → editor | editor → `.docx` | Formato simplificado: texto, negritas/cursivas/subrayado, títulos, listas. No conserva imágenes ni tablas todavía. |
| Hojas de cálculo | `.xlsx` → grid | grid → `.xlsx` | Una sola hoja, sin fórmulas todavía (solo valores). |
| Presentaciones | — | editor → `.pptx` | Aún no lee `.pptx` existentes (es la pieza más difícil de las tres); por ahora solo crea presentaciones nuevas. |

## Siguientes pasos sugeridos (por prioridad)

1. **Guardado en formato propio** (JSON ligero) para autoguardado rápido sin
   pasar por docx/xlsx/pptx en cada tecla — exportar a Office solo al final.
2. **Import de `.pptx`** — es el hueco más grande; probablemente valga la pena
   una librería distinta a pptxgenjs solo para lectura (o parsear el XML del
   pptx a mano, ya que es un zip con XML adentro).
3. **Fórmulas básicas en la hoja de cálculo** (SUMA, PROMEDIO, referencias
   tipo `A1`).
4. **Imágenes en Documentos y Presentaciones.**
5. Icono real de la app (los `icons/icon-*.png` actuales son placeholder).

Dime con cuál de los tres módulos seguimos afinando primero y le metemos
profundidad ahí.
