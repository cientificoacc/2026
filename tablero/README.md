# Tablero de control MIPG · FURAG

Tablero web del cierre de brechas del FURAG 2025 de la Agencia Catastral de Cundinamarca.
Sitio **estático**: tres archivos, sin servidor, sin base de datos y sin dependencias externas.

```
tablero/
├── index.html    estructura y estilos (identidad ACC / GOV.CO)
├── tablero.js    lógica del filtro y los gráficos (SVG dibujado a mano)
├── datos.js      los datos exportados del libro Excel (327 KB)
├── .nojekyll     evita que GitHub Pages procese el sitio con Jekyll
└── README.md     este archivo
```

## Cómo verlo sin publicar nada

Abra `index.html` con doble clic. Funciona desde el disco porque los datos van en un
`.js` y no en un `.json`: un `fetch()` de archivo local lo bloquearía el navegador.

## Publicarlo en GitHub Pages

1. Cree un repositorio (puede ser privado; Pages funciona igual en cuentas con plan de pago,
   y en cuentas gratuitas el repositorio debe ser público).
2. Suba **el contenido de esta carpeta a la raíz** del repositorio:

```bash
git init
git add .
git commit -m "Tablero MIPG FURAG"
git branch -M main
git remote add origin https://github.com/USUARIO/REPOSITORIO.git
git push -u origin main
```

3. En el repositorio: **Settings → Pages → Source: Deploy from a branch →
   Branch: `main` / `(root)` → Save.**
4. A los dos o tres minutos queda en `https://USUARIO.github.io/REPOSITORIO/`.

Si prefiere conservar la carpeta `tablero/` dentro del repositorio, publique desde
`main / docs` renombrándola a `docs`, o suba solo esta carpeta a la raíz.

## Enlaces directos por dependencia

El filtro queda guardado en la dirección, así que cada líder puede recibir su propio enlace:

```
https://USUARIO.github.io/REPOSITORIO/?area=RES-04    → Grupo de Talento Humano
https://USUARIO.github.io/REPOSITORIO/?area=RES-09    → Subgerencia de Tecnología
```

Los identificadores están en la hoja `RESPONSABLE` del libro `pruebamodeloMIPG.xlsx`.

## Actualizar los datos

Los datos no se editan a mano. Cuando cambie el libro, regenere `datos.js`:

```bash
python3 build/exportar_tablero.py
```

y vuelva a subir el archivo. La estructura de `datos.js` es una sola asignación
(`window.DATOS = {...}`), de modo que el resto del tablero no se toca.

## Decisiones que conviene conocer

- **Sin librerías de gráficos.** Todo es SVG generado en el navegador. Evita cargar un CDN
  externo, que en muchas entidades está bloqueado, y hace que el sitio funcione sin internet
  salvo por la tipografía.
- **La tipografía Nunito Sans** se carga desde Google Fonts para igualar el sitio oficial. Si
  no hay internet, cae a la tipografía del sistema sin romper nada.
- **Paleta validada.** Los colores de estado y prioridad se verificaron con un validador de
  contraste y de visión cromática deficiente (deuteranopía, protanopía, tritanopía) tanto en
  modo claro como en el de alto contraste. Ningún dato se distingue solo por color: todos los
  gráficos llevan leyenda y etiqueta directa.
- **Modo claro fijo.** El tablero abre siempre en claro, sin heredar el modo oscuro del sistema
  operativo. Los botones de contraste e impresión quedan ocultos por CSS pero siguen cableados:
  para devolverlos basta con quitar la regla `.tema{display:none}` del `<style>`.
- **Logos institucionales.** El encabezado usa los logos oficiales servidos desde
  `https://www.acc.gov.co/assets/logos/`. Si prefiere no depender del sitio externo, descargue
  `LogoGOB.png` y `LogoACC.png`, guárdelos en `assets/logos/` dentro de esta carpeta y cambie las
  dos rutas del `<img>` por `assets/logos/…`. El logo de la Gobernación y el separador se ocultan
  por debajo de 768 px, igual que en el sitio oficial.
- **Responsive.** Probado en escritorio (1440 px), iPad (768 px) e iPhone (375 px). Los
  indicadores pasan de seis a tres y a dos columnas; el filtro se vuelve de ancho completo; los
  gráficos redibujan su geometría —canal de etiquetas más corto, meses abreviados— en lugar de
  encogerse hasta volverse ilegibles; y las tablas se desplazan dentro de su propia caja sin
  arrastrar la página.
- **Impresión.** Existe una hoja de estilos de impresión que oculta filtros y navegación. Para
  usarla, imprima desde el navegador (⌘P).

## Qué muestra

| Acto | Contenido |
|---|---|
| 1 · Cómo está organizada la entidad | Las 7 dimensiones y 19 políticas del MIPG con su peso en el formulario |
| 2 · Qué nos preguntó el DAFP | Las 452 preguntas por política y por tipo de respuesta |
| 3 · Qué respondimos | Estado de cada respuesta, cobertura de soportes y el detalle de las 183 brechas |
| 4 · Qué vamos a hacer | Las 187 acciones, el cronograma mensual de agosto a diciembre de 2026 y el plan por dependencia |

El indicador «prácticas acreditadas» es una medida **interna de cobertura**, no el puntaje
oficial del DAFP: la Función Pública no publica los pesos por pregunta.
