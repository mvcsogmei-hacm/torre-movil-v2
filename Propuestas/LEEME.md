# Propuestas de identidad para "La torre"

Cada archivo es una copia completa de `index.html` con **una sola idea añadida**,
para probarla de forma aislada y decidir si pasa a la versión principal.
Ábrelos en el navegador (idealmente en el móvil) y compara.

## Propuesta 1 — Favicon (`propuesta-1-favicon.html`)

Icono de pestaña/favoritos: una torre de ajedrez en tinta oscura sobre el
verde menta de la marca. Es un SVG incrustado en el propio HTML (sin archivos
externos), así que no añade peso ni peticiones.

- **Qué aporta:** identidad en la pestaña del navegador, en favoritos y en el
  historial. Es lo primero que diferencia la página de un archivo genérico.
- **Costo:** una línea en el `<head>`. Cero impacto visual dentro de la app.
- **Nota:** si algún día quieres que al "Añadir a pantalla de inicio" en iPhone
  salga este icono, habría que exportarlo también como PNG de 180×180
  (`apple-touch-icon`), porque iOS no acepta SVG para eso.

## Propuesta 2 — Splash de bienvenida (`propuesta-2-splash.html`)

Pantalla de arranque: al abrir la página aparece el logo de la torre con el
nombre "La torre" durante ~1.5 segundos y se desvanece hacia la pantalla de
inicio. Hecha solo con CSS (sin JavaScript extra) y respeta la preferencia de
"reducir movimiento" del sistema.

- **Qué aporta:** momento de marca al estilo de las apps nativas; le da
  sensación de producto terminado.
- **Costo:** ~1.5 s de espera en cada apertura. En una demo suma; en uso
  diario real puede cansar (las apps bancarias reales lo muestran solo
  mientras cargan de verdad).

## Propuesta 3 — Ambas (`propuesta-3-completa.html`)

Favicon + splash juntos, para ver el efecto combinado.

> Nota: todas las propuestas de esta carpeta se pueden ver también en
> escritorio (se les quitó el aviso de "solo móvil" que sí conserva
> `index.html`).

## Galería de menús (`galeria-menus.html`)

Un solo archivo con **12 estilos de menú inferior** lado a lado, sobre un
esqueleto de la pantalla de inicio, para comparar de un vistazo:
flotante (actual), píldora, dock oscuro, etiquetas, menta, acento limón,
minimal, teclas, cristal, activo expandido, muesca y dividido.
Los estilos 2, 3 y 4 existen además como app completa interactiva
(archivos `menu-a`, `menu-b` y `menu-c`).

## Propuestas de menú inferior

Las tres llevan los mismos 5 botones (Inicio, Inversiones, BOT, Actividades,
Búsqueda) con BOT como botón principal al centro. La referencia es el menú
actual de `index.html`: cinco círculos flotantes sueltos.

### Menú A — Píldora (`menu-a-pildora.html`)
Barra blanca en forma de píldora con los 4 iconos dentro y el BOT
sobresaliendo por arriba del borde, como botón flotante.
Estilo muy "fintech moderna"; el BOT domina visualmente.

### Menú B — Dock oscuro (`menu-b-dock-oscuro.html`)
Barra oscura (tinta) con iconos apagados que se encienden en blanco al estar
activos, y el BOT como círculo amarillo limón al centro. Máximo contraste con
el fondo claro de la app; es la opción más llamativa.

### Menú C — Con etiquetas (`menu-c-etiquetas.html`)
Barra blanca redondeada con icono + nombre debajo de cada botón, y el BOT
flotando por encima con su etiqueta. Es la más clara para usuarios nuevos
(nadie tiene que adivinar qué es cada icono), a costa de ocupar algo más de
altura.

## Galería de tarjeta de proyecto (`galeria-card-proyecto.html`)

**Seis propuestas** de la tarjeta con que el buscador muestra cada proyecto,
con más color y jerarquía que la versión sobria. Todas conservan la fila
**PIM 2026** + **% de ejecución 2026** y colorean el avance con una sola regla
(≥60 óptimo, 25–59 en proceso, <25 crítico; gris cuando el PIM es S/ 0). Cada
estilo se ve con tres proyectos de muestra (avance alto, medio y paralizado con
PIM 0, como el caso real de Singa):

1. **Barra de estado + medidor** — barra lateral de color por estado y medidor fino del %.
2. **Cabecera menta** — cabecera con degradé menta; el % en píldora limón.
3. **Anillo de ejecución** — el avance como anillo circular protagonista.
4. **Franja de datos en tinta** — franja inferior oscura con PIM y % en limón (máximo contraste).
5. **Insignias de cartera + barra** — carteras con sus colores y % grande sobre barra.
6. **Acento limón (hero %)** — el % enorme sobre cabecera limón, estilo tarjeta de saldo.

## Buscador regional funcional (`buscador-regional.html`)

**Prototipo que sí funciona**, no una maqueta. Genera ~500 proyectos de muestra
en 5 departamentos (Amazonas, Cusco, Puno, Lima, Loreto) con su jerarquía real
departamento → provincia → distrito, para resolver el problema de "buscar
Amazonas y que aparezcan cientos de tarjetas con scroll infinito".

Combina las cuatro ideas discutidas:
- **Drill-down (A):** si un conjunto supera el umbral (25), en vez de listar
  proyectos muestra el siguiente nivel geográfico con **contadores** y barras.
  Se navega con migas de pan (Inicio › Amazonas › Bagua) que crecen al tocar.
- **Facetas (B):** chips de **cartera** (con contadores que se recalculan),
  y selects de **programa** y **rango de % de ejecución**, todo combinable.
- **Umbral (C):** nunca vuelca la lista mientras el total sea grande; muestra
  contador + invita a acotar.
- **Paginación (D):** cuando el conjunto es manejable, lista las tarjetas v5 de
  a 20 con botón **"Cargar más"**.

Toda la pantalla habla con una sola capa `filtrar(state)` (equivalente al futuro
`buscarProyectos()`), de modo que al pasar a BBDD/API solo cambia esa función,
no la interfaz. La tarjeta usada es la **propuesta 5** (insignias de cartera + barra).

## Ficha de proyecto · 5 modelos (`galeria-ficha-proyecto.html`)

Cinco maneras de presentar la **pantalla de detalle**, dibujadas con **datos
reales** de la Matriz Única de Monitoreo (31.07.2026). Arriba hay un conmutador
para ver las cinco con dos proyectos distintos —uno *Paralizada + Transferencia*
y otro *PRESET + Obra directa*— y comprobar que cada una oculta lo que no aplica.

- **A · Resumen + acordeones** — cabecera en tinta con 3 KPI y barra de avance;
  el resto en secciones plegables (solo la primera abierta). La de menos scroll.
- **B · Mosaico de cifras** — los números en mosaico de 2 columnas y los textos
  en una ficha debajo. Todo a la vista, se lee de un golpe.
- **C · Línea de proceso** — línea de tiempo PRESET → Expediente técnico →
  Proceso de selección → Ejecución → Paralizada → Cierre, marcando dónde está.
- **D · Pestañas** — Resumen · Inversión · Seguimiento · una pestaña por cartera.
- **E · Bloques con cabecera de color** — el modelo actual pero con cada bloque
  en su tarjeta y cabecera de color (las carteras usan su color). El más conservador.

Reglas que respetan las cinco: **los campos sin dato no se dibujan** (nada de
filas con "—"), **PRESET** muestra *etapa de evaluación* y *estado*, y
**Paralizadas** muestra *hito* y *fecha estimada o real de reactivación*.

## Indicador de Títulos · 5 diseños (`galeria-indicador-titulos.html`)

Cinco maneras de mostrar el indicador de la pestaña **Títulos de propiedad**
(pantalla Actividades), con los datos reales de la hoja INDICADORES TITULOS:
el **44,7%** siempre como cifra protagonista, y **meta 49,844** / **22,265
entregados** como secundarios en distinta ubicación:

1. **Franja inferior** — secundarios en la parte blanca de la tarjeta, como la franja de accesos de Inicio.
2. **Dos tarjetas debajo** — estilo "Entidades adscritas", con mini barra en Entregados.
3. **Chips en el panel** — dos píldoras en tinta con cifras en limón, dentro del amarillo.
4. **Barra con extremos** — barra de avance al 44,7%; entregados a la izquierda, meta a la derecha.
5. **Mitades en el panel** — línea divisoria dentro del amarillo y secundarios en dos mitades.

## Recomendación

La **propuesta 1 (favicon) vale la pena siempre**: costo cero, beneficio claro.
La **propuesta 2 (splash)** es cuestión de gusto: luce mucho en presentaciones
y demos, pero si la página va a usarse a diario conviene acortarla o quitarla.

Para pasar una propuesta a `index.html`, avísame y copio solo el bloque
correspondiente (la línea del favicon y/o la sección de splash).

## Galería de ranking (`galeria-ranking.html`, 18-ago-2026)

10 propuestas visuales para la tabla **Top 5** de la pantalla Ranking
(hoja `RANKING SECTOR`, tema rojo actual): 1 actual (referencia), 2 podio,
3 medallas, 4 tabla con PIM y % a la vez, 5 tarjetas por puesto, 6 gráfico de
barras, 7 Vivienda protagonista ("Puesto N / 30"), 8 código como insignia,
9 columnas, 10 número tenue de fondo. El segmentador es igual en todas.

## Galería de ranking en tarjeta (`galeria-ranking-tarjetas.html`, 18-ago-2026)

20 propuestas bajo el formato ya adoptado (tarjeta blanca con título dentro,
mini segmentador % / PIM arriba a la derecha, lista Top 5): variaciones de
filas (barra, cebra, número simple/grande, medallas, código insignia, dos
líneas con PIM, barra ancha, chip, fila Vivienda roja, compacta, tabla,
barra corta, punto semáforo, nombre en tinta, puesto/30, Vivienda fija
arriba) y del segmentador (texto, cuadrado). Datos reales de RANKING SECTOR.

## Propuesta ranking 11 (`propuesta-ranking-11.html`, 18-ago-2026)

Desarrollo de la propuesta 11 de la galería anterior (fila Vivienda en banda
roja) con puesto en cuadro redondeado y botones Ejecución / PIM. Muestra el
caso normal y 4 alternativas (A separador "···", B sin separador, C fila fija
arriba "Tu sector", D pie "puesto 9 de 30") para cuando el sector queda fuera
del Top 5.
