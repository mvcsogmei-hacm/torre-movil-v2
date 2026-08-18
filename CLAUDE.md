# La torre — Torre móvil v.2

App móvil tipo fintech ("La torre") para monitoreo de proyectos de inversión pública del Perú (CUI, programas PNSR/PNSU/PMIB, etc.).

## Estructura

- Un solo `index.html` autocontenido (HTML + CSS + JS, sin dependencias externas), más `Data/proyectos.js` con los datos reales.
- **PWA instalable** (17-ago-2026): `manifest.json` + `sw.js` (red primero, caché de respaldo) + `iconos/` (torre blanca sobre rojo, generados con `Pillow`). El botón central del **BOT se retiró del menú** (la pantalla `screen-bot` queda para el futuro); el menú inferior tiene 4 botones. Cabecera del PDF: membrete izquierdo del MVCS y a la derecha "OGMEI / La torre de control".
- Estética (desde 17-ago-2026): **colores institucionales del sector Vivienda / Estado peruano** — rojo institucional `#C8102E` (degradado a `#A50D23`) como color primario, texto blanco/rosado `#F7C8C2` sobre paneles rojos, fondos neutros cálidos. Fuente Space Grotesk, marco de teléfono en escritorio. OJO: las variables CSS conservan sus nombres históricos (`--lemon` = rojo institucional, `--mint` = rojo profundo del fondo); hay una nueva `--rosa`. El PDF descargable también usa el tema rojo. El tema anterior (menta/limón) sobrevive solo en `Propuestas/`.
- Carpeta `Propuestas/`: variaciones aisladas (favicon, splash, menús, galerías), documentadas en `Propuestas/LEEME.md`.
- Carpeta `Data/`: el Excel fuente (`31.07.2026 - Matriz Única de Monitoreo - Consolidado.xlsx`) y el `proyectos.js` generado desde él.
- Pantallas: Inicio, Inversiones, BOT, Actividades, Búsqueda + ficha de proyecto.
- La pantalla Búsqueda funciona con **datos reales**: `index.html` carga `data/proyectos.js` (define `window.PROYECTOS`, 21,528 proyectos).

## Formato de cifras en soles (regla desde 18-ago-2026)

**Sin abreviaturas** en todo el aplicativo: nada de "M", "MM", "mill." ni "millones". Las cifras agregadas (tarjetas de pliegos en Inicio, indicador y rankings de la pantalla Ranking) se muestran en millones **solo con el número** ("S/ 4,279"); la unidad se declara una sola vez en Inicio, bajo el título "A nivel de pliegos del sector" como subtítulo gris **"En millones de soles"** (`.section-head.sh-sub .sub`). En la ficha de proyecto y el PDF (`fmtMonto`) los montos van **en millones con 1 decimal** ("S/ 1,2") y la unidad se declara en el título de sección: "Información financiera (en millones de soles)" y "Transferencias 2026 (en millones de soles)". La tarjeta de resultados del buscador y el texto de "Compartir" (`fmtSoles`) siguen en soles completos.

## Importante

Esta app es **totalmente independiente**: NO tiene relación con la carpeta `Desktop\la_torre` ni con sus datos, API o base de datos (aclarado por el usuario el 14-ago-2026). No usar ese proyecto como referencia.

## Buscador con datos reales — estado (16-ago-2026)

Objetivo (acordado 13-ago-2026): hacer funcional el buscador de proyectos con datos reales. **La fase Excel → JSON ya está hecha** (14-ago-2026).

### Fuente de datos real

`Data\31.07.2026 - Matriz Única de Monitoreo - Consolidado.xlsx`: una sola hoja **CONSOLIDADO**, 21,528 proyectos × 29 columnas. Es una **matriz plana consolidada**, no el modelo normalizado que se había diseñado (hoja UNIVERSAL + una hoja por cartera): las carteras vienen como columnas con "-" cuando no aplican, y las transferencias solo como totales 2026 (transferido/ejecutado), **no una fila por transferencia**.

### `Data/proyectos.js` (generado con `Data/generar_proyectos.py`)

- Regenerar con: `py generar_proyectos.py` desde la carpeta `Data/` (openpyxl vía `py`).
- `window.PROYECTOS = [...]`, UTF-8, 21,528 proyectos, 0 CUIs duplicados.
- 23 campos comunes por proyecto: `cui, nombre, programa, uei, dep, prov, dist, modalidad, tipo, pobl, cxAgua, cxAlc, ssi, monto, devAc, pim, dev, fisico, estadoET, procSel, ssp, subSsp, fTerm`. Los vacíos ("-" en el Excel) van como `null`; `fisico` en % (el Excel lo guarda como fracción).
- Objeto `carteras` según **reglas de pertenencia definidas por el usuario (16-ago-2026)**:
  - `preset` (8,104): ETAPA DE EVALUACIÓN distinta de vacío y distinta de FINANCIADO → `{etapa, estado}`
  - `obras` (2,864): MODALIDAD = DIRECTA → `{avance}`
  - `transferencias` (1,150): MONTO TOTAL TRANSFERIDO 2026 > 0 → `{transferido, ejecutado}` (el % de ejecución se calcula en la UI)
  - `paralizadas` (247): ESTADO SSP = PARALIZADA o PARALIZADO → `{avance, hito, fecha}`
- Programas: PNSR 10,542 · PMIB 7,541 · PNSU 3,317 · PNC 66 · PASLC 47 · SEDAPAL 9 · PGSU 3 · SENCICO 3.

### Pantalla Inicio y detalle de Pliego (datos reales desde 17-ago-2026)

- **Indicador principal de Inicio** (desde 18-ago-2026): título "Sector", subtítulo "Vivienda, Construcción y Saneamiento"; el % grande sale de la nueva hoja `INDICADOR PRINCIPAL` (fila SECTOR → `INDICADORES.principal.pct`, hoy 70,8% → **redondeado a entero** 71%), sin línea secundaria y sin botón de ojo. El botón **"Ver ranking"** (y toda la parte roja de la tarjeta) abre la **pantalla Ranking** (`screen-ranking`, nueva 18-ago-2026): `balance-card` con "Sector / Vivienda, Construcción y Saneamiento", % con decimal, barra, y debajo en líneas separadas PIM y Devengado (filas PIM y DEVENGADO de la misma hoja `INDICADOR PRINCIPAL` → `INDICADORES.principal.{pct,pim,devengado}`); Debajo, **3 tarjetas** (`.rank-card`, estilo de las tarjetas de Inversiones), cada una con su título dentro y un **mini segmentador** arriba a la derecha (`.mini-seg`: **Ejecución** / **PIM**): "Ranking total Gobierno Nacional" (hoja `RANKING SECTOR` → `INDICADORES.rankingSector`), "Ranking por inversiones" (`RANKING INVERSIONES` → `rankingInversiones`) y "Ranking por actividades" (`RANKING ACTIVIDADES` → `rankingActividades`), todas `[{cod,nombre,pim,pct}]` (el generador separa el código "NN:" del nombre). Contenido: **lista Top 5** en filas (posición en **cuadro redondeado** `.pos`, "NN: NOMBRE", valor = % con semáforo o "S/ N M"), ordenada desc según el botón; la fila de VIVIENDA va como **banda roja completa** (`.dr.mio`, texto blanco, puesto blanco sobre rojo — propuesta 11 de `Propuestas/galeria-ranking-tarjetas.html`, variante A de `Propuestas/propuesta-ranking-11.html`, elegidas 18-ago-2026) y, si queda fuera del Top 5, se añade al final tras un separador "···" con su puesto real. Todo lo monta `montarRanking(lista, prefijo)` con ids `<prefijo>-seg` / `<prefijo>-tabla` (`rank`, `rankInv`, `rankAct`). El diseño podio (`Propuestas/galeria-ranking.html` n.º 2) se probó y se descartó el 18-ago-2026 por desordenado. De los 3 botones inferiores, el primero se llama **"Pliego MVCS"**. La fila PLIEGO de la hoja `MVCS` (`INDICADORES.mvcs[0]`) queda solo como resumen del detalle de Pliego MVCS; el desglose Actividades/Proyectos con decimales vive solo en el detalle.
- **Temas** (17-ago-2026): botón luna/sol en Inicio + sección "Tema de color" en Ajustes con **4 paletas** (rojo institucional por defecto, azul, verde, morado — `html[data-paleta]`) y **modo oscuro** (`html[data-tema="oscuro"]`, alto contraste: fondo #0E0F12, tarjetas #282B32 con borde, texto blanco). Persistido en `torre-ajustes` (`paleta`, `tema`). El bloque CSS de paletas va ANTES del oscuro para que en oscuro ganen las superficies oscuras. **El PDF no hereda tema**: siempre colores fijos institucionales. Iconos superiores derechos en variante `.icon-btn.solid` (círculo tinta, icono claro). La hoja Inversiones ya no tiene botón superior derecho.
- **Cabecera de Inicio**: se quitaron los botones de notificaciones y estadísticas; quedan el botón de **tema** (luna/sol) y el de **Ajustes** (engranaje) que abre `screen-ajustes`, con **funcionalidad real** (persistida en `localStorage` clave `torre-ajustes`): interruptor de Notificaciones; Tamaño de texto Pequeño/Normal/Grande (aplica `zoom` al `.phone` y se recuerda entre sesiones); Idioma (valor fijo); "Datos cargados" expandible con el resumen vivo de la matriz (total y conteo por cartera desde `PROYECTOS`) + botón "Recargar datos"; y "Acerca de" expandible.
- **Detalle de Pliego** (`screen-pliego`): tarjeta resumen con la fila PLIEGO + **un card por cada fila restante** de la hoja (ADM GENERAL, PNSU, PNSR, PASLC), cada uno con filas Todo/Actividades/Proyectos con barras semáforo (mismos helpers `seccionPct`/`filaPct`).
- **Tarjetas "A nivel de pliegos del sector"** (antes "Entidades adscritas", desde 18-ago-2026): hoja **`PLIEGOS`** (`INDICADORES.pliegos`, columnas ENTIDAD/PIM/DEVENGADO/EJECUCIÓN): 5 tarjetas — MVCS (ancho completo, `.ent.wide`), SBN, SENCICO, OTASS, COFOPRI — con % de ejecución redondeado, barra y "S/ dev M de S/ PIM M". La celda de ejecución es fórmula: el generador abre el Excel con `data_only=True` y, si viene vacía, calcula dev/PIM. La hoja `ADSCRITAS` ya no existe en el Excel.
- **Detalle de Pliego MVCS** (`screen-pliego`, título "Pliego MVCS"): hoja **`MVCS`** (antes `PLIEGO`), exportada como `INDICADORES.mvcs` (renombrada de `pliego` el 18-ago-2026 para no confundirla con `pliegos`). Se abre **solo** desde el acceso "Pliego MVCS" de Inicio.

### Pantalla Actividades

Retitulada de "Movimientos" a "Actividades". Tres pestañas estilo accesos de Inicio: **Títulos de propiedad** (activa), **Bonos** y **Wasiymi** — desde 18-ago-2026 los 3 botones van **dentro de la misma tarjeta del indicador** (como el indicador de Inicio: panel rojo arriba + `neu-actions` debajo); el bloque `#act-tabs` es único y el JS lo mueve a la tarjeta `.tit-ind` de la pestaña activa; el botón activo se marca con sombreado gris neutro (#E9EAEC — no `--key`, que es rosado), no en rojo; misma letra (12,5 px) que los accesos de Inicio. Las tres comparten estructura y helpers (`filaPct`, `seccionPct`, `indicadorPrincipal`, `segControl`), con datos del Excel exportados como `window.INDICADORES`:

**Bonos** (`INDICADORES.bonos`, hojas `INDICADORES BONOS` / `MODALIDAS BONOS` / `POR REGION BONOS`): indicador diseño 4 (desembolsados÷meta, hoy 25,3% — 8,768 de 34,685) + sub-pestañas **Por modalidad** y **Por región**, ambas con el semáforo estándar de avance (≥60 verde / 25–59 ámbar / <25 rojo) — el usuario pidió el semáforo también para modalidades (16-ago-2026).

**Wasiymi** (`INDICADORES.wasiymi`, hojas `INDICADORES WAYSIMI` / `POR REGION WAYSIMI` — **así escritas en el Excel, con "WAYSIMI"**): indicador diseño 4 (ejecución física÷meta, hoy 4,8% — 309 de 6,399) y, **sin sub-pestañas**, directamente la sección "Wasiymi por región" (14 regiones con semáforo; Piura viene como 1 en el Excel → 100%).

**Títulos** (`INDICADORES.titulos`):

- **Indicador principal (diseño 4 de `Propuestas/galeria-indicador-titulos.html`, elegido 16-ago-2026)**: % grande (entregados÷meta, hoy 44,7%) + barra de avance con extremos "22,265 entregados" / "Meta 49,844". Hoja `INDICADORES TITULOS`.
- **Sub-pestañas** (control segmentado): **Macro región** (hoja `MACROREGION TITULOS`, sección "Por macro región") y **Por regiones** (hoja `POR REGION TITULOS`, una sección por macro región). Mismo formato que la ficha de proyecto: título `section-head` + tarjeta `detail-card` con filas nombre → % coloreado por la regla ≥60/25/25 y **barra de porcentaje** bajo cada fila (filas `.drb`).
- El generador normaliza los % de esas hojas (`pct_celda`): texto "52,1%" → 52.1; fracción ≤1 → ×100; número >1 ya es %.
- El contenido demo fintech de esta pantalla se eliminó (queda demo solo en la pantalla "Pagar").

### Pantalla Inversiones

Las 4 tarjetas plegables (`.cartera-card`: Cartera priorizada, Paralizadas, Transferencias 2026 con desglose estático por programa; PRESET con desglose calculado desde `PROYECTOS`) **se mantienen** con su % a la derecha (fijos 41,5 / 41,2 / 18,2; PRESET = aptos ÷ total). Desde 18-ago-2026 el **resumen** de cada tarjeta muestra bajo el nombre: (solo PRESET) el subtítulo "Plataforma para evaluar Exp. Técnicos de los GS", **"N proyectos"** e **"Inversión: S/ N"** (millones, sin unidad), tomados de la hoja **`INVERSIONES`** (`INDICADORES.inversiones[{nombre, proyectos, inversion}]`, emparejada por nombre sin mayúsculas; JS "resumen de cada tarjeta"). La tarjeta Obras directas se retiró el 18-ago-2026. Vendrán más cambios en esta pantalla.

### Ficha de proyecto — módulos

Fijos (todos los proyectos): cabecera con el mismo formato que el indicador principal de Inicio (tarjeta `neu-card` blanca + panel limón centrado) — línea 1: departamento - provincia - distrito, línea 2: CUI, línea 3: título del proyecto, debajo chips de cartera (solo las carteras a las que pertenece el CUI: si no está en ninguna no sale chip) **más chips grises de estado** (`estadoChips(p)`, regla del usuario 18-ago-2026): si ESTADO SSP tiene valor y no es PRESET → chips [Estado SSP, Sub estado SSP]; si el proyecto está en la cartera PRESET (tiene ETAPA DE EVALUACIÓN) → solo [Etapa de evaluación] (el chip azul de cartera ya dice PRESET, no se duplica) — este caso manda porque en la matriz esos proyectos traen ESTADO SSP vacío (8,098 de 8,104), y solo 1 dice "PRESET"; si ESTADO SSP = PRESET sin etapa → [PRESET]; si ESTADO SSP está vacío o es "-" → dos chips separados: "Expediente técnico" y <ESTADO ET>. En el PDF van como línea "Estado: A · B" bajo "Carteras: …"; franja inferior con **2 acciones: "Descargar PDF"** y **"Compartir"** (Web Share API nativa — incluye WhatsApp y correo; si el navegador no la soporta, hoja propia con WhatsApp y Correo). El PDF se **genera a mano en JS** (`generarPdfBlob`, sin librerías): descarga directa y compartir como adjunto. Diseño **sobrio institucional** (17-ago-2026): página blanca, membrete gris "MINISTERIO DE VIVIENDA…" + filete rojo superior, título en negro, carteras como línea de texto, secciones en negrita con subrayado fino rojo y filas con separadores de hilo gris — sin bloques de color ni cebra. Una columna, una hoja A4, adaptativo (compacta si hay muchas secciones). No hereda el tema de la app. Las secciones de la ficha y del PDF salen de una única fuente: `fichaDatos(p)`. · Datos generales · Beneficiarios · Información financiera 2026 (con % ejec. financiera = dev/PIM calculado) · **Estado situacional** (antes "Seguimiento"; **se oculta si el proyecto pertenece a PRESET**). Condicionales: una sección por cartera a la que pertenece el CUI (PRESET, Obras directas, Transferencias 2026 con % de ejecución calculado, Paralizadas). Los encabezados de sección van sin punto de color.

### Modelo de datos normalizado (diseño de referencia, para cuando la fuente lo permita)

- **Matriz UNIVERSAL**: una fila por proyecto, CUI como llave única. Solo campos comunes. Nada específico de una cartera.
- **Una matriz por cartera**: primera columna CUI + solo los campos propios de esa cartera. Pertenencia a una cartera = que el CUI aparezca en su matriz. Un proyecto puede estar en varias carteras.
- **Transferencias es la única matriz donde el CUI se repite**: una fila por transferencia (fecha, año, monto, dispositivo legal). En las demás matrices un CUI duplicado es error.
- Los agregados (total, n° de transferencias, subtotal por año) se **calculan**, nunca se escriben en el Excel.

### Fase siguiente: BBDD

Desde cero, sin reutilizar nada previo: tablas `proyectos`, `carteras` (catálogo), `proyecto_cartera` (puente), tablas de detalle por cartera, tabla `transferencias` 1-a-muchos. Mantener una sola capa de acceso (`buscarProyectos()`) para poder cambiar JSON→API sin tocar la UI.

### Pendientes

- **Historial de transferencias**: el Excel consolidado solo trae totales 2026; el historial por fecha/año diseñado para la ficha (agrupado por año con subtotales, opción recomendada) necesita otra fuente.
- ~~Mayúsculas en la ruta~~: resuelto 17-ago-2026 — el HTML ahora pide `Data/proyectos.js`, igual que la carpeta. Repo: `mvcsogmei-hacm/torre-movil-v2`.
