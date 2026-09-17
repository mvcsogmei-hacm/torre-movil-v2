# -*- coding: utf-8 -*-
"""
Genera Data/proyectos.js a partir del Excel consolidado.

Uso:  py generar_proyectos.py  (desde la carpeta Data)

Reglas de pertenencia a carteras (definidas por el usuario, 16-ago-2026):
  - paralizadas:    ESTADO SSP = PARALIZADA o PARALIZADO
  - transferencias: MONTO TOTAL TRANSFERIDO 2026 > 0
  - preset:         ETAPA DE EVALUACIÓN distinta de vacío y distinta de FINANCIADO
  - obras:          MODALIDAD DE FINANCIAMIENTO INTERNA = DIRECTA

Desde 16-sep-2026 el consolidado y los indicadores pueden venir de archivos
distintos (EXCEL_CONSOLIDADO / EXCEL_INDICADORES) y las columnas de la matriz
se ubican por nombre de cabecera (COLUMNAS), no por posición.
"""
import json
import datetime
import unicodedata
import openpyxl

# Dos fuentes (desde 16-sep-2026):
#  - EXCEL_CONSOLIDADO: matriz de proyectos. Puede traer muchas más columnas
#    que las 29 de la matriz; las columnas se UBICAN POR NOMBRE de cabecera
#    (fila que empieza con "CUI"), no por posición.
#  - EXCEL_INDICADORES: hojas de indicadores (INDICADOR PRINCIPAL, PLIEGOS,
#    MVCS, RANKING…, INVERSIONES, TITULOS, BONOS, WAYSIMI). Si es el mismo
#    archivo que el consolidado, basta con poner el mismo nombre.
EXCEL_CONSOLIDADO = "04.09.2026 - Matriz Única de Monitoreo - Consolidado.xlsx"
EXCEL_INDICADORES = "Data.xlsx"
SALIDA = "proyectos.js"
HOJA = "CONSOLIDADO"

# Columnas de la matriz: campo -> nombres de cabecera aceptados (el primero es
# el histórico; los demás, alias vistos en versiones posteriores del Excel).
# Se comparan sin tildes ni mayúsculas (norm_cab).
COLUMNAS = {
    "cui":         ["CUI"],
    "programa":    ["PROGRAMA"],
    "nombre":      ["NOMBRE DEL PROYECTO"],
    "uei":         ["UEI"],
    "dep":         ["DEPARTAMENTO"],
    "prov":        ["PROVINCIA"],
    "dist":        ["DISTRITO"],
    "modalidad":   ["MODALIDAD DE FINANCIAMIENTO INTERNA"],
    "tipo":        ["TIPO DE PROYECTO"],
    "pobl":        ["POBLACION BENEFICIARIA"],
    "cxAgua":      ["CONEXIONES NUEVAS DE AGUA"],
    "cxAlc":       ["CONEXIONES NUEVAS DE ALCANTARILLADO/UBS"],
    "ssi":         ["ESTADO DE LA INVERSION-SSI"],
    "monto":       ["COSTO DE INVERSION"],
    "devAc":       ["DEVENGADO ACUMULADO"],
    "pim":         ["PIM 2026"],
    "dev":         ["DEVENGADO 2026"],
    "etapa":       ["ETAPA DE EVALUACION"],
    "estadoEval":  ["ESTADO"],
    # En el Excel de 04.09.2026 "ESTADO ET" pasó a llamarse "ESTADO2"
    # (grupo "ET (DIRECTAS) FUENTE SSP"): verificado 21,517/21,528 valores iguales.
    "estadoET":    ["ESTADO ET", "ESTADO2"],
    # …y "ESTADO DEL PROCESO DE SELECCIÓN" pasó a "ESTADO DEL PROCESO"
    # (grupo "ACTOS PREVIOS (FUENTE SEACE)"): 21,466/21,528 iguales.
    "procSel":     ["ESTADO DEL PROCESO DE SELECCION", "ESTADO DEL PROCESO"],
    "ssp":         ["ESTADO SSP"],
    "subSsp":      ["SUB ESTADO-SSP"],
    "fTerm":       ["FECHA REAL DE TERMINO DE OBRA"],
    "fisico":      ["% AVANCE DE OBRA REAL"],
    "hito":        ["HITO DE REACTIVACION"],
    "fReact":      ["FECHA ESTIMADA O REAL DE REACTIVACION"],
    "transferido": ["MONTO TOTAL TRANSFERIDO 2026"],
    "ejecutado":   ["MONTO TOTAL EJECUTADO 2026"],
    # Opcional (solo en el Excel nuevo): sirve para reconstruir ESTADO SSP = PRESET.
    "bdPreset":    ["BD PRESET"],
    # Opcional: estado del expediente técnico de las TRANSFERENCIAS (grupo "ET
    # TRANSFERENCIAS"); respaldo de estadoET cuando ESTADO2 viene vacío (17-sep-2026).
    "estadoETTr":  ["ESTADO Y SUB ESTADO"],
}
OPCIONALES = {"bdPreset", "estadoETTr"}

# Normalización de vocabulario que en el Excel llega con y sin tilde o con distinta
# capitalización (17-sep-2026): así el mismo valor no aparece dos veces.
TIPO_NORMAL = {
    "EXPEDIENTE TECNICO": "EXPEDIENTE TÉCNICO",
    "EXPEDIENTE TECNICO (SALDO)": "EXPEDIENTE TÉCNICO (SALDO)",
}


def normal_tipo(v):
    v = limpio(v)
    return TIPO_NORMAL.get(v.upper(), v) if v else v


def normal_proc(v):
    """'contratado' -> 'Contratado' (misma capitalización que el resto de valores)."""
    v = limpio(v)
    return v[0].upper() + v[1:] if v else v


def estado_et(r):
    """ESTADO2 (ET directas); si viene vacío, ESTADO Y SUB ESTADO (ET transferencias),
    ignorando sus marcadores vacíos '---' y 'SIN ESTADO SSP'."""
    v = limpio(r["estadoET"])
    if v:
        return v
    alt = limpio(r.get("estadoETTr")) if "estadoETTr" in r else None
    if alt and alt.upper() not in ("---", "SIN ESTADO SSP"):
        return alt
    return None


def norm_cab(v):
    """Cabecera comparable: sin tildes, mayúsculas, espacios colapsados."""
    if v is None:
        return None
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    return " ".join(s.replace("\n", " ").upper().split())


def ubicar_columnas(ws):
    """Busca la fila de cabecera (primera celda 'CUI') y devuelve
    (número de la fila de cabecera, {campo: índice de columna, base 0})."""
    for i, fila in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        if norm_cab(fila[0]) == "CUI":
            cab = [norm_cab(c) for c in fila]
            break
    else:
        raise SystemExit("No se encontró la fila de cabecera (celda 'CUI') en la hoja %s" % ws.title)
    idx = {}
    faltan = []
    for campo, nombres in COLUMNAS.items():
        for nombre in nombres:
            if nombre in cab:
                idx[campo] = cab.index(nombre)
                break
        else:
            if campo not in OPCIONALES:
                faltan.append("%s (%s)" % (campo, " / ".join(nombres)))
    if faltan:
        raise SystemExit("Faltan columnas en la hoja %s:\n  - %s" % (ws.title, "\n  - ".join(faltan)))
    return i, idx


def limpio(v):
    """None, '', '-' -> None; resto como texto sin espacios sobrantes."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s not in ("", "-") else None


def num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0
    return int(f) if f == int(f) else round(f, 2)


def pct(v):
    """El Excel guarda el avance como fracción (0.8 = 80%)."""
    try:
        return round(float(v) * 100, 1)
    except (TypeError, ValueError):
        return 0


def fecha(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%d/%m/%Y")
    return limpio(v)


def main():
    wbc = openpyxl.load_workbook(EXCEL_CONSOLIDADO, read_only=True, data_only=True)
    ws = wbc[HOJA]
    fila_cab, col = ubicar_columnas(ws)
    print("Consolidado: %s | cabecera en fila %d | columnas ubicadas (base 1): %s"
          % (EXCEL_CONSOLIDADO, fila_cab, {k: v + 1 for k, v in col.items()}))
    if EXCEL_INDICADORES == EXCEL_CONSOLIDADO:
        wb = wbc
    else:
        wb = openpyxl.load_workbook(EXCEL_INDICADORES, read_only=True, data_only=True)
    print("Indicadores:", EXCEL_INDICADORES)

    proyectos = []
    vistos = {}
    duplicados = []
    preset_derivados = 0

    for fila in ws.iter_rows(min_row=fila_cab + 1, values_only=True):
        if fila[0] is None:
            continue
        # r: valores de la fila accesibles por nombre de campo
        r = {k: (fila[i] if i < len(fila) else None) for k, i in col.items()}
        cui = str(r["cui"]).strip()
        if cui.endswith(".0"):
            cui = cui[:-2]
        if cui in vistos:
            duplicados.append(cui)
            continue
        vistos[cui] = True

        ssp = limpio(r["ssp"])
        if ssp == "---":
            ssp = None
        # ESTADO SSP = PRESET: el Excel de 31.07.2026 lo traía escrito; el de
        # 04.09.2026 ya no (esos proyectos vienen con SSP vacío y BD PRESET = SI).
        # Regla reconstruida y verificada contra el archivo anterior
        # (21,525 / 21,528 coincidencias): BD PRESET = SI y SSP vacío -> PRESET.
        if ssp is None and "bdPreset" in r and (limpio(r["bdPreset"]) or "").upper() == "SI":
            ssp = "PRESET"
            preset_derivados += 1

        p = {
            "cui": cui,
            "nombre": limpio(r["nombre"]) or "",
            "programa": limpio(r["programa"]) or "",
            "uei": limpio(r["uei"]) or "",
            "dep": limpio(r["dep"]) or "",
            "prov": limpio(r["prov"]) or "",
            "dist": limpio(r["dist"]) or "",
            "modalidad": limpio(r["modalidad"]),
            "tipo": normal_tipo(r["tipo"]),
            "pobl": num(r["pobl"]),
            "cxAgua": num(r["cxAgua"]),
            "cxAlc": num(r["cxAlc"]),
            "ssi": limpio(r["ssi"]),
            "monto": num(r["monto"]),
            "devAc": num(r["devAc"]),
            "pim": num(r["pim"]),
            "dev": num(r["dev"]),
            "fisico": pct(r["fisico"]),
            "estadoET": estado_et(r),
            "procSel": normal_proc(r["procSel"]),
            "ssp": ssp,
            "subSsp": limpio(r["subSsp"]),
            "fTerm": fecha(r["fTerm"]),
            "etapa": limpio(r["etapa"]),   # ETAPA DE EVALUACIÓN siempre (incluye FINANCIADO…); la cartera PRESET sigue con su regla
            "estadoEval": limpio(r["estadoEval"]),   # ESTADO (de la evaluación), pareja de etapa
        }

        carteras = {}

        etapa = limpio(r["etapa"])
        # Desde 04.09.2026 la etapa financiada viene como "FINANCIADO - APTO" /
        # "FINANCIADO - NO APTO" (antes solo "FINANCIADO"): se excluye todo lo que empiece así.
        if etapa is not None and not etapa.upper().startswith("FINANCIADO"):
            carteras["preset"] = {"etapa": etapa, "estado": limpio(r["estadoEval"])}

        if (p["modalidad"] or "").upper() == "DIRECTA":
            carteras["obras"] = {"avance": p["fisico"]}

        transferido = num(r["transferido"])
        if transferido > 0:
            carteras["transferencias"] = {
                "transferido": transferido,
                "ejecutado": num(r["ejecutado"]),
            }

        if (p["ssp"] or "").upper() in ("PARALIZADA", "PARALIZADO"):
            carteras["paralizadas"] = {
                "avance": p["fisico"],
                "hito": limpio(r["hito"]),
                "fecha": fecha(r["fReact"]),
            }

        if carteras:
            p["carteras"] = carteras
        proyectos.append(p)

    if duplicados:
        print("ADVERTENCIA: %d CUI duplicados omitidos: %s" % (len(duplicados), duplicados[:10]))
    if preset_derivados:
        print("ESTADO SSP = PRESET reconstruido (BD PRESET = SI y SSP vacío): %d proyectos" % preset_derivados)

    # Indicadores de títulos de propiedad (hoja INDICADORES TITULOS)
    indicadores = {}
    if "INDICADORES TITULOS" in wb.sheetnames:
        titulos = {}
        for r in wb["INDICADORES TITULOS"].iter_rows(values_only=True):
            etiqueta = str(r[1] or "").strip().upper()
            if etiqueta.startswith("META"):
                titulos["meta"] = num(r[2])
            elif etiqueta.startswith("TITULOS ENTREGADOS") or etiqueta.startswith("TÍTULOS ENTREGADOS"):
                titulos["entregados"] = num(r[2])
        if titulos:
            indicadores["titulos"] = titulos

    def num_celda(v):
        """Número entero desde celda que puede venir como texto con espacio
        duro (\\xa0) o comas: '\\xa09028' -> 9028."""
        if v is None:
            return None
        if isinstance(v, str):
            v = v.replace("\xa0", "").replace(",", "").strip()
            if not v:
                return None
        try:
            return num(v)
        except (TypeError, ValueError):
            return None

    def pct_celda(v):
        """Normaliza el % de las hojas de títulos: '52,1%' -> 52.1;
        fracción (0.56) -> 56; número >1 (22.9) ya es porcentaje."""
        if v is None:
            return None
        if isinstance(v, str):
            s = v.strip().replace("%", "").replace(",", ".")
            try:
                v = float(s)
            except ValueError:
                return None
            return round(v, 1)
        v = float(v)
        return round(v * 100, 1) if v <= 1 else round(v, 1)

    # Macro regiones (hoja MACROREGION TITULOS): %, ejecución y meta
    if "MACROREGION TITULOS" in wb.sheetnames and "titulos" in indicadores:
        macro = []
        for r in wb["MACROREGION TITULOS"].iter_rows(values_only=True):
            nombre, val = limpio(r[1]), pct_celda(r[2])
            if nombre and val is not None:
                fila = {"nombre": nombre, "pct": val}
                ejec, meta = num_celda(r[3]), num_celda(r[4])
                if ejec is not None:
                    fila["ejec"] = ejec
                if meta is not None:
                    fila["meta"] = meta
                macro.append(fila)
        indicadores["titulos"]["macro"] = macro

    # Bonos (hojas INDICADORES BONOS, MODALIDAS BONOS, POR REGION BONOS)
    if "INDICADORES BONOS" in wb.sheetnames:
        bonos = {}
        for r in wb["INDICADORES BONOS"].iter_rows(values_only=True):
            etiqueta = str(r[1] or "").strip().upper()
            if etiqueta.startswith("META"):
                bonos["meta"] = num(r[2])
            elif etiqueta.startswith("DESEMBOLSADOS"):
                bonos["desembolsados"] = num(r[2])
        if bonos:
            indicadores["bonos"] = bonos
    def filas_con_detalle(hoja):
        """Filas nombre + % + desembolso/meta (hojas de bonos)."""
        filas = []
        for r in wb[hoja].iter_rows(values_only=True):
            nombre, val = limpio(r[1]), pct_celda(r[2])
            if not nombre or val is None:
                continue
            fila = {"nombre": nombre, "pct": val}
            des, meta = num_celda(r[3]), num_celda(r[4])
            if des is not None:
                fila["desembolso"] = des
            if meta is not None:
                fila["meta"] = meta
            filas.append(fila)
        return filas

    if "MODALIDAS BONOS" in wb.sheetnames and "bonos" in indicadores:
        indicadores["bonos"]["modalidades"] = filas_con_detalle("MODALIDAS BONOS")
    if "POR REGION BONOS" in wb.sheetnames and "bonos" in indicadores:
        indicadores["bonos"]["regiones"] = filas_con_detalle("POR REGION BONOS")

    # Indicador principal de Inicio (hoja INDICADOR PRINCIPAL): fila SECTOR → %
    # filas: SECTOR (%), PIM, DEVENGADO
    if "INDICADOR PRINCIPAL" in wb.sheetnames:
        principal = {}
        for r in wb["INDICADOR PRINCIPAL"].iter_rows(values_only=True):
            nombre = (limpio(r[1]) or "").upper()
            if nombre == "PIM":
                principal["pim"] = num(r[2])
            elif nombre == "DEVENGADO":
                principal["devengado"] = num(r[2])
            elif nombre and "pct" not in principal and pct_celda(r[2]) is not None:
                principal["nombre"] = limpio(r[1])
                principal["pct"] = pct_celda(r[2])
        if "pct" in principal:
            indicadores["principal"] = principal

    # Ranking del sector frente al Gobierno Nacional (hoja RANKING SECTOR):
    # columnas nombre ("NN: SECTOR"), PIM, % avance. Orden lo hace la UI.
    # Mismo formato en RANKING INVERSIONES y RANKING ACTIVIDADES.
    def leer_ranking(hoja):
        rk = []
        for r in wb[hoja].iter_rows(values_only=True):
            nombre, pim, val = limpio(r[0]), num(r[1]), pct_celda(r[2])
            if not nombre or pim is None or val is None:
                continue
            cod = None
            if ":" in nombre:
                cod, nombre = [x.strip() for x in nombre.split(":", 1)]
            rk.append({"cod": cod, "nombre": nombre, "pim": pim, "pct": val})
        return rk
    for hoja, clave in (("RANKING SECTOR", "rankingSector"),
                        ("RANKING INVERSIONES", "rankingInversiones"),
                        ("RANKING ACTIVIDADES", "rankingActividades")):
        if hoja in wb.sheetnames:
            rk = leer_ranking(hoja)
            if rk:
                indicadores[clave] = rk

    # Inversiones (hoja INVERSIONES): cartera, n.º de proyectos, inversión (millones S/)
    if "INVERSIONES" in wb.sheetnames:
        inv = []
        for r in wb["INVERSIONES"].iter_rows(values_only=True):
            nombre, n, monto = limpio(r[1]), num(r[2]), num(r[3])
            if nombre and n is not None:
                inv.append({"nombre": nombre, "proyectos": int(n), "inversion": monto})
        if inv:
            indicadores["inversiones"] = inv

    # Pliegos (hoja PLIEGOS): tarjetas de Inicio — entidad, PIM, devengado, % ejecución
    if "PLIEGOS" in wb.sheetnames:
        pliegos = []
        for r in wb["PLIEGOS"].iter_rows(values_only=True):
            nombre, pim, dev = limpio(r[2]), num(r[3]), num(r[4])
            val = pct_celda(r[5])
            if val is None and pim:          # la celda es fórmula: calcular
                val = round((dev or 0) / pim * 100, 1)
            if nombre and val is not None and nombre.upper() != "ENTIDAD":
                pliegos.append({"nombre": nombre, "pim": pim, "devengado": dev, "pct": val})
        if pliegos:
            indicadores["pliegos"] = pliegos

    # Detalle de Pliego MVCS (hoja MVCS): fila PLIEGO = resumen;
    # las demás filas = un card por fila en el detalle
    if "MVCS" in wb.sheetnames:
        filas_pliego = []
        for r in wb["MVCS"].iter_rows(values_only=True):
            nombre = limpio(r[1])
            todo = pct_celda(r[2])
            if not nombre or todo is None:
                continue
            filas_pliego.append({
                "nombre": nombre,
                "todo": todo,
                "actividades": pct_celda(r[3]),
                "proyectos": pct_celda(r[4]),
            })
        if filas_pliego:
            indicadores["mvcs"] = filas_pliego

    # Wasiymi (hojas INDICADORES WAYSIMI y POR REGION WAYSIMI — así escritas en el Excel)
    if "INDICADORES WAYSIMI" in wb.sheetnames:
        wasiymi = {}
        for r in wb["INDICADORES WAYSIMI"].iter_rows(values_only=True):
            etiqueta = str(r[1] or "").strip().upper()
            if etiqueta.startswith("META"):
                wasiymi["meta"] = num(r[2])
            elif etiqueta.startswith("EJECUCI"):
                wasiymi["ejecutadas"] = num(r[2])
        if wasiymi:
            indicadores["wasiymi"] = wasiymi
    if "POR REGION WAYSIMI" in wb.sheetnames and "wasiymi" in indicadores:
        reg_was = []
        for r in wb["POR REGION WAYSIMI"].iter_rows(values_only=True):
            nombre, val = limpio(r[1]), pct_celda(r[2])
            if nombre and val is not None:
                reg_was.append({"nombre": nombre, "pct": val})
        indicadores["wasiymi"]["regiones"] = reg_was

    # Regiones agrupadas por macro región (hoja POR REGION TITULOS)
    if "POR REGION TITULOS" in wb.sheetnames and "titulos" in indicadores:
        regiones, macro_actual = [], None
        for r in wb["POR REGION TITULOS"].iter_rows(values_only=True):
            nombre, val = limpio(r[1]), pct_celda(r[2])
            if not nombre:
                continue
            # cabecera de grupo (NORTE, CENTRO…): trae ejecución/meta en las
            # columnas 3-4 (los mismos totales de la hoja MACROREGION) o no trae %
            if num_celda(r[3]) is not None or val is None:
                macro_actual = nombre
            else:
                regiones.append({"macro": macro_actual, "nombre": nombre, "pct": val})
        indicadores["titulos"]["regiones"] = regiones

    with open(SALIDA, "w", encoding="utf-8") as f:
        f.write("window.PROYECTOS=")
        json.dump(proyectos, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";")
        f.write("window.INDICADORES=")
        json.dump(indicadores, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";")
    print("Indicadores:", indicadores)

    n = {}
    prog = {}
    for p in proyectos:
        for k in p.get("carteras", {}):
            n[k] = n.get(k, 0) + 1
        prog[p["programa"]] = prog.get(p["programa"], 0) + 1
    print("Proyectos: %d | carteras: %s" % (len(proyectos), n))
    print("Programas:", dict(sorted(prog.items(), key=lambda kv: -kv[1])))


if __name__ == "__main__":
    main()
