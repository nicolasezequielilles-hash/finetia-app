"""
FINETIA — armado del "paquete de datos" del flujo de 13 semanas.

Lee los archivos que el equipo sube a Drive, reconoce de qué es cada uno
por su contenido (no por la carpeta ni por el nombre) y arma un único
diccionario con: cuentas a cobrar, echeqs en cartera, deuda a proveedores,
cuotas de planes de pago de ARCA y echeqs usados (para la tendencia).

Reglas de negocio que se aplican acá:
  * Deuda a proveedores: solo la tabla dinámica "Resumen HC-PM - Nico"
    (empresas HC y PM). Se ignora la columna DAB y lo intercompany.
  * Cuentas a cobrar: A-Evangelista (AESA) cobra a 30 días desde la emisión.
  * Echeqs de clientes: solo en estado PENDIENTE y con vencimiento desde 2025.
  * Planes de ARCA: se leen del Excel del estudio (HC) y de los PDF de
    "Mis Facilidades" (PM).
  * Un archivo opcional "ajustes_flujo.json" permite excluir comprobantes ya
    cobrados y marcar cuotas de planes ya pagadas.
"""
import io
import json
import re
import warnings
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional

warnings.filterwarnings("ignore")

AESA_CUIT = "30685218190"
EMP_SIGLA = {"PA": "PM", "PM": "PM", "HC": "HC"}


# ---------------------------------------------------------------- utilidades
def _iso(v) -> Optional[str]:
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return None


def _num(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


def _libro(datos: bytes):
    from openpyxl import load_workbook
    return load_workbook(io.BytesIO(datos), data_only=True, read_only=False)


# ------------------------------------------------------------ clasificación
def clasificar(nombre: str, datos: bytes) -> str:
    """Devuelve: saldos | cuentas_cobrar | pagos_realizar | planes_xlsx | plan_pdf | cheques_propios | ajustes | otro."""
    n = (nombre or "").lower()
    if n.endswith(".json"):
        return "ajustes" if "ajuste" in n else "otro"
    if n.endswith(".pdf"):
        try:
            return "plan_pdf" if _texto_pdf(datos).count("Número de Plan") else "otro"
        except Exception:
            return "otro"
    if not n.endswith((".xlsx", ".xlsm")):
        return "otro"
    try:
        wb = _libro(datos)
    except Exception:
        return "otro"
    hojas = [h.strip().lower() for h in wb.sheetnames]
    if "saldos" in hojas and any(h.startswith("echeq - clientes") for h in hojas):
        return "saldos"
    if "pagos a realizar" in hojas and any(h.startswith("resumen hc-pm") for h in hojas):
        return "pagos_realizar"
    if "cuotas por plan" in hojas:
        return "planes_xlsx"
    ws = wb.worksheets[0]
    if _es_cheques_propios(ws):
        return "cheques_propios"
    for r in range(1, 30):
        if str(ws.cell(r, 2).value or "").strip().lower() == "empresa" and "cbte" in str(ws.cell(r, 6).value or "").lower():
            return "cuentas_cobrar"
    return "otro"


# ---------------------------------------------------------- cheques propios
def _es_cheques_propios(ws) -> bool:
    for r in range(1, 6):
        for c in range(1, 4):
            if "cheques propios" in str(ws.cell(r, c).value or "").lower():
                return True
    return False


def _fecha_mdy(v) -> Optional[str]:
    """Fecha de una celda: datetime, o texto m/d/aaaa (formato del export de Tango)."""
    if isinstance(v, (datetime, date)):
        return _iso(v)
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$", str(v or ""))
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    mes, dia = (a, b) if a <= 12 else (b, a)
    try:
        return date(y, mes, dia).isoformat()
    except ValueError:
        return None


def _num_en(v) -> float:
    """Número en formato 1,234,567.89 (o ya numérico)."""
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def leer_cheques_propios(datos: bytes, emp: str) -> List[dict]:
    """Cheques propios (Tango): salidas ya comprometidas, con su fecha de cobro. Excluye anulados/rechazados."""
    wb = _libro(datos)
    ws = next(w for w in wb.worksheets if _es_cheques_propios(w))
    filas = list(ws.iter_rows(values_only=True))
    cab = next((i for i, r in enumerate(filas) if any(str(x or "").strip().lower() == "fecha del cheque" for x in r)), None)
    if cab is None:
        raise ValueError("no encuentro la columna Fecha del cheque")
    nom = [str(x or "").strip().lower() for x in filas[cab]]
    col = lambda t: next((i for i, n in enumerate(nom) if n.startswith(t)), None)
    c_fecha, c_prov, c_est, c_imp, c_nro = col("fecha del cheque"), col("razón social"), col("estado"), col("importe (reexp. cte"), col("nro. de cheque")
    if c_imp is None:
        c_imp = col("importe mon")
    out = []
    for r in filas[cab + 1:]:
        f = _fecha_mdy(r[c_fecha]) if c_fecha is not None and c_fecha < len(r) else None
        imp = _num_en(r[c_imp]) if c_imp is not None and c_imp < len(r) else 0.0
        est = str(r[c_est] or "").strip() if c_est is not None and c_est < len(r) else ""
        if not f or imp <= 0 or re.search(r"anulad|rechazad|debitad|cobrad", est, re.I):
            continue
        out.append({"emp": emp, "proveedor": str(r[c_prov] or "").strip() if c_prov is not None else "", "comprobante": "Cheque " + str(r[c_nro] or "").strip() if c_nro is not None else "Cheque",
                    "venc": f, "importe": round(imp, 2), "forma": "cheque propio"})
    return sorted(out, key=lambda x: x["venc"])


def _emp_de_nombre(nombre: str) -> Optional[str]:
    m = re.search(r"(?:^|[^A-Za-z])(HC|PM)(?:[^A-Za-z]|$)", nombre or "")
    return m.group(1) if m else None


# ---------------------------------------------------------- cuentas a cobrar
def leer_cuentas_cobrar(datos: bytes) -> dict:
    ws = _libro(datos).worksheets[0]
    cab = next(r for r in range(1, 40) if str(ws.cell(r, 2).value or "").strip().lower() == "empresa")
    corte = None
    for r in range(1, cab):
        for c in range(1, 14):
            if isinstance(ws.cell(r, c).value, datetime):
                corte = _iso(ws.cell(r, c).value)
    out = []
    for r in range(cab + 1, ws.max_row + 1):
        emp, fe, vto, cli, cb, nro, neto, imp = [ws.cell(r, c).value for c in range(2, 10)]
        if emp not in EMP_SIGLA or not isinstance(imp, (int, float)):
            continue
        out.append({"emp": EMP_SIGLA[emp], "cliente": str(cli).strip(), "comprobante": str(nro).strip(), "emision": _iso(fe), "venc": _iso(vto), "importe": round(float(imp), 2)})
    return {"corte": corte, "filas": out}


def aplicar_regla_cobro(filas: List[dict]) -> List[dict]:
    for c in filas:
        c["venc_esperado"] = c["venc"]
        if "EVANGELISTA" in c["cliente"].upper() and c.get("emision"):
            c["venc_esperado"] = (date.fromisoformat(c["emision"]) + timedelta(days=30)).isoformat()
            c["regla"] = "AESA: 30 días desde la emisión"
    return filas


# ------------------------------------------------------- deuda a proveedores
_MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12}


def leer_pagos_realizar(datos: bytes) -> List[dict]:
    wb = _libro(datos)
    nombres = {h.strip().lower(): h for h in wb.sheetnames}
    det_rows = list(wb[nombres["pagos a realizar"]].iter_rows(values_only=True))
    det = []
    for r in det_rows[1:]:
        if not r[0]:
            continue
        r = list(r) + [None] * 15
        vto = r[3]
        det.append({"grupo": str(r[0]).strip(), "vto": vto if isinstance(vto, datetime) else None, "nro": str(r[5] or "").strip(),
                    "pend": float(r[6]) if isinstance(r[6], (int, float)) else 0.0, "prov": str(r[10] or "").strip(), "forma": str(r[13] or "").strip()})
    for d in det:
        if d["forma"] in ("0", "#REF!", ""):
            d["forma"] = "(sin definir)"
    pvt = next(wb[n] for k, n in nombres.items() if k.startswith("resumen hc-pm") and "nico" in k)
    prow = list(pvt.iter_rows(values_only=True))
    anios, meses = prow[3], prow[4]
    celdas, g = [], None
    for r in prow[5:]:
        a = r[0]
        if a in ("HC", "PM", "Infesur"):
            g = a
            continue
        if a is None or a == "Total general" or g not in ("HC", "PM"):
            continue
        for j in range(1, len(r)):
            v = r[j]
            if isinstance(v, (int, float)) and v:
                try:
                    anio = int(anios[j])
                except Exception:
                    continue
                celdas.append({"grupo": g, "label": str(a).strip(), "anio": anio, "mes": _MESES.get(str(meses[j]).strip().lower(), 0), "imp": float(v)})
    salida = []
    for c in celdas:
        base = [d for d in det if d["grupo"] == c["grupo"]]
        if re.fullmatch(r"[A-Z]\d{10,}", c["label"]):
            mt = [d for d in base if d["nro"] == c["label"]]
            if mt and mt[0]["vto"]:
                venc = mt[0]["vto"].date()
            else:
                venc = date(c["anio"], c["mes"] or 12, 15)
            salida.append({"emp": c["grupo"], "proveedor": mt[0]["prov"] if mt else "(comprobante)", "comprobante": c["label"], "venc": venc.isoformat(), "importe": round(c["imp"], 2), "forma": mt[0]["forma"] if mt else ""})
        else:
            mt = [d for d in base if d["prov"].upper() == c["label"].upper() and d["vto"] and d["vto"].year == c["anio"] and (not c["mes"] or d["vto"].month == c["mes"])]
            tot = sum(d["pend"] for d in mt)
            if mt and tot > 0:
                for d in mt:
                    salida.append({"emp": c["grupo"], "proveedor": c["label"], "comprobante": d["nro"], "venc": d["vto"].date().isoformat(), "importe": round(c["imp"] * d["pend"] / tot, 2), "forma": d["forma"]})
            else:
                salida.append({"emp": c["grupo"], "proveedor": c["label"], "comprobante": "(sin detalle)", "venc": date(c["anio"], c["mes"] or 12, 15).isoformat(), "importe": round(c["imp"], 2), "forma": ""})
    return sorted(salida, key=lambda x: (x["venc"], x["emp"]))


# --------------------------------------------------------------- planes ARCA
def leer_planes_xlsx(datos: bytes, emp: str = "HC") -> List[dict]:
    wb = _libro(datos)
    hoja = next(wb[n] for n in wb.sheetnames if n.strip().lower() == "cuotas por plan")
    plan, rg, out = None, "", []
    for r in hoja.iter_rows(values_only=True):
        c = str(r[0] or "")
        m = re.match(r"Plan (\S+) — (.*)", c)
        if m:
            plan, rg = m.group(1), m.group(2).split(" - ")[0]
            continue
        if plan and re.fullmatch(r"PC|\d+", c) and r[1]:
            d, mo, y = map(int, str(r[1]).split("/"))
            out.append({"emp": emp, "plan": plan, "cuota": c, "vto": date(y, mo, d).isoformat(), "importe": round(float(r[4]), 2), "estado": str(r[8] or "A vencer")})
    return out


def _texto_pdf(datos: bytes) -> str:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(datos)) as pdf:
        return "\n".join((p.extract_text(layout=False) or "") for p in pdf.pages)


_NUM = r"[\d\.]+,\d\d"
_FECHA = r"(\d\d)/(\d\d)/(\d{4})"


def leer_plan_pdf(datos: bytes, emp: str = "PM") -> dict:
    """Lee un PDF de Mis Facilidades. Devuelve {plan, cuotas: [{cuota, vto, importe}], impagas: [n]}.
    El PDF "completo" trae todas las cuotas; el de "impagas/actual" marca el estado de cada una."""
    texto = _texto_pdf(datos)
    m = re.search(r"Número de Plan:\s*(\S+)", texto)
    if not m:
        raise ValueError("no encuentro el número de plan")
    plan = m.group(1)
    cuotas, impagas, actual = [], set(), None
    for ln in texto.splitlines():
        ln = ln.strip()
        # fila de cuota con 1° vencimiento (sin interés resarcitorio): "<n> <capital> <int.fin> - <total> <fecha>"
        m1 = re.match(r"^(\d+)\s+([\d\.,]+)\s+(%s)\s+-\s+(%s)\s+%s" % (_NUM, _NUM, _FECHA), ln)
        m2 = re.match(r"^(\d+)\s+(%s)\s+(%s)\s+(%s)\s+(%s)\s+%s(.*)$" % (_NUM, _NUM, _NUM, _NUM, _FECHA), ln)
        if m1:
            cuotas.append({"emp": emp, "plan": plan, "cuota": m1.group(1), "vto": date(int(m1.group(7)), int(m1.group(6)), int(m1.group(5))).isoformat(), "importe": round(_num(m1.group(4)), 2)})
        elif m2:
            actual = m2.group(1)
            cuotas.append({"emp": emp, "plan": plan, "cuota": actual, "vto": date(int(m2.group(8)), int(m2.group(7)), int(m2.group(6))).isoformat(), "importe": round(_num(m2.group(5)), 2)})
            if "Impaga" in (m2.group(9) or ""):
                impagas.add(actual)
        elif actual and "Cuota Impaga" in ln:
            impagas.add(actual)
    return {"plan": plan, "cuotas": cuotas, "impagas": sorted(impagas, key=int) if impagas else [], "tiene_estado": "Estado de Cuota" in texto}


def combinar_planes_pdf(lecturas: List[dict], hoy: str, emp: str = "PM") -> List[dict]:
    por_plan: Dict[str, dict] = {}
    for l in lecturas:
        p = por_plan.setdefault(l["plan"], {"cuotas": {}, "impagas": None})
        for c in l["cuotas"]:
            p["cuotas"][c["cuota"]] = c
        if l["tiene_estado"]:
            p["impagas"] = set(l["impagas"])
    out = []
    for plan, p in por_plan.items():
        for n, c in sorted(p["cuotas"].items(), key=lambda kv: int(kv[0])):
            if p["impagas"] is None:
                est = "A vencer"
            elif n in p["impagas"]:
                est = "Impaga"
            else:
                est = "Pagada" if c["vto"] < hoy else "A vencer"
            out.append(dict(c, estado=est, emp=emp))
    return out


# ------------------------------------------------------------------- saldos
def leer_saldos_echeqs(datos: bytes, desde_anio: int = 2025) -> dict:
    """Echeqs de clientes pendientes (en cartera) y echeqs ya usados (entregados/endosados)."""
    wb = _libro(datos)
    ws = next(wb[n] for n in wb.sheetnames if n.strip().lower().startswith("echeq - clientes"))
    cab = next(r for r in range(1, 10) if str(ws.cell(r, 1).value or "").strip().lower() == "id de echeq")
    pend, usados = [], []
    for r in range(cab + 1, ws.max_row + 1):
        grupo, emisor, cuit, imp, vto, accion, fuso = ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 7).value, ws.cell(r, 12).value, ws.cell(r, 14).value
        if grupo not in ("HC", "PM") or not isinstance(imp, (int, float)):
            continue
        accion = str(accion or "").strip()
        if accion == "PENDIENTE" and isinstance(vto, datetime) and vto.year >= desde_anio:
            pend.append({"emp": grupo, "emisor": str(emisor or ""), "cuit": str(cuit or "").replace(".0", ""), "importe": float(imp), "venc": _iso(vto)})
        elif accion in ("Entregado a Proveedores", "Endosado") and isinstance(fuso, datetime) and fuso.year >= 2026:
            usados.append({"emp": grupo, "cuit": str(cuit or "").replace(".0", ""), "importe": float(imp), "fecha_uso": _iso(fuso), "accion": accion})
    fecha = None
    try:
        hs = wb[next(n for n in wb.sheetnames if n.strip().lower() == "saldos")]
        fechas = []
        for r in range(1, 8):
            for c in range(1, hs.max_column + 1):
                v = hs.cell(r, c).value
                if isinstance(v, datetime) and v.year >= 2025:
                    fechas.append(v)
        fecha = _iso(max(fechas)) if fechas else None
    except Exception:
        pass
    return {"echeqs": sorted(pend, key=lambda x: x["venc"]), "echeqs_usados": usados, "fecha_saldos": fecha}


# ----------------------------------------------------------------- el paquete
def armar_paquete(archivos: List[dict], hoy: Optional[str] = None) -> dict:
    """archivos: [{nombre, datos(bytes), modificado}] — devuelve el paquete y un informe de lo leído."""
    hoy = hoy or date.today().isoformat()
    informe, cuentas, pagos, planes_hc, pdfs, saldos, ajustes = [], None, None, None, [], None, {}

    def ultimo(lista):
        return sorted(lista, key=lambda a: a.get("modificado", ""))[-1] if lista else None

    tipos = {}
    for a in archivos:
        try:
            t = clasificar(a["nombre"], a["datos"])
        except Exception as e:
            t = "otro"
            informe.append({"archivo": a["nombre"], "tipo": "error", "detalle": str(e)})
        tipos.setdefault(t, []).append(a)
        if t == "otro":
            informe.append({"archivo": a["nombre"], "tipo": "otro", "detalle": "no es un archivo que el flujo use"})
    a = ultimo(tipos.get("cuentas_cobrar", []))
    if a:
        cuentas = leer_cuentas_cobrar(a["datos"]); informe.append({"archivo": a["nombre"], "tipo": "cuentas a cobrar", "detalle": "%d comprobantes al %s" % (len(cuentas["filas"]), cuentas["corte"])})
    a = ultimo(tipos.get("pagos_realizar", []))
    if a:
        pagos = leer_pagos_realizar(a["datos"]); informe.append({"archivo": a["nombre"], "tipo": "deuda a proveedores", "detalle": "%d líneas, $ %.0f M" % (len(pagos), sum(p["importe"] for p in pagos) / 1e6)})
    a = ultimo(tipos.get("planes_xlsx", []))
    if a:
        planes_hc = leer_planes_xlsx(a["datos"], "HC"); informe.append({"archivo": a["nombre"], "tipo": "planes ARCA HC", "detalle": "%d cuotas" % len(planes_hc)})
    # PDF: el más reciente de cada (plan, con/sin estado)
    vistos = {}
    for a in sorted(tipos.get("plan_pdf", []), key=lambda x: x.get("modificado", "")):
        try:
            l = leer_plan_pdf(a["datos"]); vistos[(l["plan"], l["tiene_estado"])] = l
            informe.append({"archivo": a["nombre"], "tipo": "plan ARCA PM", "detalle": "plan %s, %d cuotas%s" % (l["plan"], len(l["cuotas"]), ", %d impagas" % len(l["impagas"]) if l["tiene_estado"] else "")})
        except Exception as e:
            informe.append({"archivo": a["nombre"], "tipo": "error", "detalle": str(e)})
    planes_pm = combinar_planes_pdf(list(vistos.values()), hoy)
    # cheques propios: el más reciente de cada empresa (HC / PM)
    cheques = []
    por_emp = {}
    for a in sorted(tipos.get("cheques_propios", []), key=lambda x: x.get("modificado", "")):
        e = _emp_de_nombre(a["nombre"])
        if not e:
            informe.append({"archivo": a["nombre"], "tipo": "error", "detalle": "no sé de qué empresa es (poné HC o PM en el nombre del archivo)"}); continue
        por_emp[e] = a
    for e, a in por_emp.items():
        try:
            ch = leer_cheques_propios(a["datos"], e); cheques += ch
            informe.append({"archivo": a["nombre"], "tipo": "cheques propios " + e, "detalle": "%d cheques, $ %.1f M" % (len(ch), sum(x["importe"] for x in ch) / 1e6)})
        except Exception as ex:
            informe.append({"archivo": a["nombre"], "tipo": "error", "detalle": "cheques propios ilegibles: %s" % ex})
    a = ultimo(tipos.get("saldos", []))
    if a:
        saldos = leer_saldos_echeqs(a["datos"]); informe.append({"archivo": a["nombre"], "tipo": "tesorería", "detalle": "%d echeqs en cartera, saldos al %s" % (len(saldos["echeqs"]), saldos["fecha_saldos"])})
    a = ultimo(tipos.get("ajustes", []))
    if a:
        try:
            ajustes = json.loads(a["datos"].decode("utf-8")); informe.append({"archivo": a["nombre"], "tipo": "ajustes", "detalle": "excluye %d comprobantes, %d cuotas pagadas" % (len(ajustes.get("excluir_comprobantes", [])), len(ajustes.get("cuotas_pagadas", [])))})
        except Exception as e:
            informe.append({"archivo": a["nombre"], "tipo": "error", "detalle": "ajustes ilegibles: %s" % e})
    # ajustes
    excl = set(ajustes.get("excluir_comprobantes", []))
    cob = aplicar_regla_cobro([c for c in (cuentas or {"filas": []})["filas"] if c["comprobante"] not in excl])
    arca = (planes_hc or []) + planes_pm
    for p in ajustes.get("cuotas_pagadas", []):
        for c in arca:
            if c["plan"] == p["plan"] and c["cuota"] == str(p["cuota"]):
                c["estado"] = "Pagada (ajuste)"
    faltan = [n for n, v in (("cuentas a cobrar", cuentas), ("deuda a proveedores", pagos), ("planes de ARCA", arca), ("tesorería (echeqs)", saldos)) if not v]
    return {
        "version": 2, "fecha": (saldos or {}).get("fecha_saldos") or (cuentas or {}).get("corte") or hoy, "generado": datetime.utcnow().isoformat() + "Z",
        "cobranzas": cob, "echeqs": (saldos or {}).get("echeqs", []), "echeqs_usados": (saldos or {}).get("echeqs_usados", []),
        "pagos": (pagos or []) + cheques, "arca": arca, "faltan": faltan, "informe": informe,
        "notas": ["Cobranzas con IVA; AESA a 30 días desde la emisión.", "Pagos: tabla dinámica del sistema (sin columna DAB ni intercompany) más los cheques propios diferidos de Tango."],
    }
