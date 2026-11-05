"""
FINETIA — lectura de la carpeta de Google Drive (solo lectura).

El backend entra a Drive con una cuenta de servicio (un "usuario robot"
con permiso de Lector sobre la carpeta de entrada). Lista las subcarpetas
y archivos, lleva un registro de qué archivos ya se leyeron y entrega los
bytes de cada archivo al dashboard, que los procesa con sus lectores.

Variables de entorno:
    GOOGLE_SERVICE_ACCOUNT_JSON  contenido completo del .json de la cuenta de servicio
                                 (también se acepta codificado en base64)
    DRIVE_FOLDER_ID              id de la carpeta de entrada (01_Entrada)

Nunca escribe, mueve ni borra nada en Drive: el permiso que se pide es drive.readonly.
"""
import base64
import json
import os
import time
from datetime import datetime, timezone
from typing import Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Response
from pydantic import BaseModel

DRIVE_API = "https://www.googleapis.com/drive/v3"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
MIME_CARPETA = "application/vnd.google-apps.folder"
MIME_SHEET = "application/vnd.google-apps.spreadsheet"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_BYTES = 40 * 1024 * 1024
CACHE_SEG = 45


def _leer_credenciales(texto: str) -> dict:
    texto = (texto or "").strip()
    if not texto:
        raise ValueError("falta GOOGLE_SERVICE_ACCOUNT_JSON")
    if not texto.startswith("{"):
        texto = base64.b64decode(texto).decode("utf-8")
    info = json.loads(texto)
    if info.get("type") != "service_account" or "private_key" not in info:
        raise ValueError("el JSON no es una clave de cuenta de servicio")
    return info


class DriveClient:
    """Cliente mínimo de la API de Drive v3 (solo lectura)."""

    def __init__(self, info: dict):
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2 import service_account

        self.email = info.get("client_email", "")
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        self.sesion = AuthorizedSession(creds)

    def _get(self, url: str, **kw):
        r = self.sesion.get(url, timeout=60, **kw)
        if r.status_code >= 400:
            try:
                msg = r.json().get("error", {}).get("message", r.text[:200])
            except Exception:
                msg = r.text[:200]
            raise RuntimeError("Drive respondió %s: %s" % (r.status_code, msg))
        return r

    def hijos(self, carpeta_id: str) -> List[dict]:
        out, token = [], None
        while True:
            p = {"q": "'%s' in parents and trashed = false" % carpeta_id, "pageSize": 200,
                 "fields": "nextPageToken, files(id,name,mimeType,modifiedTime,size,md5Checksum)",
                 "supportsAllDrives": "true", "includeItemsFromAllDrives": "true", "orderBy": "name"}
            if token:
                p["pageToken"] = token
            j = self._get(DRIVE_API + "/files", params=p).json()
            out += j.get("files", [])
            token = j.get("nextPageToken")
            if not token:
                return out

    def nombre(self, archivo_id: str) -> str:
        return self._get(DRIVE_API + "/files/" + archivo_id, params={"fields": "name", "supportsAllDrives": "true"}).json().get("name", "")

    def descargar(self, archivo: dict) -> bytes:
        if archivo["mimeType"] == MIME_SHEET:
            r = self._get(DRIVE_API + "/files/%s/export" % archivo["id"], params={"mimeType": MIME_XLSX})
        elif archivo["mimeType"].startswith("application/vnd.google-apps."):
            r = self._get(DRIVE_API + "/files/%s/export" % archivo["id"], params={"mimeType": "text/plain"})
        else:
            r = self._get(DRIVE_API + "/files/" + archivo["id"], params={"alt": "media", "supportsAllDrives": "true"})
        return r.content


def armar_arbol(cliente, raiz_id: str, max_prof: int = 3) -> dict:
    """Devuelve {carpetas: [{id, nombre, ruta, archivos: [...]}], indice: {id: archivo}}."""
    carpetas, indice = [], {}

    def visitar(cid: str, ruta: str, prof: int):
        items = cliente.hijos(cid)
        archivos = []
        for it in sorted(items, key=lambda x: x["name"].lower()):
            if it["mimeType"] == MIME_CARPETA:
                continue
            if it["name"].lower().startswith("leeme"):
                continue   # notas de instrucciones para el equipo: no son datos
            a = {"id": it["id"], "nombre": it["name"], "mime": it["mimeType"], "modificado": it.get("modifiedTime", ""),
                 "tamanio": int(it.get("size") or 0), "carpeta": ruta}
            archivos.append(a)
            indice[a["id"]] = a
        carpetas.append({"id": cid, "nombre": ruta.split("/")[-1], "ruta": ruta, "archivos": archivos})
        if prof < max_prof:
            for it in sorted(items, key=lambda x: x["name"].lower()):
                if it["mimeType"] == MIME_CARPETA:
                    visitar(it["id"], ruta + "/" + it["name"], prof + 1)

    visitar(raiz_id, cliente.nombre(raiz_id) if hasattr(cliente, "nombre") else "raíz", 0)
    return {"carpetas": carpetas, "indice": indice}


class MarcaLeido(BaseModel):
    id: str
    nombre: str = ""
    carpeta: str = ""
    modificado: str = ""
    resumen: str = ""
    ok: bool = True


class MarcasLeidas(BaseModel):
    archivos: List[MarcaLeido]


def crear_router(require_auth, get_db, cliente_factory=None):
    router = APIRouter(prefix="/api/v1/drive", tags=["drive"])
    cache: Dict[str, object] = {"t": 0.0, "arbol": None, "cliente": None}

    def _cliente():
        if cliente_factory:
            return cliente_factory()
        if cache["cliente"] is None:
            info = _leer_credenciales(os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", ""))
            cache["cliente"] = DriveClient(info)
        return cache["cliente"]

    def _raiz() -> str:
        r = os.environ.get("DRIVE_FOLDER_ID", "").strip()
        if not r:
            raise ValueError("falta DRIVE_FOLDER_ID")
        return r

    def _arbol(forzar: bool = False) -> dict:
        if not forzar and cache["arbol"] and time.time() - cache["t"] < CACHE_SEG:
            return cache["arbol"]
        cache["arbol"] = armar_arbol(_cliente(), _raiz())
        cache["t"] = time.time()
        return cache["arbol"]

    def _tabla(conn):
        conn.execute("CREATE TABLE IF NOT EXISTS drive_registro (file_id TEXT PRIMARY KEY, nombre TEXT, carpeta TEXT, "
                     "modificado TEXT, procesado_en TEXT, usuario TEXT, resumen TEXT, ok INTEGER)")

    @router.get("/estado")
    def estado(refrescar: bool = False, _user: str = Depends(require_auth)):
        try:
            _cliente(); _raiz()
        except Exception as e:
            return {"configurado": False, "motivo": str(e)}
        try:
            arbol = _arbol(forzar=refrescar)
        except Exception as e:
            return {"configurado": True, "conectado": False, "motivo": str(e)}
        with get_db() as conn:
            _tabla(conn)
            reg = {r[0]: r for r in conn.execute("SELECT file_id, modificado, procesado_en, resumen, ok FROM drive_registro")}
        carpetas, nuevos = [], 0
        for c in arbol["carpetas"]:
            archivos = []
            for a in c["archivos"]:
                r = reg.get(a["id"])
                est = "nuevo" if r is None else ("actualizado" if r[1] != a["modificado"] else ("error" if not r[4] else "leído"))
                if est in ("nuevo", "actualizado"):
                    nuevos += 1
                archivos.append(dict(a, estado=est, leido_en=r[2] if r else None, resumen=r[3] if r else None))
            carpetas.append({"id": c["id"], "nombre": c["nombre"], "ruta": c["ruta"], "archivos": archivos})
        return {"configurado": True, "conectado": True, "cuenta": getattr(_cliente(), "email", ""), "carpetas": carpetas,
                "nuevos": nuevos, "consultado": datetime.now(timezone.utc).isoformat()}

    @router.get("/archivo/{archivo_id}")
    def archivo(archivo_id: str, _user: str = Depends(require_auth)):
        try:
            arbol = _arbol()
            if archivo_id not in arbol["indice"]:
                arbol = _arbol(forzar=True)
        except Exception as e:
            raise HTTPException(status_code=503, detail="No se pudo leer Drive: %s" % e)
        a = arbol["indice"].get(archivo_id)
        if not a:
            raise HTTPException(status_code=404, detail="Ese archivo no está dentro de la carpeta de entrada.")
        if a["tamanio"] > MAX_BYTES:
            raise HTTPException(status_code=413, detail="El archivo supera los 40 MB.")
        try:
            data = _cliente().descargar(a)
        except Exception as e:
            raise HTTPException(status_code=502, detail="No se pudo descargar: %s" % e)
        nombre = a["nombre"] + (".xlsx" if a["mime"] == MIME_SHEET and not a["nombre"].lower().endswith(".xlsx") else "")
        return Response(content=data, media_type="application/octet-stream",
                        headers={"X-Nombre-Archivo": base64.b64encode(nombre.encode("utf-8")).decode("ascii"), "X-Modificado": a["modificado"]})

    @router.post("/leidos")
    def leidos(req: MarcasLeidas, user: str = Depends(require_auth)):
        ahora = datetime.now(timezone.utc).isoformat()
        with get_db() as conn:
            _tabla(conn)
            for m in req.archivos:
                conn.execute("INSERT OR REPLACE INTO drive_registro (file_id, nombre, carpeta, modificado, procesado_en, usuario, resumen, ok) VALUES (?,?,?,?,?,?,?,?)",
                             (m.id, m.nombre, m.carpeta, m.modificado, ahora, user, m.resumen[:300], 1 if m.ok else 0))
            conn.commit()
        return {"guardados": len(req.archivos)}

    memo: Dict[tuple, dict] = {}   # (id, modificado) -> {"tipo": ..., "datos": bytes|None}

    @router.get("/paquete-flujo")
    def paquete_flujo(refrescar: bool = False, _user: str = Depends(require_auth)):
        """Lee de Drive los archivos que usa el flujo de 13 semanas y arma el paquete de datos."""
        import flujo_paquete as fp
        try:
            arbol = _arbol(forzar=refrescar)
        except Exception as e:
            raise HTTPException(status_code=503, detail="No se pudo leer Drive: %s" % e)
        cliente, archivos, ignorados = _cliente(), [], []
        for a in arbol["indice"].values():
            nombre = a["nombre"] + (".xlsx" if a["mime"] == MIME_SHEET and not a["nombre"].lower().endswith(".xlsx") else "")
            if not nombre.lower().endswith((".xlsx", ".xlsm", ".pdf", ".json")) or a["tamanio"] > MAX_BYTES:
                continue
            if "/bancos" in a["carpeta"].lower():
                continue   # los extractos bancarios los lee la app; no sirven para el flujo
            clave = (a["id"], a["modificado"])
            if clave not in memo:
                try:
                    datos = cliente.descargar(a)
                    tipo = fp.clasificar(nombre, datos)
                except Exception as e:
                    ignorados.append("%s: %s" % (nombre, e)); continue
                memo[clave] = {"tipo": tipo, "datos": datos if tipo != "otro" else None}
            if memo[clave]["tipo"] != "otro":
                archivos.append({"nombre": nombre, "datos": memo[clave]["datos"], "modificado": a["modificado"]})
        for k in [k for k in memo if k not in {(x["id"], x["modificado"]) for x in arbol["indice"].values()}]:
            memo.pop(k, None)
        P = fp.armar_paquete(archivos)
        if ignorados:
            P["informe"] += [{"archivo": x, "tipo": "error", "detalle": "no se pudo descargar"} for x in ignorados]
        return P

    return router
