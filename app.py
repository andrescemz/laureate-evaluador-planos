"""
Evaluador de Planos — MVP Web
Uso: uvicorn app:app --reload --port 8000
     (con ANTHROPIC_API_KEY seteada en el entorno)
"""

import base64, json, os
from datetime import datetime, timezone
from pathlib import Path

import fitz                      # pymupdf
import anthropic
import openpyxl
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel

# ── Config ────────────────────────────────────────────────────
BASE         = Path(__file__).parent
RUBRICA_FILE = BASE / "0 LISTA DE CHEQUEO_202610_Coordinacion.xlsx"
EVAL_FILE    = BASE / "evaluaciones.json"
API_KEY      = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL        = "claude-sonnet-4-6"
MAX_PX       = 1400

app = FastAPI(title="Evaluador Planos TVIII")

# ── Rúbrica ───────────────────────────────────────────────────
def cargar_rubrica(semana: str) -> list[dict]:
    col = {"4": 3, "8": 4, "12": 5, "TF": 6}.get(str(semana), 3)
    wb  = openpyxl.load_workbook(str(RUBRICA_FILE))
    ws  = wb.active
    dim = crit = ""
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]: dim  = row[0].strip()
        if row[1]: crit = row[1].strip()
        ck = row[2]
        if not ck or str(ck).startswith("="): continue
        if row[col] == "x":
            items.append({
                "id":        len(items) + 1,
                "dimension": dim,
                "criterio":  crit,
                "checklist": ck.strip().replace("\n", " "),
            })
    return items

# ── PDF → imagen base64 ───────────────────────────────────────
def pdf_a_b64(pdf_bytes: bytes) -> str:
    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    zoom = min(MAX_PX / page.rect.width, 2.5)
    pix  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    b64  = base64.standard_b64encode(pix.tobytes("jpeg", jpg_quality=85)).decode()
    doc.close()
    return b64

# ── Persistencia ─────────────────────────────────────────────
def _load() -> dict:
    if EVAL_FILE.exists():
        return json.loads(EVAL_FILE.read_text(encoding="utf-8"))
    return {}

def _save(data: dict):
    EVAL_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

class GuardarBody(BaseModel):
    archivo:   str
    alumno:    str
    semana:    str
    criterios: dict           # id_str -> {nivel, cumple, feedback, ...}
    resumen:   dict | None = None

class DesbloquearBody(BaseModel):
    archivo: str
    semana:  str

# ── Endpoints ─────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index():
    return (BASE / "templates" / "index.html").read_text(encoding="utf-8")

@app.get("/rubrica")
async def get_rubrica(semana: str = "4"):
    return cargar_rubrica(semana)

@app.get("/evaluaciones")
async def get_evaluaciones():
    return _load()

@app.post("/guardar")
async def guardar(body: GuardarBody):
    data = _load()
    if body.archivo not in data:
        data[body.archivo] = {"alumno": body.alumno, "semanas": {}}

    semanas = data[body.archivo]["semanas"]
    ya_existe = body.semana in semanas

    # Si ya existe y está bloqueada, rechazar
    if ya_existe and semanas[body.semana].get("locked", False):
        return {"ok": False, "motivo": "bloqueada"}

    version = semanas[body.semana].get("version", 0) + 1 if ya_existe else 1
    semanas[body.semana] = {
        "criterios":  body.criterios,
        "resumen":    body.resumen,
        "timestamp":  datetime.now(timezone.utc).isoformat(),
        "locked":     True,          # se bloquea al guardar
        "version":    version,
    }
    _save(data)
    return {"ok": True, "version": version}

@app.post("/desbloquear")
async def desbloquear(body: DesbloquearBody):
    data = _load()
    try:
        data[body.archivo]["semanas"][body.semana]["locked"] = False
        _save(data)
        return {"ok": True}
    except KeyError:
        return {"ok": False, "motivo": "no encontrado"}

@app.post("/evaluar")
async def evaluar(file: UploadFile = File(...), semana: str = Form("4")):

    def ev(data: dict) -> str:
        return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

    KEEPALIVE = ": keep-alive\n\n"   # comentario SSE, no dispara eventos en el cliente

    async def stream():
        import asyncio
        try:
            yield ev({"tipo": "progreso", "msg": "Leyendo PDF…"})
            pdf_bytes = await file.read()

            yield ev({"tipo": "progreso", "msg": "Renderizando imagen…"})
            img_b64 = pdf_a_b64(pdf_bytes)

            rubrica     = cargar_rubrica(semana)
            rubrica_map = {it["id"]: it for it in rubrica}
            yield ev({"tipo": "rubrica", "items": rubrica})
            yield ev({"tipo": "progreso", "msg": f"Enviando a Claude Vision ({len(rubrica)} criterios)…"})

            criterios_txt = "\n".join(f"{it['id']}. {it['checklist']}" for it in rubrica)

            prompt = f"""Analiza VISUALMENTE este plano de Taller de Arquitectura, Semana {semana}.

INSTRUCCION CRITICA: evalua SOLO lo que observas directamente. Cita elementos visuales concretos.
Si algo no es visible, nivel=0. NO asumas ni inventes.

CRITERIOS ({len(rubrica)}):
{criterios_txt}

ESCALA: 0=No presentado | 1=Insuficiente | 2=Suficiente | 3=Logrado

Responde SOLO NDJSON (un JSON por linea, sin texto extra):
Lineas 1-{len(rubrica)}: {{"id":1,"cumple":true,"nivel":3,"evidencia":"elemento visual concreto observado","feedback":"comentario"}}
Ultima linea: {{"tipo":"resumen","fortalezas":"...","areas_mejora":"...","comentario_general":"..."}}"""

            client = anthropic.AsyncAnthropic(api_key=API_KEY)

            # ── Llamada a Claude + keep-alive concurrente ──────────────
            # Railway bufferiza SSE, así que llamamos sin streaming y
            # mandamos pings cada 3s para que no corte la conexión.
            done_event = asyncio.Event()

            async def keepalive_loop():
                while not done_event.is_set():
                    await asyncio.sleep(3)

            # Tarea paralela de keep-alive
            ka_task = asyncio.create_task(keepalive_loop())

            # Claude call (no-stream, espera respuesta completa)
            try:
                msg = await client.messages.create(
                    model=MODEL,
                    max_tokens=4000,
                    system=(
                        "Eres evaluador experto en arquitectura universitaria. "
                        "Evalúas Taller VIII (Arquitectura y Ciudad). "
                        "Solo evalúas lo que observas en la imagen. "
                        "Responde en español y en formato NDJSON estricto."
                    ),
                    messages=[{
                        "role": "user",
                        "content": [
                            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                            {"type": "text",  "text": prompt},
                        ],
                    }],
                )
            finally:
                done_event.set()
                ka_task.cancel()

            # Mandamos keep-alive real mientras esperábamos (yield)
            yield KEEPALIVE

            # ── Parsear y emitir criterios uno a uno ──────────────────
            raw = msg.content[0].text.strip()
            for line in raw.split("\n"):
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                    if "tipo" not in obj:
                        obj["tipo"] = "criterio"
                        item = rubrica_map.get(obj.get("id"), {})
                        obj["checklist"] = item.get("checklist", "")
                        obj["dimension"]  = item.get("dimension", "")
                    yield ev(obj)
                except json.JSONDecodeError:
                    pass

            # flush buffer
            if buffer.strip():
                try:
                    obj = json.loads(buffer.strip())
                    if "tipo" not in obj:
                        obj["tipo"] = "criterio"
                        item = rubrica_map.get(obj.get("id"), {})
                        obj["checklist"] = item.get("checklist", "")
                        obj["dimension"]  = item.get("dimension", "")
                    yield ev(obj)
                except Exception:
                    pass

            yield ev({"tipo": "fin"})

        except Exception as e:
            import traceback
            yield ev({"tipo": "error", "msg": str(e), "trace": traceback.format_exc()})

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache, no-transform", "X-Accel-Buffering": "no"},
    )
