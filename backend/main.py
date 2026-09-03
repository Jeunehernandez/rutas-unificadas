import json
import urllib.parse

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from backend import procesador

app = FastAPI(
    title="Validador de Rutas Unificadas",
    description="Aplicación para validar y calcular costos de rutas de transporte.",
    version="1.0.0"
)

# Durante el desarrollo local permitimos cualquier origen.
# Antes de publicar en producción, restringir a los dominios reales.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Metricas"],
)


@app.get("/")
def inicio():
    return {
        "mensaje": "Validador de Rutas Unificadas funcionando correctamente"
    }


@app.get("/salud")
def salud():
    return {
        "estado": "OK"
    }


@app.post("/procesar")
async def procesar(archivo: UploadFile = File(...)):
    """
    Recibe un Excel de rutas de transporte, aplica la regla de
    negocio (rutas unificadas 100 % / 50 % por grupo y fecha) y
    devuelve el Excel validado con sus resúmenes.
    """

    nombre_original = archivo.filename or ""

    if not nombre_original.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="El archivo debe ser un Excel (.xlsx o .xls)."
        )

    contenido = await archivo.read()

    if not contenido:
        raise HTTPException(
            status_code=400,
            detail="El archivo está vacío."
        )

    try:
        salida_bytes, nombre_salida, metricas = procesador.generar_reporte(contenido)

    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error))

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"Error inesperado al procesar el archivo: {error}"
        )

    nombre_codificado = urllib.parse.quote(nombre_salida)

    headers = {
        "Content-Disposition": (
            f'attachment; filename="{nombre_salida}"; '
            f"filename*=UTF-8''{nombre_codificado}"
        ),
        "X-Metricas": json.dumps(metricas, ensure_ascii=False),
    }

    return StreamingResponse(
        iter([salida_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/app")
def redirigir_a_app():
    # Sin esto, entrar a "/app" (sin la barra final) da "Not Found",
    # porque los archivos estáticos quedan montados en "/app/".
    return RedirectResponse(url="/app/")


# Sirve el frontend en http://127.0.0.1:8000/app/
app.mount("/app", StaticFiles(directory="frontend", html=True), name="frontend")
