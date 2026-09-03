// La app se sirve desde el mismo backend (http://127.0.0.1:8000/app),
// así que usamos una ruta relativa y evitamos problemas de CORS.
const URL_API = "/procesar";

const formulario = document.getElementById("form-subida");
const entradaArchivo = document.getElementById("archivo");
const textoArchivo = document.getElementById("texto-archivo");
const botonProcesar = document.getElementById("boton-procesar");
const estado = document.getElementById("estado");
const resultado = document.getElementById("resultado");
const enlaceDescarga = document.getElementById("enlace-descarga");

entradaArchivo.addEventListener("change", () => {
  if (entradaArchivo.files.length) {
    textoArchivo.textContent = entradaArchivo.files[0].name;
  } else {
    textoArchivo.textContent = "Selecciona o arrastra el Excel (.xlsx)";
  }
});

formulario.addEventListener("submit", async (evento) => {
  evento.preventDefault();

  if (!entradaArchivo.files.length) {
    mostrarEstado("Selecciona un archivo Excel primero.", "error");
    return;
  }

  const datosFormulario = new FormData();
  datosFormulario.append("archivo", entradaArchivo.files[0]);

  botonProcesar.disabled = true;
  resultado.hidden = true;
  mostrarEstado("Procesando archivo...", "cargando");

  try {
    const respuesta = await fetch(URL_API, {
      method: "POST",
      body: datosFormulario,
    });

    if (!respuesta.ok) {
      let mensaje = `Error del servidor (${respuesta.status})`;

      try {
        const error = await respuesta.json();
        mensaje = error.detail || mensaje;
      } catch (_) {
        // el cuerpo no era JSON, se deja el mensaje genérico
      }

      throw new Error(mensaje);
    }

    const metricasHeader = respuesta.headers.get("X-Metricas");
    const metricas = metricasHeader ? JSON.parse(metricasHeader) : null;

    const disposicion = respuesta.headers.get("Content-Disposition") || "";
    const coincidencia = disposicion.match(/filename="([^"]+)"/);
    const nombreArchivo = coincidencia ? coincidencia[1] : "Transporte_Validado.xlsx";

    const blob = await respuesta.blob();
    const url = URL.createObjectURL(blob);

    enlaceDescarga.href = url;
    enlaceDescarga.download = nombreArchivo;
    enlaceDescarga.textContent = `Descargar ${nombreArchivo}`;

    if (metricas) {
      mostrarMetricas(metricas);
    }

    resultado.hidden = false;
    mostrarEstado("Proceso completado.", "exito");

  } catch (error) {
    mostrarEstado(error.message, "error");
  } finally {
    botonProcesar.disabled = false;
  }
});

function mostrarEstado(mensaje, tipo) {
  estado.textContent = mensaje;
  estado.className = `estado ${tipo}`;
}

function mostrarMetricas(metricas) {
  document.getElementById("m-total").textContent =
    metricas.servicios_totales.toLocaleString("es-DO");

  document.getElementById("m-unificados").textContent =
    metricas.servicios_unificados.toLocaleString("es-DO");

  document.getElementById("m-grupos").textContent =
    metricas.grupos_unificados.toLocaleString("es-DO");

  document.getElementById("m-original").textContent =
    formatoMoneda(metricas.costo_original);

  document.getElementById("m-ajustado").textContent =
    formatoMoneda(metricas.costo_ajustado);

  document.getElementById("m-ahorro").textContent =
    formatoMoneda(metricas.ahorro);

  document.getElementById("m-porcentaje").textContent =
    `${(metricas.porcentaje_ahorro * 100).toFixed(2)}%`;
}

function formatoMoneda(valor) {
  return `RD$ ${Number(valor).toLocaleString("es-DO", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}
