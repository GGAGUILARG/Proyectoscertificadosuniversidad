# AGENTS.md

## Idioma

El usuario habla español: respondé en español. El código, comentarios y textos de la interfaz van en español.

## Proyecto

Academia Horizonte — sistema web de certificados por cohorte. Se arma en 3 piezas:

- **Bodega (datos)**: los Excel en `INSUMOS/`. NO se modifican nunca.
- **Cocina (backend)**: `app.py` (Flask). TODA la lógica vive acá, nunca en el HTML.
- **Salón (frontend)**: `templates/` + `static/style.css`. Solo muestran datos ya calculados. La página principal muestra contadores (19/3/1), filtro por programa, tabla con color por estado (verde Aprobación, azul Participación, rojo Sin certificado) y la sección de inconsistencias (las 3 detectadas).

Python 3.12 (instalado vía winget) + `flask` + `openpyxl` (ver `requirements.txt`). Correr: `python app.py` → http://127.0.0.1:5000. `debug=True` está activo a propósito.

Plan de construcción por etapas: ver `PLAN.md`. Números de referencia para validar (ya verificados): 19 Aprobación, 3 Participación, 1 Sin certificado, 22 certificados emitibles.

## Decisiones técnicas

- **Librerías**: Flask (servidor web) + openpyxl (lectura de Excel, `data_only=True`). Se descartaron pandas (pesado, oculta lógica) y Django/FastAPI (sobran para el tamaño).
- **Unión**: `Identificacion` como str en ambos archivos (viene como texto; no convertir a int).
- **Agrupación**: por `(Identificacion, Programa)` tomando el programa del registro de evaluaciones, no del maestro.
- **Promedios**: Promedio = Σ Notas / módulos cursados; Asistencia = promedio de `Asistencia_Pct`. Límites INCLUYEN el valor.
- **Estructura**: `app.py` (cocina) → `templates/` + `static/style.css` (salón); la interfaz nunca calcula ni decide. `validar_datos.py` = script de validación de datos (Etapa 1, solo lectura). `iniciar_app.bat` = arranque con doble clic (detener: Ctrl+C o cerrar la ventana).
- **Funciones clave en `app.py`**: `cargar_estudiantes`, `cargar_notas`, `calcular_promedio`, `decidir_certificado`, `calcular_resultados`, `calcular_inconsistencias`, `obtener_programas`. Filtro por programa vía `?programa=` (los contadores siempre son del total).

## Datos (`INSUMOS/`)

Dos Excel UTF-8, una hoja cada uno, fila 1 congelada. Datos de muestra (`@ejemplo.cr`), nunca tratarlos como PII real.

- `Maestro_Estudiantes.xlsx`: `Identificacion`, `Nombre_Completo`, `Correo`, `Programa`, `Cohorte` — 24 estudiantes, cohorte `2026-A`. `Identificacion` llega como texto (str).
- `Registro_Evaluaciones.xlsx`: `Identificacion`, `Programa`, `Modulo`, `Nota` (0-100 int), `Asistencia_Pct` (0-100 int), `Fecha_Cierre` (ISO). 88 filas, una por estudiante/módulo. `Identificacion` es la llave de unión.

## Reglas de negocio (en `app.py`, límites INCLUYEN el valor)

- Promedio = suma de Notas / cantidad de módulos cursados. Asistencia = promedio de `Asistencia_Pct`.
- Se agrupa por `Identificacion` + `Programa` → un certificado por estudiante y programa.
- Aprobación si Promedio >= 70 y Asistencia >= 80; Participación si Promedio < 70 y Asistencia >= 80; Sin certificado si Asistencia < 80.
- Módulos por programa: `Técnico en IA Aplicada` → 1–4; `Excel Avanzado para Negocios` → 1–3.

## Quirks verificados (no "arreglar" en silencio)

- `304560321` está en el maestro pero sin notas; `999880777` tiene notas pero no está en el maestro. Decisión vigente: "por ahora nada" — quedan fuera de los resultados (sin advertencias ni certificados).
- `111230984` reprobó: Nota 0 y Asistencia 0–20 en el último módulo (asistencia final < 80 → sin certificado). Caso de prueba de la regla.
- Strings acentuados (`Técnico`, `Módulo`); preservar encoding al leer/escribir.
