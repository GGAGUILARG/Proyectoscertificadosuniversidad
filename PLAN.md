# PLAN.md — Academia Horizonte: sistema de certificados

## Objetivo

Automatizar la emisión de certificados por cohorte a partir de los Excel de `INSUMOS/`, con una aplicación web en 3 piezas:

- **Bodega (datos)**: `INSUMOS/` — los Excel NO se modifican nunca.
- **Cocina (backend)**: `app.py` con Flask — toda la lógica vive acá.
- **Salón (frontend)**: `templates/` + `static/` — solo muestran datos ya calculados.

## Estado actual

Ya existe una base funcional de la sesión anterior: `app.py`, `templates/index.html`, `templates/certificado.html`, `static/style.css`, `requirements.txt`. El plan la revisa contra los mismos criterios de validación de una construcción desde cero.

## Decisiones técnicas (acordadas)

| Tema | Decisión | Motivo |
|---|---|---|
| Servidor web | **Flask** | Mínimo, un archivo, suficiente para el tamaño del proyecto |
| Lectura de Excel | **openpyxl** (con `data_only=True`) | Ligera; deja los bucles y condicionales visibles para comentar |
| Clave de unión | `Identificacion` como **str** en ambos archivos | Viene como texto; evitar que int/str rompa la unión silenciosamente |
| Agrupación | Por `(Identificacion, Programa)`, programa tomado del registro de evaluaciones | El maestro y el registro pueden discrepar algún día |
| Promedios | Promedio = Σ Notas / módulos cursados; Asistencia = promedio de Asistencia_Pct | Regla de negocio; límites INCLUYEN el valor (>= 70, >= 80) |
| Programa | `Técnico en IA Aplicada` → módulos 1–4; `Excel Avanzado para Negocios` → 1–3 | No asumir cantidad fija de módulos |
| Encoding | UTF-8 en todo (`Técnico`, `Módulo`) | Los Excel son UTF-8 |

## Números de referencia (ya verificados, sirven de aceptación)

- Maestro: 24 estudiantes, cohorte `2026-A` (16 IA + 8 Excel).
- Registro: 88 filas (una por estudiante/módulo), 24 estudiantes distintos.
- `304560321` en maestro sin notas; `999880777` con notas sin maestro → fuera de resultados ("por ahora nada", sin advertencias).
- Resultado: **19 Aprobación, 3 Participación, 1 Sin certificado** (`111230984`), **22 certificados emitibles**.

---

## Etapa 1 — Validar el cruce de datos (bodega)

**Estado: completada** — entregable: `validar_datos.py` (solo lectura). El reporte coincidió con los números de referencia; las 2 inconsistencias detectadas son las esperadas.

**Qué se hace**
- Script de solo lectura que cruza ambos Excel y reporta:
  1. filas y estudiantes por archivo;
  2. IDs en un archivo y no en el otro;
  3. grupos `(Identificacion, Programa)` y módulos por programa;
  4. duplicados de `(Identificacion, Programa, Modulo)`;
  5. Notas y Asistencia fuera del rango 0–100;
  6. resumen de estados (19 / 3 / 1) con los mismos cálculos de las reglas.

**Entregable**: reporte de validación (consola o archivo de salida; no toca los Excel).

**Cómo se valida**
- El reporte coincide **exactamente** con los números de referencia de arriba.
- Todo caso raro (IDs huérfanos, duplicados, rangos) queda documentado en el reporte; si aparece algo nuevo, se detiene la etapa y se consulta a negocio.
- Criterio de aceptación: sin discrepancias o con excepción documentada.

## Etapa 2 — Backend Flask (cocina)

**Estado: completada** — `app.py` validado con cliente de Flask: conteos 19/3/1, 22 certificados, 404 en sin-certificado/inexistente.

**Qué se hace**
- `app.py` con funciones separadas (cada una su ladrillo):
  - `cargar_estudiantes()` → dict `identificacion → {nombre, correo, programa, cohorte}`;
  - `cargar_notas()` → lista de registros con módulo, nota, asistencia;
  - `calcular_promedio(lista)` → promedio de una lista de números;
  - `decidir_certificado(promedio, asistencia)` → `Aprobado` / `Participacion` / `Sin certificado` (límites inclusivos);
  - `calcular_resultados()` → agrupa por `(Identificacion, Programa)`, calcula promedios y estado, descarta los dos casos irregulares;
  - rutas `GET /` (listado) y `GET /certificado/<identificacion>/<programa>` (vista única; **404** si no existe o no hay certificado).
- Comentarios en español marcando variables, bucles, condicionales y funciones.
- NUNCA lógica en el HTML.

**Cómo se valida** (test del cliente de Flask, sin servidor)
- `GET /` responde 200.
- `GET /certificado/<id>/<programa>` de un Aprobado y de un Participación responden 200.
- Ídem para un "Sin certificado" responde 404; id inexistente responde 404.
- Conteo de estados == 19 / 3 / 1 y certificados == 22.
- Si algún conteo no coincide, se corrige la lógica antes de pasar a la interfaz.

## Etapa 3 — Interfaz (salón)

**Estado: completada** — servidor real OK: listado 200 con 23 estudiantes y acentos correctos; certificado 200; `@media print` activo.

**Qué se hace**
- `templates/index.html`: 4 tarjetas de resumen (Aprobación, Participación, Sin certificado, Total) + **filtro por programa** + tabla con Nombre, Programa, Módulos, Promedio, Asistencia, Estado (fila de color: verde Aprobación, azul Participación, rojo Sin certificado) + enlace "Ver certificado" (solo cuando existe) + **sección de inconsistencias** (las 3 detectadas: sin evaluaciones, sin ficha en maestro, sin certificado).
- `templates/certificado.html`: certificado imprimible (Academia Horizonte, tipo de certificado, nombre, programa, cohorte, tabla de módulos, promedio, asistencia, fecha de emisión) + botones "Volver" e "Imprimir".
- `static/style.css`: identidad azul marino `#0d2b52` y dorado `#bf9b30`; `@media print` para que al imprimir solo salga el certificado.

**Cómo se valida** (servidor real)
- `python app.py` arranca sin errores; `http://127.0.0.1:5000` responde 200.
- El listado muestra los 23 estudiantes con sus estados correctos.
- Acentos correctos en pantalla (`Técnico`, `Módulo`, `Participación`).
- Un certificado se ve bien y al imprimir genera PDF limpio (solo el certificado, sin botones).
- Revisión visual del color y contraste (azul marino/dorado).

## Etapa 4 — Entrega final

**Estado: completada** — `requirements.txt` verificado, `__pycache__` limpiado, documentación reconciliada.

**Qué se hace**
- `requirements.txt` con `flask` y `openpyxl`.
- `iniciar_app.bat`: arranque con doble clic (instala dependencias si faltan, abre el navegador y corre `app.py`). Detener: Ctrl+C o cerrar la ventana.
- Recorrido completo: listado → filtro → certificado → imprimir.
- Limpieza de archivos temporales (`__pycache__`).
- Actualización de `AGENTS.md` si algo cambió durante la construcción.

**Cómo se valida**
- Checklist final: los 4 pasos de validación anteriores pasan en el estado final del repo.
- Ejecución fresca: `pip install -r requirements.txt` + `python app.py` funcionan en una máquina limpia.
