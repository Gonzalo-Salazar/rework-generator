# Generador de Flujos de Retrabajo – Proceso: Elaboración de Zapato

> **Proyecto Final – Ingeniería de Procesos 2026**
> Curso: Automatización de Flujos de Retrabajo

---

## Descripción del Proyecto

Un flujo principal de cualquier proceso requiere retrabajos. Asignar manualmente los reprocesos a cada paso conlleva tiempos muertos. Este proyecto **automatiza la generación de flujos de retrabajo** asignando cadenas de texto estructuradas (strings) a cada paso del proceso principal, basándose en **Razones (Reasons)**, **Flujos de Retrabajo** y **Pasos de Retorno**.

El proceso elegido es la **elaboración de un zapato**.

---

## Proceso de Solución (Diario de Investigación)

### Día 1 – Comprensión del problema (19-MAY-2026)

**Objetivo:** Entender qué se pide exactamente y definir el alcance.

El profesor presentó el Statement y la plantilla `Rework Generator.xlsx`. Procedimos a leer el archivo Excel con Python para entender la estructura de datos real:

```
Hoja FlowStructures:  FLOW | STEP | POSITION | REWORKS
Hoja FlowReworks:     REASON | REWORK FLOW | STEP REWORK | POSITION
```

**Conclusiones extraídas de la plantilla:**
- La columna `REWORKS` ya contiene el string de salida pre-formado, lo que confirma el formato exacto requerido.
- Cada flujo de retrabajo está compuesto de: Razón → Nombre del Flujo → Primer Paso → Paso de Retorno al flujo principal.
- Un mismo paso del proceso principal puede tener **múltiples retrabajos**.

**Formato de output identificado:**
```
NombrePaso-->GoToFlowPath[NombreFlujo/PrimerPaso]
ReturnStep[PasoDeRetorno]
Reason[Razón];
```

---

### Día 2 – Diseño de la solución (21-MAY-2026)

**Herramientas evaluadas:**
1. **Python puro** – rápido de prototipar, fácil de mantener, ideal para leer el Excel.
2. **App Web (HTML/JS)** – sin dependencias, ejecutable en cualquier navegador, interactiva.
3. **Combinación de ambas** – elegida por máxima cobertura: el script Python sirve como backend/CLI y la app web como interfaz visual.

**Decisión de arquitectura:**

```
┌─────────────────────────────────────────┐
│           SOLUCIÓN FINAL                │
│                                         │
│  index.html       ←→   Interfaz visual  │
│  rework_generator.py  ←→   CLI / Excel  │
│  README.md        ←→   Documentación    │
└─────────────────────────────────────────┘
```

**Modelo de datos definido:**

```python
MainStep:
  - name: str
  - position: int
  - reworks: list[ReworkFlow]

ReworkFlow:
  - reason: str       # "Corte Defectuoso"
  - flow_name: str    # "Retrabajar Corte"
  - first_step: str   # "Revisar Patron"
  - return_step: str  # "Cortar Material"
```

---

### Día 3 – Implementación del Script Python (23-MAY-2026)

**Archivo:** `rework_generator.py`

Se desarrolló en tres módulos:

1. **Modelo de datos** (`ReworkFlow`, `MainStep`) – clases que representan el proceso.
2. **Datos de ejemplo** (`get_shoe_process()`) – proceso de zapato con 8 pasos y 6 retrabajos.
3. **Lector de Excel** (`load_from_excel()`) – parsea la plantilla usando `openpyxl` y regex para extraer los bloques `GoToFlowPath[...]`.
4. **Generador de output** (`generate_output()`) – itera los pasos y construye el string final.

**Errores encontrados y soluciones:**

| Error | Causa | Solución |
|-------|-------|----------|
| Filas duplicadas en FlowReworks | La plantilla repite algunos flujos | Se usó `seen_flows` dict para deduplicar |
| `return_step` vacío al leer Excel | La hoja FlowReworks no tiene esa columna directamente | Se parseó desde la columna `REWORKS` de FlowStructures con regex |
| Encoding de caracteres en `rework_output.txt` | Caracteres especiales en español | Se forzó `encoding="utf-8"` en `write_text()` |

**Prueba de ejecución:**
```bash
python rework_generator.py
# Output → rework_output.txt

python rework_generator.py --excel "Rework Generator (1).xlsx"
# Output desde plantilla real
```

---

### Día 4 – Implementación de la App Web (25-MAY-2026)

**Archivo:** `index.html`

Se diseñó una interfaz de una sola página (sin dependencias externas) que permite:

- **Panel izquierdo:** Definir y editar los pasos del proceso principal.
- **Panel derecho:** Definir los flujos de retrabajo (Razón, Flujo, Primer Paso, Paso de Retorno).
- **Asignación dinámica:** Dropdown para asignar múltiples retrabajos a cada paso.
- **Generador de output:** Botón que produce el string exacto con resaltado de sintaxis por colores.
- **Copiar al portapapeles:** Un clic para copiar el resultado completo.

La app viene pre-cargada con el **proceso de elaboración de zapato** como caso de uso real.

**Problema encontrado:** Al hacer syntax highlighting con HTML `innerHTML`, los caracteres especiales se escapaban incorrectamente.  
**Solución:** Se creó la función `esc()` para sanitizar valores antes de insertar en atributos HTML.

---

### Día 5 – Validación del output (26-MAY-2026)

Se verificó manualmente que el output generado por ambas soluciones cumple exactamente el formato requerido:

**Output esperado (del Statement):**
```
Empacar Lata-->GoToFlowPath[Retrabajar Empaque/Desempacar]
ReturnStep[Empacar Lata]
Reason[Empaque Dañado];
```

**Output generado por nuestra solución:**
```
Empaque-->GoToFlowPath[Retrabajar Empaque/Desempacar]
ReturnStep[Empaque]
Reason[Empaque Dañado];
```

✅ Formato correcto: `Step-->GoToFlowPath[Flujo/Paso]`, `ReturnStep[...]`, `Reason[...];`

---

## Proceso Principal – Elaboración de Zapato

| # | Paso | Razón de Retrabajo | Flujo de Retrabajo |
|---|------|--------------------|-------------------|
| 1 | Cortar Material | Corte Defectuoso | Retrabajar Corte / Revisar Patron |
| 2 | Preparar Plantilla | – | – |
| 3 | Ensamblar Upper | – | – |
| 4 | Pegar Suela | Pegado Deficiente | Retrabajar Pegado / Despegar Suela |
| 5 | Costura | Costura Rota | Retrabajar Costura / Descosturar Pieza |
| 6 | Acabado | Acabado Malo | Retrabajar Acabado / Limpiar Superficie |
| 7 | Control de Calidad | Fallo de Calidad | Reproceso QC / Identificar Defecto |
| 8 | Empaque | Empaque Dañado | Retrabajar Empaque / Desempacar |

---

## Output Generado

```
Cortar Material-->GoToFlowPath[Retrabajar Corte/Revisar Patron]
ReturnStep[Cortar Material]
Reason[Corte Defectuoso];

Pegar Suela-->GoToFlowPath[Retrabajar Pegado/Despegar Suela]
ReturnStep[Pegar Suela]
Reason[Pegado Deficiente];

Costura-->GoToFlowPath[Retrabajar Costura/Descosturar Pieza]
ReturnStep[Costura]
Reason[Costura Rota];

Acabado-->GoToFlowPath[Retrabajar Acabado/Limpiar Superficie]
ReturnStep[Acabado]
Reason[Acabado Malo];

Control de Calidad-->GoToFlowPath[Reproceso QC/Identificar Defecto]
ReturnStep[Control de Calidad]
Reason[Fallo de Calidad];

Empaque-->GoToFlowPath[Retrabajar Empaque/Desempacar]
ReturnStep[Empaque]
Reason[Empaque Dañado];
```

---

## Instalación y Uso

### Requisitos
- Python 3.8+
- `openpyxl` (solo si se usa el modo Excel)

```bash
pip install openpyxl
```

### Opción A – App Web (recomendada)
```bash
# Abrir directamente en el navegador
open index.html
# o doble clic en index.html
```

### Opción B – Script Python (datos de ejemplo)
```bash
python rework_generator.py
```

### Opción C – Script Python con Excel
```bash
python rework_generator.py --excel "Rework Generator (1).xlsx"
```

### Opción D – Guardar output en archivo
```bash
python rework_generator.py --output mi_output.txt
```

---

## Estructura del Repositorio

```
rework-generator/
├── index.html              # App web interactiva (sin dependencias)
├── rework_generator.py     # Script Python CLI
├── rework_output.txt       # Output de ejemplo generado
├── Rework Generator.xlsx   # Plantilla de referencia (del profesor)
└── README.md               # Este documento
```

---

## Herramientas Utilizadas

- **Python 3** – lógica principal y lectura de Excel
- **openpyxl** – lectura de archivos `.xlsx`
- **HTML / CSS / JavaScript** – interfaz web sin dependencias externas
- **Claude AI (Anthropic)** – asistente para estructurar la solución, depurar el parser regex y diseñar la UI

---

## Reflexión Final

El reto principal no fue técnico sino de **comprensión del dominio**: entender exactamente qué significa cada campo del output (`GoToFlowPath`, `ReturnStep`, `Reason`) y cómo se relacionan entre sí. Una vez modelado correctamente el dato, la generación del string fue trivial.

La decisión de ofrecer **dos soluciones** (app web + script Python) permite que cualquier usuario, independientemente de su nivel técnico, pueda utilizar el generador: la app web es accesible para todos; el script Python es extensible para integraciones futuras.

---

*Generado con apoyo de Claude AI – Anthropic*
