# INDISA Consentimientos AI

Este repositorio contiene el código fuente del sistema de revisión automática de consentimientos informados de Clínica INDISA, desarrollado en Python.

El proyecto incluye dos versiones:

- **V1:** utiliza una fuente externa de atenciones para identificar al médico y resolver el timbre correspondiente.
- **V2:** procesa directamente el consentimiento sin utilizar Excel ni una imagen de timbre. Detecta presencia de campos, checkbox, firmas, timbres existentes y valida el slot disponible para timbrado.

Actualmente se soportan las siguientes plantillas:

- Endoscopía Digestiva Alta.
- Colonoscopía / Endoscopía Digestiva Baja.

# Estructura del proyecto

```text
indisa-consent-ai/
├── README.md
├── requirements.txt
├── .gitignore
├── .gitattributes
│
├── models/
│   ├── field_presence_v811_best.pt
│   ├── visual_main_v813.pt
│   └── doctor_signature_fallback_v811.pt
│
├── datos/
│   └── Consentimientos/
│       └── Plantillas base/
│           ├── Consentimiento_base_EndoscopíaDigestivaAlta_Esófago-gastr.pdf
│           └── Consentimiento_base_Colonoscopía_EndoscopíaDigestivaBaja.pdf
│
├── v1/
│   └── run_indisa_v1_1.py
│
└── v2/
    ├── indisa_core_v9_1.py
    ├── placement_zones.json
    └── run_indisa_v2_5_3.py
```

# Generar un ambiente

Ir a la carpeta raíz del proyecto y ejecutar:

```bash
python -m venv venv
```

# Activar el ambiente en Linux o Mac

```bash
source venv/bin/activate
```

# Activar el ambiente en Windows

PowerShell:

```powershell
venv\Scripts\Activate.ps1
```

Si PowerShell bloquea temporalmente la activación:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
venv\Scripts\Activate.ps1
```

# Instalar dependencias

Con el ambiente activado:

```bash
pip install -r requirements.txt
```

La ejecución con GPU requiere una instalación de PyTorch compatible con la GPU, los drivers NVIDIA y la versión de CUDA del equipo.

# Git LFS y modelos

Los tres modelos `.pt` forman parte del repositorio y se administran mediante **Git LFS**.

Antes de clonar o descargar los modelos desde GitHub se recomienda instalar Git LFS:

```powershell
git lfs install
```

Después de clonar el repositorio:

```powershell
git lfs pull
```

Para verificar que los modelos estén disponibles:

```powershell
Get-ChildItem ".\models"
```

Deben existir:

```text
field_presence_v811_best.pt
visual_main_v813.pt
doctor_signature_fallback_v811.pt
```

# Plantillas

Las dos plantillas base necesarias para el pipeline están incluidas en:

```text
datos/Consentimientos/Plantillas base/
```

# Ejecutar V2 Autónoma

La V2 Autónoma no utiliza Excel ni requiere una imagen de timbre.

Desde la raíz del repositorio:

```powershell
python .\v2\run_indisa_v2_5_3.py `
  --input "CARPETA_CON_PDFS" `
  --placement-zones ".\v2\placement_zones.json" `
  --output "resultado_v2"
```

# Ejecutar V2 sin exportar PDFs

La exportación de PDFs es opcional.

Si solo se necesitan las inferencias, JSON y CSV:

```powershell
python .\v2\run_indisa_v2_5_3.py `
  --input "CARPETA_CON_PDFS" `
  --placement-zones ".\v2\placement_zones.json" `
  --output "resultado_v2" `
  --no-pdf-export
```

Con `--no-pdf-export` se siguen generando:

```text
results.json
results_debug.json
alerts.json
audit_log.csv
summary.json
```

pero no se generan:

```text
pdfs/
manual_review/
final_pdfs/
```

Esta opción puede reducir el tiempo de ejecución cuando el resultado será consumido por una API o integración.

# Ejecutar V1 Production

La V1 utiliza una fuente externa de atenciones para conocer qué médico corresponde al consentimiento y resolver el timbre asociado.

Además del contenido incluido en el repositorio, V1 requiere localmente:

```text
datos/Excel atenciones ejemplo.xlsx
datos/timbres/
```

y una instalación de Tesseract.

Ejecución básica:

```powershell
python .\v1\run_indisa_v1_1.py `
  --input "CARPETA_CON_PDFS" `
  --output "resultado_v1"
```

Si Tesseract está instalado en otra ubicación se puede indicar:

```powershell
--tesseract "RUTA_A_TESSERACT.exe"
```

V1 también permite ejecutar sin exportar PDFs:

```powershell
python .\v1\run_indisa_v1_1.py `
  --input "CARPETA_CON_PDFS" `
  --output "resultado_v1" `
  --no-pdf-export
```

# Archivos de salida

Cuando la exportación PDF está habilitada:

```text
resultado/
├── final_pdfs/
├── pdfs/
├── manual_review/
├── previews/
├── results.json
├── results_debug.json
├── alerts.json
├── audit_log.csv
└── summary.json
```

# Convención de nombres y versiones

Las carpetas principales del proyecto son:

```text
v1/
v2/
```

Los ejecutables principales son:

```text
V1 -> v1/run_indisa_v1_1.py
V2 -> v2/run_indisa_v2_5_3.py
```

Las versiones informadas por los JSON son:

```text
V1 -> pipeline_version: "V1.1"
V2 -> pipeline_version: "V2.5.3"
```

Ambas versiones generan los mismos nombres de archivos estructurados:

```text
results.json
results_debug.json
alerts.json
audit_log.csv
summary.json
```

# results.json

`results.json` es la salida simplificada recomendada para la integración.

Su estructura principal es:

```json
{
  "schema_version": "1.0",
  "pipeline_version": "V2.5.3",
  "documents": []
}
```

Cada documento contiene principalmente:

- archivo y página;
- plantilla reconocida;
- estado y confianza de los campos;
- checkbox;
- firmas;
- estado de timbre o slot;
- revisión manual;
- ruta del PDF final, cuando se exporta.

Ejemplo simplificado:

```json
{
  "source_pdf": "consentimiento.pdf",
  "page": 1,
  "template": {
    "type": "endoscopia_alta",
    "recognized": true,
    "alignment_inliers": 212
  },
  "fields": {
    "patient_name": {
      "status": "FILLED",
      "confidence": 0.96,
      "source": "primary_field_roi"
    },
    "procedure_date": {
      "status": "EMPTY",
      "confidence": 0.94,
      "source": "primary_field_roi"
    }
  },
  "checkbox": {
    "status": "PRESENT",
    "confidence": 0.82,
    "option": "Paciente"
  },
  "signatures": {
    "patient": {
      "status": "PRESENT",
      "confidence": 0.74
    },
    "doctor": {
      "status": "PRESENT_GEOMETRIC_CONFIRMED",
      "confidence": 0.08,
      "confirmation": {
        "method": "geometric_confirmation",
        "score": 0.043
      }
    }
  },
  "review": {
    "required": true,
    "reasons": [
      {
        "component": "procedure_date",
        "reason": "missing_required_field"
      }
    ]
  }
}
```

# results_debug.json

`results_debug.json` contiene la salida técnica completa del pipeline.

Incluye información utilizada para trazabilidad y depuración, como:

- candidatos YOLO;
- bounding boxes;
- ROIs;
- scores geométricos;
- homografías;
- información de alineación;
- detalles internos del postprocesamiento.

Para una integración normal se recomienda utilizar `results.json`.

```text
Aplicación / integración -> results.json
Depuración / auditoría   -> results_debug.json
```

# alerts.json

`alerts.json` contiene únicamente los problemas o condiciones que requieren atención.

Puede utilizarse directamente para mostrar advertencias en una interfaz.

# audit_log.csv

`audit_log.csv` es la salida recomendada para Excel, Power BI, pandas o reportería.

Una fila representa un consentimiento o página procesada.

Entre sus columnas se encuentran:

```text
source_pdf
page
template_type

patient_name_status
patient_name_confidence
patient_rut_status
patient_rut_confidence
procedure_date_status
procedure_date_confidence
doctor_name_status
doctor_name_confidence
doctor_rut_status
doctor_rut_confidence

checkbox_status
checkbox_option
checkbox_confidence

patient_signature_status
patient_signature_confidence

doctor_signature_status
doctor_signature_confidence

slot_status
manual_review_required
manual_review_reasons
```

# Estados de los campos

Los campos estructurados utilizan:

```text
FILLED
EMPTY
UNCERTAIN
```

- `FILLED`: se detectó contenido.
- `EMPTY`: el campo se considera vacío.
- `UNCERTAIN`: no existe evidencia suficiente para decidir con seguridad.

# Nivel de confianza de los campos

Para aceptar directamente la predicción del clasificador se utiliza:

```text
confidence >= 0.80
```

Ejemplo:

```text
present con 0.95 -> FILLED
empty con 0.91   -> EMPTY
0.72             -> UNCERTAIN inicialmente
```

La confianza corresponde a una puntuación del modelo y no debe interpretarse literalmente como un porcentaje de probabilidad de acierto.

# Confirmación geométrica de campos

Un campo inicialmente `UNCERTAIN` puede confirmarse como `FILLED` mediante comparación con la plantilla.

Umbral:

```text
geometric score >= 0.010
```

En esos casos el resultado puede incluir:

```json
{
  "status": "FILLED",
  "confidence": 0.72,
  "source": "template_subtraction_confirmation",
  "confirmation": {
    "method": "template_subtraction",
    "score": 0.041
  }
}
```

# Nombre y RUT del paciente mediante etiqueta superior

En V2, `patient_name` y `patient_rut` pueden utilizar como evidencia alternativa una etiqueta o sello de datos ubicado en el encabezado.

Para evitar falsos positivos se exige:

```text
data_stamp confidence >= 0.80
```

y que la detección esté efectivamente dentro de la zona del encabezado del paciente.

# Checkbox y firmas

Para las detecciones visuales:

```text
confidence >= 0.20 -> detección directa
0.03 a < 0.20      -> candidato de baja confianza
< 0.03              -> se descarta
```

Un candidato de baja confianza puede ser confirmado mediante evidencia geométrica.

Firma paciente:

```text
geometric score >= 0.021232
```

Firma médico:

```text
geometric score >= 0.013727
```

Checkbox:

```text
geometric score >= 0.023835
option margin >= 0.015
opción geométrica == opción YOLO
```

Por eso una detección puede tener una confianza YOLO baja y aun así aparecer como:

```text
PRESENT_GEOMETRIC_CONFIRMED
```

si la segunda evidencia supera el umbral correspondiente.

# Reconocimiento de plantilla

V2 exige actualmente:

```text
template_orb_inliers >= 40
```

Este valor corresponde a coincidencias geométricas consistentes con la plantilla y no a una probabilidad.

# Cambio de umbrales

Los umbrales pueden modificarse a nivel de código para procesos de recalibración o futuras validaciones.

Sin embargo, los valores incluidos en esta versión corresponden a la configuración congelada del pipeline.

No se recomienda modificarlos sin volver a evaluar el sistema sobre un conjunto de validación, ya que reducir un threshold puede disminuir casos `UNCERTAIN` pero aumentar falsos positivos.

# Revisión manual

La decisión principal está resumida en:

```text
manual_review_required
```

en `audit_log.csv`, o en:

```json
"review": {
  "required": true
}
```

dentro de `results.json`.

Las razones específicas quedan disponibles en:

```text
manual_review_reasons
```

o en el arreglo `review.reasons`.

# summary.json

`summary.json` contiene los totales generales de la corrida, por ejemplo:

- PDFs procesados;
- consentimientos reconocidos;
- documentos a revisión;
- slots disponibles;
- timbres existentes;
- alertas;
- páginas duplicadas o rechazadas;
- estado de exportación PDF.

Cuando se utiliza:

```text
--no-pdf-export
```

el summary registra:

```json
"pdf_export_enabled": false
```

# PDFs de salida

Cuando la exportación está habilitada:

`final_pdfs/`
: PDF de revisión con cajas de inferencia.

`pdfs/`
: PDF limpio cuando no requiere revisión.

`manual_review/`
: PDF limpio cuando existe una condición de revisión.

Los PDFs originales no se modifican.
