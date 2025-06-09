# Proyecto: Impuestos Diarios

**Objetivo**  
Extraer a diario los impuestos marcados en un Excel y generar documentos Word/TXT con el detalle.

## Primeros pasos

1. Clonar repo y entrar en carpeta:
   ```
   git clone <URL> andres_bot
   cd andres_bot
   ```
2. Crear y activar entorno:
   ```
   python3 -m venv env
   source env/bin/activate
   ```
3. Copiar `.env.template` a `.env` y configurar rutas:
   ```
   cp .env.template .env
   ```
4. Instalar dependencias (cuando haya algo en requirements.txt):
   ```
   pip install -r requirements.txt
   ```

# 🧾 ANDRES_BOT — Tax Calendar Processor

Este proyecto automatiza la lectura y análisis de un calendario fiscal en formato Excel, para generar informes diarios por empresa, país e impuesto. Actualmente el sistema se encuentra en fase de integración, pruebas de lectura y generación de reportes.

---

## 📂 ESTRUCTURA DEL PROYECTO

andres_bot/
├── data/ # Carpeta donde se guarda el Excel original y salidas generadas
├── docs/ # Documentación futura
├── env/ # Entorno virtual Python
├── scripts/ # Scripts principales de ejecución y debugging
├── src/ # Lógica principal de lectura y parsing
├── tests/ # Pruebas unitarias y validaciones
├── requirements.txt # Dependencias del proyecto
├── README.md # Este archivo

yaml
Copy
Edit

---

## 🚀 USO GENERAL

### 1. Inicialización

```bash
# Entrar al entorno virtual
cd andres_bot
source env/bin/activate

# (Opcional) Instalar dependencias si es primera vez
pip install -r requirements.txt
2. Ejecutar un informe para un día y empresa
bash
Copy
Edit
export PYTHONPATH=$(pwd)
python scripts/generate_reports.py --date 2025-06-02 --company ENDESA
3. Ejecutar para un rango de fechas y todas las empresas
bash
Copy
Edit
python scripts/generate_reports.py --range 2025-06-01 2025-06-10
Los archivos se guardan automáticamente en:

bash
Copy
Edit
data/outputs/YYYY-MM-DD_EMPRESA.txt

📜 DESCRIPCIÓN DE SCRIPTS
scripts/
generate_reports.py: Generador principal de informes diarios.

inspect_altadia.py: Inspección visual para entender errores de lectura concretos.

convert_colors.py: 🔧 Script para convertir colores condicionales en siglas (manual, a usar en Excel antes de lanzar procesamiento).

debug_excel.py: Verifica filas y columnas útiles del Excel (cabeceras de meses, días).

debug_parse.py: Inspección detallada por celda del archivo.

test_parse.py: Lanza pruebas de lectura simuladas para una pestaña.

src/
reader.py: Motor principal de lectura. Interpreta celdas con colores o siglas.

__init__.py: Inicializador.

tests/
test_parse_calendar.py: Valida el calendario completo.

test_reader.py: Valida lectura por empresa y día.

📘 FUNCIONAMIENTO INTERNO
El archivo Excel contiene una pestaña por empresa.

Cada pestaña tiene un calendario con colores o siglas (SI, OP, SD, etc.).

El sistema reconoce estas marcas y genera informes diarios con:

Empresa

País

Tipo de impuesto

Fecha

Estado (sigla y descripción)

🔧 PROBLEMAS DETECTADOS
❗ Algunos colores no se detectaban por ser formato condicional, no relleno real.

✅ Solución: sustituir los colores por las siglas correspondientes (SI, OP, etc.).

✅ PASOS A REALIZAR
🔄 FASE ACTUAL: Sustituir colores por siglas
Ir al Excel y sustituir celdas coloreadas por su sigla textual (SI, RI, etc.).

Puedes usar la validación de datos o una macro de conversión.

Colores relevantes:

Color	Sigla	Significado
#FFFF66	SI	Send information
#9966FF	RI	Review information and doubts (EY Local)
#FF5050	SD	Send draft (EY Local)
#FF99FF	AD	Approve draft
#70AD47	OS	Official Submission Deadline
#00B0F0	OP	Official Payment Deadline
#FFFFFF	SP	Submission & Payment (same day)

📩 SIGUIENTES FASES
[✔] Limpiar y unificar el Excel → con solo siglas, sin colores condicionales.

[✔] Validar lectura y generación de informes completos.

[⏳] Automatizar el sistema de envío por email.

Se integrará con smtplib o Google Workspace API.

Se permitirá enviar el informe de cada día a los destinatarios relevantes.

[⏳] Añadir UI o dashboard para configuración visual del proceso.

🧠 CONTRIBUCIÓN Y NOTAS
Este proyecto está en entorno local (WSL).

No se sincroniza con la nube por razones de seguridad.

Todos los scripts son ejecutables desde terminal con PYTHONPATH=$(pwd) activo.

Se recomienda mantener el Excel original como plantilla, y trabajar en una copia transformada.

