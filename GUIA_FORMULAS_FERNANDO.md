"""
GUÍA: Cómo usar fórmulas complejas de Fernando en tu flujo LISP → Python → Excel
==============================================================================

Este flujo permite reutilizar las fórmulas avanzadas del ODS de Fernando
en tu sistema generador de horarios.

COMPONENTES:
============

1. extraer_y_convertir_formulas.py
   - Extrae fórmulas del ODS de Fernando
   - Las convierte de formato ODS (of:=...) a Excel (=...)
   - Genera formulas_fernando_convertidas.json

2. formulas_fernando.lisp
   - Carga las fórmulas desde JSON
   - Proporciona interfaces para LISP
   - Genera plantillas de fórmulas dinámicas

3. hoja_con_formulas.py (ACTUALIZADO)
   - Nueva función: _inject_fernando_formulas()
   - Inyecta fórmulas en las celdas correctas
   - Se integra automáticamente en el proceso

FLUJO DE USO:
=============

OPCIÓN A: Desde Python puro
----------------------------

from extraer_y_convertir_formulas import FormulaConverter
from hoja_con_formulas import generate_excel

# 1. Extraer fórmulas del ODS
converter = FormulaConverter()
formulas = converter.extract_formulas_from_ods('propuesta-de-horarios-fernando-v3-2026-04--02.ods')

# 2. Generar Excel con las fórmulas
config = {
    "sheets": [
        {
            "title": "Aulas",
            "headers": [...],
            "data": [...],
            "fernando_formulas": formulas['Aulas'],  # ✅ Inyectar aquí
            ...
        }
    ]
}

generate_excel(config, "output.xlsx")


OPCIÓN B: Desde LISP (llamando a Python)
------------------------------------------

(load "formulas_fernando.lisp")

(let ((formulas (initialize-fernando-formulas)))
  ; Las fórmulas están disponibles en *formulas-fernando*
  (format t "~a~%" (get-formulas-for-sheet "Aulas"))
  
  ; Pasar a Python vía JSON...
  )


EJEMPLO COMPLETO DE FÓRMULA INYECTADA:
=====================================

Original (ODS):
  of:=SUBSTITUTE(TRIM(COM.MICROSOFT.CONCAT(
    IF([$C111.$C$5]=[.C$2];[$C111.$B$1] & " "; "");
    IF([$C112.$C$5]=[.C$2];[$C112.$B$1] & " "; "");
    IF([$C113.$C$5]=[.C$2];[$C113.$B$1] & " "; "")
  )); " "; ",")

Convertida (Excel):
  =SUBSTITUTE(TRIM(CONCAT(
    IF(C111!C5=C2,C111!B1 & " ", ""),
    IF(C112!C5=C2,C112!B1 & " ", ""),
    IF(C113!C5=C2,C113!B1 & " ", "")
  )), " ", ",")

Resultado en Excel:
  Si C111.C5 = "AM1", C111.B1 = "Análisis I"
  Si C112.C5 = "AM1", C112.B1 = "Calculus"
  → Celda muestra: "Análisis I,Calculus"


PROCESO PASO A PASO:
=====================

1. Extrae fórmulas:
   $ python3 extraer_y_convertir_formulas.py
   → Genera: formulas_fernando_convertidas.json

2. Carga en Python:
   import json
   with open('formulas_fernando_convertidas.json') as f:
       formulas = json.load(f)

3. Usa en configuración Excel:
   config["sheets"][0]["fernando_formulas"] = formulas.get("Aulas", [])

4. Genera Excel:
   generate_excel(config, "horario_con_formulas.xlsx")

5. ¡Verifica! Las fórmulas se evaluarán en Excel automáticamente


TESTING:
========

Ver prueba_formulas_fernando.py para ejemplo completo:
  $ python3 prueba_formulas_fernando.py
  
Genera: prueba_formulas_fernando.xlsx con fórmulas inyectadas


VENTAJAS:
===========

✅ Reutiliza fórmulas complejas de Fernando
✅ Conversión automática ODS → Excel
✅ Integración directa en LISP/Python
✅ Fórmulas dinámicas (plantillas)
✅ Mantiene lógica original intacta
✅ Escalable a más hojas y fórmulas


LIMITACIONES ACTUALES:
=======================

• Las referencias a hojas (C111!C5) asumen que existen esas hojas
• Algunas funciones ODS (COM.MICROSOFT.CONCAT) se convierten a CONCAT
• Los estilos/formatos condicionales se inyectan por separado
• Requiere openpyxl >=3.0 para máxima compatibilidad


PRÓXIMOS PASOS:
================

1. Integrar carga de fórmulas directamente desde LISP
2. Agregar validación de referencias de hojas
3. Soporte para más funciones ODS → Excel
4. Caché de fórmulas para mejor rendimiento
5. Documentar casos de uso específicos
"""

if __name__ == '__main__':
    print(__doc__)
