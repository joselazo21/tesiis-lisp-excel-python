# Guía para Solucionar el Problema de Referencias a Hojas Inexistentes

## El Problema

Las fórmulas de Fernando en la hoja "Aulas" contienen referencias hardcodeadas a hojas específicas:
```excel
=SUSTITUIR(ESPACIOS((SI($C111.$C$9=H$2;$C111.$B$1 & " "; ""); 
                     SI(c112!$c$9=H$2;c112!$b$1 & " "; ""); 
                     SI(c113!$c$9=H$2;c113!$b$1 & " "; ""))); 
          " "; ",")
```

Cuando las hojas `C112` o `C113` no existen, Excel muestra errores `#REF!`.

---

## Soluciones Disponibles

### Solución 1: Desactivar Fórmulas de Fernando (Rápido) ⚡

Edita `replicar_propuesta_ods.lisp` línea 17:

```lisp
;; ANTES:
(defparameter *usar-formulas-fernando* t
  "Si es T, inyecta fórmulas complejas del ODS de Fernando")

;; DESPUÉS:
(defparameter *usar-formulas-fernando* nil
  "Si es T, inyecta fórmulas complejas del ODS de Fernando")
```

Luego regenera:
```bash
sbcl --load replicar_propuesta_ods.lisp --quit
python3 generar_propuesta_desde_lisp.py
```

**Ventajas**: Rápido, sin errores
**Desventajas**: Pierdes las fórmulas complejas de la hoja Aulas

---

### Solución 2: Usar Fórmulas Adaptadas (Recomendado) ✅

He creado `arreglar_formulas_fernando.py` que genera fórmulas dinámicas.

**Paso 1**: Edita el script para especificar las hojas que realmente generas:

```python
# En arreglar_formulas_fernando.py, línea ~35
def obtener_hojas_disponibles_por_tipo() -> Dict[str, List[str]]:
    return {
        'C': ['C111', 'C121', 'C122', 'C211', 'C212', 'C311', 'C312', 'C411', 'C412'],
        'D': ['D111', 'D211', 'D311', 'D411'],
        'M': ['M111', 'M211', 'M311', 'M411']
    }
```

**Paso 2**: Ejecuta el script:
```bash
python3 arreglar_formulas_fernando.py
```

Esto genera `formulas_fernando_adaptadas_es.json` con hojas correctas.

**Paso 3**: Modifica `replicar_propuesta_ods.lisp` línea 791:

```lisp
;; ANTES:
(format stream "    filename = f'formulas_fernando_convertidas_{locale}.json'~%")

;; DESPUÉS:
(format stream "    filename = f'formulas_fernando_adaptadas_{locale}.json'~%")
```

**Paso 4**: Regenera:
```bash
sbcl --load replicar_propuesta_ods.lisp --quit
python3 generar_propuesta_desde_lisp.py
```

---

### Solución 3: Generar Fórmulas Dinámicamente desde Lisp (Mejor) 🏆

Modificar `replicar_propuesta_ods.lisp` para que genere las fórmulas de la hoja Aulas basándose en las hojas que realmente existen.

**Ventaja**: Totalmente automático, sin archivos JSON intermedios
**Desventaja**: Requiere más cambios en el código Lisp

#### Implementación:

1. **Agregar función en Lisp para detectar hojas disponibles:**

```lisp
(defun obtener-hojas-disponibles (libro)
  "Obtiene lista de nombres de hojas disponibles en el libro."
  (loop for grupo in (grupos libro)
        collect (grupo grupo)))

(defun generar-formula-aulas-dinamica (hojas-disponibles dia-columna)
  "Genera fórmula para hoja Aulas que referencia solo hojas existentes."
  (let ((condiciones
          (loop for hoja in hojas-disponibles
                when (string-starts-with hoja "C")  ; Solo hojas tipo C
                collect (format nil "SI(~a!$C$5=~a$2;~a!$B$1 & \" \"; \"\")"
                               hoja dia-columna hoja))))
    (format nil "=SUSTITUIR(ESPACIOS(CONCATENAR(~{~a~^; ~})); \" \"; \",\")"
            condiciones)))
```

2. **Modificar la generación de la hoja Aulas:**

En la función `generate-code` para `clase-hoja-aulas`, en lugar de usar fórmulas del JSON, generar dinámicamente:

```lisp
(defmethod generate-code ((node clase-hoja-aulas)
                         (lang clase-output-python-config)
                         (stream t))
  ;; ... código existente ...
  
  ;; En lugar de cargar formulas_fernando desde JSON:
  (let ((hojas-disponibles (obtener-hojas-disponibles *libro-actual*)))
    (format stream "# Generar fórmulas dinámicas para hoja Aulas~%")
    (format stream "aulas_formulas = []~%")
    (loop for col-letra in '("C" "D" "E" "F" "G" "H" "I" "J" "K" "L")
          for fila from 4 to 50  ; Ajustar según tus necesidades
          do (let ((formula (generar-formula-aulas-dinamica hojas-disponibles col-letra)))
               (format stream "aulas_formulas.append({'row': ~a, 'col': ~a, 'value': ~s})~%"
                      fila
                      (1+ (- (char-code (char col-letra 0)) (char-code #\A)))
                      formula)))))
```

---

## Prueba de la Solución

Después de aplicar cualquier solución, verifica:

```bash
python3 << 'EOF'
from openpyxl import load_workbook

wb = load_workbook('propuesta_horarios_desde_lisp.xlsx')

if 'Aulas' in wb.sheetnames:
    ws = wb['Aulas']
    print("Verificando fórmulas en hoja Aulas...")
    
    # Buscar fórmulas con referencias a hojas
    errores = []
    for row in ws.iter_rows(min_row=4, max_row=20, min_col=3, max_col=12):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and '!' in cell.value:
                print(f"\n{cell.coordinate}: {cell.value[:80]}...")
                
                # Verificar si las hojas referenciadas existen
                import re
                hojas_ref = re.findall(r'([A-Z]\d{3})!', cell.value)
                for hoja_ref in hojas_ref:
                    if hoja_ref not in wb.sheetnames:
                        errores.append(f"{cell.coordinate} referencia hoja inexistente: {hoja_ref}")
    
    if errores:
        print("\n⚠️  ERRORES ENCONTRADOS:")
        for error in errores:
            print(f"  - {error}")
    else:
        print("\n✅ Todas las referencias a hojas son válidas")
else:
    print("⚠️  Hoja 'Aulas' no encontrada")
EOF
```

---

## Recomendación Final

1. **Inmediato**: Usa **Solución 1** (desactivar) si necesitas resultados rápidos sin errores
2. **Corto plazo**: Implementa **Solución 2** (fórmulas adaptadas) para mantener funcionalidad
3. **Largo plazo**: Implementa **Solución 3** (generación dinámica desde Lisp) para sistema robusto

---

## Scripts Disponibles

```bash
# Ver qué hojas existen en el Excel generado
python3 -c "from openpyxl import load_workbook; print(load_workbook('propuesta_horarios_desde_lisp.xlsx').sheetnames)"

# Adaptar fórmulas de Fernando
python3 arreglar_formulas_fernando.py

# Regenerar todo el sistema
sbcl --load replicar_propuesta_ods.lisp --quit && python3 generar_propuesta_desde_lisp.py
```
