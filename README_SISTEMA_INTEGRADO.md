# Sistema Integrado de Horarios con Fórmulas Dinámicas

## 🎯 Descripción

Sistema completo que genera archivos Excel con fórmulas dinámicas para horarios universitarios. Las fórmulas se generan **automáticamente** basándose en los grupos reales definidos en `variables_horario.lisp`, eliminando hardcodeos y adaptándose a cualquier configuración.

## ✨ Características Principales

### 1. **Generación Dinámica de Fórmulas**
- Las fórmulas NO están hardcodeadas (adiós C111, C112, C113 fijos)
- Se generan basándose en los grupos REALES del horario
- Soporte para español e inglés
- Extracción automática de grupos desde `variables_horario.lisp`

### 2. **Hojas Automáticas**
- **Hoja Aulas**: Con fórmulas que concatenan datos de todos los grupos
- **Hojas por Grupo**: Una hoja por cada grupo (C111, C121, D111, etc.)
- Referencias correctas entre hojas
- Sin errores #NOMBRE

### 3. **Integración Completa**
- LISP genera el código Python
- Python crea el Excel con OpenPyxl
- Fórmulas funcionan inmediatamente al abrir el archivo

## 📁 Archivos del Sistema

### Archivos LISP (Núcleo)

```
sistema_horarios_integrado.lisp    - Sistema principal integrado
generador_formulas_dinamicas.lisp  - Generador de fórmulas dinámicas
codigo-tesis.lisp                  - Clases base y utilidades
variables_horario.lisp             - Datos de horarios por grupo
```

### Archivos Python (Generados)

```
ejecutor_aulas.py                  - Generado por LISP, crea el Excel
```

### Archivos Excel (Salida)

```
Aulas_Con_Formulas_Fernando.xlsx   - Excel final con fórmulas
```

## 🚀 Uso

### Opción 1: Desde LISP (Recomendado)

```bash
sbcl
```

```lisp
(load "sistema_horarios_integrado.lisp")
(generar-excel-completo)
```

Esto:
1. Extrae todos los grupos de `variables_horario.lisp`
2. Genera fórmulas dinámicas para cada grupo
3. Crea `ejecutor_aulas.py`
4. Te indica que ejecutes: `python3 ejecutor_aulas.py`

### Opción 2: Directo desde Shell

```bash
sbcl --eval "(load \"sistema_horarios_integrado.lisp\")" \
     --eval "(generar-excel-completo)" \
     --eval "(quit)"

python3 ejecutor_aulas.py
```

### Opción 3: Un solo comando

```bash
./ru.sh
```

## 📊 Estructura del Excel Generado

### Hoja "Aulas"

```
| Turno  | Aula 1 | Aula 2 | Aula 3 | ... | Lab |
|--------|--------|--------|--------|-----|-----|
| Lunes  |        |        |        |     |     |
| 1ro    | =FORM  | =FORM  | =FORM  | ... | ... |
| 2do    | =FORM  | =FORM  | =FORM  | ... | ... |
| ...    |        |        |        |     |     |
```

**Fórmulas** (ejemplo en español):
```excel
=SUSTITUIR(ESPACIOS(CONCATENAR(
  SI(C111!$C$5=C$2;C111!$B$1 & " "; "")
  SI(C121!$C$5=C$2;C121!$B$1 & " "; "")
  SI(C211!$C$5=C$2;C211!$B$1 & " "; "")
  SI(C311!$C$5=C$2;C311!$B$1 & " "; "")
  ... (todos los grupos reales)
)); " "; ",")
```

**Fórmulas** (ejemplo en inglés):
```excel
=SUBSTITUTE(TRIM(CONCAT(
  IF(C111!$C$5=C$2,C111!$B$1 & " ", "")
  IF(C121!$C$5=C$2,C121!$B$1 & " ", "")
  IF(C211!$C$5=C$2,C211!$B$1 & " ", "")
  ... (todos los grupos reales)
)), " ", ",")
```

### Hojas de Grupos (C111, C121, D111, etc.)

```
| Turno | Profesor    | Aula   | Asignatura | Horario    |
|-------|-------------|--------|------------|------------|
| 1     | Profesor X  | Aula 6 | Física     | 8:30-10:00 |
```

## 🔧 Configuración

### Cambiar Idioma de Fórmulas

En `sistema_horarios_integrado.lisp`:

```lisp
(defparameter *locale-formulas* 'es)  ; 'es para español, 'en para inglés
```

### Cambiar Archivo de Salida

```lisp
(defparameter *archivo-salida* "Mi_Horario.xlsx")
```

### Agregar Nuevos Grupos

Solo agrega el grupo a `variables_horario.lisp`:

```lisp
(defparameter *horario-c511*
  '(("F" "AM I" "P" "L" "")
    ("Aula 3" "Aula 3" "Aula 3" "Aula 3" "")
    ...))
```

El sistema **detectará automáticamente** el nuevo grupo y generará las fórmulas correspondientes.

## 🧪 Testing

### Probar Generador de Fórmulas

```bash
sbcl --eval "(load \"generador_formulas_dinamicas.lisp\")" \
     --eval "(test-generador-formulas)" \
     --eval "(quit)"
```

Esto muestra:
- Fórmulas generadas para diferentes grupos
- Conversiones de columnas (número ↔ letra)
- Validaciones de formato

### Verificar Excel Generado

```bash
python3 << 'EOF'
import openpyxl
wb = openpyxl.load_workbook('Aulas_Con_Formulas_Fernando.xlsx')
print(f"Hojas: {wb.sheetnames}")
print(f"Fórmula C4: {wb['Aulas']['C4'].value}")
EOF
```

## 🆚 Comparación: Antes vs Ahora

### ❌ Antes (Hardcodeado)

```lisp
;; Fórmulas fijas con C111, C112, C113
(defparameter *formula-fija*
  "=SUSTITUIR(...SI(C111!..SI(C112!..SI(C113!...")
```

**Problemas:**
- Solo funciona con C111, C112, C113
- Si agregas C411, tienes que editar manualmente
- Error #NOMBRE si las hojas no existen
- No escalable

### ✅ Ahora (Dinámico)

```lisp
;; Fórmulas generadas desde grupos reales
(generar-formula-aulas-dinamica 
  (obtener-todos-los-grupos)  ; ← Detecta automáticamente
  "C$2" 
  :locale 'es)
```

**Ventajas:**
- Funciona con CUALQUIER conjunto de grupos
- Agregar grupos = automático
- Siempre crea las hojas necesarias
- 100% escalable

## 📚 API Principal

### Funciones LISP

```lisp
;; Generar sistema completo
(generar-excel-completo &key (locale 'es))

;; Obtener grupos automáticamente
(obtener-todos-los-grupos)  ; => (C111 C121 C211 ...)

;; Generar fórmula para columna específica
(generar-formula-aulas-dinamica 
  '(C111 C121 D111)  ; grupos
  "F$2"              ; columna
  :locale 'es)

;; Exportar fórmulas a Python
(exportar-formulas-a-python 
  '(C111 C121 D111) 
  "formulas.py" 
  :locale 'en)
```

## 🐛 Solución de Problemas

### Error: "No se encontraron grupos"

**Causa:** `variables_horario.lisp` no está cargado

**Solución:**
```lisp
(load "variables_horario.lisp")
(obtener-todos-los-grupos)  ; Debe retornar lista de grupos
```

### Error: #NOMBRE en Excel

**Causa:** Hojas de grupos no existen

**Solución:** El sistema ahora crea automáticamente todas las hojas. Si ves este error, regenera con:
```bash
sbcl --eval "(load \"sistema_horarios_integrado.lisp\")" \
     --eval "(generar-excel-completo)" \
     --eval "(quit)"
python3 ejecutor_aulas.py
```

### Fórmulas en idioma incorrecto

**Causa:** Locale mal configurado

**Solución:**
```lisp
;; Para español
(generar-excel-completo :locale 'es)

;; Para inglés
(generar-excel-completo :locale 'en)
```

## 📈 Extensiones Futuras

1. **Cargar datos reales de horarios** en cada hoja de grupo
2. **Formato condicional** basado en colisiones de aulas
3. **Validación automática** de conflictos
4. **Exportación a múltiples formatos** (ODS, PDF, HTML)
5. **Interfaz web** para configuración

## 📝 Changelog

### 2026-04-02 - v2.0 (Sistema Integrado)
- ✅ Generación dinámica de fórmulas basada en grupos reales
- ✅ Eliminación de hardcodeos (C111, C112, C113)
- ✅ Detección automática de grupos desde variables_horario.lisp
- ✅ Creación automática de hojas para cada grupo
- ✅ Soporte para español e inglés
- ✅ Sistema completamente integrado LISP → Python → Excel

### 2026-03-xx - v1.0 (Versión Original)
- Fórmulas hardcodeadas
- Hojas fijas
- Requería copia manual de hojas

## 👥 Autor

Sistema de tesis - Universidad de La Habana

## 📄 Licencia

Proyecto académico de tesis
