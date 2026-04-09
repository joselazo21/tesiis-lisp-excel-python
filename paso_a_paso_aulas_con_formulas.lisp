;;; paso_a_paso_aulas_con_formulas.lisp
;;; 
;;; VERSIÓN MEJORADA: Integra las fórmulas complejas del ODS de Fernando
;;; automáticamente en el flujo de generación Excel
;;;
;;; Cargar con: (load "paso_a_paso_aulas_con_formulas.lisp")

(load "codigo-tesis.lisp")

;; =============================================================================
;; CONFIGURACIÓN DE FÓRMULAS
;; =============================================================================

(defparameter *fernando-formulas-locale* 'es
  "Locale para fórmulas: 'es o 'en")

(defparameter *fernando-formulas-file* 
  (format nil "formulas_fernando_convertidas_~a.json" 
          (string-downcase (symbol-name *fernando-formulas-locale*)))
  "Archivo JSON con fórmulas convertidas")

(defparameter *fernando-formulas* nil
  "Cache de fórmulas cargadas")

;; =============================================================================
;; FUNCIONES DE CARGA DE FÓRMULAS
;; =============================================================================

(defun load-fernando-formulas (&optional (locale 'es))
  "Carga las fórmulas del ODS de Fernando desde JSON
   Retorna lista de fórmulas por hoja"
  (let ((file (format nil "formulas_fernando_convertidas_~a.json" 
                      (string-downcase (symbol-name locale)))))
    (format t "~%📥 Cargando fórmulas de Fernando desde: ~a~%" file)
    
    ;; Intentar cargar el archivo JSON
    (if (probe-file file)
        (progn
          ;; Aquí usaríamos un parser JSON real en producción
          ;; Por ahora, llamamos Python para cargar
          (format t "✅ Archivo encontrado~%")
          file)
        (progn
          (format t "⚠️  Archivo no encontrado: ~a~%" file)
          nil))))

(defun format-python-formulas-list (formulas stream)
  "Formatea lista de fórmulas para Python"
  (format stream "[")
  (loop for formula in formulas
        for i from 0
        do (when (> i 0)
             (format stream ", "))
           (format stream "{~%")
           (format stream "            'cell': '~a',~%" (getf formula :cell))
           (format stream "            'formula': ~s~%" (getf formula :formula))
           (format stream "        }"))
  (format stream "]"))

(defun python-load-fernando-formulas (locale)
  "Genera código Python para cargar fórmulas de Fernando"
  (format nil "
import json

def cargar_formulas_fernando(locale='~a'):
    '''Carga fórmulas convertidas del ODS de Fernando'''
    filename = f'formulas_fernando_convertidas_{locale}.json'
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f'⚠️  {filename} no encontrado')
        return {}

# Cargar fórmulas
_formulas_fernando = cargar_formulas_fernando('~a')
" (string-downcase (symbol-name locale))
  (string-downcase (symbol-name locale))))

;; =============================================================================
;; CLASES EXTENDIDAS
;; =============================================================================

(defclass* hoja-aulas ()
  (lunes
   martes
   miercoles
   jueves
   viernes
   ;; NUEVO: Fórmulas de Fernando por hoja
   formulas-aulas-es
   formulas-aulas-en
   formulas-c111-es
   formulas-c111-en))

(defclass output-python-config () ())
(defparameter python-config (make-instance 'output-python-config))

;; =============================================================================
;; HELPERS PYTHON
;; =============================================================================

(defun format-python-list (lst stream)
  "Formatea lista Lisp como array Python"
  (if (null lst)
      (format stream "None")
      (progn
        (format stream "[")
        (loop for x in lst
              for i from 0
              do
              (if (> i 0) (format stream ", "))
              (if (listp x)
                  (format-python-list x stream)
                  (cond ((numberp x) (format stream "~a" x))
                        ((equal x "") (format stream "\"\""))
                        (t (format stream "\"~a\"" x)))))
        (format stream "]"))))

;; =============================================================================
;; MÉTODOS GENERATE-CODE (extendidos)
;; =============================================================================

(defmethod generate-code ((node clase-tabla)
                          (lang output-python-config)
                          (stream t))
  "Genera código para tabla"
  (if (not (string-equal (id node) "Lunes"))
      (progn
        (format stream "data_aulas.append([])~%")
        (format stream "data_aulas.append(")
        (format-python-list (nombres-columnas node) stream)
        (format stream ")~%")))
        
  (loop for r-idx from 0 to (1- (filas node))
        do
        (format stream "data_aulas.append(")
        (let ((row (list (nth r-idx (nombres-filas node)))))
          (if (contenido-de-la-tabla node)
              (setf row (append row (nth r-idx (contenido-de-la-tabla node))))
              (setf row (append row (make-list (- (columnas node) 1) :initial-element ""))))
          (format-python-list row stream))
        (format stream ")~%")))

(defmethod generate-code ((node clase-hoja-aulas)
                          (lang output-python-config)
                          (stream t))
  "Genera código Python para hoja-aulas CON FÓRMULAS DE FERNANDO"
  
  ;; Importaciones y carga de fórmulas
  (format stream "import sys~%")
  (format stream "import json~%")
  (format stream "from hoja_con_formulas import generar_excel_personalizado~%~%")
  
  ;; Función para cargar fórmulas
  (format stream "~a~%" (python-load-fernando-formulas *fernando-formulas-locale*))
  
  ;; Datos de aulas
  (format stream "~%data_aulas = []~%")
  (format stream "headers_aulas = ")
  (format-python-list (nombres-columnas (lunes node)) stream)
  (format stream "~%~%")
  
  ;; Generar tablas
  (generate-code (lunes node) lang stream)
  (generate-code (martes node) lang stream)
  (generate-code (miercoles node) lang stream)
  (generate-code (jueves node) lang stream)
  (generate-code (viernes node) lang stream)
  
  ;; NUEVO: Inyectar fórmulas de Fernando
  (format stream "~%# ============================================================~%")
  (format stream "# INYECTAR FÓRMULAS COMPLEJAS DEL ODS DE FERNANDO~%")
  (format stream "# ============================================================~%~%")
  
  (format stream "fernando_formulas = _formulas_fernando.get('Aulas', [])~%~%")
  
  (format stream "# Convertir fórmulas a formato esperado~%")
  (format stream "formulas_para_inyectar = []~%")
  (format stream "for f in fernando_formulas[:10]:  # Primeras 10 de demostración~%")
  (format stream "    formulas_para_inyectar.append({~%")
  (format stream "        'cell': f.get('cell'),~%")
  (format stream "        'formula': f.get('excel')~%")
  (format stream "    })~%~%")
  
  ;; Configuración de Excel
  (format stream "config_excel = {~%")
  (format stream "    'sheets': [~%")
  (format stream "        {~%")
  (format stream "            'title': 'Aulas',~%")
  (format stream "            'headers': headers_aulas,~%")
  (format stream "            'data': data_aulas,~%")
  (format stream "            'fernando_formulas': formulas_para_inyectar,  # ✅ NUEVO~%")
  (format stream "            'column_widths': {i: 12 for i in range(1, 12)},~%")
  (format stream "            'header_style': {'bold': True, 'align': 'center', 'bg_color': 'F09E9E'}~%")
  (format stream "        }~%")
  (format stream "    ]~%")
  (format stream "}~%~%")
  
  (format stream "# Generar Excel con fórmulas inyectadas~%")
  (format stream "print('✅ Generando Excel con fórmulas de Fernando...')~%")
  (format stream "generar_excel_personalizado(config_excel, 'Aulas_Con_Formulas_Fernando.xlsx')~%")
  (format stream "print(f'✅ Archivo generado: Aulas_Con_Formulas_Fernando.xlsx')~%")
  (format stream "print(f'✅ Total de fórmulas inyectadas: {len(formulas_para_inyectar)}')~%"))

;; =============================================================================
;; DATOS DE PRUEBA
;; =============================================================================

(defun crear-tabla-dia (dia id-tabla contenido)
  "Crea tabla para un día específico"
  (tabla :id id-tabla
         :filas 6
         :columnas 11
         :nombres-columnas (list dia "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
         :nombres-filas '("1ro" "2do" "3ro" "4to" "5to" "6to")
         :contenido-de-la-tabla contenido))

(defparameter tbl-lunes (crear-tabla-dia "Lunes" "Lunes" 
                       '(("" "" "" "" "" "C111,C112" "" "" "" "C113")
                         ("" "" "" "" "C112" "C113" "" "" "" "")
                         ("" "" "" "" "" "C111,C113" "" "C112" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-martes (crear-tabla-dia "Martes" "Martes" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-miercoles (crear-tabla-dia "Miércoles" "Miércoles" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-jueves (crear-tabla-dia "Jueves" "Jueves" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-viernes (crear-tabla-dia "Viernes" "Viernes" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter mi-hoja-aulas (hoja-aulas :lunes tbl-lunes
                                        :martes tbl-martes
                                        :miercoles tbl-miercoles
                                        :jueves tbl-jueves
                                        :viernes tbl-viernes))

;; =============================================================================
;; GENERAR ARCHIVOS
;; =============================================================================

(format t "~%~%")
(format t "╔═══════════════════════════════════════════════════════════════╗~%")
(format t "║ GENERANDO EXCEL CON FÓRMULAS DE FERNANDO                     ║~%")
(format t "╚═══════════════════════════════════════════════════════════════╝~%~%")

(format t "📋 Configuración:~%")
(format t "   Locale: ~a~%" *fernando-formulas-locale*)
(format t "   Archivo de fórmulas: ~a~%" *fernando-formulas-file*)

(with-open-file (f "ejecutar_hoja_con_formulas_con_fernando.py" 
                    :direction :output :if-exists :supersede)
  (generate-code mi-hoja-aulas python-config f))

(format t "~%✅ Archivo generado: ejecutar_hoja_con_formulas_con_fernando.py~%")

(format t "~%Próximos pasos:~%")
(format t "  1. python3 ejecutar_hoja_con_formulas_con_fernando.py~%")
(format t "  2. Abre: Aulas_Con_Formulas_Fernando.xlsx~%")
(format t "  3. ¡Revisa las fórmulas de Fernando inyectadas!~%~%")
