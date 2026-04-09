;;; generador_formulas_dinamicas.lisp
;;;
;;; GENERADOR DINÁMICO DE FÓRMULAS PARA EXCEL
;;; ==========================================
;;; 
;;; En lugar de copiar fórmulas hardcodeadas (C111, C112, C113),
;;; este módulo genera fórmulas basadas en los grupos REALES del horario.
;;;
;;; Autor: Sistema de tesis
;;; Fecha: 2026-04-02

(load "codigo-tesis.lisp")

;; =============================================================================
;; EXTRACCIÓN DE GRUPOS DESDE HORARIOS
;; =============================================================================

(defun extraer-grupos-desde-variables ()
  "Extrae todos los grupos definidos en variables_horario.lisp
   Retorna una lista de símbolos como (C111 C121 C211 C311 ...)"
  (let ((grupos '()))
    (do-all-symbols (sym)
      (when (and (boundp sym)
                 (stringp (string sym))
                 (search "*HORARIO-" (string sym)))
        (let* ((sym-name (string sym))
               (grupo-name (subseq sym-name 
                                  (length "*HORARIO-")
                                  (- (length sym-name) 1))))
          (push (intern grupo-name) grupos))))
    (sort grupos #'string<)))

(defun cargar-grupos-desde-archivo (archivo)
  "Carga variables_horario.lisp y extrae grupos automáticamente"
  (load archivo)
  (extraer-grupos-desde-variables))

;; =============================================================================
;; GENERACIÓN DE FÓRMULAS DINÁMICAS
;; =============================================================================

(defun generar-formula-aulas-dinamica (grupos columna-ref &key (locale 'es))
  "Genera fórmula que concatena aulas de múltiples grupos dinámicamente
   
   Parámetros:
     grupos      - Lista de grupos (ej: '(C111 C121 C211))
     columna-ref - Referencia de columna (ej: 'C2' o 'F$2')
     locale      - 'es (español) o 'en (inglés)
   
   Ejemplo de salida (ES):
   =SUSTITUIR(ESPACIOS(CONCATENAR(
     SI(C111!$C$5=C2;C111!$B$1 & \" \"; \"\"); 
     SI(C121!$C$5=C2;C121!$B$1 & \" \"; \"\"); 
     SI(C211!$C$5=C2;C211!$B$1 & \" \"; \"\")))
   ; \" \"; \",\")
   
   La fórmula:
   1. Revisa cada grupo (C111, C121, etc.)
   2. Si el grupo está en esa columna/turno ($C$5=C2)
   3. Concatena el nombre del profesor ($B$1)
   4. Reemplaza espacios por comas al final"
  
  (let* ((funciones (if (eq locale 'es)
                       '(:sustituir "SUSTITUIR"
                         :espacios "ESPACIOS"
                         :concatenar "CONCATENAR"
                         :si "SI")
                       '(:sustituir "SUBSTITUTE"
                         :espacios "TRIM"
                         :concatenar "CONCAT"
                         :si "IF")))
         (fn-sustituir (getf funciones :sustituir))
         (fn-espacios (getf funciones :espacios))
         (fn-concatenar (getf funciones :concatenar))
         (fn-si (getf funciones :si))
         (separador (if (eq locale 'es) ";" ","))
         (partes '()))
    
    ;; Generar SI(...) para cada grupo
    (dolist (grupo grupos)
      (let* ((grupo-str (string grupo))
             (parte (format nil "~a(~a!$C$5=~a~a~a!$B$1 & \" \"~a \"\")"
                           fn-si
                           grupo-str
                           columna-ref
                           separador
                           grupo-str
                           separador)))
        (push parte partes)))
    
    ;; Ensamblar fórmula completa
    (format nil "=~a(~a(~a(~{~a~^~a~}))~a \" \"~a \",\")"
            fn-sustituir
            fn-espacios
            fn-concatenar
            (reverse partes)
            separador
            separador
            separador)))

(defun generar-formulas-para-rango (grupos 
                                   columnas-inicio 
                                   columnas-fin
                                   fila-inicio
                                   fila-fin
                                   &key (locale 'es))
  "Genera todas las fórmulas para un rango de celdas
   
   Retorna una lista de plists:
   '((:celda \"C4\" :formula \"=SUSTITUIR...\")
     (:celda \"D4\" :formula \"=SUSTITUIR...\")
     ...)"
  
  (let ((formulas '())
        (col-num columnas-inicio))
    
    (loop while (<= col-num columnas-fin) do
      (loop for fila from fila-inicio to fila-fin do
        (let* ((col-letra (numero-a-letra-columna col-num))
               (celda (format nil "~a~a" col-letra fila))
               (ref-columna (format nil "~a$2" col-letra))
               (formula (generar-formula-aulas-dinamica grupos ref-columna :locale locale)))
          
          (push (list :celda celda :formula formula) formulas)))
      
      (incf col-num))
    
    (reverse formulas)))

;; =============================================================================
;; UTILIDADES
;; =============================================================================

(defun numero-a-letra-columna (num)
  "Convierte número de columna (1-based) a letra Excel (A, B, C, ..., Z, AA, AB, ...)"
  (let ((resultado ""))
    (loop while (> num 0) do
      (let ((resto (mod (1- num) 26)))
        (setf resultado (concatenate 'string 
                                    (string (code-char (+ 65 resto)))
                                    resultado))
        (setf num (floor (1- num) 26))))
    resultado))

(defun letra-columna-a-numero (letra)
  "Convierte letra de columna Excel (A, B, C, ...) a número (1-based)"
  (let ((num 0)
        (multiplicador 1))
    (loop for i from (1- (length letra)) downto 0 do
      (let ((char (char letra i)))
        (incf num (* multiplicador (- (char-code char) 64)))
        (setf multiplicador (* multiplicador 26))))
    num))

;; =============================================================================
;; FUNCIONES DE EXPORTACIÓN PARA PYTHON
;; =============================================================================

(defun generar-codigo-python-formulas (grupos &key (locale 'es))
  "Genera código Python que crea las fórmulas dinámicamente
   
   Esto genera un diccionario Python con las fórmulas para cada celda"
  
  (with-output-to-string (s)
    (format s "# Fórmulas generadas dinámicamente desde LISP~%")
    (format s "# Grupos incluidos: ~{~a~^, ~}~%~%" grupos)
    (format s "def generar_formulas_aulas():~%")
    (format s "    \"\"\"Genera fórmulas para la hoja Aulas\"\"\"~%")
    (format s "    formulas = []~%~%")
    
    ;; Generar fórmulas para rango C4:L17 (ejemplo)
    ;; C=3, L=12, filas 4-17
    (let ((formulas (generar-formulas-para-rango grupos 3 12 4 17 :locale locale)))
      
      (dolist (formula formulas)
        (format s "    formulas.append({~%")
        (format s "        'cell': '~a',~%" (getf formula :celda))
        (format s "        'formula': ~s~%" (getf formula :formula))
        (format s "    })~%")))
    
    (format s "~%    return formulas~%~%")
    (format s "# Generar fórmulas automáticamente~%")
    (format s "_formulas_aulas_dinamicas = generar_formulas_aulas()~%")))

(defun exportar-formulas-a-python (grupos archivo-salida &key (locale 'es))
  "Exporta las fórmulas dinámicas a un archivo Python"
  (with-open-file (f archivo-salida 
                    :direction :output
                    :if-exists :supersede
                    :if-does-not-exist :create)
    (write-string (generar-codigo-python-formulas grupos :locale locale) f))
  (format t "✅ Fórmulas exportadas a: ~a~%" archivo-salida))

;; =============================================================================
;; FUNCIONES DE TEST
;; =============================================================================

(defun test-generador-formulas ()
  "Prueba el generador de fórmulas con grupos de ejemplo"
  (format t "~%=== TEST GENERADOR DE FÓRMULAS DINÁMICAS ===~%~%")
  
  ;; Test 1: Grupos de Ciencia de la Computación
  (let ((grupos-cc '(C111 C121 C211 C311 C411)))
    (format t "1. Grupos Ciencia Computación: ~{~a~^, ~}~%" grupos-cc)
    (format t "   Fórmula (ES) para C4:~%")
    (format t "   ~a~%~%" (generar-formula-aulas-dinamica grupos-cc "C$2" :locale 'es))
    
    (format t "   Fórmula (EN) para C4:~%")
    (format t "   ~a~%~%" (generar-formula-aulas-dinamica grupos-cc "C$2" :locale 'en)))
  
  ;; Test 2: Grupos de Ciencia de Datos
  (let ((grupos-cd '(D111 D211 D311 D411)))
    (format t "2. Grupos Ciencia de Datos: ~{~a~^, ~}~%" grupos-cd)
    (format t "   Fórmula (ES) para F4:~%")
    (format t "   ~a~%~%" (generar-formula-aulas-dinamica grupos-cd "F$2" :locale 'es)))
  
  ;; Test 3: Conversión número ↔ letra
  (format t "3. Conversiones columna:~%")
  (format t "   1 → ~a~%" (numero-a-letra-columna 1))
  (format t "   26 → ~a~%" (numero-a-letra-columna 26))
  (format t "   27 → ~a~%" (numero-a-letra-columna 27))
  (format t "   A → ~a~%" (letra-columna-a-numero "A"))
  (format t "   Z → ~a~%" (letra-columna-a-numero "Z"))
  (format t "   AA → ~a~%" (letra-columna-a-numero "AA"))
  
  (format t "~%✅ Test completado~%"))

;; =============================================================================
;; INICIALIZACIÓN
;; =============================================================================

(format t "~%✅ Módulo generador_formulas_dinamicas.lisp cargado~%")
(format t "   Funciones disponibles:~%")
(format t "   - (test-generador-formulas)~%")
(format t "   - (generar-formula-aulas-dinamica grupos columna :locale 'es/'en)~%")
(format t "   - (exportar-formulas-a-python grupos archivo :locale 'es/'en)~%")
(format t "   - (extraer-grupos-desde-variables)~%~%")
