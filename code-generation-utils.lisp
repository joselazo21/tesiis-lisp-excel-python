;; =============================================================================
;; code-generation-utils.lisp — Utilidades del backend Excel/Python del DSL
;;
;; RESPONSABILIDAD: funciones auxiliares y variables de configuración usadas
;; por generate-code-direct.lisp para traducir nodos AST a Excel/Python.
;; Incluye:
;;   - Helpers de escritura (xl-write, escape-python-string, col->letter)
;;   - Variables de configuración del backend (*param-cells*, *turno-dia-col-map*,
;;     *sheet-env*, *table-first-rows*, *source-table-schemas*)
;;   - Construcción de mapas de columnas y expansión de tablas
;;   - Resolución de plantillas de rangos
;;   - Helpers de estilos por tipo y por reglas condicionales
;;   - Funciones de generación y ejecución (xl-generate, xl-run-generated)
;; =============================================================================
(load "ast-def.lisp")

; =====================================================================
; UTILITIES
; =====================================================================

(defun xl-write (val s)
  (cond ((null val) (format s "None"))
        ((stringp val) (format s "~s" val))
        ((numberp val) (format s "~a" val))
        ((listp val) (progn (format s "[") (loop for i from 0 for x in val do (when (> i 0) (format s ", ")) (xl-write x s)) (format s "]")))
        (t (format s "~s" val))))

(defun escape-python-string (s)
  (with-output-to-string (out)
    (loop for c across s
          do (if (char= c #\") (format out "\\\"") (write-char c out)))))

(defun col->letter (n)
  (with-output-to-string (s)
    (loop while (> n 0)
          do (multiple-value-bind (q r) (floor (1- n) 26)
               (setf n q)
               (write-char (code-char (+ 65 r)) s)))))

(defvar *param-cells* nil
  "Alist of (param-name . cell-ref) for compile-excel-formula")

(defun build-col-map (tbl)
  (loop for name in (col-names tbl) for i from 1
        collect (cons name (col->letter i))))

(defun data-col-count (tbl)
  (let ((computed-names (mapcar #'car (computed tbl))))
    (count-if-not (lambda (n) (member n computed-names)) (col-names tbl))))

(defun data-col-names (tbl)
  (let ((computed-names (mapcar #'car (computed tbl))))
    (remove-if (lambda (n) (member n computed-names)) (col-names tbl))))

(defun expand-data-row (row num-cols cell-height)
  (let ((physical-rows (loop repeat cell-height collect (make-list num-cols :initial-element ""))))
    (loop for col-idx from 0 below num-cols
          for val = (if (< col-idx (length row)) (nth col-idx row) "")
          do (if (listp val)
                 (loop for r from 0 below cell-height
                       do (setf (nth col-idx (nth r physical-rows))
                                (or (nth r val) "")))
                 (setf (nth col-idx (first physical-rows)) val)))
    physical-rows))

(defun expand-table-content (tbl)
  (let* ((con (contenido tbl))
         (num-cols (length (col-names tbl)))
         (cell-height (max 1 (or (cell-height tbl) 1)))
         (explicit-first (table-first-row tbl)))
    (if (<= cell-height 1)
        con
        (let* ((prefix-count (if explicit-first (max 0 (1- explicit-first)) 0))
               (prefix (subseq con 0 (min prefix-count (length con))))
               (logical-data (nthcdr (min prefix-count (length con)) con))
               (expanded-data
                 (loop for row in logical-data
                       append (expand-data-row row num-cols cell-height))))
          (append prefix expanded-data)))))

;; =====================================================================
; RANGE TEMPLATE RESOLUTION
; =====================================================================

(defun replace-all (string old new)
  (let ((result string))
    (loop for pos = (search old result)
          while pos
          do (setf result (concatenate 'string
                                       (subseq result 0 pos)
                                       new
                                       (subseq result (+ pos (length old))))))
    result))

(defun resolve-range-template (template first-row last-row row-num col-map)
  (let ((result template))
    (setf result (replace-all result "{first-row}" (write-to-string first-row)))
    (setf result (replace-all result "{last-row}" (write-to-string last-row)))
    (setf result (replace-all result "{row-num}" (write-to-string row-num)))
    (loop for (name . letter) in col-map
          do (setf result (replace-all result
                                       (format nil "{col:~(~a~)}" name)
                                       letter)))
    result))

;; =============================================================================
;; INFRAESTRUCTURA PARA BÚSQUEDAS CROSS-SHEET DINÁMICAS
;;
;; Estas tres piezas colaboran para compilar collect-over / cross-cell:
;;
;;  *turno-dia-col-map*  → conocimiento de la estructura de turno-table
;;  *sheet-env*          → entorno dinámico de ligaduras de variables de hoja
;;  resolve-cross-col    → símbolo DSL → letra Excel
;; =============================================================================

;; Posición de cada columna de turno-table dentro de la hoja de grupo.
;; Este mapa ES conocimiento del backend: sabe que turno-table tiene
;; ((turno "") (lun "") (mar "") (mie "") (jue "") (vie ""))
;; y que en el Excel generado eso ocupa columnas A..F (1-indexed).
;; Si la definición de turno-table cambiara, este mapa debe actualizarse.
(defparameter *turno-dia-col-map*
  '((turno . 1) (lun . 2) (mar . 3) (mie . 4) (jue . 5) (vie . 6)))

;; Resuelve un símbolo de columna DSL a la letra Excel correspondiente.
;; Primero busca en *turno-dia-col-map*; si no está, usa el nombre del símbolo.
;; Ejemplo: LUN → "B", VIE → "F", A → "A" (fallback directo).
(defun resolve-cross-col (col-sym)
  (let ((pos (cdr (assoc col-sym *turno-dia-col-map* :test #'string-equal))))
    (if pos (col->letter pos) (symbol-name col-sym))))

;; Entorno dinámico de ligaduras de variables de hoja.
;; Estructura: alist ((SYM . "NombreHoja") ...)
;; collect-over amplía este entorno en cada iteración para que cross-cell
;; pueda resolver a qué hoja concreta hace referencia la variable.
;; Es nil fuera de un collect-over.
(defparameter *sheet-env* nil)

;; Primera fila de datos de cada tabla nombrada.
;; Responsabilidad del backend: el DSL no debe conocer la disposición física.
(defparameter *table-first-rows*
  '((turno-table     . 4)
    (stats-table     . 4)
    (aulas-table     . 4)
    (aulas-dia-table . 2)))

;; Devuelve el first-row efectivo de una tabla: busca por id en el mapa,
;; y si no encuentra (tabla inline sin id) usa el default (2, ó 3 con params).
(defun table-first-row (tbl &optional num-params)
  (or (cdr (assoc (id tbl) *table-first-rows*))
      (+ 2 (if (and num-params (> num-params 0)) 1 0))))

;; Propiedades de las tablas que pueden actuar como fuente en source-row.
;; Este es el lugar correcto para estos valores: conocimiento del backend,
;; no del DSL ni de los nodos AST.
(defparameter *source-table-schemas*
  '((turno-table . (:first-row 4 :cell-height 2))))

; =====================================================================
; HELPER — Colores por tipo-programa
; =====================================================================

(defun color-for-tipo (tipo-val)
  (let ((s (and (stringp tipo-val) (string-upcase tipo-val))))
    (cond
      ((null s)                              "#D9D9D9")
      ((search "INFORM" s :test #'char=)     "#E6B8AF")
      ((search "MUSIC" s :test #'char=)      "#A9D18E")
      ((search "CULTUR" s :test #'char=)     "#B4C7E7")
      ((search "DEPOR" s :test #'char=)      "#FFD966")
      ((search "ENTRET" s :test #'char=)     "#D5A6BD")
      ((search "EDUCA" s :test #'char=)      "#C5E0B4")
      (t                                     "#D9D9D9"))))

; =====================================================================
; HELPER — Range styles por tipo
; =====================================================================

(defun collect-range-styles-from-tipo (tbl first-row last-row)
  (let* ((dnames (data-col-names tbl))
         (tipo-idx (position 'tipo dnames :test #'string-equal))
         (raw-data (contenido tbl))
         (result ())
         current-color current-start)
    (unless tipo-idx (return-from collect-range-styles-from-tipo nil))
    (let ((stylable-cols
           (loop for col-name in '("hora-inicio" "hora-terminacion" "tipo-calc")
                 for pos = (position col-name (col-names tbl) :test #'string-equal)
                 when pos collect (1+ pos))))
      (unless stylable-cols (return-from collect-range-styles-from-tipo nil))
      (loop for i from 0 below (length raw-data)
            for row-num = (+ first-row i)
            for tipo-val = (let ((row (nth i raw-data)))
                            (if (and (listp row) (< tipo-idx (length row)))
                                (nth tipo-idx row)
                                nil))
            for color = (color-for-tipo tipo-val)
            do (cond
                 ((null current-color)
                  (setf current-color color current-start row-num))
                 ((string/= color current-color)
                  (loop for col-idx in stylable-cols
                        for col-letter = (col->letter col-idx)
                        do (push (cons (format nil "~a~a:~a~a" col-letter current-start col-letter (1- row-num)) current-color) result))
                  (setf current-color color current-start row-num))))
      (when current-color
        (loop for col-idx in stylable-cols
              for col-letter = (col->letter col-idx)
              do (push (cons (format nil "~a~a:~a~a" col-letter current-start col-letter last-row) current-color) result)))
      (nreverse result))))

; =====================================================================
; HELPER — Range styles por reglas condicionales
; =====================================================================

(defun row-value (raw-data row-idx col-idx)
  (let ((row (nth row-idx raw-data)))
    (and (listp row) (< col-idx (length row)) (nth col-idx row))))

(defun build-col-pos-map (tbl)
  (loop for name in (col-names tbl) for idx from 0
        collect (cons name idx)))

;; Evalúa un árbol de condición contra raw-data en row-idx.
;; Devuelve (values truthy match-count total-conditions)
;; match-count / total-conditions permiten clasificar :neighbor vs :both
;; en condiciones top-level _or.
(defun eval-condition-tree (node raw-data row-idx col-pos-map)
  (typecase node
    (clase-xl-expr-equals
     (let* ((a (eval-condition-tree (a node) raw-data row-idx col-pos-map))
            (b (eval-condition-tree (b node) raw-data row-idx col-pos-map))
            (match (and a b (equalp a b))))
       (values match (if match 1 0) 1)))
    (clase-xl-expr-different
     (let* ((a (eval-condition-tree (a node) raw-data row-idx col-pos-map))
            (b (eval-condition-tree (b node) raw-data row-idx col-pos-map))
            (match (or (null a) (null b) (not (equalp a b)))))
       (values match (if match 1 0) 1)))
    (clase-xl-expr-and
     (multiple-value-bind (a-match a-cnt a-tot) (eval-condition-tree (a node) raw-data row-idx col-pos-map)
       (multiple-value-bind (b-match b-cnt b-tot) (eval-condition-tree (b node) raw-data row-idx col-pos-map)
         (values (and a-match b-match)
                 (+ a-cnt b-cnt)
                 (+ a-tot b-tot)))))
    (clase-xl-expr-or
     (multiple-value-bind (a-match a-cnt a-tot) (eval-condition-tree (a node) raw-data row-idx col-pos-map)
       (multiple-value-bind (b-match b-cnt b-tot) (eval-condition-tree (b node) raw-data row-idx col-pos-map)
         (values (or a-match b-match)
                 (+ a-cnt b-cnt)
                 (+ a-tot b-tot)))))
    (clase-xl-expr-column-ref
     (let ((col-idx (cdr (assoc (name node) col-pos-map))))
       (if col-idx
           (values (row-value raw-data row-idx col-idx) 0 0)
           (values nil 0 0))))
    (clase-xl-expr-previous-row
     (let ((prev (1- row-idx)))
       (if (or (< prev 0) (>= prev (length raw-data)))
           (values nil 0 0)
           (eval-condition-tree (expr node) raw-data prev col-pos-map))))
    (clase-xl-expr-next-row
     (let ((next (1+ row-idx)))
       (if (or (< next 0) (>= next (length raw-data)))
           (values nil 0 0)
           (eval-condition-tree (expr node) raw-data next col-pos-map))))
    (t (values nil 0 0))))

(defun classify-row (truthy match-count total-count)
  (if (not truthy) nil
      (if (and (> total-count 0) (= match-count total-count))
          :both
          :neighbor)))

(defun collect-range-styles-from-rules (tbl rules first-row last-row)
  (let* ((raw-data (contenido tbl))
         (col-pos-map (build-col-pos-map tbl))
         (result ())
         (neighbor-color "#FFA500")
         (both-color "#FF0000"))
    (dolist (rule rules result)
      (let* ((targets (or (target-columns rule)
                          '("hora-inicio" "hora-terminacion" "tipo-calc")))
             (target-idxs (loop for name in targets
                                for idx = (position name (col-names tbl) :test #'string-equal)
                                when idx collect idx)))
        (when target-idxs
          (let ((status (make-array (length raw-data) :initial-element nil)))
            (loop for i from 0 below (length raw-data)
                  do (multiple-value-bind (truthy match-cnt total-cnt)
                         (eval-condition-tree (rule-condition rule) raw-data i col-pos-map)
                       (setf (aref status i)
                             (classify-row truthy match-cnt total-cnt))))
            (flet ((add-range (start end color)
                     (loop for ti in target-idxs
                           for tl = (col->letter (1+ ti))
                           do (push (cons (format nil "~a~a:~a~a" tl start tl end) color) result))))
              (let ((cur-type nil) (cur-start nil))
                (loop for i from 0 below (length raw-data)
                      for rn = (+ first-row i)
                      for st = (aref status i)
                      do (cond
                           ((and st (null cur-type))
                            (setf cur-type st cur-start rn))
                           ((and st (not (eq st cur-type)))
                            (add-range cur-start (1- rn)
                                        (if (eq cur-type :both) both-color neighbor-color))
                            (setf cur-type st cur-start rn))
                           ((null st)
                            (when cur-type
                              (add-range cur-start (1- rn)
                                          (if (eq cur-type :both) both-color neighbor-color))
                              (setf cur-type nil cur-start nil)))))
                (when cur-type
                  (add-range cur-start last-row
                             (if (eq cur-type :both) both-color neighbor-color)))))))))))

; =====================================================================
; FUNCIONES DE GENERACIÓN
; =====================================================================

(defun xl-generate (wb file)
  (with-open-file (s file :direction :output :if-exists :supersede)
    (generate-code wb xl-py s))
  (format t "Generado: ~a~%" file))

(defun xl-run-generated (python-file)
  (format t "Ejecutando python3 ~a...~%" python-file)
  #+sbcl
  (sb-ext:run-program "/bin/sh" (list "-c" (format nil "python3 ~a" python-file))
                      :output *standard-output* :error *error-output*)
  #+clisp
  (shell (format nil "python3 ~a" python-file)))

(provide "code-generation-utils")
