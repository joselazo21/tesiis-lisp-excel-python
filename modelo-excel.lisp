; modelo-excel.lisp - Original working version with xl-* names
(load "codigo-tesis.lisp")

; Core classes xl-*
(defclass* xl-table () (id cols rows contenido headers column-widths formulas))
(defclass* xl-sheet () (name tables formulas fernando-formulas table-borders border-color border-style table-ranges range-styles merge-ranges conditional-format-rules column-widths cell-size header-style))
(defclass* xl-workbook () (name sheets))

; Style classes
(defclass* xl-style () (bold color bg-color align))
(defclass* xl-range-style () (range style))
(defclass* xl-header-style () (bold color bg-color align))
(defclass* xl-formula () (row col value))
(defclass* xl-formula-list () (items))
(defclass* xl-fernando-formula () (cell formula))
(defclass* xl-conditional-rule () (tipo rango formula color))
(defclass* xl-merge-range () (range))
(defclass* xl-column-widths () (pairs))

; Output class
(defclass xl-out () ())
(defparameter xl-py (make-instance 'xl-out))

; Utilities
(defun xl-write (val s)
  (cond ((null val) (format s "None"))
        ((stringp val) (format s "~s" val))
        ((numberp val) (format s "~a" val))
        ((listp val) (progn (format s "[") (loop for i from 0 for x in val do (when (> i 0) (format s ", ")) (xl-write x s)) (format s "]")))
        (t (format s "~s" val))))

(defun xl-write-dict (plist s)
  (format s "{")
  (loop for i from 0 for (k v) on plist by #'cddr when v do (when (> i 0) (format s ", ")) (format s "~s: " k) (xl-write v s))
  (format s "}"))

; generate-code methods
(defmethod generate-code ((st clase-xl-style) (lang xl-out) (stream t))
  (format stream "{")
  (let ((first t))
    (when (bold st) (format stream "~a\"bold\": True" (if first "" ", ")) (setf first nil))
    (when (color st) (format stream "~a\"color\": ~s" (if first "" ", ") (color st)) (setf first nil))
    (when (bg-color st) (format stream "~a\"bg_color\": ~s" (if first "" ", ") (bg-color st)) (setf first nil))
    (when (align st) (format stream "~a\"align\": ~s" (if first "" ", ") (align st)) (setf first nil))
    (format stream "}")))

(defmethod generate-code ((st clase-xl-header-style) (lang xl-out) (stream t))
  (format stream "{")
  (let ((first t))
    (when (bold st) (format stream "~a\"bold\": True" (if first "" ", ")) (setf first nil))
    (when (color st) (format stream "~a\"color\": ~s" (if first "" ", ") (color st)) (setf first nil))
    (when (bg-color st) (format stream "~a\"bg_color\": ~s" (if first "" ", ") (bg-color st)) (setf first nil))
    (when (align st) (format stream "~a\"align\": ~s" (if first "" ", ") (align st)) (setf first nil))
    (format stream "}")))

(defmethod generate-code ((rs clase-xl-range-style) (lang xl-out) (stream t))
  (format stream "{\"range\": ~s, \"style\": " (range rs))
  (generate-code (style rs) lang stream)
  (format stream "}"))

(defmethod generate-code ((rs list) (lang xl-out) (stream t))
  (let ((rng (getf rs :range)) (st (getf rs :style)))
    (when (and rng st)
      (format stream "{\"range\": ~s, \"style\": " rng)
      (cond ((getf st :bg-color) (format stream "{\"bg_color\": ~s}" (getf st :bg-color)))
           ((getf st :bold) (format stream "{\"bold\": True}"))
           (t (format stream "{}")))
      (format stream "}"))))

(defmethod generate-code ((f clase-xl-formula) (lang xl-out) (stream t))
  (format stream "{\"row\": ~a, \"col\": ~a, \"value\": ~s}" (row f) (col f) (value f)))

(defmethod generate-code ((f clase-xl-fernando-formula) (lang xl-out) (stream t))
  (format stream "{\"cell\": ~s, \"formula\": ~s}" (cell f) (formula f)))

(defmethod generate-code ((r clase-xl-conditional-rule) (lang xl-out) (stream t))
  (format stream "{\"tipo\": ~s, \"rango\": ~s, \"formula\": ~s, \"color\": ~s}" (tipo r) (rango r) (formula r) (color r)))

(defmethod generate-code ((m clase-xl-merge-range) (lang xl-out) (stream t))
  (format stream "~s" (range m)))

(defmethod generate-code ((cw clase-xl-column-widths) (lang xl-out) (stream t))
  (format stream "            \"column_widths\": ")
  (xl-write-dict (pairs cw) stream)
  (format stream ",~%"))

(defmethod generate-code ((lst clase-xl-formula-list) (lang xl-out) (stream t))
  (format stream "            \"formulas\": [")
  (loop for i from 0 for f in (items lst) do
       (when (> i 0) (format stream ", "))
       (generate-code f lang stream))
  (format stream "],~%"))

(defmethod generate-code ((tbl clase-xl-table) (lang xl-out) (stream t))
  (let ((con (contenido tbl)) (hdrs (headers tbl)) (cwidths (column-widths tbl)) (frms (formulas tbl)))
    (when con (format stream "        \"data\": ") (xl-write con stream) (format stream ",~%"))
    (when hdrs (format stream "        \"headers\": ") (xl-write hdrs stream) (format stream ",~%"))
    (when cwidths
      (if (typep cwidths 'clase-xl-column-widths)
          (generate-code cwidths lang stream)
          (progn
            (format stream "        \"column_widths\": ")
            (xl-write-dict cwidths stream)
            (format stream ",~%"))))
    (when frms
      (if (typep frms 'clase-xl-formula-list)
          (generate-code frms lang stream)
          (progn
            (format stream "        \"formulas\": [")
            (loop for i from 0 for f in frms do (when (> i 0) (format stream ", ")) (generate-code f lang stream))
            (format stream "],~%"))))))

(defmethod generate-code ((sh clase-xl-sheet) (lang xl-out) (stream t))
  (format stream "        {~%")
  (format stream "            \"title\": ~s,~%" (name sh))
  (dolist (tbl (tables sh)) (generate-code tbl lang stream))
  (let ((frms (formulas sh)))
    (when frms
      (if (typep frms 'clase-xl-formula-list)
          (generate-code frms lang stream)
          (progn
            (format stream "            \"formulas\": [")
            (loop for i from 0 for f in frms do
                 (when (> i 0) (format stream ", "))
                 (generate-code f lang stream))
            (format stream "],~%")))))
  (let ((ffrms (fernando-formulas sh)))
    (when ffrms
      (format stream "            \"fernando_formulas\": [")
      (loop for i from 0 for f in ffrms do
           (when (> i 0) (format stream ", "))
           (generate-code f lang stream))
      (format stream "],~%")))
  (when (table-borders sh) (format stream "            \"table_borders\": True,~%"))
  (let ((bc (border-color sh))) (when bc (format stream "            \"border_color\": ~s,~%" bc)))
  (let ((bs (border-style sh))) (when bs (format stream "            \"border_style\": ~s,~%" bs)))
  (let ((tr (table-ranges sh))) (when tr (format stream "            \"table_ranges\": ") (xl-write tr stream) (format stream ",~%")))
  (let ((rs (range-styles sh)))
    (when rs
      (format stream "            \"range_styles\": [")
      (loop for i from 0 for r in rs do
           (when (> i 0) (format stream ", "))
           (generate-code r lang stream))
      (format stream "],~%")))
  (let ((mr (merge-ranges sh))) (when mr (format stream "            \"merge_ranges\": ") (xl-write mr stream) (format stream ",~%")))
  (let ((cr (conditional-format-rules sh)))
    (when cr
      (format stream "            \"conditional_format_rules\": [")
      (loop for i from 0 for r in cr do
           (when (> i 0) (format stream ", "))
           (generate-code r lang stream))
      (format stream "],~%")))
  (let ((cw (column-widths sh)))
    (when cw
      (if (typep cw 'clase-xl-column-widths)
          (generate-code cw lang stream)
          (progn
            (format stream "            \"column_widths\": ")
            (xl-write-dict cw stream)
            (format stream ",~%")))))
  (let ((cs (cell-size sh))) (when cs (format stream "            \"cell_size\": ~a,~%" cs)))
  (let ((hs (header-style sh))) (when hs (format stream "            \"header_style\": ") (generate-code hs lang stream) (format stream "~%")))
  (format stream "        }~%"))

(defmethod generate-code ((wb clase-xl-workbook) (lang xl-out) (stream t))
  (format stream "#!/usr/bin/env python3~%")
  (format stream "from hoja_con_formulas import generar_excel_personalizado~2%")
  (format stream "config = {~%")
  (format stream "    \"sheets\": [~%")
  (when (sheets wb) (loop for i from 0 for sh in (sheets wb) do (generate-code sh lang stream) (when (< i (1- (length (sheets wb)))) (format stream ",~%")))
  (format stream "    ]~%")
  (format stream "}~2%")
  (format stream "generar_excel_personalizado(config, ~s)~%" (name wb))
  (format stream "~%if __name__=='__main__':~%")
  (format stream "    print('OK: ~a')~%" (name wb)))
)

(defun xl-generate (wb file)
  (with-open-file (s file :direction :output :if-exists :supersede) (generate-code wb xl-py s))
  (format t "Generado: ~a~%" file))

(defun xl-header-style (&key bold color bg-color align)
  (make-instance 'clase-xl-header-style :bold bold :color color :bg-color bg-color :align align))
