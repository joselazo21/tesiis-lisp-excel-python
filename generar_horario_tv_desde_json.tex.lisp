;; generar_horario_tv_desde_json.lisp
;; Convierte variables TV (desde JSON) en clases Lisp y genera un script Python
;; que llama la API parametrica de hoja_con_formulas.py.

(load "codigo-tesis.lisp")

(defparameter *tv-tex-intervalo-minutos* 15)
(defparameter *tv-tex-alto-base-ex* 2.8)

(defun tv-intervalo-dia (dia-config &optional (default *tv-tex-intervalo-minutos*))
  (or (getf dia-config :intervalo-minutos) default))

(defun tv-safe (value &optional (default ""))
  (if value value default))

(defun tv-dia-display (dia)
  (let ((value (string-downcase (tv-safe dia "dia"))))
    (cond
      ((string= value "lunes") "Lunes")
      ((string= value "martes") "Martes")
      ((or (string= value "miercoles") (string= value "miércoles")) "Miércoles")
      ((string= value "jueves") "Jueves")
      ((string= value "viernes") "Viernes")
      ((or (string= value "sabado") (string= value "sábado")) "Sábado")
      ((string= value "domingo") "Domingo")
      (t (string-capitalize value)))))

(defun tv-parse-hora-a-minutos (hora-str)
  (let* ((sep (position #\: hora-str))
         (h (parse-integer (subseq hora-str 0 sep)))
         (m (parse-integer (subseq hora-str (1+ sep)))))
    (+ (* h 60) m)))

(defun tv-minutos-a-hora (minutos)
  (let* ((total (mod minutos 1440))
         (h (floor total 60))
         (m (mod total 60)))
    (format nil "~2,'0d:~2,'0d" h m)))

(defun tv-calcular-intervalos (duracion intervalo)
  (max 1 (ceiling duracion intervalo)))

(defun tv-tex-escape (text)
  (with-output-to-string (out)
    (loop for ch across (princ-to-string (or text ""))
          do
            (case ch
              (#\\ (princ "\\textbackslash{}" out))
              (#\& (princ "\\&" out))
              (#\% (princ "\\%" out))
              (#\$ (princ "\\$" out))
              (#\# (princ "\\#" out))
              (#\_ (princ "\\_" out))
              (#\{ (princ "\\{" out))
              (#\} (princ "\\}" out))
              (#\~ (princ "\\textasciitilde{}" out))
              (#\^ (princ "\\textasciicircum{}" out))
              (t (write-char ch out))))))

(defun programa-tv-a-fila (programa)
  (list
    (tv-safe (getf programa :hora-inicio) "")
    (tv-safe (getf programa :hora-final) "")
    (or (getf programa :duracion) 0)
    (tv-safe (getf programa :nombre) "")
    (tv-safe (getf programa :tipo-programa) "")
    (tv-safe (getf programa :tipo-publico) "")))

(defun crear-tabla-tv-dia (dia-config)
  (let* ((dia (tv-dia-display (getf dia-config :dia)))
         (programas (or (getf dia-config :programas) '()))
         (filas (mapcar #'programa-tv-a-fila programas)))
    (tabla :id (format nil "tv-~a" dia)
           :nombres-columnas '("Inicio" "Fin" "Duración (min)" "Programa" "Tipo" "Público")
           :contenido-de-la-tabla filas)))

(defun crear-hoja-tv-dia (dia-config &optional (intervalo-default *tv-tex-intervalo-minutos*))
  (let ((dia (tv-dia-display (getf dia-config :dia))))
    (hoja :grupo dia
          :horario (crear-tabla-tv-dia dia-config)
          :intervalo-minutos (tv-intervalo-dia dia-config intervalo-default)
          :asignaturas (tabla :id (format nil "meta-~a" dia)
                              :contenido-de-la-tabla '(("" ""))))))

(defun crear-libro-tv-desde-planificacion (planificacion &optional (intervalo-default *tv-tex-intervalo-minutos*))
  (libro :hojas (mapcar (lambda (dia-config)
                          (crear-hoja-tv-dia dia-config intervalo-default))
                        planificacion)))

(defun escribir-programa-tv-python (programa stream)
  (format stream "        {'nombre': ~s, 'duracion': ~a, 'hora_inicio': ~s, 'hora_final': ~s, 'tipo_programa': ~s, 'tipo_publico': ~s}"
          (tv-safe (getf programa :nombre) "")
          (or (getf programa :duracion) 0)
          (tv-safe (getf programa :hora-inicio) "")
          (tv-safe (getf programa :hora-final) "")
          (tv-safe (getf programa :tipo-programa) "")
          (tv-safe (getf programa :tipo-publico) "")))

(defun escribir-dia-tv-python (dia-config stream &optional (last-day nil))
  (let* ((dia (string-downcase (tv-safe (getf dia-config :dia) "dia")))
         (programas (or (getf dia-config :programas) '())))
    (format stream "    {'dia': ~s, 'programas': [~%" dia)
    (loop for programa in programas
          for idx from 0
          for is-last = (= idx (1- (length programas)))
          do
            (escribir-programa-tv-python programa stream)
            (format stream "~a~%" (if is-last "" ",")))
    (format stream "    ]}~a~%" (if last-day "" ","))))

(defun generar-script-python-tv (nombre-canal planificacion output-python-file output-excel-file)
  (with-open-file (stream output-python-file
                          :direction :output
                          :if-exists :supersede)
    (format stream "from hoja_con_formulas import generar_excel_horario_tv_desde_parametros~2%")
    (format stream "nombre_canal = ~s~%" nombre-canal)
    (format stream "planificacion_semanal = [~%")
    (loop for dia-config in planificacion
          for idx from 0
          for last-day = (= idx (1- (length planificacion)))
          do (escribir-dia-tv-python dia-config stream last-day))
    (format stream "]~2%")
    (format stream "generar_excel_horario_tv_desde_parametros(~%")
    (format stream "    filename=~s,~%" output-excel-file)
    (format stream "    nombre_canal=nombre_canal,~%")
    (format stream "    planificacion_semanal=planificacion_semanal,~%")
    (format stream "    incluir_resumen=True~%")
    (format stream ")~%")))

(defun tv-rango-intervalo (inicio-min intervalo)
  (format nil "~a--~a"
          (tv-minutos-a-hora inicio-min)
          (tv-minutos-a-hora (+ inicio-min intervalo))))

(defun tv-escribir-programa-escalado-tex (programa stream intervalo alto-base-ex)
  (let* ((inicio (tv-safe (getf programa :hora-inicio) ""))
         (duracion (or (getf programa :duracion) 0))
         (nombre (tv-safe (getf programa :nombre) ""))
         (inicio-min (if (> (length inicio) 0) (tv-parse-hora-a-minutos inicio) 0))
         (unidades (tv-calcular-intervalos duracion intervalo)))
    (loop for u from 0 below unidades
          for tramo-inicio = (+ inicio-min (* u intervalo))
          for tramo = (tv-rango-intervalo tramo-inicio intervalo)
          do
            (if (= u 0)
                (format stream
                        "~a & \\multirow{~a}{=}{\\parbox[c][~,1fex][c]{\\linewidth}{\\centering\\textbf{~a}}} \\\\~%"
                        (tv-tex-escape tramo)
                        unidades
                        (* alto-base-ex unidades)
                        (tv-tex-escape nombre))
                (format stream "~a &  \\\\~%" (tv-tex-escape tramo)))
            (if (< u (1- unidades))
                (format stream "\\cline{1-1}~%")
                (format stream "\\hline~%")))))

(defun tv-escribir-tabla-dia-tex (dia-config stream intervalo alto-base-ex)
  (let* ((dia (tv-dia-display (getf dia-config :dia)))
         (programas (or (getf dia-config :programas) '())))
    (format stream "\\section*{~a}~%" (tv-tex-escape dia))
    (format stream "\\begin{center}~%")
    (format stream "\\begin{adjustbox}{max totalsize={\\textwidth}{0.82\\textheight},center}~%")
    (format stream "\\begin{tabular}{|p{3.6cm}|p{11cm}|}~%")
    (format stream "\\hline~%")
    (format stream "\\textbf{Intervalo} & \\textbf{Programa} \\\\~%")
    (format stream "\\hline~%")
    (loop for programa in programas
          do (tv-escribir-programa-escalado-tex programa stream intervalo alto-base-ex))
    (format stream "\\end{tabular}~%")
    (format stream "\\end{adjustbox}~%")
    (format stream "\\end{center}~%~%")))

(defun generar-horario-tv-tex (nombre-canal planificacion output-tex-file
                                 &key
                                   (intervalo-minutos *tv-tex-intervalo-minutos*)
                                   (alto-base-ex *tv-tex-alto-base-ex*))
  (with-open-file (stream output-tex-file
                          :direction :output
                          :if-exists :supersede)
    (format stream "\\documentclass[11pt,a4paper]{article}~%")
    (format stream "\\usepackage[utf8]{inputenc}~%")
    (format stream "\\usepackage[T1]{fontenc}~%")
    (format stream "\\usepackage[spanish]{babel}~%")
    (format stream "\\usepackage{geometry}~%")
    (format stream "\\geometry{left=1.8cm,right=1.8cm,top=1.8cm,bottom=1.8cm}~%")
    (format stream "\\usepackage{adjustbox}~%")
    (format stream "\\usepackage{multirow}~%")
    (format stream "\\usepackage{array}~%")
    (format stream "\\setlength{\\parindent}{0pt}~%")
    (format stream "\\renewcommand{\\arraystretch}{1.0}~%")
    (format stream "\\begin{document}~%")
    (format stream "\\section*{~a}~%" (tv-tex-escape nombre-canal))
    (format stream "Intervalo base: ~a minutos.\\\\~%" intervalo-minutos)
    (format stream "La columna Intervalo es fija por unidad; Programa se dibuja a escala.\\\\~%~%")
    (loop for dia-config in planificacion
          for idx from 0
          for intervalo-dia = (tv-intervalo-dia dia-config intervalo-minutos)
          do
            (when (> idx 0)
              (format stream "\\clearpage~%"))
            (tv-escribir-tabla-dia-tex dia-config stream intervalo-dia alto-base-ex))
    (format stream "\\end{document}~%"))
  (format t "Archivo LaTeX '~a' generado.~%" output-tex-file)
  (format t "Para compilar a PDF: pdflatex ~a~%" output-tex-file)
  output-tex-file)

(defun generar-horario-tv-desde-json
       (&key
          (variables-file "variables_horario_tv.lisp")
          (output-python-file "generar_horario_tv_desde_lisp.py")
          (output-excel-file "horario_tv_semanal.xlsx")
          (output-tex-file nil)
          (intervalo-minutos 15)
       (intervalo-minutos-para-todos nil)
          (alto-base-ex 2.8))
  (load variables-file)
  (unless (boundp '*tv-planificacion-semanal*)
    (error "No se encontro *tv-planificacion-semanal* en ~a" variables-file))

  (unless (boundp '*tv-nombre-canal*)
    (setf *tv-nombre-canal* "Canal TV"))

    (let* ((planificacion-base (or *tv-planificacion-semanal* '()))
      (planificacion (if intervalo-minutos-para-todos
          (mapcar (lambda (dia-config)
               (let ((copia (copy-list dia-config)))
                 (setf (getf copia :intervalo-minutos) intervalo-minutos-para-todos)
                 copia))
             planificacion-base)
          planificacion-base))
      (libro-tv (crear-libro-tv-desde-planificacion planificacion intervalo-minutos)))
    (declare (ignore libro-tv))
    (setf *tv-tex-intervalo-minutos* intervalo-minutos)
    (setf *tv-tex-alto-base-ex* alto-base-ex)
    (generar-script-python-tv *tv-nombre-canal*
                              planificacion
                              output-python-file
                              output-excel-file)
    (format t "Archivo Python '~a' generado.~%" output-python-file)
    (format t "Para generar el Excel, ejecuta: python3 ~a~%" output-python-file)
    (when output-tex-file
      (generar-horario-tv-tex *tv-nombre-canal*
                              planificacion
                              output-tex-file
                              :intervalo-minutos intervalo-minutos
                              :alto-base-ex alto-base-ex))
    output-python-file))
