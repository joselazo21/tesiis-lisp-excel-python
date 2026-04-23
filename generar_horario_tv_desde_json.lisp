;; generar_horario_tv_desde_json.lisp
;; Convierte variables TV (desde JSON) en clases Lisp y genera un script Python
;; que llama la API parametrica de hoja_con_formulas.py.

(load "codigo-tesis.lisp")

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

(defun crear-hoja-tv-dia (dia-config)
  (let ((dia (tv-dia-display (getf dia-config :dia))))
    (hoja :grupo dia
          :horario (crear-tabla-tv-dia dia-config)
          :asignaturas (tabla :id (format nil "meta-~a" dia)
                              :contenido-de-la-tabla '(("" ""))))))

(defun crear-libro-tv-desde-planificacion (planificacion)
  (libro :hojas (mapcar #'crear-hoja-tv-dia planificacion)))

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

(defun generar-horario-tv-desde-json
       (&key
          (variables-file "variables_horario_tv.lisp")
          (output-python-file "generar_horario_tv_desde_lisp.py")
          (output-excel-file "horario_tv_semanal.xlsx"))
  (load variables-file)
  (unless (boundp '*tv-planificacion-semanal*)
    (error "No se encontro *tv-planificacion-semanal* en ~a" variables-file))

  (unless (boundp '*tv-nombre-canal*)
    (setf *tv-nombre-canal* "Canal TV"))

  (let* ((planificacion (or *tv-planificacion-semanal* '()))
         (libro-tv (crear-libro-tv-desde-planificacion planificacion)))
    (declare (ignore libro-tv))
    (generar-script-python-tv *tv-nombre-canal*
                              planificacion
                              output-python-file
                              output-excel-file)
    (format t "Archivo Python '~a' generado.~%" output-python-file)
    (format t "Para generar el Excel, ejecuta: python3 ~a~%" output-python-file)
    output-python-file))
