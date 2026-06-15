;; =============================================================================
;; dsl-directo.lisp — Macros del DSL de hojas de cálculo
;;
;; RESPONSABILIDAD: exponer una sintaxis de dominio limpia que construye
;; nodos del AST (definidos en ast-def.lisp). Este archivo no contiene
;; ninguna lógica de generación de código.
;;
;; PRINCIPIO CLAVE: los macros son azúcar sintáctico puro.
;;   - Cada macro construye uno o más objetos xl-* del AST.
;;   - No hay strings de Excel, no hay letras de columna, no hay fórmulas.
;;   - Los nombres son del dominio: "dia", "lun", "turno", "grupos"...
;;
;; REGLA DE QUOTING EN MACROS:
;;   - Un símbolo que es nombre fijo de columna/tabla/variable se quote: ',sym
;;     (el caller escribe sin comilla: (col lun) → sym LUN queda en el AST)
;;   - Un símbolo que puede ser variable en scope NO se quote: ,sym
;;     (el caller escribe la variable o la quotea él: :col dia  o  :col 'a)
;;
;; Carga generate-code-direct para que los métodos compile-excel-formula
;; estén disponibles al evaluar los ASTs generados.
;; =============================================================================

(load (merge-pathnames "codigo-tesis.lisp"        *load-truename*))
(load (merge-pathnames "generate-code-direct.lisp" *load-truename*))

;; =============================================================================
;; EXPRESIONES DE FLUJO Y COMPARACIÓN
;; =============================================================================

;; Condicional: (_if test then else)
;; → IF(test, then, else)
(defmacro _if (test then else)
  `(xl-expr-if :test ,test :then ,then :else ,else))

;; Verdad si la expresión no es vacía
;; (non-empty (col programa)) → programa <> ""
(defmacro non-empty (expr)
  `(xl-expr-non-empty :expr ,expr))

;; Verdad en la primera fila de datos
;; → ROW() = first-row
(defmacro it-is-the-first-row ()
  `(xl-expr-first-row))

;; Valor de la misma columna en la fila anterior / siguiente
;; (previous-row (col tipo)) → C3 cuando se está en C4
(defmacro previous-row (expr)
  `(xl-expr-previous-row :expr ,expr))

(defmacro next-of (base)
  `(xl-expr-next-row :expr ,base))

(defmacro previous-of (base)
  `(xl-expr-previous-row :expr ,base))

;; Produce "" visualmente (celda en blanco)
(defmacro show-nothing ()
  `(xl-expr-show-nothing))

;; Comparaciones: (equals a b) (different a b) (_and a b) (_or a b)
(defmacro equals (a b)    `(xl-expr-equals :a ,a :b ,b))
(defmacro different (a b) `(xl-expr-different :a ,a :b ,b))
(defmacro _and (a b)      `(xl-expr-and :a ,a :b ,b))
(defmacro _or (a b)       `(xl-expr-or :a ,a :b ,b))

;; Comparaciones aritméticas: (gt a b) (lt a b) (gte a b) (lte a b)
(defmacro gt  (a b) `(xl-expr-gt  :a ,a :b ,b))
(defmacro lt  (a b) `(xl-expr-lt  :a ,a :b ,b))
(defmacro gte (a b) `(xl-expr-gte :a ,a :b ,b))
(defmacro lte (a b) `(xl-expr-lte :a ,a :b ,b))

;; =============================================================================
;; ARITMÉTICA
;; =============================================================================

(defmacro add (a b)      `(xl-expr-add :a ,a :b ,b))
(defmacro subtract (a b) `(xl-expr-subtract :a ,a :b ,b))
(defmacro multiply (a b) `(xl-expr-multiply :a ,a :b ,b))
(defmacro divide (a b)   `(xl-expr-divide :a ,a :b ,b))

;; Promedio: (promedio (col n1) (col n2) ...)
;; Compila a ((n1 + n2 + ...) / cantidad-de-columnas)
(defmacro promedio (&rest cols)
  `(xl-expr-promedio :cols (list ,@cols)))

;; =============================================================================
;; STRINGS
;; =============================================================================

;; Literal de texto: (str "Total") → "Total"
(defmacro str (value)
  `(xl-expr-string :value ,value))

;; Concatenación: (concat (col nombre) (str " hs")) → (nombre & " hs")
(defmacro concat (a b)
  `(xl-expr-concat :a ,a :b ,b))

;; Suma de tiempos: (time-add (col hora) 90) → TEXT(TIMEVALUE(hora)+90/1440,"hh:mm")
(defmacro time-add (a b)
  `(xl-expr-time-add :a ,a :b ,b))

;; =============================================================================
;; REFERENCIA A COLUMNAS DE LA TABLA ACTUAL
;; =============================================================================

;; (col nombre)          → columna nombre, fila actual
;; (col nombre contexto) → columna nombre con contexto de fila (previous-row, etc.)
;; El nombre se quotea: es un símbolo del dominio, no una variable Lisp.
(defmacro col (name &optional context)
  (if context
      `(xl-expr-column-ref :name ',name :context ,context)
      `(xl-expr-column-ref :name ',name :context nil)))

;; (tcol tabla nombre)          → columna nombre de tabla concreta, fila actual
;; (tcol tabla nombre contexto) → ídem con desfase de fila
;; Usar cuando otra tabla de la misma región tiene una columna con el mismo nombre.
(defmacro tcol (table-id name &optional context)
  (if context
      `(xl-expr-table-col-ref :table-id ',table-id :name ',name :context ,context)
      `(xl-expr-table-col-ref :table-id ',table-id :name ',name :context nil)))

;; (trange tabla desde)        → rango de una columna de tabla concreta
;; (trange tabla desde hasta)  → rango de columnas de tabla concreta
(defmacro trange (table-id from &optional to)
  `(xl-table-range
     :table-id ',table-id
     :from-col (xl-expr-table-col-ref :table-id ',table-id :name ',from :context nil)
     :to-col   ,(if to
                    `(xl-expr-table-col-ref :table-id ',table-id :name ',to :context nil)
                    nil)))

;; (param nombre) → referencia a la celda que contiene el parámetro nombre
;; Los parámetros se pasan al instanciar la tabla: (tabla t :params ((n val)))
(defmacro param (name)
  `(xl-expr-param-ref :name ',name))

;; =============================================================================
;; RANGOS Y AGREGADOS (sobre columnas de la tabla actual)
;; =============================================================================
;;
;; El rango abarca desde first-row hasta last-row de la tabla.
;;   (range lun)       → columna lun, todas las filas
;;   (range lun vie)   → desde columna lun hasta columna vie

(defmacro range (from &optional to)
  `(xl-range :from-col (xl-expr-column-ref :name ',from :context nil)
             :to-col ,(if to
                          `(xl-expr-column-ref :name ',to :context nil)
                          nil)))

;; (countif (range lun vie) (col abrev)) → COUNTIF($C$4:$L$9, N4)
(defmacro countif (range-expr criteria)
  `(xl-expr-countif :count-range ,range-expr :criteria ,criteria))

;; (counta (range abrev)) → COUNTA($N$4:$N$9)
(defmacro counta (range-expr)
  `(xl-expr-counta :count-range ,range-expr))

;; (sum-range (range frec)) → SUM($J$4:$J$9)
(defmacro sum-range (range-expr)
  `(xl-expr-sum :count-range ,range-expr))

;; =============================================================================
;; LOOKUPS
;; =============================================================================

;; (lookup key-expr campo) → IFERROR(VLOOKUP(key, rango, col-campo, FALSE), 0)
;; Busca key-expr en la primera columna de la tabla y devuelve el campo.
(defmacro lookup (key-expr field)
  `(xl-expr-lookup :value-field ',field :key-expr ,key-expr))

;; =============================================================================
;; REFERENCIAS CROSS-SHEET SIMPLES (hoja fija)
;; =============================================================================
;;
;; Para referencias a una hoja conocida en tiempo de escritura del AST.
;; La hoja y la celda son literales — no cambian según el contexto.
;;
;; (sheet-ref "C111" "$C$5")        → C111!$C$5
;; (sheet-ref "C111" "$C${row-num}") → C111!$C<row-actual>

(defmacro sheet-ref (sheet cell-template)
  `(xl-expr-cross-sheet-ref :sheet ,sheet :cell-template ,cell-template))

;; =============================================================================
;; REFERENCIAS CROSS-SHEET DINÁMICAS — para búsquedas sobre múltiples hojas
;;
;; Estas tres expresiones trabajan juntas:
;;
;;   (collect-over lista (var) expr)
;;     Itera lista, liga var a cada hoja, compila expr para cada una,
;;     y concatena los resultados: SUBSTITUTE(TRIM(t1&t2&...&tN)," ",",")
;;
;;   (cross-cell :sheet var :col col-sym :row row-expr)
;;     Dentro de collect-over: referencia a la celda (col-sym, row-expr)
;;     en la hoja actualmente ligada a var.
;;
;;   (turno-aula-row)
;;     Fila del aula en la turno-table del grupo para el turno que se
;;     está compilando. generate-code la computa desde row-num y first-row.
;;
;; Ejemplo de uso:
;;   (collect-over *grupos* (g)
;;     (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
;;                  (str "Aula 6"))
;;          (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
;;          (str "")))
;;
;; :col QUOTING:
;;   - :col dia    → dia es una variable en scope (inst-param): se evalúa
;;   - :col 'a     → 'a quotea el literal: resuelve al símbolo A (col A)
;; =============================================================================

;; (cross-cell :sheet g :col col :row row)
;; :sheet se quotea automáticamente — es el nombre de la variable de iteración.
;; :col no se quotea — el caller decide: variable (dia) o literal ('a).
;; :row no se quotea — entero fijo (1) o expresión ((turno-aula-row)).
(defmacro cross-cell (&key sheet col row)
  `(xl-expr-cross-cell :sheet ',sheet :xcol ,col :row ,row))

;; (source-row :table TABLE :offset N)
;; Fila de una subcelda en la tabla origen TABLE, mapeada desde la fila actual.
;; TABLE debe estar registrada en *source-table-schemas* (generate-code-direct.lisp).
;; :offset subcelda dentro de la fila compuesta (0=primera, 1=segunda...) — default 0.
;; Para conceptos nombrados usa un macro explícito en este archivo (ver turno-aula-row).
(defmacro source-row (&key table (offset 0))
  `(xl-expr-source-row :table-id ',table :offset ,offset))

;; (turno-aula-row)
;; Alias de (source-row :table turno-table :offset 1).
;; No requiere clase propia — usa xl-expr-source-row directamente.
(defmacro turno-aula-row ()
  `(source-row :table turno-table :offset 1))

;; (sheet-id :sheet var)
;; Identificador canónico de la hoja ligada a var en el collect-over actual.
;; Excel: genera SHEET!$A$1 — celda donde cada hoja almacena su nombre de grupo.
(defmacro sheet-id (&key sheet)
  `(xl-expr-sheet-id :sheet ',sheet))

;; (collect-over groups (sheet-var) expr)
;; groups    : expresión que evalúa a lista de strings (nombres de hojas)
;; sheet-var : símbolo que nombra la variable de iteración — usar en cross-cell
;; expr      : una única expresión AST que puede referenciar sheet-var
(defmacro collect-over (groups (sheet-var) expr)
  `(xl-expr-collect-over :groups ,groups
                          :sheet-var ',sheet-var
                          :body ,expr))

;; =============================================================================
;; POSICIONAMIENTO RELATIVO — expresiones fijas en la hoja
;; =============================================================================
;;
;; Las expresiones fijas ya NO van en def-table: son slots de la hoja (hoja/hoja-v).
;;
;; Anclas — punto de partida sobre una columna de una tabla:
;;   (ultima-fila tabla col)   → última fila de datos de col en tabla
;;   (primera-fila tabla col)  → primera fila de datos de col en tabla
;;
;; Navegación — desplazamiento desde el ancla:
;;   (nav ancla abajo 2 izquierda 1)
;;   Direcciones: abajo  arriba  izquierda  derecha

(defmacro ultima-fila (tabla col)
  `(xl-anchor :table-id  ',tabla
              :col-name  ',col
              :anchor-type :ultima-fila))

(defmacro primera-fila (tabla col)
  `(xl-anchor :table-id  ',tabla
              :col-name  ',col
              :anchor-type :primera-fila))

;; (nav ancla dir1 n1 dir2 n2 ...) — pasos planos, pares dir+cantidad
(defmacro nav (anchor &rest steps)
  (let ((step-pairs (loop for (dir n) on steps by #'cddr
                          collect `(cons ',dir ,n))))
    `(xl-nav :anchor ,anchor :steps (list ,@step-pairs))))

;; =============================================================================
;; CUANTIFICACIÓN — exists sobre filas
;; =============================================================================
;;
;; (exists (var :all-rows)              — filas de la tabla actual
;;   :matching    (col1 col2 ...)
;;   :any-overlap (col3 col4 ...))
;;
;; (exists (var :rows-of tabla-id)      — filas de otra tabla (cross-table)
;;   :self-in  (col1 col2 ...)          — columnas de la otra tabla donde aparece la celda actual
;;   :matching (col3 col4 ...))         — columnas que definen el slot de tiempo (conflicto > 1)
;;
;;   var          : nombre de la variable (documentación)
;;   :matching    : columnas que deben coincidir entre filas del dominio
;;   :any-overlap : columnas donde algún valor debe ser compartido (same-table)
;;   :self-in     : columnas del domain-table donde buscar el valor actual (cross-table)
;;
;; El generador decide el mecanismo de comprobación según el tipo de dominio
;; y la presencia de self-in-cols — el usuario solo declara la intención.
(defmacro exists ((var &rest domain-spec) &key matching any-overlap self-in)
  (let ((domain-type (first domain-spec)))
    (ecase domain-type
      (:all-rows
       `(xl-expr-exists :bind-var     ',var
                         :domain       (xl-domain-rows :table-id nil)
                         :match-keys   ',matching
                         :overlap-cols ',any-overlap
                         :self-in-cols  nil))
      (:rows-of
       (let ((table-id (second domain-spec)))
         `(xl-expr-exists :bind-var     ',var
                           :domain       (xl-domain-rows :table-id ',table-id)
                           :match-keys   ',matching
                           :overlap-cols ',any-overlap
                           :self-in-cols  ',self-in))))))


;; =============================================================================
;; RENDERIZADO CONDICIONAL
;; =============================================================================
;;
;; (conditional-rendering :condition expr :target-columns (col1 col2))
;; Aplica estilo condicional a las columnas indicadas cuando se cumple condition.

(defmacro conditional-rendering (&key condition target-columns)
  `(xl-style-rule :rule-condition ,condition
                  :target-columns ',target-columns))

;; =============================================================================
;; DEF-TABLE — define la plantilla de una tabla reutilizable
;; =============================================================================
;;
;; (def-table nombre
;;   ((col1 "Header1") (col2 "Header2") ...)
;;   :cell-height 2          ; filas por fila lógica (para celdas paired)
;;   :cell-width  1
;;   :inst-params (param1)   ; parámetros de generación pasados al instanciar
;;   :computed               ; fórmulas por columna
;;     ((col1 expresion-dsl)
;;      (col2 (collect-over ...)))
;;   :render                 ; regla de estilo condicional (opcional)
;;     (conditional-rendering ...))
;;
;; Genera una función: (nombre &key data params [inst-params...])
;;   - data       : lista de filas de contenido
;;   - params     : lista de pares (sym . valor) para celdas auxiliares
;;   - inst-params: cada parámetro declarado en :inst-params se añade a la firma
;;
;; Las expresiones fijas van en :fixed-expressions del macro hoja, no aquí.

(defmacro def-table (name columns &body body)
  (let ((col-names   (mapcar #'first columns))
        (headers     (mapcar #'second columns))
        (computed    (getf body :computed))
        (render      (getf body :render))
        (renders     (getf body :renders))
        (cell-height (or (getf body :cell-height) 1))
        (cell-width  (or (getf body :cell-width) 1))
        (first-row   (or (getf body :first-row) 4))
        (inst-params (getf body :inst-params)))
    `(defun ,name (&key data params ,@inst-params)
       (xl-table :id ',name
                 :contenido (or data '())
                 :headers ',headers
                 :col-names ',col-names
                 :computed (list ,@(loop for (col expr) in computed
                                         collect `(cons ',col ,expr)))
                 :style-rules (list ,@(when render (list render)) ,@(or renders nil))
                 :cell-height ,cell-height
                 :cell-width ,cell-width
                 :first-row ,first-row
                 :params params))))

;; =============================================================================
;; TABLA — instancia una tabla (llama a la función definida con def-table)
;; =============================================================================
;;
;; Forma instancia (tabla definida con def-table):
;;   (tabla nombre-tabla
;;     :data *mis-datos*
;;     :params ((param-sym valor))
;;     :inst-key val ...)         ; inst-params adicionales
;;
;; Los keyword args que no son :data ni :params son inst-params: se pasan
;; quoted a la función de tabla para usarlos en generación de código.
;; Ejemplo: (tabla aulas-dia-table :dia lun :data *DATOS-AULAS-LUNES*)
;;          → llama a (aulas-dia-table :data ... :dia 'lun :params ...)
;;
;; Forma inline (tabla definida en el lugar):
;;   (tabla ((col1 "H1") (col2 "H2"))
;;     :data *datos*
;;     :computed ((col1 expr))
;;     :params ...)

(defmacro tabla (first &rest args)
  (if (listp first)
      ;; ── forma inline ──────────────────────────────────────────────────────
      (let* ((col-names   (mapcar #'first first))
             (headers     (mapcar #'second first))
             (data        (getf args :data))
             (computed    (getf args :computed))
             (params      (getf args :params))
             (cell-height (or (getf args :cell-height) 1))
             (cell-width  (or (getf args :cell-width) 1)))
        `(xl-table :contenido (or ,data '()) :headers ',headers
                   :col-names ',col-names
                   :computed (list ,@(loop for (col expr) in computed
                                           collect `(cons ',col ,expr)))
                   :style-rules (list ,@(when (getf args :render) (list (getf args :render))))
                   :cell-height ,cell-height
                   :cell-width ,cell-width
                   :params (list ,@(loop for (name val) in params
                                         collect `(cons ',name ,val)))))
      ;; ── forma instancia ───────────────────────────────────────────────────
      (let* ((data-arg    (getf args :data))
             (params-arg  (getf args :params))
             ;; inst-kwargs: cualquier keyword que no sea :data ni :params.
             ;; Se pasa quoted porque son símbolos de dominio, no variables Lisp.
             (inst-kwargs (loop for (k v) on args by #'cddr
                                unless (member k '(:data :params))
                                nconc `(,k ',v))))
        `(,first :data ,data-arg
                 ,@inst-kwargs
                 :params (list ,@(loop for (name val) in params-arg
                                        collect `(cons ',name ,val)))))))

;; =============================================================================
;; HOJA — agrupa tablas en una hoja de cálculo
;; =============================================================================
;;
;; (hoja "NombreHoja"
;;   (tabla turno-table :data *datos-turno*)
;;   (tabla stats-table :data *datos-stats*)
;;   :fixed-expressions
;;     (((nav (ultima-fila stats-table abrev) abajo 1)
;;       (counta (trange stats-table abrev)))
;;      ((nav (ultima-fila turno-table vie) abajo 2 derecha 1)
;;       (sum-range (trange stats-table frec)))))
;;
;; Las formas (tabla ...) van primero; :fixed-expressions al final (opcional).
;; Todas las tablas van en una ÚNICA región (horizontales side-by-side).
;; Para tablas apiladas verticalmente usar hoja-v.

(defmacro hoja (name &body body)
  (let* ((fe-pos      (position :fixed-expressions body))
         (table-forms (if fe-pos (subseq body 0 fe-pos) body))
         (fixed-exprs (when fe-pos (nth (1+ fe-pos) body))))
    `(xl-sheet
       :name ,name
       :regions (list
         (xl-region
           :tables (list ,@table-forms)))
       :fixed-expressions
         (list ,@(loop for (pos-form expr-form) in fixed-exprs
                        collect `(xl-sheet-fixed-expr
                                   :pos ,pos-form
                                   :expr     ,expr-form))))))

;; (hoja-v "NombreHoja" tabla1 tabla2 ...)
;; Cada tabla va en su propia región.
;; El backend Python (_unwrap_regions) las apila verticalmente con separación.
;; :fixed-expressions soportado igual que en hoja.
(defmacro hoja-v (name &body body)
  (let* ((fe-pos      (position :fixed-expressions body))
         (table-forms (if fe-pos (subseq body 0 fe-pos) body))
         (fixed-exprs (when fe-pos (nth (1+ fe-pos) body))))
    `(xl-sheet
       :name ,name
       :regions (list
         ,@(mapcar (lambda (tbl) `(xl-region :tables (list ,tbl))) table-forms))
       :fixed-expressions
         (list ,@(loop for (pos-form expr-form) in fixed-exprs
                        collect `(xl-sheet-fixed-expr
                                   :pos ,pos-form
                                   :expr     ,expr-form))))))

;; =============================================================================
;; LIBRO — construye el workbook completo
;; =============================================================================
;;
;; (libro nombre-workbook
;;   :hojas (list (hoja "H1" ...) (hoja "H2" ...) ...))
;;
;; Define el símbolo nombre-workbook con el objeto xl-workbook generado.
;; El nombre del archivo de salida lo define el generador (por defecto
;; "Archivo-Excel.xlsx").

(defmacro libro (name &body body)
  (let* ((hojas-form  (getf body :hojas)))
    `(let ((wb (xl-workbook :name "Archivo-Excel.xlsx" :sheets ,hojas-form)))
       (defparameter ,name wb)
       (format t "Libro ~a creado con ~a hojas~%" ',name (length ,hojas-form))
       wb)))

(provide "dsl-directo")
