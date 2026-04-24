# Comandos para generar `.tex` y `.pdf`

## 1) Generar `.tex` (y script Python para Excel)

```bash
sbcl --noinform --non-interactive \
  --load generar_horario_tv_desde_json.tex.lisp \
  --eval '(generar-horario-tv-desde-json
            :output-python-file "generar_horario_tv_desde_lisp.py"
            :output-excel-file "horario_tv_semanal.xlsx"
            :output-tex-file "horario_tv.tex"
            :intervalo-minutos 15
            :alto-base-ex 2.8)'
```

## 2) Compilar `.tex` a `.pdf`

```bash
pdflatex horario_tv.tex
```

## Opcional: otro intervalo (ej. 30 min)

```bash
sbcl --noinform --non-interactive \
  --load generar_horario_tv_desde_json.tex.lisp \
  --eval '(generar-horario-tv-desde-json
            :output-tex-file "horario_tv.tex"
            :intervalo-minutos 30
            :alto-base-ex 2.8)'
```
