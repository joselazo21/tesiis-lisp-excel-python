#!/usr/bin/env bash
sbcl --script replicar_propuesta_ods.lisp
python3 generar_propuesta_desde_lisp.py
