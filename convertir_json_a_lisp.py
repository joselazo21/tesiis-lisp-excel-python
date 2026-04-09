#!/usr/bin/env python3
"""
Convierte horario.json a variables Lisp para replicar_propuesta_ods.lisp
Nueva estructura: horarios -> CARRERA -> GRUPO -> turnos [{turno, hora, dias...}]
"""
import json
import re

DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes"]
DIAS_MAP = {"lunes": 0, "martes": 1, "miercoles": 2, "jueves": 3, "viernes": 4}

def parsear_entrada(valor):
    """
    Parsea una entrada como "F Aula 8" o "AM I Aula 6*" o "EF SEDER"
    Devuelve (asignatura, aula)
    """
    if not valor or valor.strip() == "":
        return ("", "")
    
    valor = valor.strip()
    
    # Casos especiales como "EF SEDER" (sin aula)
    if valor in ["EF SEDER", "EF"]:
        return ("EF", "SEDER")
    
    # Buscar patrón: ASIGNATURA AULA
    # La aula suele ser "Aula X", "c X", "cp X", "Lab", o un número
    # La asignatura puede tener espacios como "AM I", "A I", etc.
    
    # Patrones de aula
    aula_patterns = [
        r'(Aula\s*\d+\*?)',
        r'(Aula\s*\d+)',
        r'(c\s*\d+)',
        r'(cp\s*\d+)',
        r'(cp\s+Lab\d*)',
        r'(Lab\s*\d*)',
        r'(lab)',
        r'(\d+)$',  # número solo al final
    ]
    
    aula = ""
    asignatura = valor
    
    for pattern in aula_patterns:
        match = re.search(pattern, valor, re.IGNORECASE)
        if match:
            aula = match.group(1).strip()
            # Quitar el aula del valor para obtener la asignatura
            asignatura = valor[:match.start()].strip()
            break
    
    # Si no encontramos aula, verificar si hay número al final
    if not aula:
        parts = valor.rsplit(None, 1)
        if len(parts) == 2 and parts[1].isdigit():
            aula = parts[1]
            asignatura = parts[0]
    
    # Limpiar asignatura de paréntesis y notas adicionales
    # Ej: "IP Lab ICD 7 (s. impares) / F 8 (s. pares)"
    if "/" in asignatura:
        # Tomar solo la primera parte
        asignatura = asignatura.split("/")[0].strip()
    
    # Quitar notas entre paréntesis que no son parte del aula
    asignatura = re.sub(r'\s*\([^)]*\)\s*$', '', asignatura)
    
    return (asignatura.strip(), aula.strip())

def procesar_grupo(grupo_code, grupo_data):
    """
    Procesa un grupo del nuevo formato JSON.
    grupo_data tiene: turnos [{turno, hora, lunes, martes, miercoles, jueves, viernes}]
    """
    turnos = grupo_data.get("turnos", [])
    
    # Crear matriz de horario: 6 turnos x 5 días
    # Cada celda tiene (asignatura, aula)
    horario_matriz = [[("", "") for _ in range(5)] for _ in range(6)]
    
    for turno_data in turnos:
        turno_num = turno_data.get("turno", 1) - 1  # Turnos 1-6 -> índices 0-5
        if turno_num < 0 or turno_num > 5:
            continue
        
        for dia_key in DIAS:
            if dia_key in turno_data:
                valor = turno_data[dia_key]
                if valor:
                    asig, aula = parsear_entrada(valor)
                    dia_idx = DIAS_MAP[dia_key]
                    # Si ya hay algo, puede haber múltiples entradas
                    existing = horario_matriz[turno_num][dia_idx]
                    if existing[0]:
                        horario_matriz[turno_num][dia_idx] = (
                            f"{existing[0]}/{asig}",
                            f"{existing[1]}/{aula}"
                        )
                    else:
                        horario_matriz[turno_num][dia_idx] = (asig, aula)
    
    # Convertir a formato Lisp (lista de listas, cada turno = 2 filas)
    # Siempre generar 6 turnos (12 filas) para consistencia
    horario_lisp = []
    for turno_idx in range(6):
        fila_asignaturas = []
        fila_aulas = []
        for dia_idx in range(5):
            asig, aula = horario_matriz[turno_idx][dia_idx]
            fila_asignaturas.append(asig if asig else "")
            fila_aulas.append(aula if aula else "")
        horario_lisp.append(fila_asignaturas)
        horario_lisp.append(fila_aulas)
    
    return horario_lisp

def procesar_grupo_nuevo(grupo_data):
    """
    Procesa un grupo del nuevo formato JSON (array con 'tabla').
    tabla: [{Turno, Horas, Lunes, Martes, Miércoles, Jueves, Viernes}]
    """
    tabla = grupo_data.get("tabla", [])
    
    # Crear matriz de horario: 6 turnos x 5 días
    horario_matriz = [[("", "") for _ in range(5)] for _ in range(6)]
    
    dias_keys = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    
    for fila in tabla:
        turno_str = fila.get("Turno", "1")
        try:
            turno_num = int(turno_str) - 1  # Turnos 1-6 -> índices 0-5
        except ValueError:
            continue
        
        if turno_num < 0 or turno_num > 5:
            continue
        
        for dia_idx, dia_key in enumerate(dias_keys):
            valor = fila.get(dia_key, "")
            if valor:
                asig, aula = parsear_entrada(valor)
                existing = horario_matriz[turno_num][dia_idx]
                if existing[0]:
                    horario_matriz[turno_num][dia_idx] = (
                        f"{existing[0]}/{asig}",
                        f"{existing[1]}/{aula}"
                    )
                else:
                    horario_matriz[turno_num][dia_idx] = (asig, aula)
    
    # Convertir a formato Lisp
    # Siempre generar 6 turnos (12 filas) para consistencia
    horario_lisp = []
    for turno_idx in range(6):
        fila_asignaturas = []
        fila_aulas = []
        for dia_idx in range(5):
            asig, aula = horario_matriz[turno_idx][dia_idx]
            fila_asignaturas.append(asig if asig else "")
            fila_aulas.append(aula if aula else "")
        horario_lisp.append(fila_asignaturas)
        horario_lisp.append(fila_aulas)
    
    return horario_lisp

def extraer_asignaturas_por_grupo(data, grupo_code):
    """
    Extrae asignaturas para un grupo específico desde info_asignaturas.
    Mapea D111 -> D1, C211 -> C2, M311 -> M3, etc.
    """
    # Mapear grupo a año (ej: D111 -> D1, C211 -> C2)
    prefix = ""
    for char in grupo_code:
        if char.isalpha():
            prefix += char
        elif char.isdigit():
            # Tomar el primer dígito como el año
            year = char
            break
    
    anno_key = f"{prefix}{year}"
    
    info_asignaturas = data.get("info_asignaturas", [])
    
    for anno_data in info_asignaturas:
        anno = anno_data.get("anno", "")
        if anno.startswith(anno_key):
            # Encontramos el año, extraer asignaturas
            asignaturas_list = []
            for asig in anno_data.get("asignaturas", []):
                nombre = asig.get("nombre", "")
                abreviatura = asig.get("abreviatura", "")
                horas = asig.get("horas_real", asig.get("horas_plan", 64))
                # Calcular frecuencia aproximada (horas / 16 semanas)
                frecuencia = max(2, min(4, round(horas / 16 / 2)))
                if nombre:
                    # Formato: (abreviatura, nombre, frecuencia, faltan, asignadas)
                    asignaturas_list.append((abreviatura, nombre, frecuencia, 0, frecuencia))
            return asignaturas_list
    
    return []

def generar_lisp_variable(nombre, datos, indent=2):
    """Genera código Lisp para una variable"""
    ind = " " * indent
    lines = [f"(defparameter *{nombre}*"]
    lines.append(f"{ind}'(")
    
    for fila in datos:
        fila_str = " ".join(f'"{x}"' if isinstance(x, str) else str(x) for x in fila)
        lines.append(f"{ind} ({fila_str})")
    
    lines.append(f"{ind}))")
    return "\n".join(lines)

def main():
    with open("horario.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    
    output_lines = []
    output_lines.append(";; Variables generadas automáticamente desde horario.json")
    output_lines.append(";; Para usar en replicar_propuesta_ods.lisp")
    output_lines.append("")
    
    grupos_procesados = []
    
    # Nueva estructura: array de horarios
    horarios = data.get("horarios", [])
    
    for grupo_data in horarios:
        carrera = grupo_data.get("carrera", "")
        grupo_code = grupo_data.get("grupo", "")
        
        # Limpiar código de grupo (quitar paréntesis si tiene)
        grupo_code_clean = grupo_code.split(" ")[0] if " " in grupo_code else grupo_code
        
        horario = procesar_grupo_nuevo(grupo_data)
        grupos_procesados.append(grupo_code_clean)
        
        output_lines.append(f";; {carrera} - {grupo_code}")
        output_lines.append(generar_lisp_variable(f"horario-{grupo_code_clean.lower()}", horario))
        output_lines.append("")
    
    # Generar asignaturas por grupo
    output_lines.append(";; Asignaturas por grupo")
    
    for grupo_code_clean in grupos_procesados:
        asignaturas = extraer_asignaturas_por_grupo(data, grupo_code_clean)
        if asignaturas:
            output_lines.append(f";; Asignaturas para {grupo_code_clean}")
            output_lines.append(generar_lisp_variable(f"asignaturas-{grupo_code_clean.lower()}", asignaturas))
            output_lines.append("")
    
    # Asignaturas por defecto vacías
    output_lines.append(";; Variable por defecto para grupos sin asignaturas específicas")
    output_lines.append("(defparameter *asignaturas-todas* '())")
    
    # Guardar resultado
    with open("variables_horario.lisp", "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
    
    print("Archivo 'variables_horario.lisp' generado.")
    print(f"Grupos procesados: {len(grupos_procesados)}")
    
    for g in grupos_procesados:
        print(f"  - {g}")

if __name__ == "__main__":
    main()
