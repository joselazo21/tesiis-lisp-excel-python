#!/usr/bin/env python3
"""
Script para arreglar las fórmulas de Fernando que referencian hojas que no existen.
Genera fórmulas dinámicas basándose en las hojas que realmente están en el libro.
"""

import json
import re
from typing import List, Dict, Set

def extraer_referencias_hojas(formula: str) -> Set[str]:
    """Extrae todas las referencias a hojas de una fórmula."""
    # Buscar patrones como C111!, D211!, c112!, etc.
    patron = r'([A-Za-z]\d{3})!'
    referencias = re.findall(patron, formula)
    return set(ref.upper() for ref in referencias)

def obtener_hojas_disponibles_por_tipo() -> Dict[str, List[str]]:
    """
    Obtiene las hojas disponibles organizadas por tipo (C, D, M).
    Debe sincronizarse con las hojas que realmente se generan.
    """
    # NOTA: Esta lista debe coincidir con las hojas generadas en tu sistema
    # Puedes obtenerla dinámicamente del archivo Python generado
    return {
        'C': ['C111', 'C121', 'C122', 'C211', 'C212', 'C311', 'C312', 'C411', 'C412'],
        'D': ['D111', 'D211', 'D311', 'D411'],
        'M': ['M111', 'M211', 'M311', 'M411']
    }

def generar_formula_dinamica(formula_template: str, hojas_disponibles: List[str]) -> str:
    """
    Genera una fórmula dinámica basándose en las hojas disponibles.
    
    Ejemplo:
    Input: "=SUSTITUIR(ESPACIOS(CONCATENAR(SI(C111!$C$5=C$2;C111!$B$1 & \" \"; \"\"); ...)))"
    Hojas: ['C111', 'C211', 'C311']
    Output: Genera fórmula con todas las hojas disponibles
    """
    # Extraer el patrón completo de la fórmula
    # Buscar patrones como: SI(XXX!$C$5=Y$2;XXX!$B$1 & " "; "")
    patron_si = r'SI\(([A-Za-z]\d{3})!(\$[A-Z]\$\d+)=([A-Z]\$\d+);(\1)!(\$[A-Z]\$\d+)\s*&\s*" ";\s*""\)'
    
    # Buscar la primera ocurrencia para obtener el patrón
    match = re.search(patron_si, formula_template)
    if not match:
        return formula_template
    
    # Extraer componentes del patrón
    _, ref_celda_condicion, celda_destino, _, ref_celda_valor = match.groups()
    
    # Generar SI() para cada hoja disponible
    condiciones = []
    for hoja in hojas_disponibles:
        condicion = f'SI({hoja}!{ref_celda_condicion}={celda_destino};{hoja}!{ref_celda_valor} & " "; "")'
        condiciones.append(condicion)
    
    # Reconstruir la fórmula completa
    concatenacion = "; ".join(condiciones)
    
    # Determinar el tipo de fórmula externa
    if 'SUSTITUIR(ESPACIOS(CONCATENAR(' in formula_template:
        formula_nueva = f'=SUSTITUIR(ESPACIOS(CONCATENAR({concatenacion})); " "; ",")'
    elif 'SUSTITUIR(ESPACIOS(' in formula_template:
        # Buscar qué hay dentro de ESPACIOS pero no es CONCATENAR
        # Por ejemplo: SUSTITUIR(ESPACIOS((SI(...); SI(...); SI(...))))
        formula_nueva = f'=SUSTITUIR(ESPACIOS(({concatenacion})); " "; ",")'
    else:
        formula_nueva = f'=CONCATENAR({concatenacion})'
    
    return formula_nueva

def adaptar_formulas_fernando(
    formulas_originales: List[Dict],
    hojas_disponibles: List[str]
) -> List[Dict]:
    """
    Adapta las fórmulas de Fernando para usar solo las hojas disponibles.
    """
    formulas_adaptadas = []
    
    for formula_obj in formulas_originales:
        celda = formula_obj['cell']
        formula = formula_obj['excel']
        
        # Verificar si la fórmula tiene referencias a hojas
        referencias = extraer_referencias_hojas(formula)
        
        if referencias:
            # Determinar el tipo de hoja (C, D, M) según la primera referencia
            tipo_ref = referencias.pop()[0] if referencias else None
            
            if tipo_ref and tipo_ref in ['C', 'D', 'M']:
                # Filtrar hojas disponibles del mismo tipo
                hojas_tipo = [h for h in hojas_disponibles if h.startswith(tipo_ref)]
                
                if hojas_tipo:
                    # Generar fórmula dinámica
                    formula_nueva = generar_formula_dinamica(formula, hojas_tipo)
                    formulas_adaptadas.append({
                        'cell': celda,
                        'original_formula': formula,
                        'excel': formula_nueva
                    })
        else:
            # Si no tiene referencias a hojas, conservar la fórmula original
            formulas_adaptadas.append(formula_obj)
    
    return formulas_adaptadas

def main():
    """Función principal para probar el sistema."""
    # Cargar fórmulas originales
    with open('formulas_fernando_convertidas_es.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Obtener hojas disponibles
    hojas_por_tipo = obtener_hojas_disponibles_por_tipo()
    
    # Adaptar fórmulas de la hoja Aulas
    if 'Aulas' in data:
        print("Adaptando fórmulas de la hoja Aulas...")
        
        # Todas las hojas disponibles (aplanar el diccionario)
        todas_las_hojas = []
        for hojas in hojas_por_tipo.values():
            todas_las_hojas.extend(hojas)
        
        formulas_adaptadas = adaptar_formulas_fernando(data['Aulas'], todas_las_hojas)
        
        # Mostrar algunas fórmulas adaptadas
        print(f"\nTotal de fórmulas: {len(formulas_adaptadas)}")
        print("\nPrimeras 3 fórmulas con referencias:")
        count = 0
        for formula_obj in formulas_adaptadas:
            if 'original_formula' in formula_obj:
                print(f"\nCelda {formula_obj['cell']}:")
                print(f"  Original: {formula_obj['original_formula'][:100]}...")
                print(f"  Adaptada: {formula_obj['excel'][:100]}...")
                count += 1
                if count >= 3:
                    break
        
        # Guardar resultado
        output_file = 'formulas_fernando_adaptadas_es.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump({
                'C111': data.get('C111', []),
                'C112': data.get('C112', []),
                'C113': data.get('C113', []),
                'Aulas': formulas_adaptadas
            }, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ Fórmulas adaptadas guardadas en: {output_file}")
        print("\nPara usar las fórmulas adaptadas:")
        print("  1. Modifica replicar_propuesta_ods.lisp para usar 'formulas_fernando_adaptadas_es.json'")
        print("  2. O mejor aún, genera las fórmulas dinámicamente desde Lisp")

if __name__ == '__main__':
    main()
