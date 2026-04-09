"""
Extrae fórmulas complejas del ODS de Fernando y las convierte a formato Excel
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple
import re

class FormulaConverter:
    """Convierte fórmulas de formato ODS a formato Excel"""
    
    # Mapeo de funciones ODS → Excel (inglés)
    FUNCTION_MAP_EN = {
        'COM.MICROSOFT.CONCAT': 'CONCATENATE',
    }
    
    # Mapeo de funciones ODS → Excel (español)
    FUNCTION_MAP_ES = {
        'COM.MICROSOFT.CONCAT': 'CONCATENAR',
        'TRIM': 'ESPACIOS',
        'SUBSTITUTE': 'SUSTITUIR',
        'IF': 'SI',
        'COUNTIF': 'CONTAR.SI',
    }
    
    @staticmethod
    def ods_to_excel_formula(ods_formula: str, locale: str = 'en') -> str:
        """
        Convierte una fórmula ODS (of:=...) a Excel (=...).
        
        Args:
            ods_formula: Fórmula en formato ODS
            locale: 'en' para inglés, 'es' para español
        
        Ejemplos:
            of:=[.J4]-[.L4] → =J4-L4 (ambos)
            of:=COUNTIF([.$C$4:.$G$17];[.I4]) → 
                EN: =COUNTIF($C$4:$G$17,I4)
                ES: =CONTAR.SI($C$4:$G$17;I4)
        """
        # Remover prefijo ODS
        formula = ods_formula.replace('of:=', '')
        
        # PASO 1: Convertir referencias a hojas PRIMERO (más específico)
        # [$C111.$C$5] → C111!$C$5 (con $ en ambos)
        formula = re.sub(r'\[\$([A-Z0-9]+)\.\$([A-Z]+)\$(\d+)\]', r'\1!$\2$\3', formula)
        # [$C111.$D5] → C111!$D5 (con $ en columna, sin en fila)
        formula = re.sub(r'\[\$([A-Z0-9]+)\.\$([A-Z]+)(\d+)\]', r'\1!$\2\3', formula)
        # [$C111.C5] → C111!C5 (sin $ en ninguno)
        formula = re.sub(r'\[\$([A-Z0-9]+)\.([A-Z]+)(\d+)\]', r'\1!\2\3', formula)
        
        # PASO 2: Convertir referencias simples de celda (preservando absolutas)
        # [.$A$1] → $A$1
        # [.$A1]  → $A1
        # [.A$1]  → A$1
        # [.A1]   → A1
        formula = re.sub(r'\[\.\$([A-Z]+)\$(\d+)\]', r'$\1$\2', formula)
        formula = re.sub(r'\[\.\$([A-Z]+)(\d+)\]', r'$\1\2', formula)
        formula = re.sub(r'\[\.([A-Z]+)\$(\d+)\]', r'\1$\2', formula)
        formula = re.sub(r'\[\.([A-Z]+)(\d+)\]', r'\1\2', formula)
        
        # PASO 3: Convertir rangos simples (cuando hay :)
        # [$C$4:.$G$17] → $C$4:$G$17
        formula = re.sub(r'\[\.\$([A-Z]+)\$(\d+):\.\$([A-Z]+)\$(\d+)\]', r'$\1$\2:$\3$\4', formula)
        formula = re.sub(r'\[\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)\]', r'$\1$\2:$\3$\4', formula)
        
        # Seleccionar mapa de funciones según locale
        function_map = FormulaConverter.FUNCTION_MAP_ES if locale == 'es' else FormulaConverter.FUNCTION_MAP_EN
        
        # Convertir nombres de funciones (solo identificadores completos)
        # Evita reemplazos parciales como IF dentro de COUNTIF.
        for ods_func, excel_func in sorted(function_map.items(), key=lambda x: len(x[0]), reverse=True):
            pattern = rf'(?<![A-Z0-9_.]){re.escape(ods_func)}(?=\()'
            formula = re.sub(pattern, excel_func, formula)
        
        # Convertir separadores de argumentos
        if locale == 'es':
            # Español: mantener ; como separador (ya está)
            pass
        else:
            # Inglés: convertir ; a ,
            formula = formula.replace(';', ',')
        
        return f'={formula}'
    
    @staticmethod
    def extract_formulas_from_ods(ods_path: str, locale: str = 'en') -> Dict[str, List[Tuple[str, str, str]]]:
        """
        Extrae todas las fórmulas del ODS.
        
        Args:
            ods_path: Ruta al archivo ODS
            locale: 'en' para inglés, 'es' para español
        
        Returns:
            Dict con estructura:
            {
                'C111': [('A1', 'original_formula', 'excel_formula'), ...],
                'Aulas': [...],
                ...
            }
        """
        namespaces = {
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
        }
        
        result = {}
        
        with zipfile.ZipFile(ods_path, 'r') as zip_ref:
            content_xml = zip_ref.read('content.xml')
            root = ET.fromstring(content_xml)
            
            tables = root.findall('.//table:table', namespaces)
            
            for sheet in tables:
                sheet_name = sheet.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name', 'Sin nombre')
                sheet_formulas = []
                
                rows = sheet.findall('.//table:table-row', namespaces)
                
                for row_idx, row in enumerate(rows, start=1):
                    cells = row.findall('.//table:table-cell', namespaces)
                    
                    for col_idx, cell in enumerate(cells, start=1):
                        formula = cell.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula')
                        
                        if formula:
                            # Calcular referencia de celda Excel
                            from openpyxl.utils import get_column_letter
                            cell_ref = f'{get_column_letter(col_idx)}{row_idx}'
                            
                            excel_formula = FormulaConverter.ods_to_excel_formula(formula, locale=locale)
                            sheet_formulas.append((cell_ref, formula, excel_formula))
                
                if sheet_formulas:
                    result[sheet_name] = sheet_formulas
        
        return result


def main():
    """Extrae y convierte las fórmulas"""
    
    converter = FormulaConverter()
    
    # Generar ambas versiones: inglés y español
    for locale, locale_name in [('en', 'English'), ('es', 'Español')]:
        print(f"\n{'='*80}")
        print(f"EXTRAYENDO FÓRMULAS EN {locale_name.upper()}")
        print(f"{'='*80}\n")
        
        formulas_dict = converter.extract_formulas_from_ods(
            'propuesta-de-horarios-fernando-v3-2026-04--02.ods',
            locale=locale
        )
        
        for sheet_name, formulas in formulas_dict.items():
            print(f"📊 HOJA: {sheet_name}")
            print(f"   Total de fórmulas: {len(formulas)}\n")
            
            for cell_ref, original, converted in formulas[:3]:  # Mostrar primeras 3
                print(f"   Celda: {cell_ref}")
                print(f"     Original:   {original}")
                print(f"     Convertida: {converted}")
                print()
            
            if len(formulas) > 3:
                print(f"   ... y {len(formulas) - 3} fórmulas más\n")
        
        # Guardar a JSON para usar desde LISP
        import json
        
        # Convertir a formato serializable
        output = {}
        for sheet_name, formulas in formulas_dict.items():
            output[sheet_name] = [
                {
                    'cell': cell_ref,
                    'original': original,
                    'excel': converted
                }
                for cell_ref, original, converted in formulas
            ]
        
        filename = f'formulas_fernando_convertidas_{locale}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Fórmulas guardadas en: {filename}\n")


if __name__ == '__main__':
    main()
