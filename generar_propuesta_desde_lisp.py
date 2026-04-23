from hoja_con_formulas import generar_excel_desde_parametros

grupos = []
horarios_por_grupo = {}
asignaturas_por_grupo = {}
aulas_por_dia = {}

grupos.append("D111")
horarios_por_grupo["D111"] = [["F", "ICD", "AL", "AM I", "AL"], ["Aula 8", "Aula 7", "Aula 7*", "Aula 7*", "Aula 7*"], ["", "", "", "", ""], ["L", "AL", "EF", "L", ""], ["Aula 6*", "Aula 6*", "SEDER", "Aula 7", ""], ["", "", "", "", ""], ["IP", "AM I", "AM I", "IP", ""], ["Aula 7", "Aula 6*", "Aula 7*", "Lab", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["D111"] = [["AL", "Álgebra Lineal", 3, 0, 3], ["L", "Lógica", 2, 0, 2], ["IP", "Introducción a la Programación", 2, 0, 2], ["AM I", "Análisis Matemático I", 2, 0, 2], ["ICD", "Introducción a la Ciencia de Datos", 2, 0, 2], ["F", "Filosofía", 2, 0, 2], ["EF", "Educación Física I", 2, 0, 2]]

grupos.append("D211")
horarios_por_grupo["D211"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["ED", "BD", "MA (EDO) cp 6 (con", "BD", ""], ["c 7", "c 7", "C211", "cp 7", ""], ["", "", "", "", ""], ["VD", "EP", "EP", "Prb", ""], ["c 7", "c 7", "c 7", "cp 7", ""], ["", "", "", "", ""], ["MA", "Prb", "EF 4:45pm a 5:35pm", "ED", ""], ["c 6", "c 7", "", "cp 7", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["D211"] = [["MA", "Matemática y Aplicaciones", 2, 0, 2], ["Prb", "Probabilidades", 2, 0, 2], ["BD", "Bases de Datos", 2, 0, 2], ["ED", "Estructura de Datos", 2, 0, 2], ["VD", "Visualización de Datos", 2, 0, 2], ["EP", "Economía Política", 2, 0, 2], ["EF", "Educación Física III", 2, 0, 2]]

grupos.append("D311")
horarios_por_grupo["D311"] = [["AE2", "RN", "TP 4", "", ""], ["c 2", "c 2", "", "", ""], ["", "", "", "", ""], ["MDE", "PL", "MDE", "RN", ""], ["c 2", "c 2", "cp 2", "cp Lab2", ""], ["", "", "", "", ""], ["PGVD", "TP", "AE2", "PGVD", ""], ["c 2", "c 4", "cp Lab2", "cp 7", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["D311"] = [["AE2", "Análisis Estadístico II", 2, 0, 2], ["MDE", "Muestreo y Diseño de Experimentos", 2, 0, 2], ["RN", "Redes Neuronales", 2, 0, 2], ["PL", "Procesamiento del Lenguaje", 2, 0, 2], ["PGVD", "Procesamiento de Grandes Volúmenes de Datos", 2, 0, 2], ["TP", "Teoría Política", 2, 0, 2]]

grupos.append("D411")
horarios_por_grupo["D411"] = [["", "EIA", "", "SN", ""], ["CP 2", "2", "CP 2", "C4", ""], ["", "", "", "", ""], ["IN", "SN", "IN", "SN", ""], ["2", "C4", "2", "C4", ""], ["", "", "", "", ""], ["ECTS 9 (con", "CO", "EIA 2", "", ""], ["C4", "2", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["D411"] = [["IN", "Inteligencia de Negocios", 2, 0, 2], ["EIA", "Elementos de Inteligencia Artificial", 2, 0, 2], ["CP", "Ciberseguridad y Privacidad", 2, 0, 2], ["CO2", "Curso Optativo II", 2, 0, 2], ["ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2], ["SN/DN", "Seguridad Nacional / Defensa Nacional", 2, 0, 2]]

grupos.append("C111")
horarios_por_grupo["C111"] = [["F", "A I", "AM I", "F", ""], ["Aula 6", "Aula 6*", "Aula 6*", "Aula 6", ""], ["", "", "", "", ""], ["L", "A I", "EF", "L", "A I"], ["Aula 6*", "Aula 6*", "SEDER", "Aula 6*", "Aula 6*"], ["", "", "", "", ""], ["P", "AM I", "AM I", "P", "P"], ["Aula 6", "Aula 6*", "Aula 6*", "Aula 6", "Lab"], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C111"] = [["A I", "Álgebra I", 3, 0, 3], ["L", "Lógica", 2, 0, 2], ["P", "Programación", 3, 0, 3], ["AM I", "Análisis Matemático I", 2, 0, 2], ["F", "Filosofía", 2, 0, 2], ["EF", "Educación Física I", 2, 0, 2]]

grupos.append("C121")
horarios_por_grupo["C121"] = [["P", "A I", "AM I", "P", "A I"], ["Aula 5", "Aula 5", "Aula 5", "Aula 5", "Aula 5"], ["", "", "", "", ""], ["F", "AM I", "EF", "L", "P"], ["Aula 5", "Aula 5", "SEDER", "Aula 5", "Lab"], ["", "", "", "", ""], ["L", "A I", "AM I", "F", ""], ["Aula 5", "Aula 5", "Aula 5", "Aula 5", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C121"] = [["A I", "Álgebra I", 3, 0, 3], ["L", "Lógica", 2, 0, 2], ["P", "Programación", 3, 0, 3], ["AM I", "Análisis Matemático I", 2, 0, 2], ["F", "Filosofía", 2, 0, 2], ["EF", "Educación Física I", 2, 0, 2]]

grupos.append("C122")
horarios_por_grupo["C122"] = [["P", "A I", "AM I", "L", "P"], ["Aula 5", "Aula 5", "Aula 1", "Aula 1", "Lab"], ["", "", "", "", ""], ["F", "AM I", "EF", "AM I", "A I"], ["Aula 5", "Aula 5", "SEDER", "Aula 1", "Aula 1"], ["", "", "", "", ""], ["L", "A I", "P", "F", ""], ["Aula 5", "Aula 1", "Aula 1", "Aula 5", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C122"] = [["A I", "Álgebra I", 3, 0, 3], ["L", "Lógica", 2, 0, 2], ["P", "Programación", 3, 0, 3], ["AM I", "Análisis Matemático I", 2, 0, 2], ["F", "Filosofía", 2, 0, 2], ["EF", "Educación Física I", 2, 0, 2]]

grupos.append("C211")
horarios_por_grupo["C211"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["EDA I", "MD", "EDO", "MN", "AC"], ["c 6", "c6", "cp 6", "cp 6", "lab"], ["", "", "", "", ""], ["TP", "MN", "EDA I", "TP", "MD"], ["c 6", "c 6", "cp 6", "c 6", "cp 6"], ["", "", "", "", ""], ["EDO", "AC", "EF 4:45pm a 5:35pm", "", ""], ["c 6", "c 6", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C211"] = [["EDA", "Estructuras de Datos y Algoritmos I", 2, 0, 2], ["MD", "Matemática Discreta I", 2, 0, 2], ["AC", "Arquitectura de computadoras", 2, 0, 2], ["EDO", "Ecuaciones Diferenciales Ordinarias", 2, 0, 2], ["MN", "Matemática Numérica", 2, 0, 2], ["TP", "Teoría Política", 2, 0, 2], ["EF3", "Educación Física III", 2, 0, 2]]

grupos.append("C212")
horarios_por_grupo["C212"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["EDA I", "MD", "EDO I", "MN", "AC"], ["c 6", "c6", "cp 5", "cp 5", "lab"], ["", "", "", "", ""], ["TP", "MN", "EDA I", "TP", "MD"], ["c 6", "c 6", "cp 5", "c 6", "cp 5"], ["", "", "", "", ""], ["EDO", "AC", "EF 4:45pm a 5:35pm", "", ""], ["c 6", "c 6", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C212"] = [["EDA", "Estructuras de Datos y Algoritmos I", 2, 0, 2], ["MD", "Matemática Discreta I", 2, 0, 2], ["AC", "Arquitectura de computadoras", 2, 0, 2], ["EDO", "Ecuaciones Diferenciales Ordinarias", 2, 0, 2], ["MN", "Matemática Numérica", 2, 0, 2], ["TP", "Teoría Política", 2, 0, 2], ["EF3", "Educación Física III", 2, 0, 2]]

grupos.append("C311")
horarios_por_grupo["C311"] = [["BD2", "Est", "BD2 cp", "MO", ""], ["Aula 9", "Aula 9", "Aula 9", "Aula 9", ""], ["", "", "", "", ""], ["IS c", "PD c", "IS c", "PD cp", ""], ["Aula 9", "Aula 9", "Aula 9", "Aula 9", ""], ["", "", "", "", ""], ["RC", "MO", "RC cp", "Est cp", ""], ["Aula 9", "Aula 9", "Aula 9", "Aula 9", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C311"] = [["RC", "Redes de Computadoras", 2, 0, 2], ["IS", "Ingeniería de Software", 2, 0, 2], ["MO", "Modelos de Optimización", 2, 0, 2], ["BD2", "Bases de Datos II", 2, 0, 2], ["PD", "Programación Declarativa", 2, 0, 2], ["Est", "Estadística", 2, 0, 2]]

grupos.append("C312")
horarios_por_grupo["C312"] = [["BD2", "Est", "BD2 cp", "PD cp", ""], ["Aula 9", "Aula 9", "Aula 3", "Aula 3", ""], ["", "", "", "", ""], ["IS c", "PD c", "IS c", "MO", ""], ["Aula 9", "Aula 9", "Aula 3", "Aula 3", ""], ["", "", "", "", ""], ["RC", "MO", "RC cp", "Est cp", ""], ["Aula 9", "Aula 9", "Aula 3", "Aula 3", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C312"] = [["RC", "Redes de Computadoras", 2, 0, 2], ["IS", "Ingeniería de Software", 2, 0, 2], ["MO", "Modelos de Optimización", 2, 0, 2], ["BD2", "Bases de Datos II", 2, 0, 2], ["PD", "Programación Declarativa", 2, 0, 2], ["Est", "Estadística", 2, 0, 2]]

grupos.append("C411")
horarios_por_grupo["C411"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["AM", "DAA", "DAA", "SN", ""], ["9", "9", "9", "", ""], ["", "", "", "", ""], ["SD 9", "SN", "AM", "SN", ""], ["", "", "9", "", ""], ["", "", "", "", ""], ["ECTS 9", "AE", "SD", "", ""], ["", "9", "9", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C411"] = [["AM", "Aprendizaje de Máquinas", 2, 0, 2], ["DAA", "Diseño y Análisis de Algoritmos", 2, 0, 2], ["SD", "Sistemas Distribuidos", 2, 0, 2], ["AE", "Asignatura Electiva", 2, 0, 2], ["ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2], ["SN/DN", "Seguridad Nacional / Defensa Nacionaol", 2, 0, 2]]

grupos.append("C412")
horarios_por_grupo["C412"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["AM", "DAA", "DAA", "SN", ""], ["9", "9", "9", "", ""], ["", "", "", "", ""], ["SD 9", "SN", "AM", "SN", ""], ["", "", "1", "", ""], ["", "", "", "", ""], ["ECTS 9", "AE", "SD", "", ""], ["", "", "9", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["C412"] = [["AM", "Aprendizaje de Máquinas", 2, 0, 2], ["DAA", "Diseño y Análisis de Algoritmos", 2, 0, 2], ["SD", "Sistemas Distribuidos", 2, 0, 2], ["AE", "Asignatura Electiva", 2, 0, 2], ["ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2], ["SN/DN", "Seguridad Nacional / Defensa Nacionaol", 2, 0, 2]]

grupos.append("M111")
horarios_por_grupo["M111"] = [["F", "IM", "IAM", "IA", "IA"], ["Aula 8", "Aula 8", "Aula 8", "Aula 8", "Aula 8"], ["", "", "", "", ""], ["PA", "IAM", "EF", "PA Aula", "IAM"], ["Aula 8", "Aula 8", "SEDER", "Lab", "Aula 8"], ["", "", "", "", ""], ["GA", "GA", "IA", "GA 8", ""], ["Aula 8", "Aula 8", "Aula 8", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["M111"] = [["IAM", "Introducción al Análisis Matemático", 3, 0, 3], ["IA", "Introducción al Álgebra", 3, 0, 3], ["GA", "Geometría Analítica", 2, 0, 2], ["PA", "Programación y Algoritmos", 2, 0, 2], ["IM", "Introducción a la Matemática", 2, 0, 2], ["F", "Filosofía", 2, 0, 2], ["EF", "Educación Física I", 2, 0, 2]]

grupos.append("M211")
horarios_por_grupo["M211"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["CAL", "CAL", "", "", ""], ["c 3", "c 3", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["FVV", "FVV", "FVV", "CAL", ""], ["c 3", "c 3", "cp 3", "cp 3", ""], ["", "", "", "", ""], ["EP", "EP", "FVV", "", ""], ["c 7", "c 7", "cp 3", "", ""], ["", "", "", "", ""], ["EF 4:45pm a 5:35pm", "SP", "", "", ""], ["", "2", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["M211"] = [["FVV", "Funciones de Varias Variables", 4, 0, 4], ["CAL", "Complementos de Álgebra Lineal", 3, 0, 3], ["SP2", "Seminario de Problemas II", 2, 0, 2], ["AE", "Asignatura Electiva I", 2, 0, 2], ["EP", "Economía Política", 2, 0, 2], ["EF", "Educación Física III", 2, 0, 2]]

grupos.append("M311")
horarios_por_grupo["M311"] = [["MN", "EDO", "MN", "FV", "FVC"], ["Aula 4", "Aula 4", "4", "C 4", "Aula 4"], ["", "", "", "", ""], ["FVC", "OM", "EDO", "IE", "EDO"], ["Aula 4", "Aula 4", "Aula 4", "Aula 4", "Aula 4"], ["", "", "", "", ""], ["IE", "TP", "IE", "OM", ""], ["Aula 4", "Aula 4", "Aula 4", "Aula 4", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["M311"] = [["FVC", "Funciones de variable Compleja", 2, 0, 2], ["IE", "Inferencia Estadística", 3, 0, 3], ["EDO", "Ecuaciones Diferenciales Ordinarias", 3, 0, 3], ["MN", "Matemática Numérica", 2, 0, 2], ["OM", "Optimización Matemática I", 2, 0, 2], ["TP", "Teoría Política", 2, 0, 2]]

grupos.append("M411")
horarios_por_grupo["M411"] = [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["GD", "HM", "MI", "AO", ""], ["4", "4", "4", "2", ""], ["", "", "", "", ""], ["MI", "HM", "GD", "AO", ""], ["4", "4", "4", "3", ""], ["", "", "", "", ""], ["ECTS 9 (Con", "GD", "", "", ""], ["C4", "4", "", "", ""], ["", "", "", "", ""]]
asignaturas_por_grupo["M411"] = [["MI", "Medida e Integración", 2, 0, 2], ["GD", "Geometría Diferencial", 2, 0, 2], ["HM", "Historia de la Matemática", 2, 0, 2], ["ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2], ["AO2", "Asignatura Optativa II", 2, 0, 2], ["AO3", "Asignatura Optativa III", 2, 0, 2]]

aulas_por_dia["Lunes"] = [["", "", "", "", "", "C111,C112", "", "", "", "C113"], ["", "", "", "", "C112", "C113", "", "", "", ""], ["", "", "", "", "", "C111,C113", "", "C112", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""]]
aulas_por_dia["Martes"] = [["", "C112", "", "", "", "C111", "C113", "", "", ""], ["C111", "", "", "", "", "C112,C113", "", "", "", ""], ["", "", "", "", "", "C111,C112,C113", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""]]
aulas_por_dia["Miércoles"] = [["", "", "", "", "C112", "C111", "", "C113", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "C112", "", "", "C113", "C111", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""]]
aulas_por_dia["Jueves"] = [["C113", "", "", "", "C112", "C111", "", "", "", ""], ["C113", "", "", "", "C112", "C111", "", "", "", ""], ["C113", "", "", "", "C112", "C111", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""]]
aulas_por_dia["Viernes"] = [["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "C112"], ["", "", "", "", "C111", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""]]
if not grupos: grupos = ["D111", "D211", "D311", "D411", "C111", "C121", "C122", "C211", "C212", "C311", "C312", "C411", "C412", "M111", "M211", "M311", "M411"]

generar_excel_desde_parametros(
    filename="propuesta_horarios_desde_lisp.xlsx",
    grupos=grupos,
    horarios_por_grupo=horarios_por_grupo,
    asignaturas_por_grupo=asignaturas_por_grupo,
    aulas_por_dia=aulas_por_dia,
    turnos=6,
    horario_row_step=3
)
