from fileinput import filename
from pickle import GET
from random import sample
from flask import Flask, jsonify, request, render_template, send_file, send_from_directory, session
from importlib.util import spec_from_file_location
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment, Protection
from openpyxl.utils import get_column_letter
from copy import copy
from  datetime import datetime
from pkg_resources import run_script
from  werkzeug.utils import secure_filename

import threading
import openpyxl 
import sys
import os
import pandas as pd
import pathlib
import subprocess
import importlib.util
import openpyxl
import secrets
import requests


app = Flask(__name__)

app.secret_key = 'UtilTech' #LLAVE SECRETA PARA EL PASE DE DATOS
app.secret_key = secrets.token_hex(16)

#DEFINICION DE LAS VARIABLES GLOBALES
form_data = None
archivoGlo1 = None
archivoGlobal1 = None
list_Archivos = None
archivo_txt = None
workbook = None
ArchiGenerado = None
grupos = []
trimestre = '24-O'
excell_resultado = 'PlantillaAsignacionCoeficientesPT'+'-'+trimestre+'.xlsx'
profesores = []

#DIRECTORIO DE LA RUTAS
#RUTA DONDE SE ENCUENTRA EL PROYECTO
directorio_automatico = os.path.dirname(os.path.abspath(__file__))

#RUTA DONDE SE ENCUENTRAN LOS ARCHIVOS
directorio_archivos_subidos = f'{directorio_automatico}\\GruposDePTExcell\\ArchivosSubidos'
directorio_archivos_generados = f'{directorio_automatico}\\GruposDePTExcell\\ArchivosGenerados'
directorio_archiGeneradosGruposDePtExcell = f'{directorio_automatico}\\GruposDePTExcell\\ArchivosGenerados' #Añadido

#RUTA DE LA FUNCION GRUPOS DE PT
directorio_gruposDePtExcell = directorio_automatico + "\\GruposDePTExcell"
#RUTA DE LA FUNCION OFICIOS CARGA ACADEMICA
directorio_oficiosCargaAcademica = directorio_automatico + "\\OficiosCargaAcademica"
#RUTA DE LA FUNCION OFICIOS TUTORIA PT ASESORIA
directorio_oficiosDocAsesoria = directorio_automatico + "\\OficiosTutoriaPT\\docAsesoria"
#RUTA DE LA FUNCION OFICIOS TUTORIA PT PROYECTO TERMINAL
directorio_oficiosDocGenTem = directorio_automatico + "\\OficiosTutoriaPT\\documentGen"


#RUTA DE LAS PLANTILLAS A LLENAR POR LOS MODULOS
plantilla1 = directorio_gruposDePtExcell + "\\Plantilla\\PlantillaAsignacionCoeficientesPT-24-O.xlsx"

plantilla2 = directorio_oficiosCargaAcademica + "\\Plantilla\\PlantillasAsignacionCoeficientesPT-24-0.xlsx" 

plantilla3 = directorio_oficiosDocAsesoria + "\\Plantilla\\PlantillasAsignacionCoeficientesPT-24-0.xlsx"

plantilla4 = directorio_oficiosDocGenTem + "\\Plantilla\\PlantillasAsignacionCoeficientesPT-24-0.xlsx"


#RUTA HASTA LOS ARCHIVOS SUBIDOS POR LOS USUARIOS
archivosSubidos1_folder = os.path.join(os.path.dirname(__file__), 'GruposDePTExcell\\ArchivosSubidos\\')

archivosSubidos2_folder = os.path.join(os.path.dirname(__file__), 'OficiosCargaAcademica\\ArchivosSubidos\\')

archivosSubidos3_folder = os.path.join(os.path.dirname(__file__), 'OficiosTutoriaPT\\docAsesoria\\ArchivosSubidos\\')

archivosSubidos4_folder = os.path.join(os.path.dirname(__file__), 'OficiosTutoriaPT\\documentGen\\ArchivosSubidos\\')

#RUTA HASTA LOS ARCHIVOS GENERADOS POR LOS USUARIOS
archivosGenerados1_folder = os.path.join(os.path.dirname(__file__), 'GruposDePTExcell\\ArchivosGenerados\\')

archivosGenerados2_folder = os.path.join(os.path.dirname(__file__), 'OficiosCargaAcademica\\ArchivosGenerados\\')

archivosGenerados3_folder = os.path.join(os.path.dirname(__file__), 'OficiosTutoriaPT\\docAsesoria\\ArchivosGenerados\\')

archivosGenerados4_folder = os.path.join(os.path.dirname(__file__), 'OficiosTutoriaPT\\documentGen\\ArchivosGenerados\\')

#VARIABLE DEL TIPO DE CARPETA SUBIDOS/GENERADOS
tipo = None
form_id = None
list_Archivos = None


form_data = {
    'formulario1': {
        'file_key': 'asignacionesTXT',
        'subidos': archivosSubidos1_folder,
        'generados': archivosGenerados1_folder,
    },
    'formulario2': {
        'file_key': 'PAEG_TySI',
        'subidos': archivosSubidos2_folder,
        'generados': archivosGenerados2_folder,
    },
    'formulario3': {
        'file_key': 'asignaciones1TXT',
        'subidos': archivosSubidos3_folder,
        'generados': archivosGenerados3_folder,
    },
    'formulario4': {
        'file_key': 'asignaciones2TXT',
        'subidos': archivosSubidos4_folder,
        'generados': archivosGenerados4_folder,
    },
   }

# Asegurarse de que las carpetas existen
for form in form_data.values():
    os.makedirs(form['subidos'], exist_ok=True)
    os.makedirs(form['generados'], exist_ok=True)

#---------------------------------------------------------------------------
#---------------------------------------------------------------------------
#SECCION DE LA FUNCION QUE CARGA EL INDEX
@app.route('/', methods=['GET','POST'])
def index():
    return render_template('index.html', 
                           list_Archivos=listaArchivos("formulario1", "subidos"),)
                           #file_path=uploaded_file_path)

#---------------------------------------------------------------------------
#---------------------------------------------------------------------------
#SECCION DE LA FUNCION QUE SE ENCRAGA DE LA SUBIDA DE LOS ARCHIVOS
@app.route('/upload', methods=['POST'])
def upload_file():
    #Implementacion de las variables globales
    global form_data
    global archivoGlo1
    global list_Archivos

    form_id = request.form.get('form_id')
    #Impresion de la seleccion del usuario en el form a ejecutarse de archivos a subirse
    print("Este es el ID del formulario a ejecutarse: ", form_id)

    if form_id in form_data:
        file_key = form_data[form_id]['file_key']
        folder = form_data[form_id]['subidos']
        tipo = 'subidos'
        file = request.files[file_key]

        #PARTE DEL FORMATO DEL NOMBRE QUE TOMARA EL ARCHIVO SUBIDO CON EL RENOMBRE
        timestamp =datetime.now().strftime('%Y%m%d_%H%M%S')

        filename = secure_filename(file.filename)
        name, ext = os.path.splitext(filename) #SEPARA EL NOMBRE DE SU TIPO DE EXTENSION
        new_filename = f"{name}_{timestamp}{ext}" #SE PROCEDE A RENOMBRAR EL ARCHIVO SUBIDO POR EL USUARIO
        upload_path = os.path.join(directorio_automatico, folder, new_filename) 
        file.save(upload_path)
        archivoGlo1 = upload_path #SE ASIGNA LA RUTA DEL ARCHIVO GUARDADO A TRAVES DE LA VARIABLE GLOBAL

        #Impresion de los archivos subidos
        print(f"Archivo subido correctamente: {file}")
        print(f"El archivo se guardó en {upload_path}")
        print(f"Archivo con time: {new_filename}")
        print(f"Archivo con time y ruta de directorio: {archivoGlo1}") 

        return render_template('index.html',
                                #FOLDERS DONDE SE ALMACENARON LOS ARCHIVOS SUBIDOS
                                archivosGenerados_folder=folder,
                                #SE PASA EL FORM SELECCIONADO
                                form_id=form_id,
                                #FUNCION PARA LA VISUALIZACION DE LOS ARCHIVOS EN LA CARPETA
                                list_Archivos = listaArchivos(form_id, tipo),
                                #NOMBRES DE LOS ARCHIVOS SUBIDOS
                                archivo1 = file.filename,
                                #MENSAJES DONDE SE ALOJO EL ARCHIVO SUBIDO
                                message = f"Y este se guardo en:\n{upload_path}",
                                #SE DEJA VER AL USUARIO EL DIV CORRESPONDIENTE A LA RUTA DONDE SE GUARDO
                                div9_block = True
                               )

#---------------------------------------------------------------------------
#---------------------------------------------------------------------------
#SECCION DE LA FUNCION QUE SE ENCRAGA DE LA SUBIDA DE LOS ARCHIVOS
def listaArchivos(form_id, tipo):
    global form_data
    print(f"Formulario: {form_id}, Tipo: {tipo}")
    print(f"Formulario: {form_data}")
    if form_id in form_data and tipo in form_data[form_id]:
        path = form_data[form_id][tipo]
        # Si el tipo es un diccionario, se listan los archivos en cada subdirectorio
        if isinstance(path, dict):
            print(f"Archivo SUBIDO / GENERADO con time y ruta de directorio: {path}")
            return {key: os.listdir(sub_path) for key, sub_path in path.items()}
        
        # Si es una ruta simple, se lista directamente
        return os.listdir(path)
    
    return []  # Retorna una lista vacía si el form_id o tipo no son válidos

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION QUE PERMITE ELEMINAR LOS ARCHIVOS SUBIDOS POR EL USUARIO
@app.route('/deleteSub/<filename>', methods=['GET'])
def delete_file_subidos(filename):
    file_path = os.path.join(directorio_archivos_subidos, filename)

    # Aquí debes definir los valores para form_id y tipo
    form_id = 'formulario1'  # Cambia esto según tu lógica
    tipo = 'subidos'  # Cambia esto según tu lógica
    if os.path.isfile(file_path):
        try:
            os.remove(file_path)
            return render_template('index.html',
                               list_Archivos = listaArchivos(form_id, tipo),
                               show_subir = True)
        except Exception as e:
            return jsonify({'error': f'No se pudo eliminar el archivo: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Archivo no encontrado'}), 404
    
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION QUE PERMITE ELEMINAR LOS ARCHIVOS GENERADOS POR EL USUARIO
@app.route('/deleteGen/<filename>', methods=['POST'])
def delete_file_generados(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER1'], filename)

    # Aquí debes definir los valores para form_id y tipo
    form_id = 'formulario1'  # Cambia esto según tu lógica
    tipo = 'generados'  # Cambia esto según tu lógica
    if os.path.isfile(file_path):
        try:
            os.remove(file_path)
            return render_template('index.html',
                               list_Archivos = listaArchivos(form_id, tipo),
                               show_subir = True)
        except Exception as e:
            return jsonify({'error': f'No se pudo eliminar el archivo: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Archivo no encontrado'}), 404
    
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION
    
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION QUE TRANSFORMA EL ARCHIVO CSV SUBIDO A ARCHIVO TXT CON LA ESTRUCTURA
def procesar_archivo_csv(archivoGlo1):
    global archivo_txt
    # Definir el nombre del archivo TXT que se generará
    archivo_txt = f"{archivoGlo1.split('.')[0]}.txt".replace('ArchivosSubidos', 'ArchivosGenerados')
    try:
        # Ejecutar la función para la generación del archivo
        convierte_csv_a_txt(archivoGlo1)
        # Retorna la ruta completa del archivo generado TXT
    except Exception as e:
        print(f"Error al procesar el archivo CSV: {e}")
    return archivo_txt

def convierte_csv_a_txt(archivoGlo1):
    global archivo_txt
    try:
        # Lee el archivo CSV
        print(f"Leyendo el archivo CSV: {archivoGlo1}")
        df = pd.read_csv(archivoGlo1)
        print("Archivo CSV leído correctamente")
    except Exception as e:
        raise ValueError(f"Error al leer el archivo CSV: {e}")

    try:
        # Leer el archivo TXT para escritura
        print(f"Abriendo el archivo TXT para escritura: {archivo_txt}")
        with open(archivo_txt, 'w', encoding='utf-8') as f:
            for index, row in df.iterrows():
                try:
                    # Obtener el Titulo del proyecto terminal
                    if row['TituloProyecto']:
                        titulo = row['TituloProyecto']
                    else:
                        raise ValueError("El título del proyecto no está disponible")
                    print(f"Título del proyecto: {titulo}")

                    # Obtener los asesores
                    asesores = []
                    if pd.notna(row['ApellidosDelAsesor1']) and pd.notna(row['NombresDelAsesor1']):
                        asesores.append(f"{row['ApellidosDelAsesor1']} {row['NombresDelAsesor1']}")
                    if pd.notna(row['ApellidosDelAsesor2']) and pd.notna(row['NombresDelAsesor2']):
                        asesores.append(f"{row['ApellidosDelAsesor2']} {row['NombresDelAsesor2']}")
                    if pd.notna(row['ApellidosDelAsesor3']) and pd.notna(row['NombresDelAsesor3']):
                        asesores.append(f"{row['ApellidosDelAsesor3']} {row['NombresDelAsesor3']}")
                    print(f"Asesores: {asesores}")

                    # Obtener los revisores
                    revisores = []
                    if pd.notna(row['ApellidosDelRevisor1']) and pd.notna(row['NombresDelRevisor1']):
                        revisores.append(f"{row['ApellidosDelRevisor1']} ({row['NombresDelRevisor1']})")
                    if pd.notna(row['ApellidosDelRevisor2']) and pd.notna(row['NombresDelRevisor2']):
                        revisores.append(f"{row['ApellidosDelRevisor2']} ({row['NombresDelRevisor2']})")
                    print(f"Revisores: {revisores}")

                    # Obtener alumno y matricula
                    matricula_nombre = row['MatriculayNombre']
                    if not matricula_nombre:
                        raise ValueError("La matrícula y el nombre del alumno no están disponibles")
                    print(f"Matrícula y nombre: {matricula_nombre}")

                    # Separación del nombre y la matrícula
                    if ' - ' in matricula_nombre:
                        matricula, alumno = matricula_nombre.split(' - ')
                    else:
                        parts = matricula_nombre.split(' ')
                        matricula = parts[0]
                        alumno = ' '.join(parts[1:])
                    print(f"Alumno: {alumno}, Matrícula: {matricula}")

                    # Eliminación del valor PT repetitivo
                    pt = row['PT']
                    if not pt:
                        raise ValueError("El valor de PT no está disponible")
                    pt = pt.replace("PT", "").strip()
                    print(f"PT: {pt}")

                    # Escribir la estructura de cada alumno de proyecto terminal
                    f.write(f"Titulo: {titulo}\n")
                    f.write(f"Asesor(es): {', '.join(asesores)}\n")
                    f.write(f"Revisor(es): {', '.join(revisores)}\n")
                    f.write(f"Alumno: {alumno}\n")
                    f.write(f"Matrícula: {matricula}\n")
                    f.write(f"PT: {pt}\n\n")
                    print(f"Fila {index} procesada correctamente")
                except Exception as e:
                    raise ValueError(f"Error al procesar la fila {index}: {e}")
        print("Archivo TXT generado correctamente")
    except Exception as e:
        raise ValueError(f"Error al escribir el archivo TXT: {e}")

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION QUE SE ENCARGA DE COLOCAR EL NUMERO ECONOMICO A LOS PROFESORES
def get_numero_economico(nombre_del_profe, codigoGen1):
    try:
        if len(profesores) == 0:
            configProfesores()
        # Busca el número económico del profesor
        numero_economico = None
        for profe in profesores:
            if clean_name(nombre_del_profe) in clean_name(profe['nombre']):
                numero_economico = profe['numero_economico']
        if numero_economico is None:
            raise ValueError(f"Número económico no encontrado para el profesor {nombre_del_profe}")
        return numero_economico
    except Exception as e:
        raise RuntimeError(f"Error al obtener el número económico: {e}")
    """ 
    #Fila y Columna del profesor
    fila_encontrada = None
    columna_encontrada = None
    try:
        if nombre_del_profe is None:
            # Lanzar error si no se proporciona el nombre del profesor
            raise ValueError("No se ha proporcionado el nombre del profesor.")
        
        iterRowsOfProfesores = hoja_profesores.iter_rows(values_only=True)
        enumerateProfesores = enumerate(iterRowsOfProfesores)
        # print(f"Buscando el número económico del profesor: {nombre_del_profe}")
        # print(f"Recorriendo la hoja 'Profesores'")
        # print("iterRowsOfProfesores: ", iterRowsOfProfesores)
        # print("enumerateProfesores: ", enumerateProfesores)
        # Recorre la hoja 'Profesores'
        for fila_index, fila in enumerateProfesores:
            for columna_index, valor in enumerate(fila):
                if valor is None:
                    continue
                elif valor is not None and type(valor) == str and nombre_del_profe in valor:
                    #Se almacena la posicion del valor encontrado
                    fila_encontrada = fila_index + 1 #Se suma uno porque se comienza en 0
                    columna_encontrada = columna_index + 1 #Se suma uno porque se comienza en 0
                    break #Termina al encontrar el nombre el bucle
        print(f"La fila es: {fila_encontrada}")
        print(f"La columna es: {columna_encontrada}")
        # print(f"El profesor {nombre_del_profe} se encuentra en la fila {fila_encontrada} y columna {columna_encontrada}")
        if fila_encontrada is None or columna_encontrada is None:
            raise ValueError(f"No se ha encontrado el nombre del profesor: {nombre_del_profe}")

        #Coloca el numero economico al encontrar la fila y columna
        numero_economino_celda = hoja_profesores.cell(row=fila_encontrada, column=columna_encontrada + 1).value
        if numero_economino_celda is None:
            raise ValueError(f"No se ha encontrado el número económico para el profesor: {nombre_del_profe}")
        # print(f"El número económico del profesor {nombre_del_profe} es: {numero_economino_celda}")
        return int(numero_economino_celda) #Regresa el valor del numero economico
    except Exception as e:
        print(f"Error al obtener el número económico: {e}")
        return "Error al obtener el número económico." """
    
def configProfesores():
    global workbook
    workbook = openpyxl.load_workbook(plantilla1)
    #--------------------------------------------------------------
    #HOJA DE LOS PROFESORES
    hoja_profesores = workbook['Profesores']
    #Recorrer la hoja de profesores
    for cell in hoja_profesores.iter_rows(min_row=3, max_row=hoja_profesores.max_row, min_col=1, max_col=2, values_only=True):
        if cell[0] is not None and cell[1] is not None:
            profesores.append({'nombre': cell[0], 'numero_economico': cell[1]})
        else:
            break

# Función para pasar los nombres a minusculas, sin espacios y sin acentos
def clean_name(name):
    name = name.lower()
    name = name.replace(' ', '')
    name = name.replace('á', 'a')
    name = name.replace('é', 'e')
    name = name.replace('í', 'i')
    name = name.replace('ó', 'o')
    name = name.replace('ú', 'u')
    return name
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION QUE MANEJA EL NOMBRE DE LOS PROFESORES EN CASO DE NO TENER SEGUNDO NOMBRE
def manejo_de_nombre(nombre):
    try:
        if '.' in nombre:
            nombre = nombre.split(' ', 1)[-1]
        else:
            pass
        nombre_temporal = nombre.split(' ')  # Separa la estructura del nombre
        # Si el ultimo elemento es igual a SIN, se quita
        if nombre_temporal[-1] == 'SIN':
            nombre_temporal.pop(-1)
        # Juntando las secciones del nombre
        nombre_final = ' '.join(nombre_temporal)
        # Regresa el nombre ya procesado
        return nombre_final
    except Exception as e:
        print(f"Error al manejar el nombre: {e}")
        return nombre  # Retorna el nombre original en caso de error
    
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION PARA GENERAR EL DOCUMENTO TXT QUE CONTENDRA EL NOMBRE DE LOS ESTUDIANTES Y A QUE PT PERTENECEN
def generar_txt_estudiantes_PT():
    print('Generando el archivo Estudiantes de PT')
    try:
        global archivo_txt
        # Quitar extensión a archivo_txt
        archivo_txt_sin_extension = archivo_txt.split('.')[0]
        print(f"Archivo TXT sin extensión: {archivo_txt_sin_extension}")
        # Crear el archivo TXT
        with open(f"{archivo_txt_sin_extension}_Estudiantes_PT.txt", "w", encoding="UTF-8") as file:
            grupos_de_PT = ['1', '2', '3']
            # Loop para agregar la información
            for element in grupos_de_PT:
                file.write('\nPT' + element + '\n')
                for grupo in grupos:
                    if grupo['PT'] == element:
                        for alumno in grupo['Alumno(s)']:
                            file.write(manejo_de_nombre(alumno['nombre']) + ' ' + alumno['matrícula'] + '\n')
        print("Archivo Estudiantes de PT generado correctamente")
    except Exception as e:
        print(f"Error al generar el archivo Estudiantes de PT: {e}")

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION PARA EL MANEJO DE LOS GRUPOS
def manejo_grupo(proyecto : dict):
    #Agrega la informacion del proyecto actual a LA LISTA DE GRUPOS
    grupo = {
        'Asesor(es)':[],
        'Alumno(s)':[],
        'PT':''
    }
    print('Manejando el grupo')
    print('grupos:', grupos)
    print('proyecto:', proyecto)
    #Revisa si la lista de grupos ESTA VACIA
    if(len(grupos) == 0):
        config_grupos_from_0(proyecto, grupo)
    else:
        # Verificando si hay algún elemento en 'grupos' con 'Asesor(es)' y 'PT' iguales a los del proyecto
        indice_coincidencia = next((indice for indice, grupo in enumerate(grupos) if grupo['Asesor(es)'] == proyecto['Asesor(es)'] and grupo['PT'] == proyecto['PT']), None)
        print('Indice de coincidencia:', indice_coincidencia)
        if indice_coincidencia is not None:
            ### agregando el alumno al mismo grupo de PT
            grupo_coincidente = grupos[indice_coincidencia]
            grupo_coincidente['Alumno(s)'].append({'nombre' : proyecto['Alumno'], 'matrícula': proyecto['Matricula']})
            print('grupo_coincidente:', grupo_coincidente)
        else:
            #Si no son del mismo grupo de PT
            config_grupos_from_0(proyecto, grupo)

def config_grupos_from_0(proyecto, grupo):
    grupo['Asesor(es)'] = proyecto['Asesor(es)']
    grupo['Alumno(s)'].append( {'nombre' : proyecto['Alumno'], 'matrícula': proyecto['Matricula']} )
    grupo['PT'] = proyecto['PT']
    grupos.append(grupo)
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION PARA CREAR LA TABLA EN EL ARCHIVO XLSX
def crear_tabla(row_number, column_number, hoja, info):
    #Estilos globales de la tabla a crear
    alignment_header_tabla = Alignment(horizontal='center', vertical='center')
    font_header_tabla = Font(name='Arial', size=8, bold=True, italic=False, color='000000')
    font_header_info = Font(name='Arial', size=7, color='000000')
    font_tabla_info = Font(name='Arial', size=11, color='000000')
    border_header_tabla = Border(left=Side(style='medium'), 
                     right=Side(style='medium'), 
                     top=Side(style='medium'), 
                     bottom=Side(style='medium'))
    horario_del_curso='14-17'

    # Nombre del curso
    hoja.merge_cells(start_row=row_number, start_column=column_number, end_row=(row_number+1), end_column=column_number)
    celda = hoja.cell(row=row_number, column=column_number)


    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'NOMBRE DEL CURSO'

        # Clave
    hoja.merge_cells(start_row=row_number, start_column=(column_number+1), end_row=(row_number+1), end_column=(column_number+1))
    celda = hoja.cell(row=row_number, column=column_number+1)

    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'CLAVE'


    # Grupo
    hoja.merge_cells(start_row=row_number, start_column=(column_number+2), end_row=(row_number+1), end_column=(column_number+2))
    celda = hoja.cell(row=row_number, column=column_number+2)


    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'GRUPO'


    # Cupo 
    celda = hoja.cell(row=row_number, column=column_number+3)

    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'CUPO'

    # Max 
    celda = hoja.cell(row=row_number+1, column=column_number+3)

    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'MAX.'


    # Profesor
    hoja.merge_cells(start_row=row_number, start_column=(column_number+4), end_row=(row_number+1), end_column=(column_number+4))
    celda = hoja.cell(row=row_number, column=column_number+4)
    
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'PROFESOR'


    # No. eco
    hoja.merge_cells(start_row=row_number, start_column=(column_number+5), end_row=(row_number+1), end_column=(column_number+5))

    celda = hoja.cell(row=row_number, column=column_number+5)
    
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    celda.value = 'No. ECON.'


    # Horario y Aula
    hoja.merge_cells(start_row=row_number, start_column=column_number+6, end_row=(row_number), end_column=column_number+10)
    hoja.cell(row=row_number, column=(column_number+6), value='HORARIO  Y  AULA')
    celda = hoja.cell(row=row_number, column=(column_number+6))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla
    hoja.cell(row=row_number, column=(column_number+10)).border = border_header_tabla

    # LUNES, MARTES, MIERCOLES, JUEVES, VIERNES
    hoja.cell(row=row_number+1, column=(column_number+6), value='LUNES')
    celda = hoja.cell(row=row_number+1, column=(column_number+6))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla

    hoja.cell(row=row_number+1, column=(column_number+7), value='MARTES')
    celda = hoja.cell(row=row_number+1, column=(column_number+7))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla

    hoja.cell(row=row_number+1, column=(column_number+8), value='MIÉRCOLES')
    celda = hoja.cell(row=row_number+1, column=(column_number+8))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla


    hoja.cell(row=row_number+1, column=(column_number+9), value='JUEVES')
    celda = hoja.cell(row=row_number+1, column=(column_number+9))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla

    hoja.cell(row=row_number+1, column=(column_number+10), value='VIERNES')
    celda = hoja.cell(row=row_number+1, column=(column_number+10))
    celda.alignment = alignment_header_tabla
    celda.font = font_header_tabla
    celda.border = border_header_tabla



    if info['PT'] == '1': 
        pt = 'I' 
        horas_por_UEA = 9
    elif info['PT'] == '2': 
        pt = 'II'
        horas_por_UEA = 10
    else: 
        pt = 'III'
        horas_por_UEA = 10

    ### tabla observaciones
    if (info['Asesor(es)'] != 'POR ASIGNAR'):

        hoja.cell(row=row_number+1, column=(column_number+12), value='OBSERVACIONES')
        celda = hoja.cell(row=row_number+1, column=(column_number+12))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        hoja.cell(row=row_number+1, column=(column_number+13), value='Horas por alumno')
        celda = hoja.cell(row=row_number+1, column=(column_number+13))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla
        
        hoja.cell(row=row_number+1, column=(column_number+14), value='Alumnos')
        celda = hoja.cell(row=row_number+1, column=(column_number+14))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla
        
        hoja.cell(row=row_number+1, column=(column_number+15), value='Coeficiente por hr')
        celda = hoja.cell(row=row_number+1, column=(column_number+15))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        hoja.cell(row=row_number+1, column=(column_number+16), value='Coeficiente de participación')
        celda = hoja.cell(row=row_number+1, column=(column_number+16))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        hoja.cell(row=row_number+1, column=(column_number+17), value='Horas parciales')
        celda = hoja.cell(row=row_number+1, column=(column_number+17))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        hoja.cell(row=row_number+1, column=(column_number+18), value='Horas por UEA')
        celda = hoja.cell(row=row_number+1, column=(column_number+18))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        ### info llenado
        hoja.cell(row=row_number+2, column=(column_number+12), value='Asignación de profesor e inscripción de alumnos')
        celda = hoja.cell(row=row_number+2, column=(column_number+12))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        ### horas por alumnos valor
        hoja.cell(row=row_number+2, column=(column_number+13), value=2)
        celda = hoja.cell(row=row_number+2, column=(column_number+13))
        celda_horas_por_alumno = celda.value
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        ### Alumnos valor
        cantidad_de_alumnos = len(info['Alumno(s)']) 

        hoja.cell(row=row_number+2, column=(column_number+14), value=cantidad_de_alumnos)
        celda = hoja.cell(row=row_number+2, column=(column_number+14))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla
        
        ### Coeficiente por hr
        ### valor = horas por alumno / horas por UEA
        hoja.cell(row=row_number+2, column=(column_number+15), value=( round(celda_horas_por_alumno / horas_por_UEA, 2)))
        celda = hoja.cell(row=row_number+2, column=(column_number+15))
        celda_coeficiente_por_hora = celda.value
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla


        ### Coeficiente de part
        hoja.cell(row=row_number+2, column=(column_number+16), value= cantidad_de_alumnos* celda_coeficiente_por_hora)
        celda = hoja.cell(row=row_number+2, column=(column_number+16))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        ### Horas parciales
        hoja.cell(row=row_number+2, column=(column_number+17), value= cantidad_de_alumnos* 2)
        celda = hoja.cell(row=row_number+2, column=(column_number+17))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla


        ### Horas por UEA
        hoja.cell(row=row_number+2, column=(column_number+18), value=horas_por_UEA)
        celda = hoja.cell(row=row_number+2, column=(column_number+18))
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla
        celda.border = border_header_tabla

        ### tabla nombre de los alumnos:
        ### DEJAR INSCRITOS A LOS SIGUIENTES ALUMNOS:
        celda = hoja.cell(row=row_number+4, column=column_number)
        celda.value = 'DEJAR INSCRITOS A LOS SIGUIENTES ALUMNOS:'
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla

        celda = hoja.cell(row=row_number+6, column=column_number)
        celda.value = 'Nombre'
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla

        celda = hoja.cell(row=row_number+6, column=column_number+1)
        celda.value = 'Matricula'
        celda.alignment = alignment_header_tabla
        celda.font = font_header_tabla

        for index, alumno in enumerate(info['Alumno(s)']):
            celda_nombre = hoja.cell(row=row_number + 7 + index, column=column_number)
            celda_nombre.value = manejo_de_nombre(alumno['nombre'])

            celda_matricula = hoja.cell(row=row_number + 7 + index, column=column_number + 1)
            celda_matricula.value = alumno['matrícula']


    ### Llenando info
    # PROYECTO TERMINAL I
    celda = hoja.cell(row=row_number+2, column=column_number)


    celda.alignment = alignment_header_tabla
    celda.font = font_header_info
    celda.border = border_header_tabla
    celda.value = 'PROYECTO TERMINAL '+ pt

    # print('info ->', info)
    
    # Variable clave
    celda = hoja.cell(row=row_number+2, column=column_number+1)

    celda.alignment = alignment_header_tabla
    celda.font = font_header_info
    celda.border = border_header_tabla
    celda.value = int(info['CLAVE'])


    # Variable Grupo
    celda = hoja.cell(row=row_number+2, column=column_number+2)

    celda.alignment = alignment_header_tabla
    celda.font = Font(name='Calibri', size=11, color='000000')
    celda.border = border_header_tabla
    celda.value = info['GRUPO']

    

    # Variable Cupo máx
    celda = hoja.cell(row=row_number+2, column=column_number+3)

    celda.alignment = alignment_header_tabla
    celda.font = Font(name='Calibri', size=11, color='000000')
    celda.border = border_header_tabla
    celda.value = 30

    
    # Variable Profesor
    celda = hoja.cell(row=row_number+2, column=column_number+4)

    celda.alignment = alignment_header_tabla
    celda.font = Font(name='Calibri', size=11, color='000000')
    celda.border = border_header_tabla

    if info['Asesor(es)'] == 'POR ASIGNAR': 
        asesores = 'POR ASIGNAR'
        celda.value = asesores
    elif  isinstance( info['Asesor(es)'], list):
        # print('Its a list of professors')
        asesores = []

        for professor in info['Asesor(es)']:
            asesores.append(professor['nombre'])

        ### agregando el primer profe
        # print(asesores)
        celda.value = asesores[0]
        asesores.pop(0)

        for i in range(len(asesores)):
            y = i +1
            # print('valor de y -> ', y)
            # print('asesores[i] -> ', asesores[i])
            celda = hoja.cell(row=row_number+2+y, column=column_number+4)
            celda.value = asesores[i]

    else:
        asesores = info['Asesor(es)']['nombre']
        celda.value = asesores
    
    # print(asesores)
    

    # Variable Num Economico
    celda = hoja.cell(row=row_number+2, column=column_number+5)

    celda.alignment = alignment_header_tabla
    celda.font = font_tabla_info
    celda.border = border_header_tabla
    celda.value = 'N/A'

    # Blank space
    celda = hoja.cell(row=row_number+2, column=column_number+6)
    celda.border = border_header_tabla

    celda = hoja.cell(row=row_number+2, column=column_number+8)
    celda.border = border_header_tabla

    # MARTES, JUEVES y VIERNES
    celda = hoja.cell(row=row_number+2, column=column_number+7)

    celda.alignment = alignment_header_tabla
    celda.font = font_tabla_info
    celda.border = border_header_tabla
    celda.value = horario_del_curso

    celda = hoja.cell(row=row_number+2, column=column_number+9)

    celda.alignment = alignment_header_tabla
    celda.font = font_tabla_info
    celda.border = border_header_tabla
    celda.value = horario_del_curso

    celda = hoja.cell(row=row_number+2, column=column_number+10)

    celda.alignment = alignment_header_tabla
    celda.font = font_tabla_info
    celda.border = border_header_tabla
    celda.value = horario_del_curso

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION ENCARGADA EN LA GENERACION DE COEFICIENTES
def crear_coeficientes(row_number, column_number, hoja, data):
    print('Generando los coeficientes de los asesores y sus alumnos')

    #CABECERA A CREARCE EN LA SECCION DE COEFICIENTES
    header = ['PROFESOR', 'No. Económico', 'UEA', 'CLAVE', 'GRUPO', 'COEFICIENTE DE PARTICIPACIÓN', 'HORAS PARCIALES', 'HORAS TOTALES POR UEA']
    fill_header = PatternFill(start_color='a8d08d', end_color='a8d08d', fill_type='solid')
    font_header = Font(name='Arial', size=10, bold=True, italic=False, color='000000')
    alignment_header = Alignment(horizontal='center', vertical='center')
    border_header = Border(left=Side(style='medium'), 
                     right=Side(style='medium'), 
                     top=Side(style='medium'), 
                     bottom=Side(style='medium'))
    #creacion de arreglos de los PTS
    grupopt1 = []
    grupopt2 = []
    grupopt3 = []

    #creacion de conjunto de todos los arreglos
    grupos = [grupopt1, grupopt2, grupopt3]

    #Impresion de de la cabecera generada en la tabla 
    for index, i in enumerate(header):
        celda = hoja.cell(row=row_number, column=column_number+index, value=i)
        celda.fill = fill_header
        celda.font = font_header
        celda.alignment = alignment_header
        celda.border = border_header

    #Llenando la informacion de los GRUPOS
    for grupo in data:
        if grupo.get('PT') == '1':
            grupopt1.append(grupo)
        elif grupo.get('PT') == '2':
            grupopt2.append(grupo)
        else: 
            grupopt3.append(grupo)
        
    #Imprimir el data
    contador_posicion = 1

    for index, grupo in enumerate(grupos):
        for i in grupo:
            #Colocacion de los asesores / alumnos 
            asesores = i['Asesor(es)']
            cantidad_alumnos = len(i['Alumno(s)'])
            horas_parciales = cantidad_alumnos*2
            horas_por_alumno = 2
            #Colocaion de la clave / grupo de cada PT
            if i['PT'] == '1':
                pt='I'
                clave = '450218'
                grupo = 'DJ01T'
                horas_totales_por_UEA = 9
            elif i['PT'] == '2':
                pt='II'
                clave = '450219'
                grupo = 'DK01T'
                horas_totales_por_UEA = 10

            else:
                pt='III'
                clave = '450220'
                grupo = 'DL01T'
                horas_totales_por_UEA = 10
            #CALCULANDO LOS COEFICIENTES POR HORA
            coeficiente_por_hora = horas_por_alumno / horas_totales_por_UEA #Se divide las horas cada alumno y las horas totales por UEA
            
            if(coeficiente_por_hora * cantidad_alumnos > 1):
                coeficiente_de_participacion = 1 * len(asesores)
            else:
                coeficiente_de_participacion = round(coeficiente_por_hora * cantidad_alumnos, 2)

            if (isinstance(asesores, list)):
                #Lista
                suma_coeficiente_de_participacion = len(asesores) * coeficiente_de_participacion
                suma_horas_parciales = len(asesores) * horas_parciales

                for index, asesor in enumerate(asesores):
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number, value=asesor['nombre'])
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+1, value=asesor['numero_economico'])
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+2, value= 'PROYECTO TERMINAL '+ pt)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+3, value= clave)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+4, value= grupo)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+5, value= coeficiente_de_participacion)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+6, value= horas_parciales)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+7, value= horas_totales_por_UEA)
                contador_posicion = contador_posicion + index

            else:
                suma_coeficiente_de_participacion = coeficiente_de_participacion
                suma_horas_parciales = horas_parciales

                celda = hoja.cell(row=row_number+contador_posicion, column=column_number, value=asesores['nombre'])
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+1, value=asesores['numero_economico'])
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+2, value= 'PROYECTO TERMINAL '+ pt)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+3, value= clave)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+4, value= grupo)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+5, value= coeficiente_de_participacion)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+6, value= horas_parciales)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+7, value= horas_totales_por_UEA)


            #SUMA
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number, value= 'SUMA')
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number+5, value= suma_coeficiente_de_participacion)
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number+6, value= suma_horas_parciales)


            contador_posicion = celda.row + 1

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION MAIN PARA GRUPOS DE PT
def main(codigoGen1):
    with open(archivosGenerados1_folder + codigoGen1,'r',encoding="utf-8") as file:
         print(f"El archivo que se usara en el main de GRUPOS DE PT:'{codigoGen1}")
         #Se incializan las variables temporales para almacenar la informacion del proyecto
         titulo = ""
         asesores = ""
         matricula = ""
         pt = ""

         #Itera sobre cada linea del archivo
         for line in file:
            # Elimina los espacios en blanco al principio y al final de la línea
            line = line.strip()
            # Verifica si la línea contiene la informacion de TITULO
            if line.startswith('Título:'):
                titulo = line[len('Título:'):].strip()
            elif line.startswith('Asesor(es):'):
                asesores__to_get_numero_economico = []
                #Verifica si la linea que se esta leyendo  contiene coma para dividir los nombres de los revisores
                if ',' in line:
                    asesores_lista_in_line = list(set([rev.strip() for rev in line[len('Asesor(es):'):].split(',')]))
                    print("asesores_lista_in_line", asesores_lista_in_line)
                    for rev in asesores_lista_in_line:
                        asesores__to_get_numero_economico.append(rev)
                else:
                    asesor = line[len('Asesor(es):'):].strip()
                    print("asesor", asesor)
                    asesores__to_get_numero_economico.append(asesor)
                asesores = []
                for asesor in asesores__to_get_numero_economico:
                    try:
                        genero = get_genre(asesor)
                        numero_economico_asesor = get_numero_economico(manejo_de_nombre(asesor), codigoGen1)
                        asesor = {'nombre': asesor, 'genero': genero, 'numero_economico': numero_economico_asesor}
                        asesores.append(asesor)
                    except Exception as e:
                        print(f"Error al obtener el número económico del asesor: {e}")
                        continue

            elif line.startswith('Alumno:'):
                alumno = line[len('Alumno:'):].strip()

            elif line.startswith('Matrícula:'):
                matricula = line[len('Matrícula:'):].strip()
                
            elif line.startswith('PT:'):
                pt = line[len('PT:'):].strip()
                
                #Estructura a tomar de los datos recopilados
                grupo = {
                        'Título': titulo,
                        'Asesor(es)': asesores,
                        'Alumno': alumno,
                        'Matricula': matricula,
                        'PT':pt
                    }
                manejo_grupo(grupo) #Creando los grupos
         print('Grupos:', grupos)
         #Verificacion de la lista llenada
         if grupos:
            print('Generando excell con base en los grupos generados')
            nombre_de_hoja = 'GruposPt'
            #1.- Se copia la informacion de template a una nueva hoja
            #Se crea la hoja con el nombre y el numero de trimestre
            nueva_hoja = workbook.create_sheet(title=nombre_de_hoja+'-'+trimestre)

            #Se definen las variables del estilo
            font_proyecto_terminal = Font(name='Arial', size=10, bold=True, italic=False, color='000000')
            fill_proyecto_terminal = PatternFill(start_color='ffe598', end_color='ffe598', fill_type='solid')
            fill_cambios_solicitados = PatternFill(start_color='ffc000', end_color='ffc000', fill_type='solid')
            fill_proyecto_terminalII = PatternFill(start_color='c5e0b3', end_color='c5e0b3', fill_type='solid')
            fill_proyecto_terminalIII = PatternFill(start_color='99ffff', end_color='99ffff', fill_type='solid')

            #Columna A expandida
            nueva_hoja.column_dimensions['A'].width = 34
            nueva_hoja.column_dimensions['B'].width = 12
            nueva_hoja.column_dimensions['E'].width = 20

            #Primera Celda
            nueva_hoja.cell(row=2, column=1, value='REGISTRO ACTUAL PROYECTO TERMINAL I')
            celda = nueva_hoja['A2']
            celda.font = font_proyecto_terminal
            celda.fill = fill_proyecto_terminal

            #Creacion de la informacion de la tabla para los PT1, PT2 Y PT3
            info_tabla = {
            'Asesor(es)': 'POR ASIGNAR',
            'Alumno(s)': '',
            'PT': '1',
            'CLAVE': 450218,
            'GRUPO': 'DJ01T'
             }
             #Grupos de PT 1
            crear_tabla(3, 1, nueva_hoja, info_tabla)
        
            celda = nueva_hoja['A7']
            celda.font = font_proyecto_terminal
            celda.fill = fill_cambios_solicitados
            celda.value = 'CAMBIOS SOLICITADOS'

            grupopt1 = []
            grupopt2 = [{
            'Asesor(es)': 'POR ASIGNAR',
            'Alumno(s)': '',
            'PT': '2',
            'CLAVE': 450219,
            'GRUPO': 'DK01T'
            }]
        
            grupopt3 = [{
            'Asesor(es)': 'POR ASIGNAR',
            'Alumno(s)': '',
            'PT': '3',
            'CLAVE': 450220,
            'GRUPO': 'DL01T'
            }]

    #Incroporacion de los grupos en el diccionario
    for diccionario in grupos:
        if diccionario.get('PT') == '1':
            grupopt1.append(diccionario)
        elif diccionario.get('PT') == '2':
            grupopt2.append(diccionario)
        elif diccionario.get('PT') == '3':
            grupopt3.append(diccionario)
        marcador = 0

        for i in range(len(grupopt1)):
                y = i * 12
                info_tabla = {
                    'Asesor(es)': grupopt1[i]['Asesor(es)'],
                    'Alumno(s)': grupopt1[i]['Alumno(s)'],
                    'PT': grupopt1[i]['PT'],
                    'CLAVE': 450218,
                    'GRUPO': 'DJ01T'   
                }
                crear_tabla(9+y, 1, nueva_hoja, info_tabla)
                marcador = y

    #Se envia la informacion de PT II 
    marcador_continua = marcador+20
    celda = nueva_hoja.cell(row=marcador_continua, column=1, value='REGISTRO ACTUAL PROYECTO TERMINAL II') 
    celda.font = font_proyecto_terminal
    celda.fill = fill_proyecto_terminalII

    #Se manda la informacion del PT II
    for i in range(len(grupopt2)):
                y = i * 12
                #Creacion de la tabla con los valores
                info_tabla = {
                    'Asesor(es)': grupopt2[i]['Asesor(es)'],
                    'Alumno(s)': grupopt2[i]['Alumno(s)'],
                    'PT': grupopt2[i]['PT'],
                    'CLAVE': 450219,
                    'GRUPO': 'DK01T'
                    }
                marcador_pt_ii = y
                #Creacion de la tabla
                crear_tabla(marcador_continua+2+y, 1, nueva_hoja, info_tabla)


    #Se envia la informacion de PT II
    marcador_continua_pt_ii = marcador_pt_ii+38+marcador
    celda = nueva_hoja.cell(row=marcador_continua_pt_ii, column=1, value='REGISTRO ACTUAL PROYECTO TERMINAL III')
    celda.font = font_proyecto_terminal
    celda.fill = fill_proyecto_terminalIII

    #Se manda la informacion del PT III
    for i in range(len(grupopt3)):
            y = i * 12
            
            info_tabla = {
                    'Asesor(es)': grupopt3[i]['Asesor(es)'],
                    'Alumno(s)': grupopt3[i]['Alumno(s)'],
                    'PT': grupopt3[i]['PT'],
                    'CLAVE': 450219,
                    'GRUPO': 'DK01T'
                    }
            crear_tabla(marcador_continua_pt_ii+2+y, 1, nueva_hoja, info_tabla)

    #3.- Generacion de coeficientes
    nombre_de_hoja = 'Coeficientes'
    nueva_hoja = workbook.create_sheet(title=nombre_de_hoja+'-'+trimestre)
   
    #Creando la tabla con base en la informacion
    crear_coeficiente(1,1,nueva_hoja,grupos)
        
    nueva_hoja.column_dimensions['A'].width = 34
    nueva_hoja.column_dimensions['B'].width = 15
    nueva_hoja.column_dimensions['C'].width = 34
    nueva_hoja.column_dimensions['D'].width = 15
    nueva_hoja.column_dimensions['E'].width = 15
    nueva_hoja.column_dimensions['F'].width = 35
    nueva_hoja.column_dimensions['G'].width = 30
    nueva_hoja.column_dimensions['H'].width = 30

    nueva_hoja.row_dimensions[1].height = 30

    #4.- Guardar los cambios en el libro de trabajo
    try:
        timestamp =datetime.now().strftime('%Y%m%d_%H%M%S')
        excell_resultado_without_extension = excell_resultado.split('.')[0]
        workbook_file_to_save = f"{directorio_archiGeneradosGruposDePtExcell}\\{excell_resultado_without_extension}_{timestamp}.xlsx"
        print(f"Guardando el archivo en: {workbook_file_to_save}")
        workbook.save(workbook_file_to_save)
        generar_txt_estudiantes_PT()
        print(f"Archivo generados con éxito en: {workbook_file_to_save}")
        return workbook_file_to_save
    except Exception as e:
        print(f"Error al guardar el archivo o generar el TXT: {e}")
        raise Exception(f"Error al guardar el archivo o generar el TXT: {e}")
    

def get_genre(name):
    indice = name.index(' ')
    titulado = indice - 1
    if name[titulado] == 'a':
        return 'F'
    else:
        return 'M'

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION GENERADORA DEL DOCUMENTO EXCEL Y EL TXT DEL PRIMER MODULO GRUPOS DE PT (ARCHIVO ESPECIFICO)
@app.route('/GruposPT/<filename>', methods=['GET'])
def generar_GruposPT(filename):
    global directorio_automatico
    fileGenerated = doGenerarGruposPT(filename)
    if fileGenerated.success:
            return render_template('generar.html',
                                            directorio_automatico2=directorio_automatico,
                                            archivosGenerados1_folder2=archivosGenerados1_folder,
                                            #list_Archivos2=listar_archivos(form_id, tipo),
                                            message=f"El archivo se guardó en {fileGenerated.txt_filename}",
                                            archivosGlobales=[fileGenerated.archivoGlobal1],
                                            codigosGen=[fileGenerated.codigoGen1],
                                            #generated_files=generated_files,
                                            list_Archivos=listar_archivos_subidos(),
                                            archivos=[fileGenerated.archivo],
                                            message2=f"El archivo: {fileGenerated.txt_filename}\n se generó con éxito en {directorio_archiGeneradosGruposDePtExcell}",
                                            txt_filenames=[fileGenerated.txt_filename],
                                            #excel_filename=excel_filename,
                                            div9_block=True)
    else:
        return render_template('index.html',
                                            directorio_automatico2=directorio_automatico,
                                            archivosGenerados1_folder2=archivosGenerados1_folder,
                                            message=f"Hubo una falla al procesar el archivo",
                                            archivosGlobales=[],
                                            codigosGen=[],
                                            list_Archivos=listar_archivos_subidos(),
                                            message2=f"Archivo CSV no encontrado",
                                            txt_filenames=[],
                                            div9_block=True)

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION GENERADORA DEL DOCUMENTO EXCEL Y EL TXT DEL PRIMER MODULO GRUPOS DE PT (MULTI ARCHIVOS)
@app.route('/GruposPTMulti', methods=['POST'])
def generar_GruposPTMulti():
    global directorio_automatico
    print('request.form.items()',request.form.items())
    options = []
    for key, value in request.form.items():
        options.append(value)
    print('options',options)
    filesGenerated = []
    for option in options:
        filesGenerated.append(doGenerarGruposPT(option))
    print('filesGenerated',filesGenerated)
    successedFiles = []
    failedFiles = []
    for fileGenerated in filesGenerated:
        print('fileGenerated',fileGenerated)
        if fileGenerated.get("success") is not None and fileGenerated.get("success") is True:
            successedFiles.append(fileGenerated)
        else:
            failedFiles.append(fileGenerated)
    message = ''
    if successedFiles and len(successedFiles) > 0:
        message = f"Los archivos se guardaron en {', '.join([file.get('txt_filename') for file in successedFiles if file.get('txt_filename')])}"
    if failedFiles and len(failedFiles) > 0:
        message = message + f"Hubo una falla al procesar los siguientes archivos: {', '.join([file.get('archivoGlobal1') for file in failedFiles if file.get('archivoGlobal1')])}"
    message2 = ''
    if successedFiles and len(successedFiles) > 0:
        message2 = f"Los archivos {', '.join([file.get('txt_filename') for file in successedFiles if file.get('txt_filename')])} se generaron con éxito en {directorio_archiGeneradosGruposDePtExcell}"
    if failedFiles and len(failedFiles) > 0:
        message2 = message2 + f"Los archivos {', '.join([file.get('archivoGlobal1') for file in failedFiles if file.get('archivoGlobal1')])}, tuvieron el siguiente eeror: Archivo CSV no encontrado"
    archivosGlobales = [fileGenerated.get("archivoGlobal1") for fileGenerated in successedFiles]
    codigosGen = [fileGenerated.get("codigoGen1") for fileGenerated in successedFiles]
    archivos = [fileGenerated.get("archivo") for fileGenerated in successedFiles]
    txt_filenames = [fileGenerated.get("txt_filename") for fileGenerated in successedFiles]
    list_Archivos = listaArchivos("formulario1", "subidos")
    list_Archivos_Generados = listaArchivos("formulario1", "generados")
    print('message',message)
    print('message2',message2)
    print('successedFiles',successedFiles)
    print('failedFiles',failedFiles)
    print('directorio_automatico',directorio_automatico)
    print('archivosGenerados1_folder',archivosGenerados1_folder)
    print('list_Archivos',list_Archivos)
    print('archivosGlobales',archivosGlobales)
    print('codigosGen',codigosGen)
    print('archivos',archivos) #No trae campos
    print('txt_filenames',txt_filenames)
    print('list_Archivos_Generados',list_Archivos_Generados)
    return render_template('index.html', 
                            directorio_automatico2=directorio_automatico,
                            archivosGenerados1_folder2=archivosGenerados1_folder,
                            archivosGlobales=archivosGlobales,
                            codigosGen=codigosGen,
                            list_Archivos=list_Archivos,
                            archivos=archivos,
                            message=message,
                            message2=message2,
                            txt_filenames=txt_filenames,
                            list_Archivos_Generados = list_Archivos_Generados,
                            div9_block=True)

def doGenerarGruposPT(filename):
    global directorio_automatico
    global archivoGlo1
    global archivoGlobal1
    global ArchiGenerado

    form_id = 'asignacionesTXT'
    tipo = 'subidos'

    archivoGlobal1 = filename
    archivoGlo1 = directorio_automatico + "\\GruposDePTExcell\\ArchivosSubidos\\" + archivoGlobal1
    print(f"ArchivoGlobal1 que se usará en el main: {archivoGlobal1}")
    print(f"ArchivoGlo1 que se usará en el main: {archivoGlo1}")  # RUTA COMPLETA DEL CSV SUBIDO POR EL USUARIO

    try:
        # Revisar que archivoGlo1 incluya .csv
        hasDotCSV = archivoGlo1.endswith('.csv')
        if not hasDotCSV:
            raise ValueError("El archivo no tiene la extensión .csv")

        if os.path.exists(archivoGlo1):  # csv_path
            txt_filename = procesar_archivo_csv(archivoGlo1)  # csv_path

            # Asegúrate de que el archivo TXT existe antes de procesarlo
            if not os.path.exists(txt_filename):
                raise FileNotFoundError(f"El archivo TXT no se generó correctamente: {txt_filename}")

            print(f"El archivo txt tiene la ruta de: {txt_filename}")
            txt_name = os.path.basename(txt_filename)
            name, ext = os.path.splitext(txt_name)
            new_filename = f"{name}{ext}"
            codigoGen1 = new_filename
            print(f"El archivo txt tiene el nombre de: {codigoGen1}")

            # Llama a la función en main.py que procesará el archivo TXT
            coeficientes = main(codigoGen1)  # Suponiendo que 'main' es la función que deseas llamar en main.py
            print(f"El archivo coeficientes tiene el nombre de: {coeficientes}")

            # Renderizar la plantilla con los datos necesarios
            return {
                "txt_filename": txt_filename,
                "archivoGlobal1": archivoGlobal1,
                "codigoGen1": codigoGen1,
                "archivo": coeficientes,
                "success": True
            }
        else:
            raise FileNotFoundError(f"El archivo CSV no se encontró: {archivoGlo1}")

    except Exception as e:
        print(f"Error al generar los grupos PT: {e}")
        return {
            "archivoGlobal1": archivoGlobal1,
            "success": False
        }


#FUNCION QUE LISTA LOS ARCHIVOS
@app.route('/archivo/<filename>', methods=['GET'])
def view_file_generated(form_id, tipo):
    archivos_lista = []

    # Inicializar archivos_subidos_lista de manera segura
    archivos_subidos_lista = []

    # Verificar si los archivos han sido subidos (comprobando form_data)
    if form_id in form_data and tipo in form_data[form_id]:
        archivos_subidos_lista = [
            file for file in os.listdir(directorio_archivos_subidos)
            if file.endswith('.txt') or file.endswith('.xlsx')# and file != '~$PlantillaAsignacionCoeficientesPT-24-O.xlsx'
        ]
    
    # Listar archivos subidos
    archivos_lista.extend(archivos_subidos_lista)

    # Listar archivos generados
    archivos_generados_lista = [
        file for file in os.listdir(directorio_archivos_generados)
        if file.endswith('.txt') or file.endswith('.xlsx')# and file != '~$PlantillaAsignacionCoeficientesPT-24-O.xlsx'
    ]
    archivos_lista.extend(archivos_generados_lista)

    # Imprimir listas para depuración
    print(f"Archivos subidos: {archivos_subidos_lista}")
    print(f"Archivos generados: {archivos_generados_lista}")

    return archivos_lista

def listar_archivos_subidos():
    # Listar todos los archivos en la carpeta que sean .txt o .xlsx, excluyendo el archivo temporal
    archivos_lista = [
        file for file in os.listdir(directorio_archivos_subidos)
        if (file.endswith('.txt') or file.endswith('.xlsx')) #and file != '~$PlantillaAsignacionCoeficientesPT-24-O.xlsx'
    ]

    # Imprimir lista para depuración
    print(f"Archivos subidos: {archivos_lista}")

    return archivos_lista


def crear_coeficiente(row_number, column_number, hoja, data):
    print('generando coeficiente...')

    header = ['PROFESOR', 'No. Económico', 'UEA', 'CLAVE', 'GRUPO', 'COEFICIENTE DE PARTICIPACIÓN', 'HORAS PARCIALES', 'HORAS TOTALES POR UEA']
    fill_header = PatternFill(start_color='a8d08d', end_color='a8d08d', fill_type='solid')
    font_header = Font(name='Arial', size=10, bold=True, italic=False, color='000000')
    alignment_header = Alignment(horizontal='center', vertical='center')
    border_header = Border(left=Side(style='medium'), 
                     right=Side(style='medium'), 
                     top=Side(style='medium'), 
                     bottom=Side(style='medium'))
    grupopt1 = []
    grupopt2 = []
    grupopt3 = []

    grupos = [grupopt1, grupopt2, grupopt3]



    ### imprimiendo la cabecera de la tabla
    for index, i in enumerate(header):
        celda = hoja.cell(row=row_number, column=column_number+index, value=i)
        celda.fill = fill_header
        celda.font = font_header
        celda.alignment = alignment_header
        celda.border = border_header

    ### llenando la información de los grupos 
    for grupo in data:
        if grupo.get('PT') == '1':
            grupopt1.append(grupo)
        elif grupo.get('PT') == '2':
            grupopt2.append(grupo)
        else: 
            grupopt3.append(grupo)

    ### imprimir data
    contador_posicion = 1

    for index, grupo in enumerate(grupos):
        for i in grupo:
            
            asesores = i['Asesor(es)']
            cantidad_alumnos = len(i['Alumno(s)'])
            horas_parciales = cantidad_alumnos*2
            horas_por_alumno = 2

            if i['PT'] == '1':
                pt='I'
                clave = '450218'
                grupo = 'DJ01T'
                horas_totales_por_UEA = 9
            elif i['PT'] == '2':
                pt='II'
                clave = '450219'
                grupo = 'DK01T'
                horas_totales_por_UEA = 10

            else:
                pt='III'
                clave = '450220'
                grupo = 'DL01T'
                horas_totales_por_UEA = 10

            coeficiente_por_hora = horas_por_alumno / horas_totales_por_UEA

            if (coeficiente_por_hora * cantidad_alumnos > 1):
                coeficiente_de_participacion = 1 * len(asesores)
            else:
                coeficiente_de_participacion = round(coeficiente_por_hora * cantidad_alumnos, 2)


            if (isinstance(asesores, list)):
                # print('list')
                suma_coeficiente_de_participacion = len(asesores) * coeficiente_de_participacion
                suma_horas_parciales = len(asesores) * horas_parciales

                for index, asesor in enumerate(asesores):
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number, value=asesor['nombre'])
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+1, value=asesor['numero_economico'])
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+2, value= 'PROYECTO TERMINAL '+ pt)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+3, value= clave)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+4, value= grupo)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+5, value= coeficiente_de_participacion)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+6, value= horas_parciales)
                    celda = hoja.cell(row=row_number+contador_posicion+index, column=column_number+7, value= horas_totales_por_UEA)
                contador_posicion = contador_posicion + index

            else:
                # print('not a list')

                suma_coeficiente_de_participacion = coeficiente_de_participacion
                suma_horas_parciales = horas_parciales

                celda = hoja.cell(row=row_number+contador_posicion, column=column_number, value=asesores['nombre'])
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+1, value=asesores['numero_economico'])
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+2, value= 'PROYECTO TERMINAL '+ pt)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+3, value= clave)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+4, value= grupo)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+5, value= coeficiente_de_participacion)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+6, value= horas_parciales)
                celda = hoja.cell(row=row_number+contador_posicion, column=column_number+7, value= horas_totales_por_UEA)


                ### SUMA
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number, value= 'SUMA')
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number+5, value= suma_coeficiente_de_participacion)
            celda = hoja.cell(row=row_number+contador_posicion+1, column=column_number+6, value= suma_horas_parciales)


            contador_posicion = celda.row + 1


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#FUNCION MAIN 
if __name__ == "__main__":
    app.run(debug=True)