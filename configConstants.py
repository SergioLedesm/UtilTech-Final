""" ===> VARIABLES NO VARIABLES <== """
import os

#Se establece la clave secreta manualmente con la cadena
key = 'UtilTech' 

#######################################################################################
#######################################################################################
#######################################################################################
#DIRECTORIO DE LA RUTAS
#RUTA DONDE SE ENCUENTRA EL PROYECTO
directorio_automatico = os.path.dirname(os.path.abspath(__file__))

#RUTA DONDE SE ENCUENTRAN LOS ARCHIVOS DEL PROGRAMA 1 (GRUPOS DE PT)
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


#Ruta direccionada a la seleccion del Programa a usar (1, 2, 3) y si se trata de un archivo SUBIDO o GENERADO
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


#RUTA DE LA DIRECCION DEL ARCHIVO RESULTADO DEL PROGRAMA 2
path_resultado = directorio_automatico + '\Resultado.xlsx'

