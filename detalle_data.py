from datetime import datetime
import locale

Romano = ["I.","II.","III.","IV."]

Actividades = {
    "Primera_Actividad" : "a) Realizar acciones de monitoreo y evaluación ambiental en áreas de influencia de las actividades económicas fiscalizables de competencia del OEFA.",
    "Segunda_Actividad" : "b) Apoyo en la elaboración de documentos técnicos y/o sistematizar información de los monitoreos de calidad del aire en el marco de las evaluaciones ambientales a cargo de la STEC.",
    "Tercera_Actividad" : "c) Apoyo en el mantenimiento preventivo (verificaciones intermedias, control operativo de los analizadores, monitores ambientales, estaciones meteorológicas; y de sus calibraciones en campo) de las estaciones de vigilancia ambiental de la calidad del aire.",
    "Cuarta_Actividad" : "d) Elaborar requerimientos logísticos (planes de trabajo), presentaciones (Ppts, cartillas informativas y ayudas memorias), entre otros en el marco de las evaluaciones ambientales."
}
        
Orden = {
    "RESUMEN EJECUTIVO" : [
        "Durante el periodo contratado, se realizaron actividades de campo y gabinete para el mantenimiento orientados a la ejecución de evaluaciones ambientales de causalidad por encargo de la Subdirección Técnica Científica de la Dirección de Evaluación Ambiental, tales como: ",
        "La información detallada a cada una de las actividades se describe a continuación, y la documentación de sustento de las actividades en mención se encuentran como anexos a este documento."
    ],
    "ACTIVIDADES" : [
        "Actividades realizadas o en ejecución",
        "Principales dificultades encontradas",
        "Modificaciones con relación a lo inicialmente previsto y comentarios",
        "Resultados obtenidos en el periodo"
    ],
    "RECOMENDACIONES Y SUGERENCIAS" : [
        "Durante el periodo contratado se ha logrado el cumplimiento de las actividades encomendadas por el área sin mayores dificultades."
    ],
    "ANEXOS" : [
        """Se adjuntan los siguientes anexos:
        
Anexo 1: Documentos administrativos
Anexo 2: Carpeta que contiene los archivos digitales de los documentos de sustento descritos en el apartado II.4
        
Atentamente"""
    ]
}

orden_llave = list(Orden.keys())

Sustentos = {
    1 : {
        "Producto":   "Participación de la comisión de servicios: «Evaluación ambiental focal de calidad del aire en atención al incendio forestal ocurrido en la provincia de Coronel Portillo, departamento de Ucayali, setiembre de 2024» del 21 de septiembre al 15 de octubre con código de acción 0008-9-2024-411.",
        "Evidencia":   "Se adjunta el cronograma de las comisiones, sobre las acciones de evaluación a cargo de la Subdirección Técnica Científica."
    },
    2 : {
        "Producto":   "Elaboración del Reporte de Actividades N° 007-2024-EAZL: En este reporte se presenta la creación del Módulo 1, el cual permite la visualización de los datos crudos de las estaciones de monitoreo mediante gráficas. El desarrollo de este módulo se realizó entre el 30 de septiembre y el 10 de octubre de 2024. El reporte fue remitido por correo electrónico al coordinador de Vigilancia Ambiental, Ing. Andrés Brios, el 15 de octubre de 2024.",
        "Evidencia":   "Se adjunta el reporte de actividades N° 007-2024-EAZL y correo electrónico, sobre las accionesde evaluación a cargo de la Subdirección Técnica Científica."

    },
    3 : {
        "Producto":   "Elaboración de los reportes de verificación del muestreado continuo de partículas GRIMM de las estaciones CA-HU-01 Nievería, CA-HU-04 El paraíso y CA-HU-09 Santa María de Huachipa, instalados en el marco de la evaluación ambiental de seguimiento, ámbito del centro poblado menor Santa María de Huachipa, El Paraíso y Nievería, distrito Lurigancho, provincia y departamento de Lima, el 18 y 19 septiembre de 2024.",
        "Evidencia":   "Se adjunta los reportes de verificación realizados en versión digital, sobre las acciones de evaluación a cargo de la Subdirección Técnica Científica; el cual se encuentra en el archivo adjunto."
    },
    4 : {
        "Producto":   "-	No se realizó actividad en este Ítem.",
        "Evidencia":   ""
    }
}

Rellenos = [
    "No se han encontrado dificultades para la realización de las actividades programadas.",
    "No se han observado modificaciones con respecto a las actividades contenidas en el detalle de Actividades",
    "Se ha elaborado satisfactoriamente los siguientes productos:",
    "Tales documentos fueron realizados según las exigencias de calidad requeridas por el área usuaria, que se refleja en que dichos documentos fueron remitidos dentro de los plazos establecidos, y que pueden ser verificados en los Anexos 1 y 2 que contienen el sustento de cada uno de los productos señalados."
]
