# Si quieres hacer colocar un Enter (salto de linea) entre cualquier frase, solo coloca "\n\r" 
from datetime import datetime
import locale

Nombre_Ano = 'Decenio de la Igualdad de Oportunidades para Mujeres y Hombres \n“Año del Bicentenario, de la consolidación de nuestra Independencia, y de la conmemoración de las heroicas batallas de Junín y Ayacucho”'

fechas_entregables = {
    "16/09/24" : "Primer",
    "15/10/24" : "Segundo",
    "15/11/24" : "Tercer",
    "15/12/24" : "Cuarto"}

fecha = min(fechas_entregables, key=lambda x: abs(datetime.strptime(x, "%d/%m/%y") - datetime.now()))

class Margenes_Carta:
    superior = 1.10
    inferior = 0.50
    derecho = 2.50
    izquierdo = 2.80

Margenes_Carta = Margenes_Carta()

class Margenes_Informe:
    superior = 3.5
    inferior = 2.25
    derecho = 1.80
    izquierdo = 2.10

Margenes_Informe = Margenes_Informe()

class fecha_departamento:
    locale.setlocale(locale.LC_TIME, 'es')  # Para sistemas Windows
    departamento = "Lima"
    fecha = datetime.now().strftime("%d de %B de %Y")

fecha_departamento = fecha_departamento()

class receptor:
    director_deam = "LÁZARO WALTHER FAJARDO VARGAS"
    direccion = "Dirección de Evaluación Ambiental"
    director_genero = "Señor"
    
    director_tercero = "Violeta Jhicenia Rivera Minaya"

receptor = receptor()

class emisor:
    nombre = "Edwin Alexander Zuñiga Lujan"
    dni = "72746974"
    contrato = "N° 00141-2024-OEFA/OAD-UAB"
    Adenda = "N° 0002"
    n_carta = "012"
    inicio_contrato = "15/08/24"
    ano = datetime.now().strftime("%Y")
    iniciales = ''.join([palabra[0] for palabra in nombre.split()])
    entregable = fechas_entregables[fecha]

emisor = emisor()

class tabla_cabecera:
    celda1_ancho= 4
    celda2_ancho= 8.7
    celda3_ancho= 4
    celda_alto = 3 * 2
    
    texto_mapro= "MAPRO-OAD-PA-0240"
    texto_version= "Versión: 03\nFecha: 15/11/2021"
    
def Contrato_Adenda():
    if "N°" in emisor.Adenda:
        conexion =" de la "
        texto_contrato = f"Adenda {emisor.Adenda} al Contrato {emisor.contrato}"

    else:
        conexion =" del "
        texto_contrato = f"Contrato {emisor.contrato}"
        
    return conexion, texto_contrato