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

class Margenes:
    superior = 1.10
    inferior = 0.50
    derecho = 2.50
    izquierdo = 2.80

Margenes = Margenes()

class fecha_departamento:
    locale.setlocale(locale.LC_TIME, 'es')  # Para sistemas Windows
    departamento = "Lima"
    fecha = datetime.now().strftime("%d de %B de %Y")

fecha_departamento = fecha_departamento()

class receptor:
    director = "LÁZARO WALTHER FAJARDO VARGAS"
    direccion = "Dirección de Evaluación Ambiental"
    genero = "Señor"

receptor = receptor()

class emisor:
    nombre = "Edwin Alexander Zuñiga Lujan"
    dni = "72746974"
    contrato = "N° 00141-2024-OEFA/OAD-UAB"
    Adenda = "N° 0002"
    n_carta = "022"
    ano = datetime.now().strftime("%Y")
    iniciales = ''.join([palabra[0] for palabra in nombre.split()])
    entregable = fechas_entregables[fecha]

emisor = emisor()
