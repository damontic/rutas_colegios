import getopt
import math
import sys
import config

def uso():
	print("""
USO:
\tpython""", sys.argv[0], """

OPCIONES:

\t-h, --help
\t\tMuestra esta ayuda y termina.
	""")

class Ruta:
	def __init__(self):
		self.ruta = []
		self.L = 0 # Distancia recorrida
		self.T = 0 # Tiempo de llegada del bus

	def __str__(self):
		return "Ruta:{" + \
		"\n\t\t\"L\":" + str(self.L) + "," + \
		"\n\t\t\"T\":" + str(self.T) + "," + \
		"\n\t\t\"ruta\":" + str(self.ruta) + "," + \
		"\n}"

class RuteoSolver:
	def __init__(self, archivo_enunciado):
		self.Q = config.capacidad_buses # Capacidad de los buses
		self.coordenada_salida_buses = config.coordenada_salida_buses
		self.coordenadas_ninos = config.coordenadas_ninos
		self.N = len(self.coordenadas_ninos) # Cantidad de niños

		if(self.N <= 0):
			print("La cantidad de niños debe ser mayor a 0.")
			sys.exit(1)

		if(self.Q <= 0):
			print("La capacidad de los buses debe ser mayor a 0.")
			sys.exit(1)

		self.NB = math.ceil(self.N / self.Q) # Buses Objetivo
		self.grupos_ninos = self.__crear_grupos(self.coordenadas_ninos, math.ceil( self.N / self.NB ))
		self.distancias = self.__leer_matriz(config.archivo_matriz_distancias)
		self.costos = self.__leer_matriz(config.archivo_matriz_costos)
		self.tiempos = self.__leer_matriz(config.archivo_matriz_tiempos)
		self.solucion = self.__solve(self.grupos_ninos)

	def __str__(self):
		soluciones = [ str(ruta) for ruta in self.solucion ]
		return "RuteoSolver:{" + \
		"\n\t\"Q\":" + str(self.Q) + "," + \
		"\n\t\"N\":" + str(self.N) + "," + \
		"\n\t\"NB\":" + str(self.NB) + "," + \
		"\n\t\"coordenada_salida_buses\":" + str(self.coordenada_salida_buses) + "," + \
		"\n\t\"coordenadas_ninos\":" + str(self.coordenadas_ninos) + "," + \
		"\n\t\"grupos_ninos\":" + str(self.grupos_ninos) + "," + \
		"\n\t\"distancias\":" + str(self.distancias) + "," + \
		"\n\t\"costos\":" + str(self.costos) + "," + \
		"\n\t\"tiempos\":" + str(self.tiempos) + "," + \
		"\n\t\"solucion\":[\n\t" + "\n\t".join(soluciones) + "]" + \
		"\n}"

	def __crear_grupos(self, coordenadas, tamanio_grupos):
		coordenadas = sorted(coordenadas, key=lambda x: x[1])
		return [ coordenadas[i:i + tamanio_grupos] for i in range(0, len(coordenadas), tamanio_grupos) ]

	def __promedio(self, lista):
		return sum(lista) / len(lista)

	def __varianza(self, lista):
		promedio = self.__promedio(lista)
		return sum([ math.pow(elem - promedio, 2) for elem in lista ]) / (len(lista))

	def __desviacion_estandard(self, lista):
		return math.sqrt(self.__varianza(lista))

	def __leer_matriz(self, archivo):
		f = open(archivo, "r")
		lineas = f.readlines()
		f.close()
		lineas = [ linea.replace("\n", "").split("\t") for linea in lineas ]
		for indice_fila in range(0, len(lineas)):
			for indice_columna in range(0, len(lineas[indice_fila])):
				lineas[indice_fila][indice_columna] = float(lineas[indice_fila][indice_columna])
		return lineas

	def __solve(self, grupos_ninos):
		rutas = []
		for grupo in grupos_ninos:
			rutas.append( Ruta() )
		return rutas

if __name__ == '__main__':

	# Procesa todos los argumentos del programa
	try:
		opts, args = getopt.getopt(sys.argv[1:], "h", ["help"])
	except getopt.GetoptError as err:
		# dibuja ayuda y sale:
		print(str(err))  # dibuja qué opción es inválida
		uso()
		sys.exit(2)

	# Define las variables necesarias
	archivo_enunciado = None
	for option, value in opts:
		if option in ("-h", "--help"):
			uso()
			sys.exit()
		else:
			assert False, "unhandled option"

	ruteoSolver = RuteoSolver(archivo_enunciado)
	print(ruteoSolver)
