import getopt
import math
import sys
import excel

def uso():
	print("""
USO:
\tpython""", sys.argv[0], """ -e <archivo_excel> -i <numero_instancia>

OPCIONES:

\t-e, --excel <archivo_excel>
\t\tEspecifica el archivo de excel sobre el cual se trabaja.

\t-i, --instancia
\t\tEspecifica el número de la instancia especificada en el archivo excel.

\t-h, --help
\t\tMuestra esta ayuda y termina.
	""")

class Ruta:
	def __init__(self, coordenada_salida):
		self.ruta = []
		self.L = 0 # Distancia recorrida
		self.T = 0 # Tiempo de llegada del bus
		self.coordenada_salida = coordenada_salida

	def __str__(self):
		return "Ruta:{" + \
		"\n\t\t\"L\":" + str(self.L) + "," + \
		"\n\t\t\"T\":" + str(self.T) + "," + \
		"\n\t\t\"ruta\":" + str(self.ruta) + "," + \
		"\n}"

class RuteoSolver:
	def __init__(self, archivo_excel, numero_instancia):
		f = excel.OpenExcel(archivo_excel)
		indice_instancia = numero_instancia + 1
		self.N = int(f.read("C"+str(indice_instancia))) # Cantidad de niños
		self.Q = int(f.read("N"+str(indice_instancia))) # Capacidad de los buses

		if(self.N <= 0):
			print("La cantidad de niños debe ser mayor a 0.")
			sys.exit(1)

		if(self.Q <= 0):
			print("La capacidad de los buses debe ser mayor a 0.")
			sys.exit(1)

		self.NB = math.ceil(self.N / self.Q) # Buses Objetivo

		indice_nodo_colegio = int(f.read("J"+str(indice_instancia))) + 5
		cantidad_buses = int(f.read("E"+str(indice_instancia)))
		indice_inicio_ninos_hoja_2 = cantidad_buses + 6
		indice_inicio_matrices_fila = cantidad_buses + 3
		indice_inicio_matrices_columna = cantidad_buses + 2

		f = excel.OpenExcel(archivo_excel, sheet = 2)
		self.coordenada_salida_buses = ( int(f.read("B6")), int(f.read("C6")))
		self.coordenadas_ninos = []
		for i in range(indice_inicio_ninos_hoja_2, indice_inicio_ninos_hoja_2 + self.N):
			self.coordenadas_ninos.append( ( int(f.read("B"+str(i))), int(f.read("C"+str(i))) ) ) 
		i = i + 1
		self.coordenada_autopista = ( int(f.read("B"+str(i))), int(f.read("C"+str(i))))
		i = i + 1
		self.coordenada_colegio = ( int(f.read("B"+str(i))), int(f.read("C"+str(i))))

		self.grupos_ninos = self.__crear_grupos(self.coordenadas_ninos, math.ceil( self.N / self.NB ))
		self.distancias = self.__leer_matriz(archivo_excel, 3, indice_inicio_matrices_fila, indice_inicio_matrices_columna, self.N)
		self.costos = self.__leer_matriz(archivo_excel, 4, indice_inicio_matrices_fila, indice_inicio_matrices_columna, self.N)
		self.tiempos = self.__leer_matriz(archivo_excel, 5, indice_inicio_matrices_fila, indice_inicio_matrices_columna, self.N)
#		self.solucion = self.__solve(self.grupos_ninos, self.coordenada_salida_buses)
		self.solucion = []

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

	def __colnum_string(self, n):
		div=n
		string=""
		temp=0
		while div>0:
			module=(div-1)%26
			string=chr(65+module)+string
			div=int((div-module)/26)
		return string


	def __leer_matriz(self, archivo_excel, numero_hoja, indice_inicio_matrices_fila, indice_inicio_matrices_columna, cantidad_ninos):
		f = excel.OpenExcel(archivo_excel, numero_hoja)
		matriz = []
		for i in range(indice_inicio_matrices_fila, indice_inicio_matrices_fila + cantidad_ninos + 2):
			fila = []
			for j in range(indice_inicio_matrices_columna, indice_inicio_matrices_columna + cantidad_ninos + 2):
				cell_name = self.__colnum_string(j)+str(i)
				print(cell_name)
				fila.append( float( f.read(cell_name).replace(",",".") ) )
			matriz.append(fila)
		return matriz

	def __solve(self, grupos_ninos, coordenada_salida_buses):
		rutas = []
		for grupo in grupos_ninos:
			ruta = Ruta(coordenada_salida_buses)
			for nino in grupo:
				print("comparando " + str(nino) + " con ruta " + str(coordenada_salida_buses))
			rutas.append( ruta )
		return rutas

if __name__ == '__main__':

	# Procesa todos los argumentos del programa
	try:
		opts, args = getopt.getopt(sys.argv[1:], "hi:e:", ["help","--instancia=","--excel="])
	except getopt.GetoptError as err:
		# dibuja ayuda y sale:
		print(str(err))  # dibuja qué opción es inválida
		uso()
		sys.exit(2)

	# Define las variables necesarias
	archivo_excel = None
	instancia = None
	for option, value in opts:
		if option in ("-h", "--help"):
			uso()
			sys.exit(1)
		elif option in ("-e", "--excel-file"):
			archivo_excel = value
		elif option in ("-i", "--instancia"):
			try:
				instancia = int(value)
			except:
				print("El argumento que sigue a -i debe ser numérico.")
				sys.exit(1)
		else:
			assert False, "unhandled option"

	if(archivo_excel == None or instancia == None):
		uso()
		sys.exit(1)

	ruteoSolver = RuteoSolver(archivo_excel, instancia)
	print(ruteoSolver)