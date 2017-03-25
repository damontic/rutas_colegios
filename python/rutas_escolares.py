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
	def __init__(self, nodo_inicial):
		self.ruta = [nodo_inicial]
		self.L = 0 # Distancia recorrida
		self.T = 0 # Tiempo de llegada del bus
		self.nodo_actual = nodo_inicial

	def __str__(self):
		return "Ruta:{" + \
		"\n\t\t\"L\":" + str(self.L) + "," + \
		"\n\t\t\"T\":" + str(self.T) + "," + \
		"\n\t\t\"ruta\":" + str(self.ruta) + "," + \
		"\n\t\t\"nodo_actual\":" + str(self.nodo_actual) + "," + \
		"\n}"

	def recoger_nino(self, nodo_nino, tiempos, tiempos_recogida):
		self.L = self.L + 1
		self.T = self.T + tiempos[self.nodo_actual][nodo_nino] + tiempos_recogida[nodo_nino]
		self.ruta.append(nodo_nino)
		self.nodo_actual = nodo_nino

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

		cantidad_buses = int(f.read("E"+str(indice_instancia)))
		indice_inicio_ninos_hoja_2 = cantidad_buses + 6

		f = excel.OpenExcel(archivo_excel, sheet = 2)
		self.nodo_salida_buses = int(f.read("A6"))
		datos_ninos = []
		for i in range(indice_inicio_ninos_hoja_2, indice_inicio_ninos_hoja_2 + self.N):
			datos_ninos.append( ( int(f.read("A"+str(i))), int(f.read("B"+str(i))), int(f.read("C"+str(i))) ) ) 
		i = i + 1
		self.nodo_autopista = int(f.read("A"+str(i)))
		i = i + 1
		self.nodo_colegio = int(f.read("A"+str(i)))

		self.grupos_ninos = self.__crear_grupos(datos_ninos, math.ceil( self.N / self.NB ))
		self.distancias = self.__leer_matriz(archivo_excel, 3, self.N, cantidad_buses)
		self.costos = self.__leer_matriz(archivo_excel, 4, self.N, cantidad_buses)
		self.tiempos = self.__leer_matriz(archivo_excel, 5, self.N, cantidad_buses)

		f = excel.OpenExcel(archivo_excel, sheet = 6)
		self.tiempos_recogida = []
		for i in range(3, 3 + cantidad_buses + self.N + 2):
			tiempo = int(f.read("B"+str(i)))
			self.tiempos_recogida.append(tiempo)
		
		self.solucion = self.__solve(self.grupos_ninos, self.nodo_salida_buses, self.distancias, self.tiempos, self.tiempos_recogida)

	def __str__(self):
		soluciones = [ str(ruta) for ruta in self.solucion ]
		return "RuteoSolver:{" + \
		"\n\t\"Q\":" + str(self.Q) + "," + \
		"\n\t\"N\":" + str(self.N) + "," + \
		"\n\t\"NB\":" + str(self.NB) + "," + \
		"\n\t\"nodo_salida_buses\":" + str(self.nodo_salida_buses) + "," + \
		"\n\t\"nodo_autopista\":" + str(self.nodo_autopista) + "," + \
		"\n\t\"nodo_colegio\":" + str(self.nodo_colegio) + "," + \
		"\n\t\"grupos_ninos\":" + str(self.grupos_ninos) + "," + \
		"\n\t\"distancias\":" + str(self.distancias) + "," + \
		"\n\t\"costos\":" + str(self.costos) + "," + \
		"\n\t\"tiempos\":" + str(self.tiempos) + "," + \
		"\n\t\"tiempos_recogida\":" + str(self.tiempos_recogida) + "," + \
		"\n\t\"solucion\":[\n\t" + "\n\t".join(soluciones) + "]" + \
		"\n}"

	def __crear_grupos(self, datos_ninos, tamanio_grupos):
		datos_ninos = [ dato_nino[0] for dato_nino in sorted(datos_ninos, key=lambda x: x[2]) ]
		return [ datos_ninos[i:i + tamanio_grupos] for i in range(0, len(datos_ninos), tamanio_grupos) ]

	def __colnum_string(self, n):
		div=n
		string=""
		temp=0
		while div>0:
			module=(div-1)%26
			string=chr(65+module)+string
			div=int((div-module)/26)
		return string


	def __leer_matriz(self, archivo_excel, numero_hoja, cantidad_ninos, cantidad_buses):
		f = excel.OpenExcel(archivo_excel, numero_hoja)
		matriz = []
		for i in range(3, 3 + cantidad_ninos + cantidad_buses + 2):
			fila = []
			for j in range(2, 2 + cantidad_ninos + cantidad_buses + 2):
				cell_name = self.__colnum_string(j) + str(i)
				fila.append( int( f.read(cell_name).split(",")[0] ) )
			matriz.append(fila)
		return matriz

	def __solve(self, grupos_ninos, nodo_salida_buses, distancias, tiempos, tiempos_recogida):
		rutas = []
		for grupo in grupos_ninos:
			ruta = self.__encontrar_ruta(grupo, nodo_salida_buses, distancias, tiempos, tiempos_recogida)
			rutas.append(ruta)
		return rutas

	def __encontrar_ruta(self, grupo_ninos, nodo_salida_bus, distancias, tiempos, tiempos_recogida):
		ruta = Ruta(nodo_salida_bus)
		while(len(grupo_ninos) > 0):
			nino_a_recoger = self.__encontrar_siguiente_nino_a_recoger(ruta.nodo_actual, grupo_ninos, distancias)
			ruta.recoger_nino(nino_a_recoger, tiempos, tiempos_recogida)
			grupo_ninos.remove(nino_a_recoger)
		return ruta

	def __encontrar_siguiente_nino_a_recoger(self, nodo_ruta, grupo_ninos, distancias):
		a_recoger = grupo_ninos[0]
		distancia_a_nino_a_recoger = distancias[nodo_ruta][a_recoger]
		for nino in grupo_ninos[1:]:
			if(distancias[nodo_ruta][nino] < distancia_a_nino_a_recoger):
				a_recoger = nino
		return a_recoger

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