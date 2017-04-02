import getopt
import math
import sys
import excel
import json

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

class Ruta():
	def __init__(self, nodo_inicial, capacidad_bus):
		self.ruta = [nodo_inicial]
		self.L = 0 # Distancia recorrida
		self.T = 0 # Tiempo de llegada del bus
		self.D = 0 # Distancia recorrida por el bus
		self.R = capacidad_bus # Capacidad del bus
		self.nodo_actual = nodo_inicial
		self.tiempo_llegada_autopista = None

	def __str__(self):
		return json.dumps(self.toJSON())

	def recoger_nino(self, nodo_nino, tiempos, tiempos_recogida, distancias):
		if(self.R <= self.L):
			raise ValueError('A BUS with capacity', self.R, 'is full with', self.L, 'kids on board.')
		self.L = self.L + 1
		self.T = self.T + tiempos[self.nodo_actual][nodo_nino] + tiempos_recogida[nodo_nino]
		self.D = self.D + distancias[self.nodo_actual][nodo_nino]
		self.ruta.append(nodo_nino)
		self.nodo_actual = nodo_nino

	def terminar_recorrido(self, nodo_autopista, nodo_colegio, tiempos, tiempos_recogida, distancias):
		self.tiempo_llegada_autopista = self.T
		self.T = self.T + tiempos[self.nodo_actual][nodo_autopista] + tiempos_recogida[nodo_autopista]
		self.D = self.D + distancias[self.nodo_actual][nodo_autopista]
		self.ruta.append(nodo_autopista)
		self.nodo_actual = nodo_autopista

		self.T = self.T + tiempos[self.nodo_actual][nodo_colegio] + tiempos_recogida[nodo_colegio]
		self.D = self.D + distancias[self.nodo_actual][nodo_colegio]
		self.ruta.append(nodo_colegio)
		self.nodo_actual = nodo_colegio

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

class RuteoSolver():
	def __init__(self, archivo_excel, numero_instancia):
		f = excel.OpenExcel(archivo_excel)
		indice_instancia = numero_instancia + 1
		self.N = int(f.read("C"+str(indice_instancia))) # Cantidad de niños
		self.Q = int(f.read("N"+str(indice_instancia))) # Capacidad de los buses
		self.R = int(f.read("K"+str(indice_instancia))) # Capacidad de los buses

		if(self.N <= 0):
			print("La cantidad de niños debe ser mayor a 0.", file=sys.stderr)
			sys.exit(1)

		if(self.Q <= 0):
			print("La capacidad de los buses debe ser mayor a 0.", file=sys.stderr)
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
		self.tiempos_recogida = [None]
		for i in range(3, 3 + cantidad_buses + self.N + 2):
			tiempo = int(f.read("B"+str(i)))
			self.tiempos_recogida.append(tiempo)

		self.ventana = (150, 210)
		
		self.solucion = self.__solve(self.grupos_ninos, self.nodo_salida_buses, self.nodo_autopista, self.nodo_colegio, self.distancias, self.tiempos, self.tiempos_recogida, self.R)



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
		matriz = [ [ None for i in range(0, cantidad_buses + cantidad_ninos + 2)] ]
		for i in range(3, 3 + cantidad_ninos + cantidad_buses + 2):
			fila = [ None ]
			for j in range(2, 2 + cantidad_ninos + cantidad_buses + 2):
				cell_name = self.__colnum_string(j) + str(i)
				fila.append( int( f.read(cell_name).split(",")[0] ) )
			matriz.append(fila)
		return matriz

	def __solve(self, grupos_ninos, nodo_salida_buses, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus):
		rutas = []
		for grupo in grupos_ninos:
			ruta = self.__encontrar_ruta(grupo, nodo_salida_buses, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus)
			rutas.append(ruta)
		return rutas

	def __encontrar_ruta(self, grupo_ninos, nodo_salida_bus, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus):
		ruta = Ruta(nodo_salida_bus, capacidad_bus)
		try:
			while(len(grupo_ninos) > 0):
				nino_a_recoger = self.__encontrar_siguiente_nino_a_recoger(ruta.nodo_actual, grupo_ninos, distancias)
				ruta.recoger_nino(nino_a_recoger, tiempos, tiempos_recogida, distancias)
				grupo_ninos.remove(nino_a_recoger)
		except ValueError as e:
			print(e, file=sys.stderr)
		ruta.terminar_recorrido(nodo_autopista, nodo_colegio, tiempos, tiempos_recogida, distancias)
		return ruta

	def __encontrar_siguiente_nino_a_recoger(self, nodo_ruta, grupo_ninos, distancias):
		a_recoger = grupo_ninos[0]
		distancia_a_nino_a_recoger = distancias[nodo_ruta][a_recoger]
		for nino in grupo_ninos[1:]:
			if(distancias[nodo_ruta][nino] < distancia_a_nino_a_recoger):
				a_recoger = nino
		return a_recoger

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

	def __str__(self):
		return json.dumps(self.toJSON())

if __name__ == '__main__':

	# Procesa todos los argumentos del programa
	try:
		opts, args = getopt.getopt(sys.argv[1:], "hi:e:", ["help","--instancia=","--excel="])
	except getopt.GetoptError as err:
		# dibuja ayuda y sale:
		print(str(err), file=sys.stderr)  # dibuja qué opción es inválida
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
				print("El argumento que sigue a -i debe ser numérico.", file=sys.stderr)
				sys.exit(1)
		else:
			assert False, "unhandled option"

	if(archivo_excel == None or instancia == None):
		uso()
		sys.exit(1)

	ruteoSolver = RuteoSolver(archivo_excel, instancia)
	print(ruteoSolver)
