import getopt
import math
import sys
import excel
import json
import os
import time

def uso():
	print("""
USO:
\tpython""", sys.argv[0], """ -e <archivo_excel> -i <numero_instancia>

OPCIONES:

\t-e, --excel <archivo_excel>
\t\tEspecifica el archivo de excel sobre el cual se trabaja.

\t-i, --instancia
\t\tEspecifica el número de la instancia especificada en el archivo excel.

\t-o, --output <archivo_salida>
\t\tEspecifica el nombre del archivo resultado.

\t-h, --help
\t\tMuestra esta ayuda y termina.
	""")

class Ruta():
	def __init__(self, nodo_inicial, capacidad_bus):
		self.ruta = [nodo_inicial]
		self.L = 0 # Distancia recorrida
		self.T = 0 # Tiempo de llegada del bus
		self.D = 0 # Distancia recorrida por el bus
		self.Q = capacidad_bus # Capacidad del bus
		self.nodo_actual = nodo_inicial
		self.tiempo_llegada_autopista = None
		self.tiempo_llegada_autopista_ventana = None
		self.tiempos = [0]
		self.tiempos_ventana = []

	def __str__(self):
		return json.dumps(self.toJSON())

	def recoger_nino(self, nodo_nino, tiempos, tiempos_recogida, distancias):
		if(self.Q <= self.L):
			raise ValueError('A BUS with capacity', self.R, 'is full with', self.L, 'kids on board.')
		self.L = self.L + 1
		self.T = self.T + tiempos[self.nodo_actual][nodo_nino] + tiempos_recogida[nodo_nino]
		self.tiempos.append(self.T)
		self.D = self.D + distancias[self.nodo_actual][nodo_nino]
		self.ruta.append(nodo_nino)
		self.nodo_actual = nodo_nino

	def terminar_recorrido(self, nodo_autopista, nodo_colegio, tiempos, tiempos_recogida, distancias):
		self.tiempo_llegada_autopista = self.T + tiempos[self.nodo_actual][nodo_autopista]
		self.T = self.T + tiempos[self.nodo_actual][nodo_autopista] + tiempos_recogida[nodo_autopista]
		self.D = self.D + distancias[self.nodo_actual][nodo_autopista]
		self.tiempos.append(self.T)
		self.ruta.append(nodo_autopista)
		self.nodo_actual = nodo_autopista

		self.T = self.T + tiempos[self.nodo_actual][nodo_colegio] + tiempos_recogida[nodo_colegio]
		self.D = self.D + distancias[self.nodo_actual][nodo_colegio]
		self.tiempos.append(self.T)
		self.ruta.append(nodo_colegio)
		self.nodo_actual = nodo_colegio

	def calcular_tiempos_ventana_primera_ruta(self, tiempo_final):
		self.tiempos.reverse()
		diff = tiempo_final - self.tiempos[0]
		for tiempo in self.tiempos:
			self.tiempos_ventana.append(tiempo + diff)
		self.tiempo_llegada_autopista_ventana = self.tiempo_llegada_autopista + diff
		self.tiempos.reverse()
		self.tiempos_ventana.reverse()
		return self.tiempo_llegada_autopista_ventana

	def calcular_tiempos_ventana_otras_rutas(self, tiempo_entrada_autopista_siguiente_ruta, tiempo_intervalo_obligatorio):
		self.tiempos.reverse()
		diff = tiempo_entrada_autopista_siguiente_ruta - self.tiempo_llegada_autopista
		for tiempo in self.tiempos:
			self.tiempos_ventana.append(tiempo + diff)
		self.tiempo_llegada_autopista_ventana = self.tiempo_llegada_autopista + diff
		self.tiempos.reverse()
		self.tiempos_ventana.reverse()
		return self.tiempo_llegada_autopista_ventana - tiempo_intervalo_obligatorio

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

class RuteoSolver():
	def __init__(self, archivo_excel, numero_instancia, archivo_salida):
		f = excel.OpenExcel(archivo_excel)
		indice_instancia = numero_instancia + 1
		self.N = int(f.read("C"+str(indice_instancia))) # Cantidad de niños
		self.Q = int(f.read("P"+str(indice_instancia))) # Capacidad de los buses
		self.cantidad_colegios = int(f.read("I"+str(indice_instancia))) # Cantidad de Colegios
		self.R = int(f.read("M"+str(indice_instancia))) # Tiempo Mínimo de Secuenciamiento en Autopista
		self.CF = int(f.read("O"+str(indice_instancia))) # Costo Fijo
		self.CU = int(f.read("Q"+str(indice_instancia))) # Costo Unitario

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
		self.colegios = []
		for j in range (i, i + self.cantidad_colegios):
			self.colegios.append(( int(f.read("A"+str(j))), int(f.read("B"+str(j))), int(f.read("C"+str(j))) ))
		j = j + 1
		self.nodo_final = int(f.read("A"+str(j)))

		self.grupos_ninos = self.__crear_grupos(datos_ninos, math.ceil( self.N / self.NB ))
		self.distancias = self.__leer_matriz(archivo_excel, 3, self.N, cantidad_buses, self.cantidad_colegios)
		self.costos = self.__leer_matriz(archivo_excel, 4, self.N, cantidad_buses, self.cantidad_colegios)
		self.estudiante_colegio = self.__leer_matriz(archivo_excel, 5, self.N, cantidad_buses, self.cantidad_colegios)
		self.tiempos = self.__leer_matriz(archivo_excel, 6, self.N, cantidad_buses, self.cantidad_colegios)

		f = excel.OpenExcel(archivo_excel, sheet = 7)
		self.tiempos_recogida = [None]
		for i in range(3, 3 + cantidad_buses + self.N + self.cantidad_colegios + 2):
			tiempo = int(f.read("B"+str(i)))
			self.tiempos_recogida.append(tiempo)

		self.ventana = (150, 210)

		"""		
		tiempo_inicio = time.time()
		self.rutas = self.__encontrar_rutas(self.grupos_ninos, self.nodo_salida_buses, self.nodo_autopista, self.nodo_colegio, self.distancias, self.tiempos, self.tiempos_recogida, self.Q)
		self.rutas = self.__cuadrar_ventana(self.rutas, self.ventana, self.R)
		self.Z = self.CF * len(self.rutas) + self.CU * sum([ ruta.D for ruta in self.rutas ])
		tiempo_total = time.time() - tiempo_inicio

		self.__escribir_resultado(archivo_salida, self.Z, self.rutas, tiempo_total)
		"""

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


	def __leer_matriz(self, archivo_excel, numero_hoja, cantidad_ninos, cantidad_buses, cantidad_colegios):
		f = excel.OpenExcel(archivo_excel, numero_hoja)
		matriz = [ [ None for i in range(0, cantidad_buses + cantidad_ninos + cantidad_colegios + 3)] ]
		for i in range(3, 3 + cantidad_ninos + cantidad_buses + cantidad_colegios + 2):
			fila = [ None ]
			for j in range(2, 2 + cantidad_ninos + cantidad_buses + cantidad_colegios + 2):
				cell_name = self.__colnum_string(j) + str(i)
				try:
					fila.append( int( f.read(cell_name).split(",")[0] ) )
				except:
					fila.append( int( f.read(cell_name) ) )
			matriz.append(fila)
		return matriz

	def __encontrar_rutas(self, grupos_ninos, nodo_salida_buses, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus):
		rutas = []
		for grupo in grupos_ninos:
			rutas_nuevas = self.__encontrar_rutas_nuevas(grupo, nodo_salida_buses, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus)
			rutas = rutas + rutas_nuevas
		return rutas

	def __encontrar_rutas_nuevas(self, grupo_ninos, nodo_salida_bus, nodo_autopista, nodo_colegio, distancias, tiempos, tiempos_recogida, capacidad_bus):
		nuevas_rutas = []
		cantidad_rutas = math.ceil(len(grupo_ninos) / capacidad_bus)
		for ruta_n in range(cantidad_rutas):
			ruta = Ruta(nodo_salida_bus, capacidad_bus)
			while(len(grupo_ninos) > 0 and ruta.L < capacidad_bus):
				nino_a_recoger = self.__encontrar_siguiente_nino_a_recoger(ruta.nodo_actual, grupo_ninos, distancias)
				ruta.recoger_nino(nino_a_recoger, tiempos, tiempos_recogida, distancias)
				grupo_ninos.remove(nino_a_recoger)
			ruta.terminar_recorrido(nodo_autopista, nodo_colegio, tiempos, tiempos_recogida, distancias)
			nuevas_rutas.append(ruta)
		return nuevas_rutas

	def __encontrar_siguiente_nino_a_recoger(self, nodo_ruta, grupo_ninos, distancias):
		a_recoger = grupo_ninos[0]
		distancia_a_nino_a_recoger = distancias[nodo_ruta][a_recoger]
		for nino in grupo_ninos[1:]:
			if(distancias[nodo_ruta][nino] < distancia_a_nino_a_recoger):
				a_recoger = nino
				distancia_a_nino_a_recoger = distancias[nodo_ruta][nino]
		return a_recoger

	def __cuadrar_ventana(self, rutas, ventana, tiempo_intervalo_obligatorio):
		rutas = sorted(rutas, key=lambda x: x.T, reverse=True)
		tiempo_entrada_autopista_siguiente_ruta = rutas[0].calcular_tiempos_ventana_primera_ruta(ventana[1]) - tiempo_intervalo_obligatorio
		for ruta in rutas[1:]:
			tiempo_entrada_autopista_siguiente_ruta = ruta.calcular_tiempos_ventana_otras_rutas(tiempo_entrada_autopista_siguiente_ruta, tiempo_intervalo_obligatorio)
		return rutas
	def __escribir_resultado(self, archivo, Z, rutas, tiempo_ejecuion):
		f = open(archivo, "w")
		f.write("Z: " + str(Z) + os.linesep)
		f.write("Cantidad Rutas: " + str(len(rutas))  + os.linesep)
		f.write("========================================="  + os.linesep)
		i = 0
		for ruta in rutas:
			i = i+1
			f.write("Ruta["+ str(i) + "]:" + os.linesep)
			f.write("Nodos Visitados: " + str(ruta.ruta) + os.linesep)
			f.write("Tiempos: " + str(ruta.tiempos_ventana) + os.linesep)
			f.write("Tiempo Legada Autopista: " + str(ruta.tiempo_llegada_autopista_ventana) + os.linesep)
			f.write("========================================="  + os.linesep)
		f.write("Tiempo Ejecución: " + str(tiempo_ejecuion)  + " segundos" + os.linesep)
		f.close()

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

	def __str__(self):
		return json.dumps(self.toJSON())

if __name__ == '__main__':

	# Procesa todos los argumentos del programa
	try:
		opts, args = getopt.getopt(sys.argv[1:], "hi:e:o:", ["help","--instancia=","--excel=","--output="])
	except getopt.GetoptError as err:
		# dibuja ayuda y sale:
		print(str(err), file=sys.stderr)  # dibuja qué opción es inválida
		uso()
		sys.exit(2)

	# Define las variables necesarias
	archivo_excel = None
	instancia = None
	archivo_resultado = None
	for option, value in opts:
		if option in ("-h", "--help"):
			uso()
			sys.exit(1)
		elif option in ("-e", "--excel-file"):
			archivo_excel = value
		elif option in ("-o", "--output"):
			archivo_resultado = value
		elif option in ("-i", "--instancia"):
			try:
				instancia = int(value)
			except:
				print("El argumento que sigue a -i debe ser numérico.", file=sys.stderr)
				sys.exit(1)
		else:
			assert False, "unhandled option"

	if(archivo_excel == None or instancia == None or archivo_resultado == None):
		uso()
		sys.exit(1)

	r = RuteoSolver(archivo_excel, instancia, archivo_resultado)
	print(r)