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

class Nino():
	def __init__(self, id, x, y, colegio):
		self.id = id
		self.x = x
		self.y = y
		self.colegio = colegio

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

	def __str__(self):
		return json.dumps(self.toJSON())

class Colegio():
	def __init__(self, id, x, y):
		self.id = id
		self.x = x
		self.y = y

	def toJSON(self):
		return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)

	def __str__(self):
		return json.dumps(self.toJSON())

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
		self.tiempos_salida = [0]
		self.tiempos_llegada = [0]
		self.tiempos_ventana = []
		self.colegios_a_visitar = []

	def __str__(self):
		return json.dumps(self.toJSON())

	def recoger_nino(self, nino, tiempos, tiempos_recogida, distancias, colegios):
		if(self.Q <= self.L):
			raise ValueError('Un bus con capacidad', self.R, 'está lleno con', self.L, 'niños a bordo.')
		self.L = self.L + 1
		self.T = self.T + tiempos[self.nodo_actual][nino.id]
		self.tiempos_llegada.append(self.T)
		self.T = self.T + tiempos_recogida[nino.id]
		self.tiempos_salida.append(self.T)
		self.D = self.D + distancias[self.nodo_actual][nino.id]
		self.ruta.append(nino.id)
		try:
			self.colegios_a_visitar.index(colegios[nino.colegio-1])
		except:
			self.colegios_a_visitar.append(colegios[nino.colegio-1])
		self.nodo_actual = nino.id

	def terminar_recorrido(self, nodo_autopista, tiempos, tiempos_recogida, distancias, colegios):
		self.tiempo_llegada_autopista = self.T + tiempos[self.nodo_actual][nodo_autopista]
		self.T = self.T + tiempos[self.nodo_actual][nodo_autopista]
		self.tiempos_llegada.append(self.T)
		self.T = self.T + tiempos_recogida[nodo_autopista]
		self.tiempos_salida.append(self.T)
		self.D = self.D + distancias[self.nodo_actual][nodo_autopista]
		self.ruta.append(nodo_autopista)
		self.nodo_actual = nodo_autopista

		# solamente pasar por los colegios donde hay niños de la ruta
		colegios_a_visitar = [ colegio for colegio in colegios if colegio in self.colegios_a_visitar ]

		while len(colegios_a_visitar) > 0:
			colegio_a_visitar = colegios_a_visitar[0]
			for i in range(1,len(colegios_a_visitar)):
				if (distancias[self.nodo_actual][colegios_a_visitar[i].id] < distancias[self.nodo_actual][colegio_a_visitar.id]):
					colegio_a_visitar = colegios_a_visitar[i]
			self.__visitar_colegio(colegio_a_visitar, tiempos, tiempos_recogida, distancias)
			colegios_a_visitar.remove(colegio_a_visitar)

	def __visitar_colegio(self, colegio, tiempos, tiempos_recogida, distancias):
		self.T = self.T + tiempos[self.nodo_actual][colegio.id]
		self.tiempos_llegada.append(self.T)
		self.T = self.T + tiempos_recogida[colegio.id]
		self.tiempos_salida.append(self.T)
		self.D = self.D + distancias[self.nodo_actual][colegio.id]
		self.ruta.append(colegio.id)
		self.nodo_actual = colegio.id

	def calcular_tiempos_ventana_primera_ruta(self, ventana):
		tiempo_min = ventana[0]
		tiempo_max = ventana[1]
		self.tiempos_llegada.reverse()
		diff = tiempo_max - self.tiempos_llegada[0]
		for tiempo in self.tiempos_llegada:
			tiempo_ventana = tiempo + diff
			if(tiempo_ventana < 0):
				raise ValueError("Tiempo Negativo")
			self.tiempos_ventana.append(tiempo_ventana)
		self.tiempo_llegada_autopista_ventana = self.tiempo_llegada_autopista + diff
		self.tiempos_llegada.reverse()
		if(self.tiempos_ventana[0] < tiempo_min or self.tiempos_ventana[0] > tiempo_max):
			self.tiempos_ventana.reverse()
			raise ValueError('La solución no es factible. Revisar la capacidad de los buses.')
		self.tiempos_ventana.reverse()
		return self.tiempo_llegada_autopista_ventana

	def calcular_tiempos_ventana_otras_rutas(self, tiempo_entrada_autopista_siguiente_ruta, tiempo_intervalo_obligatorio, ventana):
		tiempo_min = ventana[0]
		tiempo_max = ventana[1]
		self.tiempos_llegada.reverse()
		diff = tiempo_entrada_autopista_siguiente_ruta - self.tiempo_llegada_autopista
		for tiempo in self.tiempos_llegada:
			tiempo_ventana = tiempo + diff
			if(tiempo_ventana < 0):
				raise ValueError("Tiempo Negativo")
			self.tiempos_ventana.append(tiempo_ventana)
		self.tiempo_llegada_autopista_ventana = self.tiempo_llegada_autopista + diff
		self.tiempos_llegada.reverse()
		if(self.tiempos_ventana[0] < tiempo_min or self.tiempos_ventana[0] > tiempo_max):
			self.tiempos_ventana.reverse()
			raise ValueError('La solución no es factible. Revisar la capacidad de los buses.')
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
			datos_ninos.append( Nino( int(f.read("A"+str(i))), int(f.read("B"+str(i))), int(f.read("C"+str(i))), int(f.read("E"+str(i))) ) )
		i = i + 1
		self.nodo_autopista = int(f.read("A"+str(i)))
		i = i + 1
		self.nodos_colegios = []
		for j in range (i, i + self.cantidad_colegios):
			self.nodos_colegios.append(Colegio( int(f.read("A"+str(j))), int(f.read("B"+str(j))), int(f.read("C"+str(j))) ))
		j = j + 1
		self.nodo_final = int(f.read("A"+str(j)))

		self.grupos_ninos = self.__crear_grupos(datos_ninos, math.ceil( self.N / self.NB ))
		self.distancias = self.__leer_matriz(archivo_excel, 3, self.N, cantidad_buses, self.cantidad_colegios)
		self.costos = self.__leer_matriz(archivo_excel, 4, self.N, cantidad_buses, self.cantidad_colegios)
		# self.estudiante_colegio = self.__leer_matriz(archivo_excel, 5, self.N, cantidad_buses, self.cantidad_colegios)
		self.tiempos = self.__leer_matriz(archivo_excel, 6, self.N, cantidad_buses, self.cantidad_colegios)

		f = excel.OpenExcel(archivo_excel, sheet = 7)
		self.tiempos_recogida = [None]
		for i in range(3, 3 + cantidad_buses + self.N + self.cantidad_colegios + 2):
			tiempo = int(f.read("B"+str(i)))
			self.tiempos_recogida.append(tiempo)

		self.ventana = (150, 210)

		resultado = False
		iteraciones = 0
		tiempo_inicio = time.time()
		while(not resultado):
			iteraciones = iteraciones + 1
			print("Iteración", iteraciones)
			resultado = self.__iteracion(datos_ninos)
		self.Z = self.CF * len(self.rutas) + self.CU * sum([ ruta.D for ruta in self.rutas ])
		tiempo_total = time.time() - tiempo_inicio

		self.__escribir_resultado(archivo_salida, self.Z, self.rutas, tiempo_total)

	def __iteracion(self, datos_ninos):
		try:
			self.grupos_ninos = self.__crear_grupos(datos_ninos, math.ceil( self.N / self.NB ))
			self.rutas = self.__encontrar_rutas(self.grupos_ninos, self.nodo_salida_buses, self.nodo_autopista, self.distancias, self.tiempos, self.tiempos_recogida, self.Q, self.nodos_colegios)
			self.rutas = self.__cuadrar_ventana(self.rutas, self.ventana, self.R)
		except ValueError as e:
			if(e.args[0] != "Tiempo Negativo"):
				raise e
			self.NB = self.NB  + 1
			if(self.NB > self.N):
				raise ValueError("No pueden haber más rutas que niños")
			return False
		return True

	def __crear_grupos(self, datos_ninos, tamanio_grupos):
		datos_ninos = [ dato_nino for dato_nino in sorted(datos_ninos, key=lambda x: x.y) ]
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

	def __encontrar_rutas(self, grupos_ninos, nodo_salida_buses, nodo_autopista, distancias, tiempos, tiempos_recogida, capacidad_bus, colegios):
		rutas = []
		for grupo in grupos_ninos:
			rutas_nuevas = self.__encontrar_rutas_nuevas(grupo, nodo_salida_buses, nodo_autopista, distancias, tiempos, tiempos_recogida, capacidad_bus, colegios)
			rutas = rutas + rutas_nuevas
		return rutas

	def __encontrar_rutas_nuevas(self, grupo_ninos, nodo_salida_bus, nodo_autopista, distancias, tiempos, tiempos_recogida, capacidad_bus, colegios):
		nuevas_rutas = []
		cantidad_rutas = math.ceil(len(grupo_ninos) / capacidad_bus)
		for ruta_n in range(cantidad_rutas):
			ruta = Ruta(nodo_salida_bus, capacidad_bus)
			while(len(grupo_ninos) > 0 and ruta.L < capacidad_bus):
				nino_a_recoger = self.__encontrar_siguiente_nino_a_recoger(ruta.nodo_actual, grupo_ninos, distancias)
				ruta.recoger_nino(nino_a_recoger, tiempos, tiempos_recogida, distancias, colegios)
				grupo_ninos.remove(nino_a_recoger)
			ruta.terminar_recorrido(nodo_autopista, tiempos, tiempos_recogida, distancias, colegios)
			nuevas_rutas.append(ruta)
		return nuevas_rutas

	def __encontrar_siguiente_nino_a_recoger(self, nodo_ruta, grupo_ninos, distancias):
		a_recoger = grupo_ninos[0]
		distancia_a_nino_a_recoger = distancias[nodo_ruta][a_recoger.id]
		for nino in grupo_ninos[1:]:
			if(distancias[nodo_ruta][nino.id] < distancia_a_nino_a_recoger):
				a_recoger = nino
				distancia_a_nino_a_recoger = distancias[nodo_ruta][nino.id]
		return a_recoger

	def __cuadrar_ventana(self, rutas, ventana, tiempo_intervalo_obligatorio):
		rutas = sorted(rutas, key=lambda x: x.T, reverse=True)
		tiempo_entrada_autopista_siguiente_ruta = rutas[0].calcular_tiempos_ventana_primera_ruta(ventana) - tiempo_intervalo_obligatorio
		for ruta in rutas[1:]:
			tiempo_entrada_autopista_siguiente_ruta = ruta.calcular_tiempos_ventana_otras_rutas(tiempo_entrada_autopista_siguiente_ruta, tiempo_intervalo_obligatorio, ventana)
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
			f.write("Tiempos Llegada: " + str(ruta.tiempos_ventana) + os.linesep)
			f.write("Tiempo Llegada Autopista: " + str(ruta.tiempo_llegada_autopista_ventana) + os.linesep)
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

	RuteoSolver(archivo_excel, instancia, archivo_resultado)