import pandas as pd
from datetime import datetime
import re

def leer_datos(archivo_txt):
    datos = []
    with open(archivo_txt, 'r') as file:
        for line in file:
            # Separar el código y las horas
            codigo, horas = line.strip().split(' - ')
            entrada, salida = horas.split('-')
            datos.append((codigo, entrada, salida))
    return datos

def calcular_porcentajes(datos):
    resultados = {}
    total_minutos = 0

    # Calcular la diferencia en minutos para cada registro
    for codigo, entrada, salida in datos:
        entrada_dt = datetime.strptime(entrada, '%H:%M')
        salida_dt = datetime.strptime(salida, '%H:%M')
        minutos = int((salida_dt - entrada_dt).total_seconds() / 60)
        total_minutos += minutos

        # Si el código ya existe, acumula los minutos
        if codigo in resultados:
            resultados[codigo] += minutos
        else:
            resultados[codigo] = minutos

    # Calcular el porcentaje de cada registro respecto al total de minutos
    for codigo in resultados:
        resultados[codigo] = (resultados[codigo] / total_minutos) * 100

    return resultados

def separar_codigo(codigo):
    # Usar regex para dividir el código en dos partes
    match = re.match(r"(.+?)-(.*)", codigo)
    if match:
        codigo_parte1 = match.group(1)
        producto_nombre = match.group(2)
        return codigo_parte1, producto_nombre
    return codigo, ''  # Si no hay coincidencia, retornar el código completo

def exportar_a_excel(resultados, archivo_salida):
    filas = []

    # Preparar los datos para exportación
    for codigo, porcentaje in resultados.items():
        codigo_parte1, producto_nombre = separar_codigo(codigo)
        filas.append([codigo_parte1, producto_nombre, f"{porcentaje:.2f}%"])

    # Crear el DataFrame y exportarlo a Excel
    df = pd.DataFrame(filas, columns=['Código', 'Producto', 'Porcentaje'])
    df.to_excel(archivo_salida, index=False)

# Ejecución del script
archivo_txt = 'Horas.txt'
archivo_salida = 'Horas_Eficiencia.xlsx'

# Leer, procesar y exportar datos
datos = leer_datos(archivo_txt)
resultados = calcular_porcentajes(datos)
exportar_a_excel(resultados, archivo_salida)

print("Exportación completa a", archivo_salida)
