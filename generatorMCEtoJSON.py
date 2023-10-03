import json
import os
import tkinter as tk
import time
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

# Variables globales para rastrear el estado del procesamiento
archivos_procesados = 0
archivos_generados = 0
errores = False
tiempo_inicio = 0

def procesar_carpeta(carpeta):
    global archivos_procesados, archivos_generados, errores, tiempo_inicio
    archivos_procesados = 0
    archivos_generados = 0
    errores = False
    tiempo_inicio = time.time()


    for archivo in os.listdir(carpeta):
        if archivo.endswith(".xlsx") or archivo.endswith(".xls"):
            ruta_archivo = os.path.join(carpeta, archivo)
            archivos_procesados += 1
            try:
                cargar_excel(ruta_archivo)
                archivos_generados += 1
            except Exception as e:
                print(f"Error al procesar {ruta_archivo}: {e}")
                errores = True

    mostrar_resumen()

def cargar_excel(ruta):
    try:

        df = pd.read_excel(ruta, skiprows=1)
        df.fillna(0, inplace=True)
        
        # Obtener el nombre base del archivo Excel (sin extensión)
        nombre_base = os.path.splitext(os.path.basename(ruta))[0]
        # Obtener los primeros 11 caracteres del nombre del archivo Excel
        #Tengo que obtener a partir de lo de adentro del excel cl
        primera_fila_excel = pd.read_excel(ruta, header=None, nrows=1).iloc[0]

        #Todo esto me sirve para obtener el cuit de dentro del mismo archivo excel
        invertir = (primera_fila_excel[0])
        invertir = invertir[::-1]
        invertir = invertir[:11]
        invertir = invertir[::-1]
        #print(invertir)

        cuit = invertir

        # Construir la ruta completa del archivo JSON en la misma carpeta
        ruta_json = os.path.join(os.path.dirname(ruta), f"{nombre_base}.json")

        data_dict = df.to_dict(orient="records")

        # Crear la estructura JSON deseada
        estructura_json = {
            "CUIT": cuit,
            "datosFacturacion": data_dict
        }

        with open(ruta_json, "w", encoding="utf-8") as json_file:
            json.dump(estructura_json, json_file, ensure_ascii=False)

        print(f"Los datos se han guardado en {ruta_json}")

    except Exception as e:
        print("Error al cargar el archivo Excel:", e)

def mostrar_resumen():
    global archivos_procesados, archivos_generados, errores, tiempo_inicio
    tiempo_fin = time.time()
    tiempo_transcurrido = tiempo_fin - tiempo_inicio
    tiempo_transcurrido_formateado = time.strftime("%H:%M:%S", time.gmtime(tiempo_transcurrido))

    mensaje = f"Archivos procesados: {archivos_procesados}\nArchivos generados: {archivos_generados}\n"
    mensaje += f"Tiempo empleado: {tiempo_transcurrido_formateado}\n"
    if errores:
        mensaje += "\nHubo errores durante el procesamiento."

    messagebox.showinfo("Resumen del Proceso", mensaje)

#FUNCION PARA SELECCIONAR LA CARPETA DE TRABAJO
def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        procesar_carpeta(carpeta)

# Crear la ventana de la aplicación
ventana = tk.Tk()
ventana.title("Seleccionar Carpeta con Archivos Excel")

ventana.geometry("400x200")
instrucciones = tk.Label(ventana, text="Script para transformar MCE a JSON (y generar informes PDF)")
instrucciones.pack(pady=10)

# Botón para seleccionar la carpeta
boton_seleccionar = tk.Button(ventana, text="Seleccionar Carpeta", command=seleccionar_carpeta)
boton_seleccionar.pack(pady=20)
boton_cancelar = tk.Button(ventana, text="Cancelar", command=ventana.quit)
boton_cancelar.pack(pady=10)

ventana.mainloop()
