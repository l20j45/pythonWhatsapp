import pandas as pd
import pywhatkit as pwk
import time
from datetime import datetime
import pyautogui
import webbrowser
from urllib.parse import quote


def enviar_mensajes_whatsapp_instantaneo(
    archivo_excel, columna_telefono, columna_mensaje, hoja=0
):
    try:
        df = pd.read_excel(archivo_excel, sheet_name=hoja)
        print(
            f"Archivo Excel cargado correctamente. Se encontraron {len(df)} contactos."
        )
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")
        return
    if columna_telefono not in df.columns or columna_mensaje not in df.columns:
        print(f"Error: Una o ambas columnas no existen en el Excel.")
        print(f"Columnas disponibles: {df.columns.tolist()}")
        return

    enviados = 0
    for index, row in df.iterrows():
        try:
            numero = str(row[columna_telefono]).strip()
            mensaje = str(row[columna_mensaje]).strip()
            if pd.isna(numero) or numero == "":
                print(f"Fila {index + 2}: Número de teléfono vacío, saltando...")
                continue
            if numero.startswith("+52"):
                numero = numero[1:]
            numero = numero.replace(" ", "").replace("-", "")

            print(f"Enviando mensaje a {numero}...")
            try:
                mensaje_codificado = quote(mensaje)
                url = f"https://web.whatsapp.com/send?phone={numero}&text={mensaje_codificado}"
                webbrowser.open(url)

                print("Esperando a que cargue WhatsApp Web (15 segundos)...")
                time.sleep(15)
                pyautogui.press("enter")
                time.sleep(1.5)
                pyautogui.hotkey("ctrl", "w")
                enviados += 1
                print(f"Mensaje enviado exitosamente a {numero}")
            except Exception as e:
                print(f"Error en el envío automático: {e}")
            time.sleep(5)

        except Exception as e:
            print(f"Error al procesar el número {numero}: {e}")

    print(f"\nProceso completado. Se enviaron {enviados} de {len(df)} mensajes.")


if __name__ == "__main__":
    archivo_excel = "contactos.xlsx"
    columna_telefono = "telefono"
    columna_mensaje = "mensaje"
    hoja = 0

    # Ejecutar la función principal
    enviar_mensajes_whatsapp_instantaneo(
        archivo_excel, columna_telefono, columna_mensaje, hoja
    )
