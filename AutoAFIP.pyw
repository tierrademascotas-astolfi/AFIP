import tkinter as tk
from tkinter import messagebox
import subprocess
import datetime

def Ejecutar_Script_AFIPy():

    """
    Ejecuta el script AFIPy.py utilizando el intérprete de Python.
    Si hoy no es miércoles o domingo, solicita confirmación al usuario.
    
    """

    # Obtener el día actual de la semana (0 = Lunes, 6 = Domingo).
    Dia_Actual = datetime.datetime.now().weekday()

    # Verificar si hoy es miércoles (2) o domingo (6).
    if Dia_Actual in [2, 6]:
        Ejecutar_Script = True
    else:
        # Preguntar al usuario si desea ejecutar el script.
        Ejecutar_Script = messagebox.askyesno("", 
            "Hoy no es miércoles ni domingo. ¿Querés continuar?")

    if not Ejecutar_Script:
        messagebox.showinfo("", "El script no va a correr.")
        Ventana_Principal.destroy()
        return

    try:
        # Ocultar la ventana principal antes de ejecutar el script.
        Ventana_Principal.withdraw()

        # Ejecutar el script de Python usando subprocess.
        subprocess.run(
            ["python", "C:/Users/tomas/Documents/Programación/Github/Programacion/Forrager/AFIP/AFIPy.pyw"],
            check=True
        )
        messagebox.showinfo("", "BotAFIP ejecutado correctamente")
    except Exception as Error:
        messagebox.showerror("Error", f"Ocurrió un error:\n{Error}")

# Crear la ventana principal de Tkinter.
Ventana_Principal = tk.Tk()
Ventana_Principal.title("")
Ventana_Principal.geometry("300x150")

# Centrar la ventana en la pantalla.
Ancho_Pantalla = Ventana_Principal.winfo_screenwidth()  # Obtener ancho de pantalla.
Alto_Pantalla = Ventana_Principal.winfo_screenheight()  # Obtener alto de pantalla.
Ancho_Ventana = 500  
Alto_Ventana = 150  
Posicion_Superior = int(Alto_Pantalla / 2 - Alto_Ventana / 2)  # Calcular posición Y.
Posicion_Izquierda = int(Ancho_Pantalla / 2 - Ancho_Ventana / 2)  # Calcular posición X.
Ventana_Principal.geometry(f'{Ancho_Ventana}x{Alto_Ventana}+{Posicion_Izquierda}+{Posicion_Superior}')  # Establecer posición.

# Agregar un ícono a la ventana.
Ventana_Principal.iconbitmap('C:/Users/tomas/Documents/Programación/Github/Programacion/Forrager/AFIP/Icon.ico') 

# Agregar una etiqueta para recordar al usuario.
Etiqueta_Recordatorio = tk.Label(Ventana_Principal, text="Momento de facturar en AFIP", 
                                font=("Calibri", 14))
Etiqueta_Recordatorio.pack(pady=20)

# Agregar un botón para iniciar el script.
Boton_Inicio = tk.Button(
    Ventana_Principal, text="Iniciar", font=("Calibri", 10), 
    command=Ejecutar_Script_AFIPy
)
Boton_Inicio.pack(pady=10)

# Ejecutar el bucle principal de Tkinter.
Ventana_Principal.mainloop()

