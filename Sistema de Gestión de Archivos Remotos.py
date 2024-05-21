import os
import stat
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import paramiko

def limpiar_nombre(nombre):
    """Elimina caracteres no válidos del nombre del archivo."""
    caracteres_no_validos = ['\\', '/', '*', '[', ']', ':', '?']
    for caracter in caracteres_no_validos:
        nombre = nombre.replace(caracter, '_')
    return nombre[:30]

#Listado de archivos:

#Linux

def listar_archivos_remotos_linux(sftp_client, carpeta):
    archivos = []
    try:
        # Obtener la lista de archivos y carpetas en la carpeta remota
        for entry in sftp_client.listdir_attr(carpeta):
            ruta_completa = carpeta + "/" + entry.filename  # Utiliza "/" como separador de ruta en lugar de os.path.join()
            # Si es una carpeta, explorar recursivamente
            if stat.S_ISDIR(entry.st_mode):
                archivos.extend(listar_archivos_remotos_linux(sftp_client, ruta_completa))
            else:
                # Si es un archivo, agregar a la lista
                fecha_modificacion = datetime.fromtimestamp(entry.st_mtime)
                fecha_acceso = datetime.fromtimestamp(entry.st_atime)
                nombre, extension = os.path.splitext(entry.filename)
                ruta_padre = os.path.dirname(ruta_completa)
                archivos.append((nombre, extension, fecha_modificacion, fecha_acceso, ruta_padre))
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"No se encontró la carpeta remota: {carpeta} - {e}")
    except PermissionError as e:
        messagebox.showerror("Error", f"Error de permisos con la carpeta: {carpeta} - {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado al listar archivos remotos: {e}")
    return archivos

#Windows

def listar_archivos_remotos_windows(sftp_client, carpeta):
    archivos = []
    try:
        # Obtener la lista de archivos y carpetas en la carpeta remota
        for entry in sftp_client.listdir_attr(carpeta):
            ruta_completa = os.path.join(carpeta, entry.filename.replace('/', '\\'))
            # Si es una carpeta, explorar recursivamente
            if stat.S_ISDIR(entry.st_mode):
                archivos.extend(listar_archivos_remotos_windows(sftp_client, ruta_completa))
            else:
                # Si es un archivo, agregar a la lista
                fecha_modificacion = datetime.fromtimestamp(entry.st_mtime)
                fecha_acceso = datetime.fromtimestamp(entry.st_atime)
                nombre, extension = os.path.splitext(entry.filename)
                ruta_padre = os.path.dirname(ruta_completa)
                archivos.append((nombre, extension, fecha_modificacion, fecha_acceso, ruta_padre))
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"No se encontró la carpeta remota: {carpeta} - {e}")
    except PermissionError as e:
        messagebox.showerror("Error", f"Error de permisos con la carpeta: {carpeta} - {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado al listar archivos remotos: {e}")
    return archivos

def agregar_hoja_excel(libro_excel, nombre_hoja):
    """Agrega una nueva hoja al archivo Excel."""
    nombre_hoja = limpiar_nombre(nombre_hoja)
    if nombre_hoja in libro_excel.sheetnames:
        messagebox.showinfo("Información", f"La hoja '{nombre_hoja}' ya existe en el archivo de Excel. Es necesario cambiar el nombre.")
        nuevo_nombre = input("Ingrese un nuevo nombre para la hoja: ")
        return agregar_hoja_excel(libro_excel, nuevo_nombre)
    else:
        libro_excel.create_sheet(title=nombre_hoja)
        return nombre_hoja

def ajustar_ancho_columnas(hoja):
    """Ajusta el ancho de las columnas en la hoja de Excel."""
    for col in hoja.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        hoja.column_dimensions[column].width = adjusted_width

def guardar_en_excel(archivos, nombre_archivo, nombre_hoja_nueva=None):
    """Guarda la lista de archivos en un archivo Excel."""
    if not nombre_archivo:
        nombre_archivo = "Bitacora.xlsx"
    if not nombre_archivo.lower().endswith('.xlsx'):
        nombre_archivo += '.xlsx'
    if os.path.exists(nombre_archivo):
        libro_excel = load_workbook(nombre_archivo)
    else:
        libro_excel = Workbook()
    nombre_hoja = nombre_hoja_nueva if nombre_hoja_nueva else "Datos"
    nombre_hoja_final = agregar_hoja_excel(libro_excel, nombre_hoja)
    hoja = libro_excel[nombre_hoja_final]

    # Agregar solo las columnas necesarias
    hoja.append(["ID", "Nombre de dato / archivo", "Formato", "Fecha de modificación", "Fecha de acceso", "Ruta"])

    for id_archivo, archivo in enumerate(archivos, start=1):
        hoja.append([id_archivo] + list(archivo[:5]))  # Seleccionar solo las primeras 5 columnas
    ajustar_ancho_columnas(hoja)
    
    try:
        libro_excel.save(nombre_archivo)
        messagebox.showinfo("Éxito", "Datos guardados exitosamente en el archivo de Excel.")
    except PermissionError as e:
        messagebox.showerror("Error", f"Error de permisos al guardar el archivo: {nombre_archivo} - {e}")


def establecer_conexion(host, username, use_private_key, private_key_path=None, password=None):
    """Establece la conexión SSH con el servidor remoto."""
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        if use_private_key:
            private_key = paramiko.RSAKey.from_private_key_file(private_key_path)
            ssh_client.connect(hostname=host, username=username, pkey=private_key)
        else:
            ssh_client.connect(hostname=host, username=username, password=password)
        return ssh_client
    except paramiko.AuthenticationException:
        messagebox.showerror("Error", "Error de autenticación. Verifica tus credenciales e intenta de nuevo.")
        return None
    except paramiko.SSHException as e:
        messagebox.showerror("Error", f"Error al establecer la conexión SSH: {e}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")
        return None

def iniciar_operacion():
    """Inicia la operación de análisis de la carpeta remota y exportación a Excel."""
    host = host_entry.get()
    username = username_entry.get()
    use_private_key = use_private_key_var.get()
    private_key_path = private_key_path_entry.get() if use_private_key else None
    password = password_entry.get() if not use_private_key else None
    carpeta = carpeta_entry.get()
    nombre_archivo_excel = nombre_archivo_entry.get()
    nombre_hoja_nueva = nombre_hoja_entry.get()

    ssh_client = establecer_conexion(host, username, use_private_key, private_key_path, password)

    if ssh_client:
        try:
            # Determinar qué función utilizar para listar archivos remotos según el sistema operativo
            if seleccion_sistema.get() == "linux":
                archivos = listar_archivos_remotos_linux(ssh_client.open_sftp(), carpeta)
            else:
                archivos = listar_archivos_remotos_windows(ssh_client.open_sftp(), carpeta)

            ssh_client.close()

            guardar_en_excel(archivos, nombre_archivo_excel, nombre_hoja_nueva)
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar la operación: {e}")
    else:
        messagebox.showerror("Error", "No se pudo establecer la conexión SSH.")


# Función para seleccionar el sistema operativo
def seleccionar_sistema_operativo():
    global seleccion_sistema
    sistema_operativo = seleccion_sistema.get()
    iniciar_programa(sistema_operativo)

# Función principal para iniciar el programa según el sistema operativo seleccionado
def iniciar_programa(sistema_operativo):
    if sistema_operativo == "linux":
        usar_barra = "/"
        listar_archivos_remotos = listar_archivos_remotos_linux  # Usar la función para Linux
    else:
        usar_barra = "\\"
        listar_archivos_remotos = listar_archivos_remotos_windows  # Usar la función para Windows
    
    print(f"Se seleccionó el sistema operativo {sistema_operativo}. Se usará la barra '{usar_barra}' para las rutas.")
    # Aquí puedes llamar a las funciones necesarias para el programa, pasando la variable 'usar_barra'

    # Por ejemplo, llamar a la función para listar archivos remotos
    # archivos = listar_archivos_remotos(sftp_client, carpeta)

def mostrar_ocultar_llave_privada():
    """Muestra u oculta la entrada de la ruta de la clave privada."""
    if use_private_key_var.get():
        private_key_path_label.grid()
        private_key_path_entry.grid()
    else:
        private_key_path_label.grid_remove()
        private_key_path_entry.grid_remove()

def actualizar_barra_progreso():
    """Actualiza la barra de progreso."""
    for i in range(101):
        progress_bar["value"] = i
        root.update_idletasks()
        root.after(100)  # Simula el tiempo que tarda en completarse una tarea (en milisegundos)

    # Restablece la barra de progreso una vez completada la tarea
    progress_bar["value"] = 0


# Inicialización de la variable modo_oscuro
modo_oscuro = False

def cambiar_modo_oscuro():
    """Cambia entre el modo oscuro y claro."""
    global modo_oscuro
    modo_oscuro = not modo_oscuro
    if modo_oscuro:
        root.config(background="black")
        for widget in root.winfo_children():
            if isinstance(widget, tk.Label) or isinstance(widget, tk.Button):
                widget.config(background="black", foreground="white")
            elif isinstance(widget, tk.Entry) or isinstance(widget, tk.Checkbutton):
                widget.config(background="black", foreground="white")
            elif isinstance(widget, tk.Radiobutton):
                widget.config(background="black", foreground="white", selectcolor="black")
            elif isinstance(widget, tk.Toplevel):
                widget.config(background="black")
            elif isinstance(widget, ttk.Progressbar):
                widget.config(style="alt.Horizontal.TProgressbar")  # Cambiar el estilo del ttk.Progressbar
    else:
        root.config(background="white")
        for widget in root.winfo_children():
            if isinstance(widget, tk.Label) or isinstance(widget, tk.Button):
                widget.config(background="white", foreground="black")
            elif isinstance(widget, tk.Entry) or isinstance(widget, tk.Checkbutton):
                widget.config(background="white", foreground="black")
            elif isinstance(widget, tk.Radiobutton):
                widget.config(background="white", foreground="black", selectcolor="white")
            elif isinstance(widget, tk.Toplevel):
                widget.config(background="white")
            elif isinstance(widget, ttk.Progressbar):
                widget.config(style="alt.Horizontal.TProgressbar")  # Cambiar el estilo del ttk.Progressbar

def mostrar_ocultar_nombre_archivo_excel():
    """Muestra u oculta la entrada del nombre del archivo Excel."""
    if nombre_archivo_var.get():
        nombre_archivo_label.grid()
        nombre_archivo_entry.grid()
    else:
        nombre_archivo_label.grid_remove()
        nombre_archivo_entry.grid_remove()

# ENTORNO VISUAL

root = tk.Tk()
root.title("необнаружимый шпионский код")

# Configuración de la interfaz de usuario
for i in range(10):
    root.grid_rowconfigure(i, weight=1)
    root.grid_columnconfigure(1, weight=1)

# Entradas y etiquetas

# Etiqueta para el sistema operativo
tk.Label(root, text="Sistema Operativo:").grid(row=0, column=0, sticky=tk.W)

# Radiobutton para Linux
seleccion_sistema = tk.StringVar(value="linux")  # Valor predeterminado
tk.Radiobutton(root, text="Linux", variable=seleccion_sistema, value="linux", command=seleccionar_sistema_operativo).grid(row=0, column=1, padx=(0, 0), sticky=tk.W)

# Radiobutton para Windows
tk.Radiobutton(root, text="Windows", variable=seleccion_sistema, value="windows", command=seleccionar_sistema_operativo).grid(row=1, column=1, padx=(0, 0), sticky=tk.W)

# Entradas y etiquetas

tk.Label(root, text="Host:").grid(row=2, column=0, sticky=tk.W)
host_entry = tk.Entry(root)
host_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.EW)

tk.Label(root, text="Nombre de usuario:").grid(row=3, column=0, sticky=tk.W)
username_entry = tk.Entry(root)
username_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.EW)

use_private_key_var = tk.BooleanVar()
use_private_key_checkbutton = tk.Checkbutton(root, text="Usar clave privada", variable=use_private_key_var, command=mostrar_ocultar_llave_privada)
use_private_key_checkbutton.grid(row=4, columnspan=2, sticky=tk.W)

private_key_path_label = tk.Label(root, text="Ruta de la llave privada:")
private_key_path_entry = tk.Entry(root)

private_key_path_label.grid(row=5, column=0, sticky=tk.W)
private_key_path_label.grid_remove()
private_key_path_entry.grid(row=5, column=1, padx=5, pady=5, sticky=tk.EW)
private_key_path_entry.grid_remove()

tk.Label(root, text="Contraseña:").grid(row=6, column=0, sticky=tk.W)
password_entry = tk.Entry(root, show="*")
password_entry.grid(row=6, column=1, padx=5, pady=5, sticky=tk.EW)

tk.Label(root, text="Carpeta remota a escanear:").grid(row=7, column=0, sticky=tk.W)
carpeta_entry = tk.Entry(root)
carpeta_entry.grid(row=7, column=1, padx=5, pady=5, sticky=tk.EW)

tk.Label(root, text="¿Usar nombre de archivo Excel?").grid(row=8, columnspan=2, sticky=tk.W)
nombre_archivo_var = tk.BooleanVar()
nombre_archivo_checkbutton = tk.Checkbutton(root, text="Usar archivo Excel diferente a 'Bitacora.xlsx'. En ese caso, seleccione la casilla y agregue el nombre del archivo (debe estar en la misma carpeta que el programa)", variable=nombre_archivo_var, command=mostrar_ocultar_nombre_archivo_excel)
nombre_archivo_checkbutton.grid(row=8, columnspan=2, sticky=tk.W)

nombre_archivo_label = tk.Label(root, text="Nombre del archivo Excel:")
nombre_archivo_entry = tk.Entry(root)
nombre_archivo_label.grid(row=9, column=0, sticky=tk.W)
nombre_archivo_label.grid_remove()
nombre_archivo_entry.grid(row=9, column=1, padx=5, pady=5, sticky=tk.EW)
nombre_archivo_entry.grid_remove()

tk.Label(root, text="Nombre de la nueva hoja:").grid(row=10, column=0, sticky=tk.W)
nombre_hoja_entry = tk.Entry(root)
nombre_hoja_entry.grid(row=10, column=1, padx=5, pady=5, sticky=tk.EW)

# Botón para iniciar operación
tk.Button(root, text="Iniciar Operación", command=lambda: [actualizar_barra_progreso(), iniciar_operacion()]).grid(row=11, columnspan=2, pady=10)

# Modo oscuro
boton_modo_oscuro = tk.Button(root, text="Cambiar Modo Oscuro", command=cambiar_modo_oscuro)
boton_modo_oscuro.grid(row=0, column=3, padx=10, pady=10)

# Barra de progreso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.grid(row=12, column=0, columnspan=2, padx=10, pady=10)


root.mainloop()
