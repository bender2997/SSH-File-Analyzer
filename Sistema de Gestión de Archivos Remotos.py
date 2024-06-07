import os
import stat
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import paramiko
from smb.SMBConnection import SMBConnection

def listar_archivos_remotos_smb(smb_conn, carpeta):
    archivos = []
    try:
        file_list = smb_conn.listPath("SharedFolder", carpeta)
        for file in file_list:
            if file.isDirectory:
                archivos.extend(listar_archivos_remotos_smb(smb_conn, os.path.join(carpeta, file.filename)))
            else:
                nombre, extension = os.path.splitext(file.filename)
                archivos.append((nombre, extension, None, None, carpeta))
    except Exception as e:
        messagebox.showerror("Error", f"Error al listar archivos remotos a través de SMB: {e}")
    return archivos

def establecer_conexion_smb(host, username, password):
    try:
        smb_conn = SMBConnection(username, password, "", "local_machine_name", use_ntlm_v2=True)
        smb_conn.connect(host, 445)
        return smb_conn
    except Exception as e:
        messagebox.showerror("Error", f"Error al establecer la conexión SMB: {e}")
        return None

def iniciar_operacion():
    host = host_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    carpeta = carpeta_entry.get()
    nombre_archivo_excel = nombre_archivo_entry.get()
    nombre_hoja_nueva = nombre_hoja_entry.get()
    smb_conn = establecer_conexion_smb(host, username, password)
    if smb_conn:
        try:
            archivos = listar_archivos_remotos_smb(smb_conn, carpeta)
            smb_conn.close()
            guardar_en_excel(archivos, nombre_archivo_excel, nombre_hoja_nueva)
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar la operación SMB: {e}")
    else:
        messagebox.showerror("Error", "No se pudo establecer la conexión SMB.")

def limpiar_nombre(nombre):
    caracteres_no_validos = ['\\', '/', '*', '[', ']', ':', '?']
    for caracter in caracteres_no_validos:
        nombre = nombre.replace(caracter, '_')
    return nombre[:30]

def listar_archivos_remotos_linux(sftp_client, carpeta):
    archivos = []
    try:
        for entry in sftp_client.listdir_attr(carpeta):
            ruta_completa = carpeta + "/" + entry.filename
            if stat.S_ISDIR(entry.st_mode):
                archivos.extend(listar_archivos_remotos_linux(sftp_client, ruta_completa))
            else:
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

def listar_archivos_remotos_windows(sftp_client, carpeta):
    archivos = []
    try:
        for entry in sftp_client.listdir_attr(carpeta):
            ruta_completa = os.path.join(carpeta, entry.filename.replace('/', '\\'))
            if stat.S_ISDIR(entry.st_mode):
                archivos.extend(listar_archivos_remotos_windows(sftp_client, ruta_completa))
            else:
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
    nombre_hoja = limpiar_nombre(nombre_hoja)
    if nombre_hoja in libro_excel.sheetnames:
        messagebox.showinfo("Información", f"La hoja '{nombre_hoja}' ya existe en el archivo de Excel. Es necesario cambiar el nombre.")
        nuevo_nombre = input("Ingrese un nuevo nombre para la hoja: ")
        return agregar_hoja_excel(libro_excel, nuevo_nombre)
    else:
        libro_excel.create_sheet(title=nombre_hoja)
        return nombre_hoja

def ajustar_ancho_columnas(hoja):
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
    hoja.append(["ID", "Nombre de dato / archivo", "Formato", "Fecha de creación", "Fecha de modificación", "Ruta",
                 "Responsable del dato / archivo", "Propósito del dato / archivo",
                 "¿Quién tiene acceso al archivo / dato?", "Transferencia con 3ero / externos",
                 "Responsable de respaldo", "Fecha de respaldo", "Responsable de eliminación", "Fecha de eliminación"])
    for id_archivo, archivo in enumerate(archivos, start=1):
        hoja.append([id_archivo] + list(archivo))
    ajustar_ancho_columnas(hoja)
    try:
        libro_excel.save(nombre_archivo)
        messagebox.showinfo("Éxito", "Datos guardados exitosamente en el archivo de Excel.")
    except PermissionError as e:
        messagebox.showerror("Error", f"Error de permisos al guardar el archivo: {nombre_archivo} - {e}")

def establecer_conexion(host, username, use_private_key, private_key_path=None, password=None, port=22):
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        if use_private_key:
            private_key = paramiko.RSAKey.from_private_key_file(private_key_path)
            ssh_client.connect(hostname=host, username=username, pkey=private_key, port=port)
        else:
            ssh_client.connect(hostname=host, username=username, password=password, port=port)
        return ssh_client
    except paramiko.AuthenticationException:
        messagebox.showerror("Error", "Error de autenticación. Verifica tus credenciales.")
        return None
    except paramiko.SSHException as e:
        messagebox.showerror("Error", f"Error estableciendo la conexión SSH: {e}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado al establecer la conexión: {e}")
        return None

def iniciar_operacion():
    host = host_entry.get()
    username = username_entry.get()
    password = password_entry.get() if not use_private_key_var.get() else None
    private_key_path = private_key_path_entry.get() if use_private_key_var.get() else None
    port = int(port_entry.get())
    carpeta = carpeta_entry.get()
    nombre_archivo_excel = nombre_archivo_entry.get()
    nombre_hoja_nueva = nombre_hoja_entry.get()
    ssh_client = establecer_conexion(host, username, use_private_key_var.get(), private_key_path, password, port)
    if ssh_client:
        try:
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

def seleccionar_sistema_operativo():
    global seleccion_sistema
    sistema_operativo = seleccion_sistema.get()
    iniciar_programa(sistema_operativo)

def iniciar_programa(sistema_operativo):
    if sistema_operativo == "linux":
        usar_barra = "/"
        listar_archivos_remotos = listar_archivos_remotos_linux
    elif sistema_operativo == "windows":
        usar_barra = "\\"
        listar_archivos_remotos = listar_archivos_remotos_windows
    else:
        messagebox.showerror("Error", "Sistema operativo no soportado.")
        return
    iniciar_operacion()

root = tk.Tk()
root.title("Análisis de Carpeta Remota y Exportación a Excel")

style = ttk.Style(root)
style.theme_use('clam')

style.configure('TLabel', font=('Arial', 12), padding=10)
style.configure('TButton', font=('Arial', 12, 'bold'), padding=10)
style.configure('TEntry', font=('Arial', 12), padding=10)
style.configure('TCheckbutton', font=('Arial', 12), padding=10)
style.configure('TRadiobutton', font=('Arial', 12), padding=10)

notebook = ttk.Notebook(root)

frame_conexion = ttk.Frame(notebook)
notebook.add(frame_conexion, text="Conexión")

frame_configuracion = ttk.Frame(notebook)
notebook.add(frame_configuracion, text="Configuración")

notebook.pack(expand=True, fill="both", padx=20, pady=20)

# --- Pestaña de Conexión ---

ttk.Label(frame_conexion, text="Host:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
host_entry = ttk.Entry(frame_conexion)
host_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame_conexion, text="Username:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
username_entry = ttk.Entry(frame_conexion)
username_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

use_private_key_var = tk.BooleanVar()
use_private_key_checkbutton = ttk.Checkbutton(frame_conexion, text="Usar clave privada", variable=use_private_key_var)
use_private_key_checkbutton.grid(row=2, columnspan=2, padx=5, pady=5)

ttk.Label(frame_conexion, text="Ruta de la clave privada:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
private_key_path_entry = ttk.Entry(frame_conexion)
private_key_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame_conexion, text="Password:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
password_entry = ttk.Entry(frame_conexion, show="*")
password_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame_conexion, text="Port:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
port_entry = ttk.Entry(frame_conexion)
port_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
port_entry.insert(0, "22")

frame_conexion.columnconfigure(1, weight=1)

# --- Pestaña de Configuración ---

ttk.Label(frame_configuracion, text="Sistema operativo:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
seleccion_sistema = tk.StringVar()
seleccion_sistema.set("linux")
ttk.Radiobutton(frame_configuracion, text="Linux", variable=seleccion_sistema, value="linux").grid(row=0, column=1, padx=5, pady=5, sticky="w")
ttk.Radiobutton(frame_configuracion, text="Windows", variable=seleccion_sistema, value="windows").grid(row=0, column=2, padx=5, pady=5, sticky="w")

ttk.Label(frame_configuracion, text="Carpeta remota:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
carpeta_entry = ttk.Entry(frame_configuracion)
carpeta_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame_configuracion, text="Nombre del archivo Excel:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
nombre_archivo_entry = ttk.Entry(frame_configuracion)
nombre_archivo_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame_configuracion, text="Nombre de la nueva hoja:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
nombre_hoja_entry = ttk.Entry(frame_configuracion)
nombre_hoja_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

frame_configuracion.columnconfigure(1, weight=1)

iniciar_button = ttk.Button(root, text="Iniciar Operación", command=iniciar_operacion)
iniciar_button.pack(pady=20)

root.update_idletasks()
root.geometry(f'{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}')
root.minsize(root.winfo_width(), root.winfo_height())

root.mainloop()
