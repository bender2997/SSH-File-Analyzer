import os
import stat
import paramiko
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

def listar_archivos_remotos_linux(sftp_client, carpeta):
    """
    Lista los archivos en una carpeta remota en un sistema Linux a través de una conexión SFTP.
    
    :param sftp_client: Cliente SFTP activo para la conexión SSH.
    :param carpeta: Ruta de la carpeta remota a listar.
    :return: Lista de tuplas con información de archivos (nombre, extensión, fecha de modificación, fecha de acceso, ruta del padre).
    """
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
        print(f"No se encontró la carpeta remota: {carpeta} - {e}")
    except PermissionError as e:
        print(f"Error de permisos con la carpeta: {carpeta} - {e}")
    except Exception as e:
        print(f"Error inesperado al listar archivos remotos: {e}")
    return archivos

def listar_archivos_remotos_windows(sftp_client, carpeta):
    """
    Lista los archivos en una carpeta remota en un sistema Windows a través de una conexión SFTP.
    
    :param sftp_client: Cliente SFTP activo para la conexión SSH.
    :param carpeta: Ruta de la carpeta remota a listar.
    :return: Lista de tuplas con información de archivos (nombre, extensión, fecha de modificación, fecha de acceso, ruta del padre).
    """
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
        print(f"No se encontró la carpeta remota: {carpeta} - {e}")
    except PermissionError as e:
        print(f"Error de permisos con la carpeta: {carpeta} - {e}")
    except Exception as e:
        print(f"Error inesperado al listar archivos remotos: {e}")
    return archivos

def agregar_hoja_excel(libro_excel, nombre_hoja):
    """
    Agrega una nueva hoja a un libro de Excel si el nombre de la hoja no está ya en uso.
    Si el nombre ya existe, solicita al usuario un nuevo nombre.

    :param libro_excel: Objeto Workbook de openpyxl.
    :param nombre_hoja: Nombre deseado para la nueva hoja.
    :return: Nombre de la hoja que fue creada.
    """
    nombre_hoja = limpiar_nombre(nombre_hoja)
    if nombre_hoja in libro_excel.sheetnames:
        print(f"La hoja '{nombre_hoja}' ya existe en el archivo de Excel. Es necesario cambiar el nombre.")
        nuevo_nombre = input("Ingrese un nuevo nombre para la hoja: ")
        return agregar_hoja_excel(libro_excel, nuevo_nombre)
    else:
        libro_excel.create_sheet(title=nombre_hoja)
        return nombre_hoja

def ajustar_ancho_columnas(hoja):
    """
    Ajusta el ancho de las columnas de una hoja de Excel para que el contenido se ajuste automáticamente.

    :param hoja: Objeto Worksheet de openpyxl.
    """
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
    """
    Guarda la lista de archivos en un archivo de Excel. Crea un nuevo archivo si no existe.
    También crea una nueva hoja si se proporciona un nombre para ella.

    :param archivos: Lista de tuplas con información de archivos.
    :param nombre_archivo: Nombre del archivo Excel donde se guardarán los datos.
    :param nombre_hoja_nueva: Nombre de la hoja donde se guardarán los datos (opcional).
    """
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
        print("Datos guardados exitosamente en el archivo de Excel.")
    except PermissionError as e:
        print(f"Error de permisos al guardar el archivo: {nombre_archivo} - {e}")

def limpiar_nombre(nombre):
    """
    Limpia el nombre de una hoja de Excel eliminando caracteres no válidos.

    :param nombre: Nombre original de la hoja.
    :return: Nombre de la hoja limpio y truncado a 30 caracteres.
    """
    caracteres_no_validos = ['\\', '/', '*', '[', ']', ':', '?']
    for caracter in caracteres_no_validos:
        nombre = nombre.replace(caracter, '_')
    return nombre[:30]

def establecer_conexion(host, username, use_private_key, private_key_path=None, password=None, port=22):
    """
    Establece una conexión SSH al servidor especificado utilizando autenticación con clave privada o contraseña.

    :param host: Dirección del servidor SSH.
    :param username: Nombre de usuario para la autenticación.
    :param use_private_key: Booleano que indica si se debe usar una clave privada para la autenticación.
    :param private_key_path: Ruta al archivo de clave privada (si se usa autenticación con clave privada).
    :param password: Contraseña para la autenticación (si no se usa clave privada).
    :param port: Puerto del servidor SSH (por defecto es 22).
    :return: Cliente SSH conectado o None si no se pudo establecer la conexión.
    """
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        if use_private_key:
            print(f"Intentando cargar la clave privada desde: {private_key_path}")
            private_key = paramiko.RSAKey.from_private_key_file(private_key_path)
            ssh_client.connect(hostname=host, username=username, pkey=private_key, port=port)
        else:
            ssh_client.connect(hostname=host, username=username, password=password, port=port)
        return ssh_client
    except paramiko.AuthenticationException:
        print("Error de autenticación. Verifica tus credenciales.")
        return None
    except paramiko.SSHException as e:
        print(f"Error estableciendo la conexión SSH: {e}")
        return None
    except Exception as e:
        print(f"Error inesperado al establecer la conexión: {e}")
        return None

def iniciar_operacion():
    """
    Función principal que inicia la operación de conexión SSH, listado de archivos y guardado en Excel.
    """
    host = "10.10.163.163"
    username = "prueba"
    use_private_key = 's'
    private_key_path = r"C:/Users/LENOVO/Desktop/ANALIZADOR/bkusr1_openssh.pub"
    password = "prueba"
    port = 22
    carpeta = "/home/prueba/Prueba_SSH"
    nombre_archivo_excel = input("Ingrese el nombre del archivo Excel (deje vacío para 'Bitacora.xlsx'): ") or "Bitacora.xlsx"
    nombre_hoja_nueva = input("Ingrese el nombre de la nueva hoja: ")
    sistema_operativo = input("Ingrese el sistema operativo (linux/windows): ").strip().lower()
    
    ssh_client = establecer_conexion(host, username, use_private_key, private_key_path, password, port)
    if ssh_client:
        try:
            if sistema_operativo == "linux":
                archivos = listar_archivos_remotos_linux(ssh_client.open_sftp(), carpeta)
            elif sistema_operativo == "windows":
                archivos = listar_archivos_remotos_windows(ssh_client.open_sftp(), carpeta)
            else:
                print("Sistema operativo no soportado.")
                return
            ssh_client.close()
            guardar_en_excel(archivos, nombre_archivo_excel, nombre_hoja_nueva)
        except Exception as e:
            print(f"Error al procesar la operación: {e}")
    else:
        
        print("No se pudo establecer la conexión SSH.")

if __name__ == "__main__":
    iniciar_operacion()
