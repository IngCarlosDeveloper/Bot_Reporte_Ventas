import os
import win32com.client
import subprocess
import fnmatch


# Configuración de la carpeta para guardar archivos adjuntos

ruta_base = os.path.dirname(os.path.abspath(__file__))

output_folder = os.path.join(ruta_base, "Data")
os.makedirs(output_folder, exist_ok=True)

# Conectar con Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Seleccionar la bandeja de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 representa la bandeja de entrada
processed_folder_name = "RVM_Completados"  # Cambia esto por el nombre de tu carpeta
processed_folder = inbox.Folders(processed_folder_name)

# Iterar a través de los correos electrónicos
messages = inbox.Items
for message in messages:
    try:
        subject = message.Subject  # Título del correo
        sender = message.SenderName  # Remitente
        print(f"Título: {subject} | Remitente: {sender}")

        if fnmatch.fnmatch(subject, "*VENTAS*"):

            # Descargar y renombrar archivos adjuntos
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    # Crear un nombre de archivo basado en el asunto del correo
                    filename = f"{subject} - {attachment.FileName}".replace(":", "-").replace("/", "-")
                    #filename = attachment.FileName.replace(":", "-").replace("/", "-")
                    #filename = f"{attachment.FileName} - {subject}".replace(":", "-").replace("/", "-")
                    file_path = os.path.join(output_folder, filename)

                    # Evitar sobrescritura
                    counter = 1
                    original_path = file_path
                    while os.path.exists(file_path):
                        base, ext = os.path.splitext(original_path)
                        file_path = f"{base} ({counter}){ext}"
                        counter += 1

                    # Guardar el archivo adjunto
                    attachment.SaveAsFile(file_path)
                    print(f"Archivo descargado: {file_path}")

            # Mover el correo procesado a la carpeta "Completados"
            message.Move(processed_folder)
            print(f"Correo movido a la carpeta '{processed_folder_name}'")

        else:
            print("No se encontraron reportes de ventas mensuales.")

    except Exception as e:
        print(f"Error procesando un correo: {e}")
            

for folder in inbox.Folders:
    print(folder.Name)

next = os.path.join(ruta_base, "2-generador.py")

print("1-lector.py ejecutador exitosamente\n\n")

subprocess.run(["python", next])