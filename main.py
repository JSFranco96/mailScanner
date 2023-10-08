import win32com.client
import re
import pandas as pd

def verifica_estructura(texto):
    elementos = texto.split(";")
    return len(elementos) >= 3

def read_outlook_emails(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print('primero')
    folder = outlook.Folders["javierfranco0904@hotmail.com"]

    print(folder)
    print('************************')
    # Obtener la carpeta deseada
    folder_path_parts = folder_path.split(";")
    for folder_name in folder_path_parts:
        folder = folder.Folders[folder_name]
        print(folder)

    # Leer los correos electrónicos en la carpeta
    emails = folder.Items
    data = []
    for email in emails:
        # Filtrar correos
        subject = email.Subject
        do_match = verifica_estructura(subject)
        
        if do_match:
            # Extraer el código y nombre de la empresa del asunto
            asunto_parts = email.Subject.split(";")
            codigo = asunto_parts[2]
            nombre_empresa = asunto_parts[1]
            data.append({
                'Fecha recibido':  email.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') ,
                'Empresa': nombre_empresa,
                'Factura': codigo
            })

    df = pd.DataFrame(data)
    # df['Fecha recibido'] = df['Fecha recibido'].dt.strftime('%Y-%m-%d %H:%M:%S')

    df.to_excel('correosEDM.xlsx', index=False)

if __name__ == "__main__":
    folder_path = "Bandeja de entrada;@2022;@Facturas;11 - Noviembre;Registro atrazado"
    read_outlook_emails(folder_path)
