import PySimpleGUI as sg
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import traceback

layout = [
    [sg.Text("Selecione o arquivo TXT com endereços de e-mail:")],
    [sg.InputText(key='-TXT_FILE-'), sg.FileBrowse()],
    [sg.Text("Configurações do Remetente:")],
    [sg.Text("Endereço de E-mail:"), sg.InputText(key='-FROM_EMAIL-')],
    [sg.Text("Senha:"), sg.InputText(key='-EMAIL_PASSWORD-', password_char='*')],
    [sg.Text("Assunto:"), sg.InputText(key='-EMAIL_SUBJECT-')],
    [sg.Text("Corpo do E-mail:"), sg.Multiline(key='-EMAIL_BODY-', size=(40, 10))],
    [sg.Text("Anexo:"), sg.InputText(key='-ATTACHMENT_PATH-'), sg.FileBrowse()],
    [sg.Button("Enviar E-mails"), sg.Button("Sair")]
]

window = sg.Window("Enviar E-mails em Massa", layout)

def send_email(from_email, password, recipient, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    print(from_email)
    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            att = MIMEApplication(attachment.read(), _subtype="pdf")
            att.add_header("Content-Disposition", "attachment", filename=attachment_path)
            msg.attach(att)
            print(attachment_path)
    try:
        server = smtplib.SMTP('smtp.hostinger.com', 587)
    except smtplib.SMTPConnectError as e:
        print(f"Erro de conexão com o servidor SMTP: {e}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Erro de autenticação no servidor SMTP: {e}")
    except Exception as e:
        print(f"Erro ao criar conexão SMTP: {e}")
            
    server.starttls()
    print('server starttls 1')    
    server.login(from_email, password)
    server.sendmail(from_email, recipient, msg.as_string())
    server.quit()

while True:
    event, values = window.read()
    try:
        if event == sg.WINDOW_CLOSED or event == "Sair":
            break
        elif event == "Enviar E-mails":
            txt_file_path = values['-TXT_FILE-']
            from_email = values['-FROM_EMAIL-']
            email_password = values['-EMAIL_PASSWORD-']
            email_subject = values['-EMAIL_SUBJECT-']
            email_body = values['-EMAIL_BODY-']
            attachment_path = values['-ATTACHMENT_PATH-']
            
            with open(txt_file_path, 'r') as file:
                email_addresses = file.readlines()
                print(email_addresses)
            for email_address in email_addresses:
                print(email_address)
                try:
                    send_email(from_email, email_password, email_address.strip(), email_subject, email_body, attachment_path)
                except Exception as e:
                    sg.popup_error(f"Erro ao enviar e-mail para {email_address.strip()}:\n{str(e)}")
                    traceback.print_exc()  # Imprime a stack trace no console

            sg.popup("E-mails enviados com sucesso!")
        
    except Exception as e:
        sg.popup_error(f"Erro geral:\n{str(e)}")
        traceback.print_exc()  # Imprime a stack trace no console

window.close()
