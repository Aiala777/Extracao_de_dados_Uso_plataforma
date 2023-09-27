import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email import encoders
from email.mime.image import MIMEImage


def envio():
    # email details
    sender = '.com.br'
    recipient = '.com.br'
    subject = 'Subject line of the email'
    body = """Olá, Segue em anexo a quantidade de chats do dia de hoje no (WhatsApp Oficial)!

    <br>  <br> 

    <span style="font-size: 15px;">Atenciosamente,</span> <br>
    <span style="font-size: 15px;">BB-8</span> <br>
    <span style="font-size: 12pt;">Robô auxiliar de correio eletrônico</span>
    <br>
    <img src="cid:image1">
    """

    image_path = 'C:/Users/arthuraiala/Desktop/uso_gup/assinatura.png'
    # attachment details
    attachment_path = r"C:/Users/arthuraiala/Desktop/uso_gup/Quantidade_de_chats.txt"
    attachment_name = 'Quantidade_de_chats.txt'

    # create a MIME message object
    msg = MIMEMultipart()
    msg['From'] = '.com.br'
    msg['To'] = '.com.br'
    msg['Subject'] = 'Chats diários WhatsApp Oficial'
    msg.attach(MIMEText(body, 'html'))
    

    # attach the file to the message
    with open(attachment_path, 'rb') as f:
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename=attachment_name)
        msg.attach(attachment)
    
    with open(image_path, 'rb') as f:
        image = MIMEImage(f.read())
        image.add_header('Content-ID', '<image1>')
        msg.attach(image)

    # send the email
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587
    smtp_username = 'com.br'
    smtp_password = ''
    smtp_tls = True

    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
        smtp.ehlo()
        if smtp_tls:
            smtp.starttls()
            smtp.ehlo()
        smtp.login(smtp_username, smtp_password)
        smtp.sendmail(sender, COMMASPACE.join([recipient]), msg.as_string())
        print('e-mail enviado!')


def emailInsusesso():
    sender_email = ".com.br"
    receiver_email = ".com.br"
    password = "AAaa22**"
    message = "Subject: Erro BB-8\n\nEncontrei uma excecao no envio da quantidade de chats no whatsapp oficial, F no chat :("

    server = smtplib.SMTP("smtp-mail.outlook.com", 587)
    server.starttls()
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message)
    server.quit()

