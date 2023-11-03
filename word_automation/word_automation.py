import os, sys
from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert
import pdb
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from os.path import join, dirname
from dotenv import load_dotenv
import threading

class EmailSender:

    def __init__(self):
        self.from_email = os.environ.get("SMTP_HOST_USER")
        self.password = os.environ.get("SMTP_HOST_PASSWORD")
        self.smtp_port = os.environ.get("SMTP_PORT")  # Standard secure SMTP port
        self.smtp_server = os.environ.get("SMTP_HOST")  # Google SMTP Server
        self.file_path = join(dirname(__file__), f'pdf/') # Path to the files being attached


    def send_email(self, to, subject, body, filename):
        thread = threading.Thread(target=self._send_email_thread, args=(to, subject, body, filename))
        thread.start()

    def _send_email_thread(self, to, subject, body, filename):
        print(self.from_email)
        # make a MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = "faoe2003@gmail.com"
        msg['Subject'] = subject

        # Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))

        # Define the file to attach
        file = self.file_path + filename
    

        # Open the file in python as a binary
        attachment= open(file, 'rb')  # r for read and b for binary

        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload((attachment).read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', "attachment; filename= " + filename)
        msg.attach(attachment_package)

        # Cast as string
        text = msg.as_string()

        # Connect with the server
        print("Connecting to server...")
        TIE_server = smtplib.SMTP(self.smtp_server, self.smtp_port)
        TIE_server.starttls()
        TIE_server.login(self.from_email, self.password)
        print("Succesfully connected to server")
        print()


        # Send emails to "person" as list is iterated
        print(f"Sending email to: {to}...")
        TIE_server.sendmail(self.from_email, to, text)
        print(f"Email sent to: {to}")
        print()

        # Close the port
        TIE_server.quit()

class InvitationSender:
    def __init__(self, template_name, excel_file, sheet_name, municipalities, event_settings, email_settings):
        self.template_name = template_name
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.municipalities = municipalities

        # event settings
        self.settings = event_settings

        # email settings
        self.email_settings = email_settings

    def read_excel(self):
        df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
        df_filtered = df[df['MUNICIPIO'].isin(self.municipalities)]
        result = df_filtered[['NOMBRE_ESCUELA', 'MUNICIPIO', 'CORREO', 'CCT']]
        result = result.drop_duplicates(subset=['NOMBRE_ESCUELA'], keep='first')
        self.dict_result = result.to_dict(orient='records')

    def send_email(self, to, subject, body, file):
        email_sender = EmailSender()
        print(email_sender.send_email(to, subject, body, file))

    def send_invitations(self):

        context = self.settings 
        for item in self.dict_result:
            doc = DocxTemplate(self.template_name)
            context['school_name'] = item['NOMBRE_ESCUELA']
            doc.render(context)

            print(f"Saving new invitation instance on docx/invitación_{item['CCT']}.docx")
            # save as doc
            doc.save(f"docx/invitación_{item['CCT']}.docx")
            
            # Convert docx to pdf
            convert(f"docx/invitación_{item['CCT']}.docx", f"pdf/invitacion_{item['CCT']}.pdf")
            print(f"exported invitación_{item['CCT']}.pdf")

            print(f'Sending email to {item['CORREO']}')
            self.email_settings['body'] = f"""
            En el grupo educativo Phionira nos complacemos en extender una cordial invitación a la comunidad estudiantil de los últimos semestres en {item['NOMBRE_ESCUELA']} a nuestra serie de seminarios web “APRENDE MÁS EN MENOS TIEMPO. PREPÁRATE PARA TU EXAMEN UAEMéx”, mismo que se comenzará el próximo jueves 26 de octubre de 2023 en punto de las 3 de la tarde y continuará con una serie de pláticas distribuidas a lo largo del mes de noviembre
            Los asistentes podrán iniciar su registro a través del enlace: go.phionira.com/aprende-mas
            
            Adjuntamos el siguiente oficio, al tiempo que agradeceríamos que puedan confirmar su interés en recibir información sobre los futuros eventos de esta serie de seminarios web a través de la siguiente liga: https://forms.gle/WDh8EF88jCvBCGoSA
            Para cualquier confirmación, duda o comentario, quedamos a sus órdenes, esperando contar con la participación de la comunidad de esta prestigiosa institución.  

            Coordinación de Phionira
            55 8704 6358
            coordinacion@phionira.com
            """


            self.send_email("faoe2003@gmail.com", 
                            self.email_settings['subject'],
                            self.email_settings['body'],
                            f'invitacion_{item['CCT']}.pdf')



if __name__ == "__main__":
    os.chdir(sys.path[0])
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    # Modificaciones a partir de esta línea -----------!

    # Ajustes del evento > nombre / fecha del oficio / fecha del evento / hora
    event_settings = {
        'event_name' : 'APRENDE MÁS EN MENOS TIEMPO. PREPÁRATE PARA TU EXAMEN UAEMéx',
        'date' : '23 de octubre de 2023',
        'date_of_event' : 'jueves 26 de octubre',
        'hour' : '3:00 p.m',
        'anita': "grupo phionira"
    }
    # Ajustes del correo > asunto / cuerpo
    email_settings = {
        'subject': "Invitación al webinar: APRENDE MÁS EN MENOS TIEMPO. PREPÁRATE PARA TU EXAMEN UAEMéx",
    }

    municipalities = ['TOLUCA', 'ZINACANTEPEC', 'ALMOLOYA DE JUAREZ', 'AMANALCO', 'VALLE DE BRAVO', 'TIANGUISTENCO', 'TENANCINGO', 'VILLA VICTORIA', 'LERMA', 'SAN MATEO ATENCO', 'METEPEC']
    sender = InvitationSender('Template.docx', "word_automation.xlsm", "DIRECTORIO", municipalities, event_settings, email_settings)
    sender.read_excel()
    sender.send_invitations()