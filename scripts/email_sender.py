import os
from typing import List
from collections import deque

import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

def get_name(path: str) -> str:
    return os.path.basename(path)

class EmailSender:
    def __init__(self, email, password, smtp, smtp_port=587):
        self.sender_email = email
        self.sender_password = password
        self.smtp_server = smtp
        self.smtp_port = smtp_port
        self.msg_queue = deque()

    def push_message(self, subscribers: List[str], subject: str):
        message = {
            'Subs':subscribers, 
            'Payload': MIMEMultipart()
        }
        message['Payload']['Subject'] = subject
        self.msg_queue.append(message)

    def attach_text(self, txt: str):
        if self.msg_queue:
            txt_data = MIMEText(txt, "plain")
            self.msg_queue[0]['Payload'].attach(txt_data)

    def attach_image(self, image, html_injected: str=''):
        if self.msg_queue:
            img_data = MIMEImage(image.read())
            img_data.add_header('Content-ID', '<displayed_img>')
            img_data.add_header('Content-Disposition', 'inline', filename='Advice.png')
            self.msg_queue[0]['Payload'].attach(img_data)
            body = f"""
            <html>
                <body>
                    {html_injected}
                    <table style="width: 100%; text-align: center;">
                        <tr>
                            <td>
                                <img src="cid:displayed_img" alt="Imagen" style="height: auto;">
                            </td>
                        </tr>
                    </table>
                </body>
            </html>
            """
            self.msg_queue[0]['Payload'].attach(MIMEText(body, 'html'))
            
    def attach_file(self, filename: str):
        if self.msg_queue:
            with open(filename, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {get_name(filename)}")
                self.msg_queue[0]['Payload'].attach(part)

    def send(self):
        while self.msg_queue:
            msg = self.msg_queue.popleft()
            msg['Payload']['To'] = msg['Subs'][0]
            msg['Payload']['From'] = self.sender_email
            if len(msg['Subs']) > 1:
                msg['Payload']['Cc'] = ', '.join(msg['Subs'][1:])

            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.sendmail(self.sender_email, msg['Subs'], msg['Payload'].as_string())