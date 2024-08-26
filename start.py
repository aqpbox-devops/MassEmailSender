import scripts.email_sender as ems
import scripts.input_formater as ifo
import pandas as pd
import argparse

from PIL import Image

from enum import Enum, auto

class COLS(Enum):
    EMAIL = auto()
    CC = auto()
    LNNS = auto()
    LCAT = auto()
    NCAT = auto()
    NSAL = auto()
"""
ASUNTO: 
CORREO - Cc , APELLIDOS&NOMBRES - CAMPOS DE REEMPLAZO...
"""

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Email mass sender')
    parser.add_argument('input', type=str, help='Path to the excel file')
    parser.add_argument('template', type=str, help='Path to the template image')
    parser.add_argument('htmlf', type=str, help='Path to the html injected to the message')

    args = parser.parse_args()
    sender = ems.EmailSender('compensaciones@cajaarequipa.pe',
                             'Caja2015', 'webmail.cajaarequipa.pe')
    in_data = pd.read_excel(args.input, header=None, engine='openpyxl')
    in_data.columns = [COLS.EMAIL, COLS.CC, COLS.LNNS, COLS.LCAT, COLS.NCAT, COLS.NSAL]
    image = Image.open(args.template)

    subject = in_data.iloc[0, 0]
    data_table = in_data.iloc[2:]

    meta = [
        ['', (139, 293), 15, 'Calibri', '#18243d', {'weight': 'bold', 'style': 'italic'}],
        ['', (458, 514), 12, 'Calibri', '#ffffff', {'weight': 'normal', 'style': 'normal'}],
        ['', (438, 548), 12, 'Calibri', '#ffffff', {'weight': 'normal', 'style': 'normal'}],
        ['', (408, 582), 12, 'Calibri', '#ffffff', {'weight': 'normal', 'style': 'normal'}]
    ]

    with open(args.htmlf, 'r', encoding='utf-8') as file:
        html_injected = file.read()

    print(f"Subject: [{subject}], #mails: {data_table.shape[0]}")
    print(data_table.head())

    for index, row in data_table.iterrows():
        sender.push_message([row[COLS.EMAIL], row[COLS.CC]], subject)
        meta[0][0] = row[COLS.LNNS]
        meta[1][0] = row[COLS.LCAT]
        meta[2][0] = row[COLS.NCAT]
        meta[3][0] = row[COLS.NSAL]
        img = ifo.add_text_to_image(image, meta)
        sender.attach_image(img, html_injected)
        #imgpil = Image.open(img)
        #imgpil.convert("RGB").save('input/test.jpg', "JPEG")

    sender.send()