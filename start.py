import scripts.input_formater as ifo
import pandas as pd
import argparse, os, glob, re
import configparser

from PIL import Image

from enum import Enum, auto
from exchangelib import Mailbox, Account, HTMLBody, Message, Credentials, FileAttachment

def check_file(auth_file_path):
    if not os.path.isfile(auth_file_path):
        directory = os.path.dirname(auth_file_path)
        new_files = glob.glob(os.path.join(directory, '*.xlsx'))
        if new_files:
            auth_file_path = new_files[0]
    return auth_file_path

def print_message(message):
    print('='*80)
    print(f"To: {message.to_recipients}, Subject: {message.subject}")
    print("Body:", str(message.body)[:500])
    print('='*80)

class COLS(Enum):
    NAME = auto()
    EMAIL = auto()
    SALARY = auto()
    CC = auto()

keep = ['de', 'para', 'la', 'el', 'en']

def keep_cap(string, it):
    '''
    Returns a generator by tokenizing a string and checking each word before capitalizing
    '''
    string_tokens = string.lower().split()
    for i in string_tokens:
        if i in it:
            yield i
        else:
            yield i.capitalize()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Email mass sender')
    parser.add_argument('auth', type=str, help='Path to the configuration file')

    args = parser.parse_args()

    config = configparser.ConfigParser()
    config.read(args.auth)

    config = config['ALL']

    from_email = config['EMAIL']
    from_password = config['PASSWORD']
    
    credentials = Credentials(from_email, from_password)
    account = Account(from_email, credentials=credentials, autodiscover=True)
    
    fn_input = check_file(config['INPUT_PATH'])

    try:
        data_table = pd.read_excel(fn_input, engine='openpyxl')
        if os.path.isfile(fn_input):
            os.remove(fn_input)

    except FileNotFoundError as e:
        print("Mass email sender: Can not find file input")
        exit(0)

    data_table.columns = [COLS.NAME, COLS.EMAIL, COLS.SALARY, COLS.CC]
    image = Image.open('template/bonus.png')

    subject = re.sub(r'[^ -~¡¿]+', '', config['SUBJECT'])

    meta = [
        ['', (334,360), 19, 'Calibri', '#001c56', {'weight': 'bold', 'style': 'italic'}],
        ['', (702,665), 14, 'Calibri', '#18243d', {'weight': 'bold', 'style': 'italic'}]
    ]

    print(f"Subject: [{subject}], #mails: {data_table.shape[0]}")
    print(data_table.info())

    data_table[COLS.NAME] = data_table[COLS.NAME].apply(lambda x: str(x).title())
    data_table[COLS.SALARY] = data_table[COLS.SALARY].apply(lambda x: f"{float(x):,.2f}")

    for index, row in data_table.iterrows():
        meta[0][0] = row[COLS.NAME]
        meta[1][0] = row[COLS.SALARY]
        img_buffer = ifo.add_text_to_image(image, meta)
        """
        img = Image.open(img_buffer)

        img.save("imagen.png", "PNG")

        continue"""

        attachment = FileAttachment(
            name='Advice.png',
            content=img_buffer.getvalue(),
            content_type='image/png'
        )
        
        to_recipients = [Mailbox(email_address=row[COLS.EMAIL])]
        cc_recipients = [Mailbox(email_address=row[COLS.CC])]

        message = Message(
            account=account,
            subject=subject,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            body=HTMLBody(f"""
                <html>
                <body>
                    <div>
                        <img src="cid:Advice.png" alt="AdviceImg" style="max-width: 100%; height: auto;">
                    </div>
                </body>
                </html>
            """),
            attachments=[attachment]
        )

        print_message(message)

        message.clean()

        #message.send_and_save()