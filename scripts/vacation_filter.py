import exchangelib as exlib
import pandas as pd
from dotenv import load_dotenv
from constants import *

import os
import io

def clean_lna(text):
    ls = str(text).splitlines()
    filtered = [l for l in ls if l.strip() != '']
    return '[&&]'.join(filtered)
    
if __name__ == '__main__':
    load_dotenv()

    email = os.getenv('CAJAAQP_EMAIL')
    password = os.getenv('CAJAAQP_PASSWORD')
    vacati_dir = os.getenv('VACATIONS_DIR')
    dataub_dir = os.getenv('DATAUB_DIR')

    credentials = exlib.Credentials(email, password)
    account = exlib.Account(email, credentials=credentials, autodiscover=True)

    inbox = account.inbox

    dataub_df = None
    for item in (inbox / dataub_dir).all().order_by('-datetime_received')[:1]:
        if isinstance(item, exlib.Message):
            for attachment in item.attachments:
                if attachment.name.endswith('.xlsx') and attachment.name.startswith("Data ubicación"):
                    with io.BytesIO(attachment.content) as file:
                        dataub_df = pd.read_excel(file)

    recv_emails = {}

    for item in (inbox / vacati_dir).all().order_by('-datetime_received'):
        if isinstance(item, exlib.Message):
            recv_emails[str(item.id)] = {
                'Subject': item.subject,
                'From': item.sender.email_address,
                'Recv. Date': item.datetime_received
            }
            item.is_read = True

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    print(recv_emails)

    filtered_df = dataub_df[dataub_df[CKEY_EMAIL].isin({email_data['From'] for email_data in recv_emails.values()})]
    #dataub_df[CKEY_USERNAME] = dataub_df[CKEY_EMAIL].apply(lambda x: str(x).split('@')[0].lower())
    print(filtered_df[[CKEY_EMAIL, CKEY_EMPID]].head(15))