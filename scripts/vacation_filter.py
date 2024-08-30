import exchangelib as exlib
from dotenv import load_dotenv
import os

def clean_lna(text):
    ls = str(text).splitlines()
    filtered = [l for l in ls if l.strip() != '']
    return '[&&]'.join(filtered)
    
load_dotenv()

email = os.getenv('CAJAAQP_EMAIL')
password = os.getenv('CAJAAQP_PASSWORD')
email_dir = os.getenv('EMAIL_DIRECTORY_READ')

print(email, '*'*len(password), email_dir, sep='|')

credentials = exlib.Credentials(email, password)
account = exlib.Account(email, credentials=credentials, autodiscover=True)

inbox = account.inbox
folder = inbox / email_dir

for item in folder.all():#filter(is_read=False):
    if isinstance(item, exlib.Message):
        print(f"[ID:{item.conversation_id}]")
        print(f"\tAsunto: {item.subject}")
        print(f"\tDe: {item.sender.email_address}")
        print(f"\tFecha de recibo: {item.datetime_received}")
       # print(f"\tXML: {item.to_xml}")
        print(f"\tCuerpo: {clean_lna(item.text_body)[:100]}")
        print('*'*200)
        item.is_read = True