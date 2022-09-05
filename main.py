"""
1) pip install gspread oauth2client (lub tylko gspread)

"""
import gspread
from pprint import pprint
import os
from email.message import EmailMessage
import ssl # standard tech for keeping mail secure
import smtplib

# Importing libraries
import imaplib
import email
from email.header import decode_header
import webbrowser

def RefreshChanges():

    sa = gspread.service_account() #nie musze uzywac filename= bo wrzuciłem api do ~/.config/gspread/service_account.json
    sh = sa.open("label Lukas")

    ws = sh.worksheet("All Labels")
    # print('Rows: ', ws.row_count)
    print(len(ws.get_values()))
    # print('Cols: ', ws.col_count)

    # print(ws.acell('F5').value)
    # print(ws.cell(3,4).value)
    # print(ws.get('C2:F4'))

    #pprint(ws.get_all_records())
    #pprint(ws.get_values())
    #print(ws.get_all_values())

    #ws.update('F2', 'Pomoc - Wodzisław Śląski')
    #ws.update('O2:P4',[['666666666','Projekt Python'],['555555555','Projekt 2'],['444444444','Pr Nowy']])
    #ws.update('U2', '=LEWY(N2,3)', raw=False)
    #ws.delete_rows(17)

def NewOrders():
    email_sender = 'pepco.analyst@gmail.com'
    email_password = os.environ.get('EMAIL_PWD') # Win | Linux nano .bash_profile
    email_receiver = 'borqs11@gmail.com'

    subject = 'New Order'
    body = """
    New order has been send
    """
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject

    em.set_content(body)
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())

def ReadEmail():
    """
    More details such as attachment/sending mails to folder etc. find here: https://www.thepythoncode.com/article/reading-emails-in-python

    :return:
    """

    user = 'pepco.analyst@gmail.com'
    password = os.environ.get('EMAIL_PWD')
    imap_url = 'imap.gmail.com'
    imap = imaplib.IMAP4_SSL(imap_url)
    imap.login(user, password)
    imap.select("INBOX") # You can use the imap.list() method to see the available mailboxes
    key = 'FROM'
    value = 'borycki01@gmail.com'
    _, data = imap.search(None, key, value)  # Search for emails with specific key and value
    messages = data[0].split() # total number of emails

    for i in messages:
        res, msg = imap.fetch(i, "(RFC822)") # fetch the email message by ID
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1]) # parse a bytes email into a message object
                subject, encoding = decode_header(msg["Subject"])[0] # decode the email subject
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding) # if it's a bytes, decode to str
                _from, encoding = decode_header(msg.get("From"))[0] # decode email sender
                if isinstance(_from, bytes):
                    _from = _from.decode(encoding)
                print("Subject:", subject)
                print("From:", _from)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            print(body)
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                print("=" * 100)    # close the connection and logout

    imap.close()
    imap.logout()

ReadEmail()