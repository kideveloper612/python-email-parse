import imaplib
import email
from email.header import decode_header
import re
import os
import csv
import dotenv

dotenv.load_dotenv()
username = os.environ.get('EMAIL')
password = os.environ.get('PASSWORD')
From = os.environ.get('FROM')


def write(lines):
    with open(file='result.csv', encoding='utf-8', mode='a', newline='') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerows(lines)


def main():
    global From
    if not os.path.isfile(save_name):
        header = [['Subject', 'From', 'Date', 'Name', 'Phone', 'DOB', 'Office', 'Note']]
        write(header)
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)

    imap.select("INBOX")

    sub_status, messages = imap.search(None, 'FROM {}'.format(From))
    if sub_status == 'OK':
        for message in messages[0].split():
            res, msg = imap.fetch(message, "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])
                    subject = decode_header(msg["Subject"])[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode()
                    From, encoding = decode_header(msg.get("From"))[0]
                    if isinstance(From, bytes):
                        From = From.decode(encoding)
                    if msg.is_multipart():
                        for part in msg.walk():
                            payload = part.get_payload(decode=True)
                            if payload is not None:
                                text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                                date = re.search('=================(.*)======================================', text).group(1).strip()
                                name = re.search('Name:(.*)Ph#:', text).group(1).strip()
                                phone = re.search('Ph#:(.*)DOB:', text).group(1).strip()
                                birthday = re.search('DOB:(.*)Which office were you seen at', text).group(1).strip()
                                office = re.search('Which office were you seen at(.*)Msg:', text).group(1)[2:].strip()
                                note = re.search('Msg(.*)CID:', text).group(1)[1:].strip()
                                line = [subject, msg['from'], date, name, phone, birthday, office, note]
                                print(line)
                                write([line])
                                break
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload is not None:
                            text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                            date = re.search('=================(.*)======================================', text).group(1)
                            name = re.search('Name:(.*)Ph#:', text).group(1).strip()
                            phone = re.search('Ph#:(.*)DOB:', text).group(1).strip()
                            birthday = re.search('DOB:(.*)Which office were you seen at', text).group(1).strip()
                            office = re.search('Which office were you seen at(.*)Msg:', text).group(1)[2:].strip()
                            note = re.search('Msg(.*)CID:', text).group(1)[1:].strip()
                            line = [subject, msg['from'], date, name, phone, birthday, office, note]
                            print(line)
                            write([line])

    imap.close()
    imap.logout()


if __name__ == '__main__':
    save_name = 'result.csv'
    main()
