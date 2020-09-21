import email
import imaplib
import smtplib
import datetime
import email.mime.multipart
# import config
import base64
import os
import re
from typing import Tuple, List
#from bs4 import BeautifulSoup
import mimetypes
from pathlib import Path
from datetime import datetime
import logging
import argparse
import getpass



parser = argparse.ArgumentParser()
parser.add_argument("sender")
args = parser.parse_args()

# Move to config.py file
imap_server = "imap-mail.outlook.com"
imap_port = 993
smtp_server = "smtp-mail.outlook.com"
smtp_port = 587

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter("%(asctime)s %(levelname)s:%(name)s:%(message)s")

p = Path.home()
dir_parts = ["Emails_Logs"]
log_path = p.joinpath(*dir_parts)
if not log_path.exists():
    os.makedirs(log_path)
assert log_path.is_dir() == True, f"{log_path} is not a folder"

file_handler = logging.FileHandler(log_path.joinpath("outlook.log"))
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stream_handler)

class Outlook():
    def __init__(self):
        pass

    def login(self, username: str = "", password: str = "", env_var: bool = True) -> None:
        """Gets credentials and attempts establish a connection with the IMAP server.

        Args:
            username (str, optional): The full email address, e.g. "name.surname@company.com". Defaults to "".
            password (str, optional): The password in plain-text, e.g. "notsosecurepassword123". Defaults to "".
            env_var (bool, optional): Whether to look for environment variable to get credentials or not. Defaults to True.
        """
        
        # Getting credentials, so far two possibilities, environment variables and direct user input.
        # Potentially adding a third one in config.

        if env_var:
            try:
                username = os.environ["OUTLOOK_USER"]
                password = os.environ['OUTLOOK_PW']
            except KeyError:
                ### TODO
                # Add instructions to set env variables maybe?
                # Mac/Linux export VAR=VALUE from terminal
                # Windows https://www.computerhope.com/issues/ch000549.htm
                if not username:
                    logger.error(f"Please set the environment variable OUTLOOK_PW")
                if not password:
                    logger.error(f"Please set the environment variable OUTLOOK_USER")
                return
        else:
            username = input(f"Please enter your username: ")
            password = getpass.getpass("Please enter your password: ") #~CG}FrV+Q%m[=~q/*7Ckcts}nF#~YH
        
        self.username = username
        self.password = password
        
        # Attempts to login 3 times.
        login_attempts = 0
        while True:
            try:
                self.imap = imaplib.IMAP4_SSL(imap_server, imap_port)
                response, data = self.imap.login(username, password)
                assert response == "OK", f"Login failed: {response}"
                logger.info(f"{data[0].decode()} Signed in as {self.username}")
                return
            except Exception as error:
                logger.exception(f"Sign in error: {type(error)}")
                login_attempts += 1
                if login_attempts < 3:
                    continue
                assert False, "Login Failed"
                
    def logout(self) -> str:
        self.imap.logout()
        # TODO
        # Maybe add some further checks that the session was closed correctly, but shoudln't be necessary
        return print(f"Successfully logged out")
    
    @staticmethod
    def parse_list_folders(data):
        names = []
        descriptions = []
        for i in data:
            payload = i.decode().split(' "/" ')
            name = payload[1].strip('"')
            description = payload[0]
            names.append(name)
            descriptions.append(description)
        return names, descriptions

    def list_folders(self,
                    has_children: bool = False,
                    get_status: bool = False,
                    verbose: bool = False
                    ) -> Tuple[List[str], List[str]]:
        response, data = self.imap.list()
        names, descriptions = Outlook.parse_list_folders(data)
        logger.debug(f"Response code: {response}")
        logger.debug(f"Payload: {data}")
        if has_children:
            # Returns a tuple (name, description)
            logger.debug(*zip(names, descriptions), sep="\n")
        elif get_status:
            for name in names:
                try:
                    logger.debug(self.imap.status(f'"{name}"', "(UNSEEN RECENT MESSAGES UIDNEXT)")[1][0].decode())
                except Exception as err:
                    logger.exception(name, err)
        elif verbose:
            # Returns an unpacked list of names
            print(*names, sep="\n")
        return names, descriptions
    
    def select_folder(self, folder: str):
        try:
            response, data = self.imap.select(f'"{folder}"')
            assert response == "OK", f"Selection failed, response: {response} - {data[0].decode()}"
            logger.info(f"Selected folder: {folder}, there are {int(data[0])} messages.")
        except Exception as error:
            logger.exception(f"{error}")
        return folder
    
    def get_status(self, folder: str):
        return print(self.imap.status(f'"{folder}"', "(UNSEEN RECENT MESSAGES UIDNEXT)")[1][0].decode())
    
    def search(self, sender = "", recipient = "", subject = "", flags = []):
        # Searching
        
        # Search parameters:
        # FROM, TO, SUBJECT
        # Message Flags, can be zero or more for each message.
        # DELETED, SEEN, ANSWERED, FLAGGED, DRAFT, RECENT
        from_ = f'FROM "{sender}"' if sender else ""
        to_ = f'TO "{recipient}"' if recipient else ""
        subject_ = f'SUBJECT "{subject}"' if subject else ""
        flags_ls = []
        valid_flags = ["DELETED", "SEEN", "UNSEEN", "ANSWERED", "FLAGGED", "DRAFT", "RECENT"]
        
        if flags:
            for f in flags:
                if f.upper() not in valid_flags:
                    logger.error(f"{f.upper()} is not a valid flag.\nValid flags are: {valid_flags}")
                    return
                flags_ls.append(f.upper())
                
        flags_ = " ".join([i for i in flags_ls])
        params = [from_, to_, subject_, flags_]
        query = " ".join([i for i in params if i])
        
        response, data = self.imap.search(None, query)
        assert response == "OK", f"Search failed, response: {response} - {data[0].decode()}"
        
        logger.info(f"Query: {query}")
        logger.info(f"Found {len(data[0].split())} messages matching the query.")
        return response, data
    
    def fetch(self, sender = "", recipient = "", subject = "", flags = [], verbose=True):
        # Searching
        
        # Search parameters:
        # FROM, TO, SUBJECT
        # Message Flags, can be zero or more for each message.
        # DELETED, SEEN, ANSWERED, FLAGGED, DRAFT, RECENT
        
        
        response, data = self.search(sender, recipient, subject, flags)
        assert response == "OK", f"Search failed, response: {response} - {data[0].decode()}"
        
        uids = data[0].split()
        logger.info(f"{len(uids)} emails found from '{sender}' in the selected folder.")
        
        saved_files = 0
        print(f"Saving files ...")
        for uid in uids:

            # Fetching
            response, data = self.imap.fetch(uid, '(RFC822)')
            assert response == "OK", f"Fetching failed, response: {response} - {data[0].decode()}"

            raw_email = data[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)
            to_ = email_message["To"]
            from_ = email_message["From"]
            subject_ = email_message["Subject"]
            date_ = email_message["Date"]
            # Converting string with date and time into datetime object
            date_format = "%a, %d %b %Y %H:%M:%S %z"
            dt = datetime.strptime(date_, date_format)
            date = datetime.strftime(dt, "%Y %m %d %a")
            counter = 0
            for part in email_message.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                filename = part.get_filename()
                content_type = part.get_content_type()
                #print(content_type)
                if not filename:
                    extension = mimetypes.guess_extension(content_type)
                    if not extension:
                        extension = ".bin"
                    if "text/plain" in content_type:
                        extension = ".txt"
                    elif "text/html" in content_type:
                        extension = ".html"
                    filename = f"msg_part_{counter}{extension}"
                counter += 1
                #print(filename)

                # Saving content
                p = Path.home()
                dir_parts = ["Emails", date, sender]
                save_path = p.joinpath(*dir_parts)
                if not save_path.exists():
                    os.makedirs(save_path)
                assert save_path.is_dir() == True, f"{save_path} is not a folder!"

                with open(save_path.joinpath(filename), "wb") as f:
                    f.write(part.get_payload(decode=True))
                    saved_files += 1
        logger.info(f"Saved {saved_files} files in {save_path.parent.parent}.")

        
        #print(subject_)
        #print(content_type)
        if "plain" in content_type:
            logger.debug(part.get_payload())
        elif "html" in content_type:
            html_ = part.get_payload()
            #soup = BeautifulSoup(html_, 'html.parser')
            #text = soup.get_text()
        return

def main():
    mail = Outlook()
    mail.login(env_var=False)
    mail.select_folder("INBOX")
    mail.fetch(sender=args.sender)
    mail.logout()

if __name__ == '__main__':
    main()