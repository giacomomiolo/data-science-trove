import email
import imaplib
import os
import re

# To be moved to config.py file
imap_server = "imap-mail.outlook.com"
imap_port = 993
smtp_server = "smtp-mail.outlook.com"
smtp_port = 587

class Outlook():
    def __init__(self):
        pass

    def login(self, username: str, env_var: bool = True) -> None:
        """Authentication through IMAP over SSL (port 993).

        Args:
            username (str): The full email adress, e.g. name.surname@company.com
            env_var (bool, optional): Whether to use an environmental variable to access the password or input a string. Defaults to True.
        """
        self.username = username
        
        if env_var:
            try:
                password = os.environ['OUTLOOK_PW']
            except KeyError:
                ### TODO
                # Add instructions to set env variables maybe?
                # Mac/Linux export VAR=VALUE from terminal
                # Windows https://www.computerhope.com/issues/ch000549.htm
                print(f"Please set the environment variable OUTLOOK_PW")
                return
        else:
            password = input(f"Please enter your password: ")
        
        self.password = password
        login_attempts = 0
        
        while True:
            try:
                self.imap = imaplib.IMAP4_SSL(imap_server, imap_port)
                response, data = self.imap.login(username, password) # returns "OK" if successful and the payload
                assert response == "OK", f"Login failed: {response}"
                print(f"> {data[0].decode()} Signed in as {self.username}")
                return
            except Exception as error:
                print(f"> Sign in error: {type(error)}")
                login_attempts += 1
                if login_attempts < 3:
                    continue
                assert False, 'Login Failed'
                
    def logout(self) -> str:
        self.imap.logout()
        # TODO
        # Maybe add some further checks that the session was closed correctly, but shoudln't be necessary
        return print(f"Successfully logged out")

if __name__ == '__main__':
    mail = Outlook()
    username = input(f"Please enter your username: ")
    mail.login(username, env_var=False)
    print(f"... Some operations ...")
    mail.logout()