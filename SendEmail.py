from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Robocorp.Vault import Vault
import logging


class SendEmail:

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)

    def send_email(self, recipient, subject, body):
        try:
            secret = Vault().get_secret("emailCredentials")
            gmail_account = secret["username"]
            gmail_password = secret["password"]
            mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
            mail.authorize(account=gmail_account, password=gmail_password)
            mail.send_message(
                sender=gmail_account,
                recipients=recipient,
                subject=subject,
                body=body,
            )
        except Exception as err:
            self.logger.error("SendEmail unexpectedly failed: " + str(err))
            pass