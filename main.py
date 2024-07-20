from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd
import configparser
import logging


logging.basicConfig(
    filename="execution.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


config = configparser.ConfigParser()
config.read("./config.ini")
config_data = config["EmailConfig"]

password = config_data.get("password", "")
sender_email = config_data.get("sender_email", "")
sender_name = config_data.get("sender_name", "")
attachment_path = config_data.get("attachment_path", "")
filename = config_data.get("filename", "")


logger = logging.getLogger(__name__)


class EmailData:
    def __init__(self, receiver_email, receiver_name, salutation, company_name):
        self.receiver_email = receiver_email
        self.receiver_name = receiver_name
        self.salutation = salutation
        self.company_name = company_name

        with open("./email_subject_template.txt", "r") as template:
            email_subject_template = template.read()
        self.subject = email_subject_template.format(company_name=company_name)

        with open("./email_body_template.txt", "r") as template:
            email_body_template = template.read()
        self.email_body = email_body_template.format(
            receiver_name=receiver_name,
            salutation=salutation,
            company_name=company_name,
        )


def send_email(email_data):
    try:
        em = EmailMessage()
        em.add_header("From", f"{sender_name} <{sender_email}>")
        em["TO"] = email_data.receiver_email
        em["Subject"] = email_data.subject
        em.set_content(email_data.email_body)

        with open(attachment_path, "rb") as attachment_file:
            attachment_data = attachment_file.read()

        em.add_attachment(
            attachment_data,
            maintype="application",
            subtype="pdf",
            filename=filename,
        )

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL("smtp.gmail.com", "465", context=context) as smtp:
            smtp.login(user=sender_email, password=password)
            smtp.send_message(em)

        logger.info(f"EMAIL SENT TO {email_data.receiver_email}")
        print(f"EMAIL SENT TO {email_data.receiver_email}")

    except FileNotFoundError:
        logger.error(f"Attachment file not found: {attachment_path}")
        logger.error(f"EMAIL NOT SENT TO {email_data.receiver_email}")

    except Exception as e:
        logger.error(f"ERROR: {e}")
        logger.error(f"EMAIL NOT SENT TO {email_data.receiver_email}")


def main():
    try:
        data = pd.read_excel("./data.xlsx")

        for index, receiver_data in data.iterrows():
            if receiver_data["SENT"] == True:
                pass
            else:
                ed = EmailData(
                    receiver_email=receiver_data["receiver_email"],
                    receiver_name=receiver_data["receiver_name"],
                    salutation=receiver_data["salutation"],
                    company_name=receiver_data["company_name"],
                )
                send_email(ed)

                data.at[index, "SENT"] = True

        data.to_excel("./data.xlsx", index=False)
        logger.info("Script execution completed successfully")
        print("Script execution completed successfully")

    except FileNotFoundError as fnfe:
        logger.error(f"ERROR: {fnfe}")
        logger.error("File not found: data.xlsx or template files")
        logger.error("Script execution failed")
        print("Script execution failed")

    except Exception as e:
        logger.error(f"ERROR: {e}")
        logger.error("Script execution failed")
        print("Script execution failed")


if __name__ == "__main__":
    main()

    input("PRESS ANY KEY TO EXIT")
