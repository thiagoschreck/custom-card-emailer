import random
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from PIL import Image
from html2image import Html2Image
from unidecode import unidecode

recipient_name_token = "[Recipient Name]"
directory = ".\\generated\\"

# ref: https://github.com/vgalin/html2image/blob/master/README.md
hti = Html2Image(output_path=directory + "\\screenshots",
                 custom_flags=['--hide-scrollbars'], size=(1440, 1082))


# this generated ID will try to prevent files having the same filename
def generate_id():
    return ''.join(random.choices('123456789', k=4))


def send_email(recipient_data, image_data):
    address = recipient_data.email
    name = recipient_data.name
    surname = recipient_data.surname

    with open("./email_template.html", encoding='utf-8') as fp:
        msg = MIMEMultipart()
        read = fp.read()
        read = read.replace(recipient_name_token, name + " " + surname)
        msg.attach(MIMEText(read, "html"))

    with open(image_data.path, 'rb') as fp:
        attachment = MIMEApplication(fp.read())
    attachment.add_header("Content-Disposition", f"attachment; filename=AttachedFile.png", )
    msg.attach(attachment)

    msg['Subject'] = f"Subject text"
    msg['From'] = "Your Username or Email"
    msg['To'] = address
    smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    try:
        # FIXME avoid logging in for each email sent, if possible
        smtp_server.login(msg['From'], 'Your Password')
        smtp_server.send_message(msg)
        print("Sending email to " + address)
        smtp_server.quit()
    except Exception as e:
        # do something else
        print(e)


class RecipientData:
    def __init__(self, row_number, array_with_data):
        # the row number is just useful to have
        self.row_number = row_number
        # array indexes are according to your source .xlsx
        self.email = array_with_data[0]
        # for some reason, Html2Image will not render unicode properly
        self.name = unidecode(array_with_data[1])
        self.surname = unidecode(array_with_data[2])


class ImageData:
    def __init__(self, path, recipient_data):
        self.path = path
        self.recipient_data = recipient_data
        self.path = None

    # This will generate the custom .html file, take a screenshot of it, crop it, and email it to the recipient
    def generate_and_send(self):
        html_string = self.generate_html()
        image_location = self.screenshot_generated_html(html_string)
        self.path = image_location
        image = Image.open(image_location)
        image = image.crop((0, 0, 720, 541))
        image.save(image_location)
        send_email(self.recipient_data, self)

    def screenshot_generated_html(self, html_string):
        file_name = self.get_file_name()
        file = open(self.path + file_name + ".html", "w")
        file.write(html_string)
        file.close()
        hti.screenshot(html_file=self.path + file_name + ".html", css_file="style.css",
                       save_as=file_name + ".png")
        image_location = self.path + "\\screenshots\\" + file_name + ".png"
        return image_location

    def generate_html(self):
        template = open("template.html", "r")
        html_string = template.read()
        template.close()
        html_string = html_string \
            .replace(recipient_name_token, unidecode(self.recipient_data.name + " " + self.recipient_data.surname))
        return html_string

    def get_file_name(self):
        return (generate_id() + "_"
                + self.recipient_data.name + "_" + self.recipient_data.surname) \
            .replace(" ", "_").replace("__", "_")


data_list = []
values_list = pd.read_excel("sample_dataset.xlsx").values
for row, data in enumerate(values_list):
    data_list.append(RecipientData(row, data))

image_data_list = []

for data in data_list:
    image_data_list.append(ImageData(directory, data))

for img_data in image_data_list:
    img_data.generate_and_send()
