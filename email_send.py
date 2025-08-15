import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def send_email_with_pdf(pdf_path, body,
                        to_email="КОМУ ВІДПРАВЛЯЄМО@gmail.com",
                        subject="ЗАГОЛОВОК Лист від програми"
                        ):
    from_email = "ВІД КОГО@gmail.com"
    from_password = "________________"  # Пароль додатку з Gmail

    # Витягуємо тільки ім'я файлу з повного шляху
    filename = os.path.basename(pdf_path)

    # Створюємо лист
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(pdf_path, 'rb') as file:
        part = MIMEApplication(file.read(), Name=filename)
        part['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(from_email, from_password)
        server.sendmail(from_email, to_email, msg.as_string())

    print(f"Лист з PDF '{filename}' надіслано до {to_email}.")
