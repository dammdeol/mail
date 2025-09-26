import win32com.client as win32
import datetime
import os
import pandas as pd

def leer_destinatarios(path_excel):
    if not os.path.exists(path_excel):
        raise FileNotFoundError(f"Excel file not found: {path_excel}")
    try:
        df = pd.read_excel(path_excel)
    except Exception as e:
        raise RuntimeError(f"Error reading Excel file: {e}") from e

    to_list = df['To'].dropna().tolist()
    cc_list = df['CC'].dropna().tolist()

    to = "; ".join(to_list)
    cc = "; ".join(cc_list)

    return to, cc


def send_advanced_email(
    to, subject, html_body, cc=None, attachments=None, send_at=None, images=None
):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    if cc:
        mail.CC = cc

    mail.Display()
    signature = mail.HTMLBody
    mail.Close(0)

    # Adjuntar imágenes embebidas
    if images:
        for cid, path in images.items():
            if os.path.exists(path):
                attachment = mail.Attachments.Add(Source=path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
                )
            else:
                print(f"⚠️ Imagen no encontrada: {path}")

    # Agregar firma al cuerpo
    mail.HTMLBody = html_body + signature

    # Adjuntos
    if attachments:
        for file in attachments:
            if os.path.exists(file):
                mail.Attachments.Add(file)
            else:
                print(f"⚠️ Archivo no encontrado: {file}")

    # Envío programado
    if send_at and isinstance(send_at, datetime.datetime):
        mail.DeferredDeliveryTime = send_at

    mail.Display()
    #mail.Send()
    print("✅ Correo enviado o programado.")


# Ejemplo de uso en el bloque principal
if __name__ == "__main__":
    base_dir = os.path.dirname(__file__)
    send_time = datetime.datetime.now() 
    report_date = "11/09/2025"
    report_date_str = datetime.datetime.strptime(report_date, "%d/%m/%Y")
    report_date_en = report_date_str.strftime("%B %d, %Y")
    report_date_short = report_date_str.strftime("%d %b %y")
    to, cc = leer_destinatarios(os.path.join(base_dir, "mail.xlsx"))

    with open(os.path.join(base_dir, "mail.html"), "r", encoding="utf-8") as f:
        html_body = f.read()
    
    created_at = datetime.datetime.now().strftime("%d %b %Y %H:%M")
    html_body = html_body.replace("{{report_date_en}}", report_date_en)
    html_body = html_body.replace("{{created_at}}", created_at)
    images = {
        "img1": os.path.join(base_dir, "img1.png"),
        "img2": os.path.join(base_dir, "img2.png"),
        "img3": os.path.join(base_dir, "img3.png"),
        "img4": os.path.join(base_dir, "img4.png"),
    }

    for cid, path in images.items():
        html_body = html_body.replace(f"cid:{cid}", f"file:///{path.replace(os.sep, '/')}")

    send_advanced_email(
        to=to,
        cc=cc,
        subject="Report – " + report_date_short,
        html_body=html_body,
        attachments=None,
        send_at=None,
        images=images
    )