import datetime
import win32com.client as client


def all_seller_outlook():
    date = datetime.datetime.now().strftime("%d-%m-%Y")
    print(date)
    att_file = r"D:\Durai\GMB\Reviews_count\Save Fiels\Google Ratings & Reviews " + date + ".xlsx"
    print(att_file)

    body = """
    <html>
        <body>
            <p>Hi Team,</p>
            <p>I have attached the Google Ratings & Reviews here.</p>
            <p>Thanks & Regards,<br>
                Duraikannan.R<br>
                Phone: 8682997570</p>
            <p><img src = "D:\Durai\GMB\Reviews_count\Poorvika_logo.png"><br>
                Poorvika Mobiles Pvt Ltd.</p>
        </body>
    </html>
    """

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.To = 'ganapathy.k@poorvika.com; gowrisankar.p@poorvika.com'
    message.CC = 'hepzi@poorvika.com'
    message.Subject = "Google Ratings & Reviews " + date
    message.HTMLBody = body
    message.Attachments.Add(att_file)


all_seller_outlook()
