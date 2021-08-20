import smtplib, ssl, csv, codecs, logging
from email.mime.text import MIMEText

logging.basicConfig(filename='RAEmailBlastLog.log', level=logging.INFO)

# Set up log in information
port = 465
gmail_user = ""
gmail_password = ""

# Set up address variables
sent_from = ""
recipient_email = ""

# Open HTML file, read in
html = codecs.open("Body.html", 'r')
bodyOrig = html.read()
html.close()

# Prepare server
context = ssl.create_default_context()
server = smtplib.SMTP_SSL("smtp.gmail.com", port, context=context)
server.login(gmail_user, gmail_password)

# Open CSV and send emails
with open('', encoding="utf-8") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')

    for row in csv_reader:

        # If row[6] is empty then the line contains no contact information and should be skipped
        if (len(row[6]) > 1):

            # Find [insert addressee here] and replace with proper addressee
            body = bodyOrig.replace('[INSERT ADDRESSEE HERE]', row[6])

            # Fill in subject, sender, recipient
            msg = MIMEText(body, 'html')
            msg['Subject'] = ""
            msg['From'] = ""

            # Determine if there is a CC
            if (len(row[4]) > 1): # if CC
                recipient_email = row[4]
                msg['To'] = recipient_email
                try:
                    msg['cc'] = row[7]
                    to = [recipient_email] + [msg['cc']]
                except:
                    to = [recipient_email]
                # Send
                try:
                    server.sendmail(sent_from, to, msg.as_string())
                except smtplib.SMTPSenderRefused as s:
                    logging.info(s)
                # Log
                try:
                    logging.info("Email sent to %s at %s, cc'd %s", row[6], recipient_email, row[7])
                except:
                    logging.info("Email sent to %s at %s", row[6], recipient_email)

            elif (len(row[4]) <= 0): # if no CC
                recipient_email = row[7]
                msg['To'] = recipient_email
                to = [recipient_email]
                # Send
                try:
                    server.sendmail(sent_from, to, msg.as_string())
                except SMTPRecipientsRefused as s:
                    logging.info(s)
                # Log
                logging.info("Email sent to %s at %s", row[6], recipient_email)
        else:
            logging.info("No addressee on row %s", row)


server.quit()
