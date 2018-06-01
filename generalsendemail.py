SERVER = "casarray.gwl.bz"
FROM = "West.Wang@investorsgroup.com"
TO = ["West.Wang@investorsgroup.com"] # must be a list

SUBJECT = "Subject"
TEXT = "Your Text"

import smtplib
# Prepare actual message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

fromaddr = "Business Reporting (IG)"
#fromaddr = "shrepan@investorsgroup.com"
toaddr = "West.Wang@investorsgroup.com"
msg = MIMEMultipart()
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Python email"

body = "Python test mail"
msg.attach(MIMEText(body, 'plain'))

message = msg.as_string()

# Send the mail

server = smtplib.SMTP(SERVER)
#server.login("wangwe5", "Comeback15!")
server.sendmail(FROM, TO, message)
server.quit()