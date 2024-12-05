import re

fp = open('message.txt', 'rb')
msg = MIMEText(fp.read())
fp.close()
msg['Subject'] = 'Hello This is an Test Email'
msg['From'] = 'muralik@vector97.com'
server = smtplib.SMTP('smtp.office365.com', 587)
server.starttls()
server.login('muralik@vector97.com', 'Munn@1989')
email_data = csv.reader(open('email.csv', 'rb'))
email_pattern= re.compile("^.+@.+\..+$")
for row in email_data:
  if( email_pattern.search(row[1]) ):
    del msg['To']
    msg['To'] = row[1]
    try:
      server.sendmail('test@gmail.com', [row[1]], msg.as_string())
    except SMTPException:
      print "An error occured."
server.quit()