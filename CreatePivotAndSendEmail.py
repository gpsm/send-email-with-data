import pandas as pd
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re

df = pd.read_excel('issues_list.xlsx',engine='openpyxl')

pivot = pd.pivot_table(df,values = 'JIRA ID',index = ['Assignee'],aggfunc='count').reset_index().rename(columns={'JIRA ID': 'Issues Count'})

df2 = {'Assignee':'Total Issues Count','Issues Count': pivot.sum()[1]}

pivot = pivot.append(df2, ignore_index = True)

pivot.style.applymap('font-weight: bold')

recipient = 'XXXX@XXXX.com'

subject = 'Team Pending JIRA Issues'

cc = 'abc@abc.com,avz@abc.com'

sender = 'XXXX@XXXX.com'  
# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = subject
msg['From'] = sender
msg['To'] = recipient
#msg['Cc'] = cc
 
myText=pivot.to_html()

result=re.sub('<th>\d*<\/th>','',myText)

result=re.sub('<th><\/th>','',result) 
 
# Create the body of the message (a plain-text and an HTML version).
text = '<html><body>Hello All, <br> Refer below count of issues pending at your end. Kindly fix the same asap. <br><br>' + result + '<br> Regards, <br> Your Name</body></html>'

part1 = MIMEText(text, 'html')

msg.attach(part1)


attach_file_name = 'issues_list.xlsx'

with open(attach_file_name, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {attach_file_name}",
)

# Add attachment to message 
msg.attach(part)
# Send the message via local SMTP server.
s = smtplib.SMTP('10.1.61.58',25)
# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.

s.sendmail(sender,recipient,msg.as_string())
s.quit()
