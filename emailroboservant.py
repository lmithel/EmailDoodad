import poplib
import os
import email
import getpass
from email import parser


pw = getpass.getpass()
Mailbox = poplib.POP3_SSL('pop.gmail.com',995)
Mailbox.user('recent:email.filer@gobaci.com') #lord knows why this doesn't work on a mac
Mailbox.pass_(pw)
#print Mailbox.stat()
Mailbox.set_debuglevel(0)
#Get messages from server:
numMessages = len(Mailbox.list()[1])
messages = [Mailbox.retr(i) for i in range(1, len(Mailbox.list()[1]) + 1)]
# Concat message pieces:
messages = ["\n".join(mssg[1]) for mssg in messages]
#Parse message intom an email object:
messages = [parser.Parser().parsestr(mssg) for mssg in messages]
allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document","application/pdf"]
for message in messages:
	if message['subject'] == '':
		path = 'misc'
	else:
		path = message['subject']
	date = message['Date']
	#time = message['Received']
	who = message['From']
	messageid = message['Message-ID']

	if not os.path.exists(path):
		os.mkdir(path)

	for part in message.walk():
		if part.get_content_type() == 'text/plain':
			msg = str(part)
		if part.get_content_type() in allowed_mimetypes:
			attachment_name = part.get_filename()
			data = part.get_payload(decode=True)
			if os.path.isfile(path + "/" + attachment_name):
				os.mkdir(path + "/" + who)
				f = open(path + "/" + who + "/" + attachment_name,'wb')
				print f
			else:
				f = open(path + "/" + attachment_name, 'wb')
			f.write(data)
			f.close()

	if os.path.isfile(path + "/" + date + " " + who):
		text_file = open(path + "/" + date + " " + who + messageid, "w")
	else:
		text_file = open(path + "/" + date + " " + who, "w")
	text_file.write(msg)
	text_file.close()


Mailbox.quit()