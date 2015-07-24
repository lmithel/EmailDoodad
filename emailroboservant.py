import poplib
import os
import email
import datetime
import getpass
from email import parser

#user = raw_input("Username: ")
#pw = getpass.getpass()

Mailbox = poplib.POP3_SSL('pop.gmail.com',995)
Mailbox.user('recent:email.filer@gobaci.com')
Mailbox.pass_('Filesfordays') 

Mailbox.set_debuglevel(0)
#Get messages from server:
numMessages = len(Mailbox.list()[1])
messages = [Mailbox.retr(i) for i in range(1, len(Mailbox.list()[1]) + 1)]
# Concat message pieces:
messages = ["\n".join(mssg[1]) for mssg in messages]
#Parse message into an email object:
messages = [parser.Parser().parsestr(mssg) for mssg in messages]

allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document",
"application/vnd.openxmlformats-officedocument.wordprocessingml.template","application/msword"
"application/vnd.ms-word.document.macroenabled.12","application/vnd.ms-word.template.macroenabled.12",
"image/jpeg","application/pdf","image/png","image/tiff","application/vnd.ms-excel","application/zip",
"application/x-7z-compressed","application/vnd.ms-excel","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]

for message in messages:
	if message['subject'] == '':
		path = 'misc'
	else:
		subject = message['subject'].split(':');
		if subject[0] == 'Fwd' or subject[0] == 'Re':
			subject.pop(0) #Remove the Fwd or Re
			subject[0] = subject[0][1:] #Remove the first space from the name after the colon
		filename = subject[0] #Filename is the name of the individual + What It's Regarding
		path = subject[0].split('-')[0] #Directory is the name of the individual


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
				now = str(datetime.datetime.now())
				f = open(path + "/" + now + " " + attachment_name,'wb')
			else:
				f = open(path + "/" + attachment_name, 'wb')
			f.write(data)
			f.close()

	if os.path.isfile(path + "/" + date + " " + who):
		text_file = open(path + "/" + filename + messageid, "w")
	else:
		text_file = open(path + "/" + filename, "w")
	text_file.write(msg)
	text_file.close()

Mailbox.quit()