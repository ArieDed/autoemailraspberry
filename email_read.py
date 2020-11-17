#!/usr/bin/python

import outlook
import time
import base64
import os
import subprocess
from datetime import datetime
from email.mime.text import MIMEText
import json
import RPi.GPIO as GPIO

GPIO.setmode(GPIO.BCM)
GPIO.setup(23, GPIO.IN, pull_up_down=GPIO.PUD_UP)#Button to GPIO23

def Say(text):
  os.popen( 'espeak -s110 -g10 -vnl+m6 "'+text+'" --stdout | aplay 2>/dev/null' )

def get_len(filename):
   result = subprocess.Popen(["ffprobe", filename, '-print_format', 'json', '-show_streams', '-loglevel', 'quiet'],
     stdout = subprocess.PIPE, stderr = subprocess.STDOUT)
   return float(json.loads(result.stdout.read())['streams'][0]['duration'])

num2words = {1: 'Een', 2: 'Twee', 3: 'Drie', 4: 'Vier', 5: 'Vijf', \
            6: 'Zes', 7: 'Zeven', 8: 'Acht', 9: 'Negen', 10: 'Tien', \
            11: 'Elf', 12: 'Twaalf', 13: 'Dertien', 14: 'Viertien', \
            15: 'Vijftien', 16: 'Zestien', 17: 'Zeventien', 18: 'Achttien', 19: 'Negentien'}

def mail_read_unread():
	mail = outlook.Outlook()
	mail.login('YOUR_OUTLOOK_EMAIL','YOUR_OUTLOOK_PASSWORD')
	mail.select('OUTLOOK_MAIL_FOLDER')
	unread_emails = mail.unreadIds()
	#unread_are = mail.hasUnread()
	#all_emails = mail.allIds()
	#i_1 = int(unread_emails[0])
	#i_last = int(unread_emails[-1])
	#i = int(unread_emails[0])
	#print unread_emails
	#print i_1
	#print i_last
	list_from = []
	list_message = []
	list_email = []
	list_photo_from = []
	list_video_from = []
	list_foto_files = []
	list_video_files = []
	try:
		for i in unread_emails:
			mail.getEmail(i)
			subject_receive = mail.mailsubject()
			if subject_receive == 'SUBJECT_FOR_TEXT':
				words = mail.mailbody()
				words_split = words.split()
				id_start = words_split.index('*Bericht*')+1
				from_start = words_split.index('*Naam*')+1
				from_end = words_split.index('*Email')
				email_start = words_split.index('adres*')+1
				email_end = words_split.index('*Bericht*')
				message = " ".join(words_split[id_start:-4])
				from_name = " ".join(words_split[from_start:from_end])
				email_from = "".join(words_split[email_start:email_end])
				list_from.append(from_name)
				list_message.append(message)
				list_email.append(email_from)
			if subject_receive == 'SUBJECT_FOR_PHOTO_OR_VIDEO':
				msg = mail.getEmail(i)
				attachment_test = msg.get_payload()[1]
				type_test = attachment_test.get_content_type()
				#print type_test
				if type_test[:5] != "image":
					words = mail.mailbody()
					words_split = words.split()
					name_start = words_split.index('van')+1
					name_end = words_split.index('Groetjes')
					name_from = " ".join(words_split[name_start:name_end])
					name_from = name_from[:-1]
					list_video_from.append(name_from)
					amount_videos = range(1,len(msg.get_payload()))
					for m in amount_videos:
						attachment = msg.get_payload()[m]
						timenow1 = datetime.now()
						timenow = timenow1.strftime("%H:%M:%S_%d-%m-%Y")
						filename = "/home/pi/videos/%s_%s_%d.mp4" % (name_from, timenow, m)
						list_video_files.append(filename)
						open(filename, 'wb').write(attachment.get_payload(decode=True))
				else:
					try:
						words = base64.b64decode(mail.mailbody())
					except:
						words = mail.mailbody()
					#print words
					amount_pic = range(1,len(msg.get_payload()))
					words_split = words.split()
					name_start = words_split.index('van')+1
					name_end = words_split.index('Groetjes')
					name_from = " ".join(words_split[name_start:name_end])
					name_from = name_from[:-1]
					list_photo_from.append(name_from)
					#print name_from
					#print amount_pic
					for n in amount_pic:
						attachment = msg.get_payload()[n]
						filetype = attachment.get_content_type()
						filetype = filetype[6:]
						timenow1 = datetime.now()
						timenow = timenow1.strftime("%H:%M:%S_%d-%m-%Y")
						#print timenow
						#print type(name_from)
						#print type(filetype)
						#print type(timenow)
						#print filetype
						filename = "/home/pi/fotos/%s_%s_%d.%s" % (name_from, timenow, n, filetype)
						list_foto_files.append(filename)
						open(filename, 'wb').write(attachment.get_payload(decode=True))
	except:
		print "No new emails received, Problem encountered"
		raise SystemExit(0)
	#mail.getEmail(all_emails[3])
	#words = mail.mailbody()
	#words_split = words.split()
	#id_start = words_split.index('*Bericht*')+1
	#from_start = words_split.index('*Naam*')+1
	#from_end = words_split.index('*Email')
	#message = " ".join(words_split[id_start:-4])
	#from_name = " ".join(words_split[from_start:from_end])
	#print "Bericht van " + from_name + ". " + message
	size_list = list(range(len(list_email)))

	if len(list_email) != 0:
		print "Received email with message"
		email_message = True
	else:
		email_message = False
	if len(list_photo_from) != 0:
		print "Received email with photo"
		email_photo = True
	else:
		email_photo = False
	if len(list_video_from) != 0:
		print "Received email with video"
		email_video = True
	else:
		email_video = False
	if email_message == False and email_photo == False and email_video == False:
		print "No new emails received"

	#if email_message == True:
	#	amount_message = range(0,len(list_message))
	#	Say("Hi opa, u heeft " + num2words[len(list_message)] + " nieuwe berichtjes.")
	#	for j in amount_message:
	#		Say("Berichtje van " + list_from[j] + ".")
	#		Say(list_message[j])

	if email_photo == True:
		amount_fotos = range(0,len(list_foto_files))
		for l in amount_fotos:
			print "Showing foto" + list_foto_files[l]
			os.system('sudo fbi --autozoom --noverbose -t20 -1 --vt 1 "{}"'.format(list_foto_files[l]))
			time.sleep(20)
		os.system("sudo pkill fbi")

	if email_video == True:
		amount_videos = range(0,len(list_video_files))
		for i in amount_videos:
			video_length = get_len(list_video_files[i])
			print video_length
			omxprocess = subprocess.Popen(['omxplayer',list_video_files[i]],stdin=subprocess.PIPE)
			time.sleep(video_length+3)

	#raise SystemExit(0)

	print list_from
	print list_message
	print list_email

	for j in size_list:
		recipient = list_email[j]
		subject = "REPLY_EMAIL_SUBJECT"
		to_who = list_from[j]

		return_message_text = "Dear " + to_who + ",\n\nMESSAGEFOREMAIL"
		return_message_html = """\
	<html>
	  <head>
		<meta name="viewport" content="width=device-width">
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<title>Simple Transactional Email</title>
		<style>
	@media only screen and (max-width: 620px) {{
	  table[class=body] h1 {{
		font-size: 28px !important;
		margin-bottom: 10px !important;
	  }}

	  table[class=body] p,
	table[class=body] ul,
	table[class=body] ol,
	table[class=body] td,
	table[class=body] span,
	table[class=body] a {{
		font-size: 16px !important;
	  }}

	  table[class=body] .wrapper,
	table[class=body] .article {{
		padding: 10px !important;
	  }}

	  table[class=body] .content {{
		padding: 0 !important;
	  }}

	  table[class=body] .container {{
		padding: 0 !important;
		width: 100% !important;
	  }}

	  table[class=body] .main {{
		border-left-width: 0 !important;
		border-radius: 0 !important;
		border-right-width: 0 !important;
	  }}

	  table[class=body] .btn table {{
		width: 100% !important;
	  }}

	  table[class=body] .btn a {{
		width: 100% !important;
	  }}

	  table[class=body] .img-responsive {{
		height: auto !important;
		max-width: 100% !important;
		width: auto !important;
	  }}
	}}
	@media all {{
	  .ExternalClass {{
		width: 100%;
	  }}

	  .ExternalClass,
	.ExternalClass p,
	.ExternalClass span,
	.ExternalClass font,
	.ExternalClass td,
	.ExternalClass div {{
		line-height: 100%;
	  }}

	  .apple-link a {{
		color: inherit !important;
		font-family: inherit !important;
		font-size: inherit !important;
		font-weight: inherit !important;
		line-height: inherit !important;
		text-decoration: none !important;
	  }}

	  #MessageViewBody a {{
		color: inherit;
		text-decoration: none;
		font-size: inherit;
		font-family: inherit;
		font-weight: inherit;
		line-height: inherit;
	  }}

	  .btn-primary table td:hover {{
		background-color: #34495e !important;
	  }}

	  .btn-primary a:hover {{
		background-color: #34495e !important;
		border-color: #34495e !important;
	  }}
	}}
	</style>
	  </head>
	  <body class="" style="background-color: #f6f6f6; font-family: sans-serif; -webkit-font-smoothing: antialiased; font-size: 14px; line-height: 1.4; margin: 0; padding: 0; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%;">
		<span class="preheader" style="color: transparent; display: none; height: 0; max-height: 0; max-width: 0; opacity: 0; overflow: hidden; mso-hide: all; visibility: hidden; width: 0;">Berichtje terug van opa</span>
		<table role="presentation" border="0" cellpadding="0" cellspacing="0" class="body" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f6f6f6; width: 100%;" width="100%" bgcolor="#f6f6f6">
		  <tr>
			<td style="font-family: sans-serif; font-size: 14px; vertical-align: top;" valign="top">&nbsp;</td>
			<td class="container" style="font-family: sans-serif; font-size: 14px; vertical-align: top; display: block; max-width: 580px; padding: 10px; width: 580px; margin: 0 auto;" width="580" valign="top">
			  <div class="content" style="box-sizing: border-box; display: block; margin: 0 auto; max-width: 580px; padding: 10px;">

				<!-- START CENTERED WHITE CONTAINER -->
				<table role="presentation" class="main" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; background: #ffffff; border-radius: 3px; width: 100%;" width="100%">

				  <!-- START MAIN CONTENT AREA -->
				  <tr>
					<td class="wrapper" style="font-family: sans-serif; font-size: 14px; vertical-align: top; box-sizing: border-box; padding: 20px;" valign="top">
					  <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;" width="100%">
						<tr>
						  <td style="font-family: sans-serif; font-size: 14px; vertical-align: top;" valign="top">
							<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; margin-bottom: 15px;">Dear {to_who},</p>
							<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; margin-bottom: 15px;">MESSAGEFOREMAIL.</p>
							<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; margin-bottom: 15px;">Regards,</p>
	<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; margin-bottom: 15px;">SENDEROFEMAIL</p><p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; margin-bottom: 15px;">
						  </p></td>
						</tr>
					  </table>
					</td>
				  </tr>

				<!-- END MAIN CONTENT AREA -->
				</table>
				<!-- END CENTERED WHITE CONTAINER -->

				<!-- START FOOTER -->
				<div class="footer" style="clear: both; margin-top: 10px; text-align: center; width: 100%;">
				  <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;" width="100%">
					<tr>
					  <td class="content-block" style="font-family: sans-serif; vertical-align: top; padding-bottom: 10px; padding-top: 10px; color: #999999; font-size: 12px; text-align: center;" valign="top" align="center">
						<br> FOOTER_TEXT <br>
				<a href="FOOTER_HYPERLINK" style="text-decoration: underline; color: #999999; font-size: 12px; text-align: center;"> <img src="FOOTER_IMAGE" alt="ALTERNATIVE_FOOTER" width="50" border="0" class="style="border:0;" outline:none;="" text-decoration:none;="" display:block;"="" style="border: none; -ms-interpolation-mode: bicubic; max-width: 100%;"> </a>
					  </td>
					</tr>
				  </table>
				</div>
				<!-- END FOOTER -->

			  </div>
			</td>
			<td style="font-family: sans-serif; font-size: 14px; vertical-align: top;" valign="top">&nbsp;</td>
		  </tr>
		</table>
	  </body>
	</html>
	""".format(to_who=to_who)

		part1 = MIMEText(return_message_text, 'plain')
		part2 = MIMEText(return_message_html, 'html')
		mail.sendEmailMIME(recipient,subject,part1,part2)
		print "Send email to " + list_from[j]
		time.sleep(1)

try:
	while True:
		button_state = GPIO.input(23)
		if button_state == False:
			print "Button Pressed..."
			mail_read_unread()
except:
	GPIO.cleanup()