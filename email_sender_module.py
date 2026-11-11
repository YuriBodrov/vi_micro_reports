# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         		 		 email_sender_module.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is designed to Send Email Messages to Recipients + Attachment within!
"""
#########################################################################################
import mimetypes
import os
import smtplib
from email.message import EmailMessage

class EmailSenderClass:
	pass
	
	def email_sndr_func(self, smtp_relay, smtp_port, recip_addrs, email_attch): ###########
		pass
		sndr_email_addr = "vmware_automation@alfastrah.ru"

		try:
			pass
			# Set an Email Message's Structure:
			mssg = EmailMessage()
			mssg["Subject"] = "VI MicroReport by Default"
			mssg["From"]    = sndr_email_addr
			mssg["To"]      = ", ".join(recip_addrs) # Send to Multiple Recipients
			mssg_body       = "Hello!" + "\n" + "Report File in Attachment. Good luck!"
			mssg_sgntr      = "\n" + "\n" + "---" + "\n" + "VI MicroReports. osis@alfastrah.ru"
			mssg.set_content(mssg_body + mssg_sgntr)

			# Detection of File's Type and Attachment Addition:
			if os.path.exists(email_attch):
				pass
				get_mime_type = mimetypes.guess_type(email_attch)
				#print(f"{get_mime_type}")
				#input("Press any Key to Continue...")
				
				if (get_mime_type is None):
					get_mime_type = "application/octet-stream" # Default Type for Binary Data
				
				main_type, sub_type = get_mime_type[0].split("/", 1) # type: ignore
				#att_file_name = os.path.basename(email_attch)

			# Attach the File to the Letter:
			with open(email_attch, "rb") as f:
				pass
				mssg.add_attachment(f.read(), \
				maintype = main_type, subtype = sub_type, filename = email_attch) # type: ignore
			
			# Connection to Email Server:
			try:
				pass
				with smtplib.SMTP(smtp_relay, smtp_port) as email_server:
					pass
					email_server.send_message(mssg)
				print(f"---")
				print(f"The Letter has been sent successfully!")
				print(f"")
				
			except Exception as smtp_xex:
				pass
				print(f"We have an Error!")
				print(f"---")
				print(f"{smtp_xex}")

				

		except Exception as xex:
			pass
			print(f"Error! Unable to perform an Operation!")
			print(f"---")
			print(xex)