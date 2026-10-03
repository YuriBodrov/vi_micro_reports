# Project Name  : Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         									 getallvms.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module Allows Verification if vCenter Server Instance is Available 
and returns True if OK. Overwise, returns False and stops all next Steps.
"""
##############################################################################
# Modules for ease the Connect/Disconnect Operations with vCenter Servers 
from pyVim.connect import SmartConnect, Disconnect
from dataclasses import dataclass # Namespace for Operations with Data Classes
import ssl                        # Namespace For a Connection with vCenter 
																	# Server using SSL
import time                       # Namespace For the Time Delaying and 
																	# Sleep() Function
from pyVmomi import vim           # Namespace for a Core Operations with VI 
																	# Objects based on VMware Products

@dataclass # Main Class of this Module 'vcconnect.py'
class VcConnectionSteps(): ###################################################

	vcaddr:  str # Get vCenter Server Address
	vclogin: str # Get Username
	vcpassw: str # Get Password

	# By default these Variables set to 'None'
	is_vc_alive     = None	# Boolean/Logical Variable for the Connection with 
													# vCenter Server State
	vc_conn         = None	# Returns the vCenter Server's Connection Instance 
													# for Function_01
	primary_vc_conn = None	# Returns the vCenter Server's Connection Instance
													# for Function_02

	def verify_vc_connect_func(self): ##########################################

		print("")
		
		# Bypass SSL Certificate Verification for Self-Signed 
		# Certificates when connecting with vCenter Server Instance
		ssl_context = ssl._create_unverified_context()

		print("")
		print("Verifying vCenter Server Connection...")
		print("-" * 78)

		try: # This Block Starts when Connection with vCenter Server is OK 
			vc_conn      = SmartConnect(host=self.vcaddr, user=self.vclogin, \
															 		pwd=self.vcpassw, sslContext=ssl_context)
			vc_conn_time = vc_conn.CurrentTime()
			print(f"vCenter Server  : {self.vcaddr}".upper())
			print(f"Username        : {self.vclogin}".upper())
			print(f"Connection Time : {vc_conn_time}".upper())
			is_vc_alive  = True 
			time.sleep(2) # Program Scenario Sleeps for 2 Seconds

		# This Block Starts when Connection with vCenter Server is NOT OK
		except Exception as xex:
			print(f"Connection Error Occurred : {xex}".upper())
			is_vc_alive = False
			time.sleep(2)       # Program Scenario Sleeps for 2 Seconds
			Disconnect(vc_conn) # Disconnects from vCenter Server Instance

		finally: # This Block Starts anyway 
			print("-" * 78)
			print(" " * 70 + "...Done.")

		print("")

		return is_vc_alive

	def primary_connect_param_func(self): ######################################
		ssl_context = ssl._create_unverified_context()
		try:
			self.primary_vc_conn = SmartConnect(host=self.vcaddr, \
									user=self.vclogin, pwd=self.vcpassw, sslContext=ssl_context)
		except Exception as primary_xex:
			print(f"Connection Error Occurred : {primary_xex}".upper())

			Disconnect(self.primary_vc_conn)
		return self.primary_vc_conn