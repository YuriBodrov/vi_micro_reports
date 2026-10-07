# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         	 vm_get_ha_mntr_actn_evnts.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is used to Collect VMs with a Triggered Alarms about 
			 HA Monitoring Action Events.
"""

#########################################################################################
from dataclasses import dataclass  # Namespace for Operations with Data Classes
from time import sleep             # Getting a Sleep() Function for an Output Delay 
from pyVmomi import vim            # Namespace for a Core Operations with VI Objects
from sys import stdout						 # Function for an Output Print Options
import openpyxl                    # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module : 'vm_get_ha_mntr_actn_evnts.py'
class GetVmsHaTrigAlrmClass(): ##########################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	
	def get_vm_trig_alrm_func(self): ######################################################
		pass
		vi_content = self.vc_instance.RetrieveContent()
		if ((vi_content is not None) and (self.connstate)):
			
			container = vi_content.rootFolder          # Root/Parent Directory Starting Point
			trig_alrms = container.triggeredAlarmState # Searching for Triggered Alarms 
		
			# Open and Initialize the Existing(!) MS Excel File
			xlsx_wb = openpyxl.load_workbook("report.xlsx")

			# Create a Thematic Sheet and Set it as Active
			if (not ("vm_ha_monitor_errors" in xlsx_wb.sheetnames)):
				pass
				xlsx_wb.create_sheet(title = "vm_ha_monitor_errors")
				active_xlsx_sheet = xlsx_wb["vm_ha_monitor_errors"] # type: ignore

				# Set the Column Headers of Active XLSX Sheet
				column_header_list = ["VM Name", "Event Data", "Status/Color", "Event's Datetime"]
				active_xlsx_sheet.append(column_header_list)				# type: ignore
				xlsx_wb.save("report.xlsx")                  				# type: ignore

			else:
				active_xlsx_sheet = xlsx_wb["vm_ha_monitor_errors"] # type: ignore

			stdout.write(f"[INFO] : VM. Collecting HA Monitoring Action Triggered Alarms...")
			stdout.flush()
			for alrm_state in trig_alrms:     				 						# type: ignore
				pass
				# Get only Useful Information from all Alarms Data:
				alarm      = alrm_state.alarm
				alarm_info = alarm.info

				if ("virtual machine monitoring" in alarm_info.name.lower()):
					pass
					vm_entity_name = alrm_state.entity.name if hasattr(alrm_state.entity, "name") \
					else str(alrm_state.entity)
					xlsx_data = (vm_entity_name, alarm_info.name, alrm_state.overallStatus, \
									alrm_state.time.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3])
					active_xlsx_sheet.append(xlsx_data)
					xlsx_wb.save("report.xlsx")

			container.Destroy() # to Avoid Memory Accumulation in the vCenter Server
			print("Done.")
			sleep(2)
			print ("")
			