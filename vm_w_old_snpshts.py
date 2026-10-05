# Project Name  : 										Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         								vmswoldsnaps.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is used to retrieve an Old VMs Snapshots.
This Module's Function returns "VM_NAME|SNPSHT_ID|SNPSHT_NAME|C_DATE|DAYS_GONE" List.
"""
#########################################################################################
from sys import stdout						                 # Function for an Output Print Options
from datetime import datetime, timedelta, timezone # For Datetime Operations
from dataclasses import dataclass                  # Operations with Data Classes
from time import sleep                             # Time Delaying and Sleep() Function
from pyVmomi import vim                            # Core Operations with VI Objects
from openpyxl import Workbook                      # For an Operations with MS Excel File

@dataclass # Main Class of this Module 'vmswoldsnaps.py'
class GetOldVmSnpshts(): ################################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	
	"""
	NOTE : This Function Recursively figuring the Snapshot Tree and returns Snapshots 
																							Older than a Specified Number of Days
	"""
	def get_snpsht_data(self, snpsht_lst, days_gone): # type: ignore ######################
		pass
		old_snpshts_lst = [] # Define this Variable as VM's Snapshots List
		date_as_limit   = datetime.now(timezone.utc) - timedelta(days = days_gone)

		for snpsht in snpsht_lst:
			pass
			if (snpsht.createTime < date_as_limit):
				pass
				# List Item Format : JSON "Key:Value" Format
				old_snpshts_lst.append({
					"name": snpsht.name,
					"created": snpsht.createTime,
					"id": snpsht.id
				})

			# NOTE : If there are Child Snapshots -> Go to Recursion
			if (hasattr(snpsht, "childSnapshotList") and (snpsht.childSnapshotList)):
				pass
				old_snpshts_lst.extend(self.get_snpsht_data(snpsht.childSnapshotList, days_gone)) 

		return old_snpshts_lst

	"""
	NOTE : This Function runs as Interaction Bridge between Snapshot Data Processing
				 																							and vSphere APIs Environment
	"""
	def get_old_vm_snpshts(self): #########################################################
		pass

		try:
			stdout.write(f"[INFO] : Collecting an Old VM Snapshots...")
			stdout.flush() # NOTE! Forces Text to Appear Immediately from Memory
		
			vi_content = self.vc_instance.RetrieveContent()
			if ((vi_content is not None) and (self.connstate)):
				container = vi_content.rootFolder # Root/Parent Directory 
																					# Starting Point
				view_type = [vim.VirtualMachine]  # Look for Object(s) : VM(s)
				recursive = True                  # Include SubDirs within Root/Parent Directory

				assert vi_content.viewManager is not None 
				container_view = vi_content.viewManager.CreateContainerView\
					(container, view_type, recursive)

				days_limit = 30 # Snapshot Age Threshold in Days
				
				# NOTE : Declare and Initialize MS Excel File
				xlsx_wb = Workbook()

				# NOTE : Create a Thematic Sheet and Set it as Active
				xlsx_wb.create_sheet(title="old_snapshots")
				active_xlsx_sheet = xlsx_wb["old_snapshots"] 

				# NOTE : Set the Column Headers of Active XLSX Sheet 
				column_header_list = ["VM Name", "Snapshot ID", "Snapshot Name", \
													"Creation Date", "Age in Days"]
				active_xlsx_sheet.append(column_header_list) # type: ignore
				xlsx_wb.save("report.xlsx")                  # type: ignore

				for vm in container_view.view:
					pass
					# Check If there are any Snapshots of a Particular VM
					if ((vm.snapshot) and (vm.snapshot.rootSnapshotList)):
						pass
						vm_old_snpshts_list = self.get_snpsht_data\
							(vm.snapshot.rootSnapshotList, days_limit)

						if (vm_old_snpshts_list): # Check If this List exists
							for snpsht in vm_old_snpshts_list:
								age_in_days = (datetime.now(timezone.utc) - snpsht['created']).days
								xlsx_data = (vm.name, snpsht["id"], snpsht["name"], \
										 snpsht["created"].strftime("%Y-%m-%d"), age_in_days)
								active_xlsx_sheet.append(xlsx_data) # type: ignore
								xlsx_wb.save("report.xlsx")         # type: ignore				

				container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server
				print(f"Done.")
				print("")
				sleep(2)

		except Exception as xex:
			print(f"Exception Occured : {xex}".upper())
