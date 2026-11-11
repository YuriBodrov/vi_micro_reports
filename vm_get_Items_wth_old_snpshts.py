# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         vm_get_Items_wth_old_snpshts.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is used to retrieve an Old VMs Snapshots.
			 By default, Older than 30 Days.
"""
#########################################################################################
from interaction_methods_set import InteractMethodsClass as IMClass # Apply Text Styles
from sys import stdout						                 # Function for an Output Print Options
from datetime import datetime, timedelta, timezone # For Datetime Operations
from dataclasses import dataclass                  # Operations with Data Classes
from time import sleep                             # Time Delaying and Sleep() Function
from time import perf_counter      								 # Function to Get Method's Runtime
from pyVmomi import vim                            # Core Operations with VI Objects
from openpyxl import Workbook                      # For an Operations with MS Excel File
from openpyxl.styles import Font									 # This Class provides All 
																									 # Font Customization Types
#from openpyxl.styles import DEFAULT_FONT					 # This Class sets the Default Font for
																									 # all Worksheets within Workbook

@dataclass # Main Class of this Module 'vm_get_Items_wth_old_snpshts.py'
class GetOldVmSnpshts(): ################################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	
	"""
	NOTE : This Function Recursively figuring the Snapshot Tree and returns Snapshots 
																							Older than a Specified Number of Days
	"""

	#func_start_time = 0.0
	#func_start_time = perf_counter() # Record the Start Time
	"""
	#func_start_time = datetime.now().strftime("%H:%M:%S.%f")[:-3] # Record the Start Time
	#func_start_time = datetime.now().strftime("%H:%M:%S.%f")[:-3] # Record the Start Time
	# Record the Start Time
	start_date_string = str(datetime.now())
	func_start_time = datetime.strptime(start_date_string, "%H:%M:%S.%f").time()
	"""
	
	def get_snpsht_data(self, snpsht_lst, days_gone): # type: ignore #######################
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

		func_start_time = 0.0
		func_start_time = perf_counter() # Record the Start Time

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
				container_view = vi_content.viewManager.CreateContainerView(container, \
				                                                            view_type, recursive)

				days_limit = 30 # Snapshot Age Threshold in Days
								
				# NOTE : Declare and Initialize MS Excel File
				xlsx_wb = Workbook()
				
				# NOTE : Create a Thematic Sheet and Set it as Active
				xlsx_wb.create_sheet(title="vm_old_snapshots")
				active_xlsx_sheet = xlsx_wb["vm_old_snapshots"]
				
				# NOTE : Set the Column Headers of Active XLSX Sheet 
				column_header_list = ["VM Name", "Snapshot ID", "Snapshot Name", \
													"Creation Date", "Age in Days"]
				active_xlsx_sheet.append(column_header_list) 					 # type: ignore
				xlsx_wb.save("vi_micro_report.xlsx")         					 # type: ignore

				for vm in container_view.view:
					pass
					# Check If there are any Snapshots of a Particular VM
					if (vm.snapshot) and (vm.snapshot.rootSnapshotList): # type: ignore
						pass

						"""
						# TODO : Add some Test Logic:
						total_snapshot_size = 0
						if (vm.layoutEx and vm.layoutEx.file):						 # type: ignore
							pass
							for file_info in vm.layoutEx.file:							 # type: ignore
								pass
								is_snapshot_file = (
									file_info.type == "snapshotData" or "delta" in file_info.name.lower() \
									or any(f"00000{i}" in file_info.name for i in range(1, 10))
								)
								if is_snapshot_file:
									pass
									total_snapshot_size += file_info.size
									print(f"vm_name             : {vm.name}")
									print(f"total_snapshot_size : {total_snapshot_size}")
									print("---")
						"""
						
						vm_old_snpshts_list = self.get_snpsht_data\
							(vm.snapshot.rootSnapshotList, days_limit)			 # type: ignore

						if (vm_old_snpshts_list): # Check If this List exists
							for snpsht in vm_old_snpshts_list:
								age_in_days = (datetime.now(timezone.utc) - snpsht['created']).days
								xlsx_data = (vm.name, snpsht["id"], snpsht["name"], \
										 snpsht["created"].strftime("%Y-%m-%d"), age_in_days)
								active_xlsx_sheet.append(xlsx_data)  					 # type: ignore
								xlsx_wb.save("vi_micro_report.xlsx") 					 # type: ignore

				imclass_instance = IMClass(xlsx_wb, active_xlsx_sheet, "vi_micro_report.xlsx")
				imclass_instance.add_text_style_func()

				container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server
								
			func_stop_time = perf_counter() 											 # Record the Stop Time
			func_exec_time = func_stop_time - func_start_time # Calculate the Difference   

			print(f"Done. Execution Time is {func_exec_time:.3f} seconds")
			print("")
				
			sleep(2)

		except Exception as xex:
			print(f"Exception Occured : {xex}".upper())
