# Project Name  : 										Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         				 vm_swap_balloon_val.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE
-----------------------------------------------------------------------------------------
if vm.runtime.consolidationNeeded:
    print(f"Triggering disk consolidation for {vm.name}...")
    task = vm.ConsolidateVMDisks_Task()
"""

"""
NOTE : This Module is used to collect VMs with Consolidation Needed Status 
"""
#########################################################################################
from sys import stdout						 # Function for an Output Print Options
from dataclasses import dataclass  # Namespace for Operations with Data Classes
import time                        # Namespace for the Time Delaying and Sleep() Function
from pyVmomi import vim            # Namespace for a Core Operations with VI Objects
import openpyxl                    # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'vm_swap_balloon_val.py'
class GetVmCnsldStatClass(): ############################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance

	def get_vm_w_cnsld_ndd_func(self): ####################################################
		pass

		try:
			stdout.write("[INFO] : Collecting VMs With Consolidation Needed Status...")
			stdout.flush()
		
			vi_content = self.vc_instance.RetrieveContent()
			if ((vi_content is not None) and (self.connstate)):
				container = vi_content.rootFolder # Root/Parent Directory's Starting Point
				view_type = [vim.VirtualMachine]  # Look for Object(s) : VM(s)
				recursive = True                  # Include SubDirs within Root/Parent Directory
				
				assert vi_content.viewManager is not None 
				container_view = vi_content.viewManager.CreateContainerView(container, \
																																view_type, recursive)
				
				# Open and Initialize the Existing(!) MS Excel File
				xlsx_wb = openpyxl.load_workbook("report.xlsx")

				# Create a Thematic Sheet and Set it as Active
				if (not ("vms_with_cnsldn_ndd_stat" in xlsx_wb.sheetnames)):
					pass
					xlsx_wb.create_sheet(title = "vms_with_cnsldn_ndd_stat")
					active_xlsx_sheet = xlsx_wb["vms_with_cnsldn_ndd_stat"] 

					# Set the Column Headers of Active XLSX Sheet
					column_header_list = ["VM Name", "Consolidation Needed"]
					active_xlsx_sheet.append(column_header_list) 						
					xlsx_wb.save("report.xlsx")                  					 	
				else:
					active_xlsx_sheet = xlsx_wb["vms_with_cnsldn_ndd_stat"] 

				for vm in container_view.view:
					pass
					# Avoid crashing if the Runtime Data is Temporarily Unavailable
					if (vm.runtime):
						pass
						cnsldtn_status = vm.runtime.consolidationNeeded

						# Print VM Names Only with this Status
						if (cnsldtn_status):						
							xlsx_data = (vm.name, str(cnsldtn_status))
							active_xlsx_sheet.append(xlsx_data) # type: ignore
							xlsx_wb.save("report.xlsx")         

				container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server
				print("Done.")
				time.sleep(2)
				print("")

		except Exception as xex:
			print(f"Exception Occurred : {xex}")
