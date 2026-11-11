# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         		 vm_get_cnsldtn_ndd_stat.py 
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
from interaction_methods_set import InteractMethodsClass as IMClass # Apply Text Styles
from sys import stdout						 # Function for an Output Print Options
from dataclasses import dataclass  # Namespace for Operations with Data Classes
import time                        # Namespace for the Time Delaying and Sleep() Function
from time import perf_counter      # Function to Get Method's Runtime
from pyVmomi import vim            # Namespace for a Core Operations with VI Objects
import openpyxl                    # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'vm_get_cnsldtn_ndd_stat.py'
class GetVmCnsldStatClass(): ############################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance

	def get_vm_w_cnsld_ndd_func(self): ####################################################
		pass

		func_start_time = 0.0
		func_start_time = perf_counter() # Record the Start Time
		
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
				xlsx_wb = openpyxl.load_workbook("vi_micro_report.xlsx")

				# Create a Thematic Sheet and Set it as Active
				if (not ("vm_consolidation_needed" in xlsx_wb.sheetnames)):
					pass
					xlsx_wb.create_sheet(title = "vm_consolidation_needed")
					active_xlsx_sheet = xlsx_wb["vm_consolidation_needed"] 

					# Set the Column Headers of Active XLSX Sheet
					column_header_list = ["VM Name", "Consolidation Needed"]
					active_xlsx_sheet.append(column_header_list) 						
					xlsx_wb.save("vi_micro_report.xlsx")                  					 	
				else:
					active_xlsx_sheet = xlsx_wb["vm_consolidation_needed"] 

				for vm in container_view.view:
					pass
					# Avoid crashing if the Runtime Data is Temporarily Unavailable
					if (vm.runtime):
						pass
						cnsldtn_status = vm.runtime.consolidationNeeded # type: ignore

						# Print VM Names Only with this Status
						if (cnsldtn_status):						
							xlsx_data = (vm.name, str(cnsldtn_status))
							active_xlsx_sheet.append(xlsx_data) 					# type: ignore
							xlsx_wb.save("vi_micro_report.xlsx")

				imclass_instance = IMClass(xlsx_wb, active_xlsx_sheet, "vi_micro_report.xlsx")
				imclass_instance.add_text_style_func()        

				container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server

				func_stop_time = perf_counter() 									# Record the Stop Time
				func_exec_time = func_stop_time - func_start_time # Calculate the Difference

				print(f"Done. Execution Time is {func_exec_time:.3f} seconds")
				time.sleep(2)
				print("")

		except Exception as xex:
			print(f"Exception Occurred : {xex}")
