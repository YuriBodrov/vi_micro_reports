# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         		 vm_get_resources_limits.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : 'assert' Keyword is used to check whether a condition is True during 
program execution. If the condition is True, the program continues normally. 
If the condition is False, Python raises an AssertionError and stops the 
program.
---
Ensure that values satisfy required conditions before further processing.
---
Can be used inside a function to validate input values before performing a 
calculation. If the condition evaluates to False, the function stops and
raises an AssertionError.
---
Using the Colored Output. Example:
color_code_cyan     = "\033[36m"
color_code_reset    = "\033[0m"
print(f"{self.color_code_cyan}INFO : {self.color_code_reset}{vm.name}
---

"""

"""
NOTE : This Module is used for the VMware vSphere VMs Resource Limits 
Collection from VMware Infrastructure. This Module's Function returns 
VM Names List and Its Resource Limits.
"""
#########################################################################################
from interaction_methods_set import InteractMethodsClass as IMClass # Apply Text Styles
from sys import stdout						# Function for an Output Print Options
from dataclasses import dataclass # Namespace for Operations with Data Classes
import time                       # Namespace for the Time Delaying and Sleep() Function
from time import perf_counter     # Function to Get Method's Runtime
from pyVmomi import vim           # Namespace for a Core Operations with VI Objects
import openpyxl                   # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'vm_get_resources_limits.py'
class GetVmsResLimitClass(): ############################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance

	def get_vms_res_limit_func(self): #####################################################
		pass

		func_start_time = 0.0
		func_start_time = perf_counter() # Record the Start Time

		try:
			stdout.write("[INFO] : Collecting VM Resources Limits...")
			stdout.flush() # NOTE! Forces Text to appear immediately from Memory
		
			vi_content = self.vc_instance.RetrieveContent()
			if ((vi_content is not None) and (self.connstate)):
				container = vi_content.rootFolder # Root/Parent Directory 
																					# Starting Point
				view_type = [vim.VirtualMachine]  # Look for Object(s) : VM(s)
				recursive = True                  # Include SubDirs within Root/Parent Directory
				"""
				NOTE : Check this Out! 'Assert' Keyword. If 'assert' statement is absend, then
												function 'CreateContainerView' returns 'Null' for some reason.
				"""
				assert vi_content.viewManager is not None 
				container_view = vi_content.viewManager.CreateContainerView(container, \
																																view_type, recursive)
				
				# Open and Initialize the Existing(!) MS Excel File
				xlsx_wb = openpyxl.load_workbook("vi_micro_report.xlsx")

				# Create a Thematic Sheet and Set it as Active
				if (not ("vm_resources_limits" in xlsx_wb.sheetnames)):
					pass
					xlsx_wb.create_sheet(title = "vm_resources_limits")
					active_xlsx_sheet = xlsx_wb["vm_resources_limits"] 				 # type: ignore

					# Set the Column Headers of Active XLSX Sheet
					column_header_list = ["VM Name", "CPU Limit (MHz)", "CPU Reservation (MHz)", \
													"CPU Shares Level", "CPU Shares Value", "RAM Limit (MB)", \
														"RAM Reservation (MB)", "RAM Shares Level", \
															"RAM Shares Value", "RAM Locked to Max?"]
					active_xlsx_sheet.append(column_header_list) 							 # type: ignore
					xlsx_wb.save("vi_micro_report.xlsx")                  		 # type: ignore
				else:
					active_xlsx_sheet = xlsx_wb["vm_resources_limits"] 				 # type: ignore

				for vm in container_view.view:
					pass
					# Get VM's CPU Allocation Summary
					# Get VM's Memory Allocation Summary
					# Get VM's "All Guest Memory is Locked" Status
					vm_cpu_alloc = vm.config.cpuAllocation                     # type: ignore
					vm_ram_alloc = vm.config.memoryAllocation                  # type: ignore
					vm_ram_are_locked = vm.config.memoryReservationLockedToMax # type: ignore
					
					# If we have some CPU/RAM Limits (.limit = -1 : "Unlimited")
					if ((vm_cpu_alloc.limit != -1) or (vm_ram_alloc.limit != -1)):
						pass
						# CPU Limits and Shares Data:
						vm_cpu_limit      = vm_cpu_alloc.limit   
						vm_cpu_rsrvtn     = vm_cpu_alloc.reservation  
						vm_cpu_shares_lvl = vm_cpu_alloc.shares.level  
						vm_cpu_shares_val = vm_cpu_alloc.shares.shares

						# Memory Limits and Shares Data:
						vm_ram_limit      = vm_ram_alloc.limit 
						vm_ram_rsrvtn     = vm_ram_alloc.reservation 
						vm_ram_shares_lvl = vm_ram_alloc.shares.level
						vm_ram_shares_val = vm_ram_alloc.shares.shares

						xlsx_data = (vm.name, vm_cpu_limit, vm_cpu_rsrvtn, vm_cpu_shares_lvl, \
									 vm_cpu_shares_val, vm_ram_limit, vm_ram_rsrvtn, vm_ram_shares_lvl, \
										vm_ram_shares_val, vm_ram_are_locked)
						active_xlsx_sheet.append(xlsx_data) 										 # type: ignore
						xlsx_wb.save("vi_micro_report.xlsx") 										 # type: ignore

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
