# Project Name  : 										Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         					 vm_get_res_limits.py 
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
from sys import stdout						# Function for an Output Print Options
from dataclasses import dataclass # Namespace for Operations with Data Classes
import time                       # Namespace for the Time Delaying and Sleep() Function
from pyVmomi import vim           # Namespace for a Core Operations with VI Objects
import openpyxl                   # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'vm_get_res_limits.py'
class GetVmsResLimitClass(): ############################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	#vm_collection_list = []

	def get_vms_res_limit_func(self): #####################################################
		pass

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
				xlsx_wb = openpyxl.load_workbook("report.xlsx")

				# Create a Thematic Sheet and Set it as Active
				if (not ("vm_resources_limits" in xlsx_wb.sheetnames)):
					pass
					xlsx_wb.create_sheet(title = "vm_resources_limits")
					active_xlsx_sheet = xlsx_wb["vm_resources_limits"] # type: ignore

					# Set the Column Headers of Active XLSX Sheet
					column_header_list = ["VM Name", "CPU Limit (MHz)", "CPU Reservation (MHz)", \
													"CPU Shares Level", "CPU Shares Value", "RAM Limit (MB)", \
														"RAM Reservation (MB)", "RAM Shares Level", \
															"RAM Shares Value", "RAM Locked to Max?"]
					active_xlsx_sheet.append(column_header_list) # type: ignore
					xlsx_wb.save("report.xlsx")                  # type: ignore
				else:
					active_xlsx_sheet = xlsx_wb["vm_resources_limits"] # type: ignore

				for vm in container_view.view:
					pass
					"""
					# Set Initial Values for these Variables
					vm_cpu_limit      = "Not set"
					vm_cpu_rsrvtn     = "Not set"
					vm_cpu_shares_lvl = "Not set"
					vm_cpu_shares_val = "Not set"
					#vm_ram_are_locked = "Not set"
					vm_ram_limit      = "Not set"
					vm_ram_rsrvtn     = "Not set"
					vm_ram_shares_lvl = "Not set"
					vm_ram_shares_val = "Not set"
					"""
					
					vm_cpu_alloc = vm.config.cpuAllocation                     # Get VM's CPU 
																																		 # Allocation Summary
					vm_ram_alloc = vm.config.memoryAllocation                  # Get VM's Memory
																																		 # Allocation Summary
					vm_ram_are_locked = vm.config.memoryReservationLockedToMax # If Memory Locked?
					# If we have some CPU/RAM Limits (.limit = -1 : "Unlimited")
					if ((vm_cpu_alloc.limit != -1) or (vm_ram_alloc.limit != -1)):
					#if (vm_ram_alloc.limit != -1):  
						pass
						vm_cpu_limit      = vm_cpu_alloc.limit         #
						vm_cpu_rsrvtn     = vm_cpu_alloc.reservation   # CPU Limits
						vm_cpu_shares_lvl = vm_cpu_alloc.shares.level  # and Shares Data
						vm_cpu_shares_val = vm_cpu_alloc.shares.shares #

						vm_ram_limit      = vm_ram_alloc.limit         # 
						vm_ram_rsrvtn     = vm_ram_alloc.reservation   # Memory Limits 
						vm_ram_shares_lvl = vm_ram_alloc.shares.level  # and Shares data
						vm_ram_shares_val = vm_ram_alloc.shares.shares # 

						xlsx_data = (vm.name, vm_cpu_limit, vm_cpu_rsrvtn, vm_cpu_shares_lvl, \
									 vm_cpu_shares_val, vm_ram_limit, vm_ram_rsrvtn, vm_ram_shares_lvl, \
										vm_ram_shares_val, vm_ram_are_locked)
						active_xlsx_sheet.append(xlsx_data) # type: ignore
						xlsx_wb.save("report.xlsx")         # type: ignore

						"""
						print(f"vm_name : {vm.name}")
						print(f"---")
						print(f"cpu_limit_value   : {vm_cpu_limit}")
						print(f"cpu_rsrvtn_value  : {vm_cpu_rsrvtn}")
						print(f"cpu_shares_lvl    : {vm_cpu_shares_lvl}")
						print(f"cpu_shares_value  : {vm_cpu_shares_val}")
						print(f"---")
						print(f"vm_ram_are_locked : {vm_ram_are_locked}")
						print(f"ram_limit_value   : {vm_ram_limit}")
						print(f"ram_rsrvtn_value  : {vm_ram_rsrvtn}")
						print(f"ram_shares_lvl    : {vm_ram_shares_lvl}")
						print(f"ram_shares_val    : {vm_ram_shares_val}")
						print("")
						"""

				container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server
				print("Done.")
				time.sleep(2)
				print("")

				#return self.vm_collection_list

		except Exception as xex:
			#print(f"Exception Occurred : {xex}".upper())
			print(f"Exception Occurred : {xex}")
