# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         		 vm_get_swap_balloon_val.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE
-----------------------------------------------------------------------------------------
"stats.balloonedMemory" : Displays the Amount of Memory (in MB) that the Hypervisor has 
													"reclaimed" from the guest OS using the Ballooning Technique 
													(via the vmmemctl driver)"

"stats.swappedMemory"   : Displays the Amount of Memory (in MB) that the Hypervisor 
													itself has forcibly Swapped Out to the ".vswp" File on the 
													Datastore. This occurs when the Ballooning Mechanism is 
													insufficient or when VMware Tools is not installed inside 
													the VM. Host swapping severely degrades VM performance!

"vm.summary.quickStats" : This is an object in the VMware vSphere API 
													(vim.vm.Summary.QuickStats) that contains a set of basic, 
													Real-time Statistics about a VM.
"""

"""
NOTE : This Module is used to collect these Metrics : VM's Swap and Ballooning Values. 
"""
#########################################################################################
from interaction_methods_set import InteractMethodsClass as IMClass # Apply Text Styles
from sys import stdout						 # Function for an Output Print Options
from dataclasses import dataclass  # Namespace for Operations with Data Classes
import time                        # Namespace for the Time Delaying and Sleep() Function
from time import perf_counter      # Function to Get Method's Runtime
from pyVmomi import vim            # Namespace for a Core Operations with VI Objects
import openpyxl                    # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'vm_get_swap_balloon_val.py'
class GetVmsSwapBalClass(): #############################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance

	def get_vm_swap_bal_func(self): #######################################################
		pass

		func_start_time = 0.0
		func_start_time = perf_counter() # Record the Start Time		
		
		try:
			stdout.write("[INFO] : Collecting VM's Swap and Ballooning Values...")
			stdout.flush()
		
			vi_content = self.vc_instance.RetrieveContent()
			if ((vi_content is not None) and (self.connstate)):
				container = vi_content.rootFolder # Root/Parent Directory 
																					# Starting Point
				view_type = [vim.VirtualMachine]  # Look for Object(s) : VM(s)
				recursive = True                  # Include SubDirs within Root/Parent Directory
				
				assert vi_content.viewManager is not None 
				container_view = vi_content.viewManager.CreateContainerView(container, \
																																view_type, recursive)
				
				# Open and Initialize the Existing(!) MS Excel File
				xlsx_wb = openpyxl.load_workbook("vi_micro_report.xlsx")

				# Create a Thematic Sheet and Set it as Active
				if (not ("vm_swap_balloon_values" in xlsx_wb.sheetnames)):
					pass
					xlsx_wb.create_sheet(title = "vm_swap_balloon_values")
					active_xlsx_sheet = xlsx_wb["vm_swap_balloon_values"] 		# type: ignore

					# Set the Column Headers of Active XLSX Sheet
					column_header_list = ["VM Name", "Swap Value (MB)", "Balloon Value (MB)"]
					active_xlsx_sheet.append(column_header_list) 							# type: ignore
					xlsx_wb.save("vi_micro_report.xlsx")                  		# type: ignore
				else:
					active_xlsx_sheet = xlsx_wb["vm_swap_balloon_values"] 		# type: ignore

				for vm in container_view.view:
					pass

					q_stats         = vm.summary.quickStats 									# type: ignore
					q_stats_swap    = q_stats.swappedMemory
					q_stats_balloon = q_stats.balloonedMemory 

					if ((q_stats_swap != 0) or (q_stats_balloon != 0)):
						pass

						xlsx_data = (vm.name, q_stats_swap, q_stats_balloon)
						active_xlsx_sheet.append(xlsx_data) 										# type: ignore
						xlsx_wb.save("vi_micro_report.xlsx")         						# type: ignore

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
