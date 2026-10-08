# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         esxi_get_items_with_sngl_dsk.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is designed for the Single-Disk ESXi Hypervisors Collection from 
			 																											 VMware Infrastructure.
"""

#########################################################################################
from interaction_methods_set import InteractMethodsClass as IMClass # Apply Text Styles
from sys import stdout            # Function for an Output Print Options
from dataclasses import dataclass # Function for Operations with Data Classes
import time                       # Namespace for the Time Delaying and Sleep() Function
from pyVmomi import vim           # Namespace for a Core Operations with VI Objects
import openpyxl                   # For an Operations with MS Excel File(s)
from time import perf_counter     # Function to Get Method's Runtime

@dataclass # Main Class of this Module 'esxi_get_items_with_sngl_dsk.py'
class GetESXiWithSingleDskClass(): ######################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	#esxi_w_sngl_dks_list = []

	func_start_time = perf_counter() # Record the Start Time
	def get_esxi_w_sngl_dsk_func(self): ###################################################
		pass
		vi_content = self.vc_instance.RetrieveContent()
		if ((vi_content is not None) and (self.connstate)):
			container = vi_content.rootFolder # Root/Parent Directory Starting Point
			view_type = [vim.HostSystem]      # Look for Object(s) : Host(s)
			recursive = True                  # Include SubDirs within Root/Parent Directory

			"""
			NOTE : Check this Out! 'Assert' Keyword. If 'assert' statement is absend, then 
			function 'CreateContainerView' may return 'Null' for some reason.
			"""
			assert vi_content.viewManager is not None 
			container_view = vi_content.viewManager.CreateContainerView(container, view_type, \
																															 recursive)
		
			# Open and Initialize the Existing(!) MS Excel File
			xlsx_wb = openpyxl.load_workbook("report.xlsx")

			# Create a Thematic Sheet and Set it as Active
			if (not ("esxi_with_single_disk" in xlsx_wb.sheetnames)):
				xlsx_wb.create_sheet(title="esxi_with_single_disk")
				active_xlsx_sheet = xlsx_wb["esxi_with_single_disk"] # type: ignore

				# Set the Column Headers of Active XLSX Sheet
				column_header_list = ["ESXi Server", "Datastore Name", "Datastore Size (GB)"]
				active_xlsx_sheet.append(column_header_list) 				 # type: ignore
				xlsx_wb.save("report.xlsx")                  				 # type: ignore
			else:
				active_xlsx_sheet = xlsx_wb["esxi_with_single_disk"] # type: ignore

			stdout.write(f"[INFO] : Collecting Single-Disk ESXi Servers...")
			stdout.flush()
			for host in container_view.view:
				# NOTE! len(host.datastore) : Count ESXi Server's Datastores
				for ds in host.datastore: 													 # type: ignore
					if (len(host.datastore) <= 1): 										 # type: ignore
						pass
						capacity = round(ds.summary.capacity / (1024**3), 1)
						xlsx_data = (host.name, ds.name, capacity)
						active_xlsx_sheet.append(xlsx_data) 						 # type: ignore
						xlsx_wb.save("report.xlsx")         						 # type: ignore

			imclass_instance = IMClass(xlsx_wb, active_xlsx_sheet, "report.xlsx")
			imclass_instance.add_text_style_func()
			
			container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server

			func_stop_time = perf_counter() 											 # Record the Start Time
			func_exec_time = func_stop_time - self.func_start_time # Calculate the Difference

			print(f"Done. Execution Time is {func_exec_time:.3f} seconds")
			time.sleep(2)
			print("")

