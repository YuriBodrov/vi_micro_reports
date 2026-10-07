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

"""
from pyVim.connect import Disconnect, SmartConnectNoSSL
from pyVmomi import vim


def get_datastores_per_host(content):
  # Create a container view for all HostSystem objects
  container = content.viewManager.CreateContainerView(
      content.rootFolder, [vim.HostSystem], True
  )

  host_datastores = {}

  # Iterate through each ESXi host in the view
  for host in container.view:
    host_name = host.name
    # Each host object has a 'datastore' attribute listing attached datastores
    datastores = [ds.name for ds in host.datastore]
    host_datastores[host_name] = datastores

    print(f'ESXi Host: {host_name}')
    for ds in host.datastore:
      print(f'  - Datastore: {ds.name} (Capacity: {ds.summary.capacity / (1024**3):.2f} GB)')

  container.Destroy()
  return host_datastores


def main():
  # Connect to vCenter or standalone ESXi host
  si = SmartConnectNoSSL(
      host='vcenter.yourdomain.local',
      user='administrator@vsphere.local',
      pwd='your_password',
      port=443,
  )

  try:
    content = si.RetrieveContent()
    get_datastores_per_host(content)
  finally:
    Disconnect(si)


if __name__ == '__main__':
  main()
```

### Key Concepts
* **`vim.HostSystem`**: Represents an ESXi host in the vSphere inventory.
* **`host.datastore`**: Returns a list of `vim.Datastore` managed object 
references directly mounted or accessible by that specific ESXi host.
* **`host.summary.capacity`**: Provides total storage size properties for 
individual datastore objects if you need disk utilization stats.

"""
#########################################################################################
#from pyVim.connect import SmartConnect, Disconnect
from sys import stdout            # Function for an Output Print Options
from dataclasses import dataclass # Function for Operations with Data Classes
import time                       # Namespace for the Time Delaying and Sleep() Function
from pyVmomi import vim           # Namespace for a Core Operations with VI Objects
import openpyxl                   # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module 'esxi_get_items_with_sngl_dsk.py'
class GetESXiWithSingleDskClass(): ###################################################@@@
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	#esxi_w_sngl_dks_list = []

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

			#print(f"[INFO] : Collecting Single-Disk ESXi Servers...")
			#print("-" * 47)
			stdout.write(f"[INFO] : Collecting Single-Disk ESXi Servers...")
			stdout.flush()
			for host in container_view.view:
				# NOTE! len(host.datastore) : Count ESXi Server's Datastores
				for ds in host.datastore: # type: ignore
					#print(f"  - Datastore: {ds.name} (Capacity: {ds.summary.capacity / \
					#(1024**3):.2f} GB) | Total : {len(host.datastore)}") # type: ignore
					if (len(host.datastore) <= 1): # type: ignore
						pass
						#self.esxi_w_sngl_dks_list.append(host.name)
						capacity = round(ds.summary.capacity / (1024**3), 1)
						xlsx_data = (host.name, ds.name, capacity)
						active_xlsx_sheet.append(xlsx_data) # type: ignore
						xlsx_wb.save("report.xlsx")         # type: ignore

			container_view.Destroy() # to Avoid Memory Accumulation in the vCenter Server
			print("Done.")
			time.sleep(2)
			print("")

