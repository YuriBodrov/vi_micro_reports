# Project Name  :  Virtual Infrastructure's Micro Reports  
# ------------------------------------------------------- 
# Module Name   :         								 main_module.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is so-called 'Core Module'. We can run all Modules from here. 
														 We will use a Parallel Processing Tasks here later.
"""

"""
C:\\Users\\bodrovyp\\AppData\\Roaming\\Python\\Python314\\Scripts\\pyinstaller.exe
"""

# TODO : Create more Detailed Description!

# NOTE : Limit all lines to a Maximum of 90 characters (Best Practices!)
#########################################################################################
import sys
import keyring
from vcsa_connect_module    				import VcConnectionSteps as VcConnSteps
from clu_get_ha_dsbld_items 				import GetClustersHADisabledClass as GetDisHAClu
from esxi_get_ha_agnt_dsbld_items 	import GetESXisHAAgentDisabledClass as GetDisHAESXi
from esxi_get_items_with_sngl_dsk 	import GetESXiWithSingleDskClass as SingleDskESXi
from esxi_get_ntp_anomalies     		import GetEsxiNtpSettings as GetNtpAnomEsxi
from vm_get_Items_wth_old_snpshts 	import GetOldVmSnpshts as GetVmOldSnapshots
from vm_get_resources_limits      	import GetVmsResLimitClass as GetVmResLimit
from vm_get_swap_balloon_val    		import GetVmsSwapBalClass as GetVmsSwapBal
from vm_get_ha_mntr_actn_evnts 			import GetVmsHaTrigAlrmClass as GetTrigAlarms
from vm_get_cnsldtn_ndd_stat    		import GetVmCnsldStatClass as GetCnsldtnStatus
from vm_get_vm_tools_status  				import GetVmToolsStatClass as GetVmToolsStat
from email_sender_module						import EmailSenderClass as EmailSender

# Module for Prompting a Password with Echo Turned Off
from getpass import getpass 

def main(): #############################################################################
	pass

	print("")
	
	print(f"Project Name    :  Virtual Infrastructure's Micro Reports")
	print(f"---------------------------------------------------------")
	print(f"Created by      :                          Yuri P. Bodrov")
	print(f"Email           :                    bodrovyp@hotmail.com")
	print(f"Phone Number    :                            +79259929596")
	print(f"Telegram        :                             @YuriBodrov")
	print(f"LinkedIn Page   : https://www.linkedin.com/in/yuribodrov/")
	print(f"---------------------------------------------------------")
	print(f"")

	#input_vcaddr  = input("vCenter Server  : ")
	#input_vclogin = input("Username        : ")
	# Get the Password Value from Keyboard
	#input_vcpassw = getpass("Password        : ", echo_char="*")
	
	creds = keyring.get_credential("vmc-01-new", None)
	input_vcaddr = "vmc-01-new.vesta.ru"
	input_vc_login = creds.username # type: ignore
	input_vc_passw = creds.password # type: ignore
	
	vc_con_step_instance = VcConnSteps(input_vcaddr, input_vc_login, input_vc_passw)
	conn_state = vc_con_step_instance.verify_vc_connect_func()
	vc_service = vc_con_step_instance.primary_connect_param_func()
     
	assert vc_service is not None

	
	# 01 : Get an Old (Older than 30 Days) VMs Snapshots +
	vm_old_snapshots = GetVmOldSnapshots(conn_state, vc_service)
	vm_old_snapshots.get_old_vm_snpshts()
	
	
	# 02 : Get the ESXi Server's NTP Settings Anomalies +
	get_esxi_ntp_set = GetNtpAnomEsxi(conn_state, vc_service)
	get_esxi_ntp_set.get_esxi_ntp_set_func()
	
	
	# 03 : Get the List of ESXi Servers with a Single Datastore +
	get_single_dsk_esxi = SingleDskESXi(conn_state, vc_service)
	get_single_dsk_esxi.get_esxi_w_sngl_dsk_func()

			
	# 04 : Get the List of ESXi Servers with a Disabled HA Agent +
	get_dis_ha_esxi_instance = GetDisHAESXi(conn_state, vc_service)
	get_dis_ha_esxi_instance.get_esxi_ha_agent_dis_func()


	# 05 : Get HA Disabled Clusters List +
	get_dis_ha_cl_instance = GetDisHAClu(conn_state, vc_service)
	get_dis_ha_cl_instance.get_ha_disabled_clusters_func()
	
	
	# 06 : Get VM Resource Limits (CPU/RAM) +
	get_vm_res_limit_instance = GetVmResLimit(conn_state, vc_service)
	get_vm_res_limit_instance.get_vms_res_limit_func()

	
	# 07 : Get VM Performance Stats : Swap, Balloon Data +
	get_vms_swap_bal_instance = GetVmsSwapBal(conn_state, vc_service)
	get_vms_swap_bal_instance.get_vm_swap_bal_func()

	# 08 : Get vSphere HA VM Monitoring Action Errors +
	get_trig_alarms_instance = GetTrigAlarms(conn_state, vc_service)
	get_trig_alarms_instance.get_vm_trig_alrm_func()

	# 09 : Get VMs with Snapshot's Consolidation Needed Status +
	get_cnsldtn_ndd_instance = GetCnsldtnStatus(conn_state, vc_service)
	get_cnsldtn_ndd_instance.get_vm_w_cnsld_ndd_func()
	
  # 10 : Get VMware Tools Status +
	get_vm_tools_stat_instance = GetVmToolsStat(conn_state, vc_service)
	get_vm_tools_stat_instance.get_vm_tools_stat_func()
	

	email_recpnts = ["bodrovyp@alfastrah.ru"]
	email_sender_instance = EmailSender()
	email_sender_instance.email_sndr_func("relay.alfastrah.ru", 25, email_recpnts, \
                                        "vi_micro_report.xlsx")

	print(f"")
	input("Press Any Key to Quit...")
  
  
if __name__ == "__main__":
  main()
	