# Project Name  : 										Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         									  raascore.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is so-called 'Core Module'. We can run all Modules from here. 
														 We will use a Parallel Tasks Processing here later.
"""
# TODO : Create more Detailed Description!

# NOTE : Limit all lines to a Maximum of 90 characters (Best Practices!)
#########################################################################################
import sys
from vcsa_connect_a          import VcConnectionSteps as VcConnSteps
from clu_get_ha_dis_items_a  import GetClustersHADisabledClass as GetDisHAClu
from esxi_ha_agent_dis_a     import GetESXisHAAgentDisabledClass as GetDisHAESXi
from esxi_w_single_disk_a    import GetESXiWithSingleDskClass as SingleDskESXi
from esxi_ntp_settings_a     import GetEsxiNtpSettings as GetNtpAnomEsxi
from vm_w_old_snpshts_a      import GetOldVmSnpshts as GetVmOldSnapshots
from vm_get_res_limits_a     import GetVmsResLimitClass as GetVmResLimit
from vm_swap_balloon_val_a   import GetVmsSwapBalClass as GetVmsSwapBal
from vm_ha_mon_action_evt_a  import GetVmsHaTrigAlrmClass as GetTrigAlarms
from vm_cnsldtn_ndd_stat_a   import GetVmCnsldStatClass as GetCnsldtnStatus
from vm_get_vmtools_status_a import GetVmToolsStatClass as GetVmToolsStat

# Module for Prompting a Password with Echo Turned Off
from getpass import getpass 

def main(): #############################################################################
	pass

	print("")
	input_vcaddr  = input("vCenter Server : ")
	input_vclogin = input("Username       : ")
	# Get the Password Value from Keyboard
	input_vcpassw = getpass("Password       : ", echo_char="*") 

	vc_con_step_instance = VcConnSteps(input_vcaddr, input_vclogin, input_vcpassw)
	conn_state = vc_con_step_instance.verify_vc_connect_func()
	vc_service = vc_con_step_instance.primary_connect_param_func()
     
	assert vc_service is not None

	"""
	# 01 : Get an Old VMs Snapshots +
	vm_old_snapshots = GetVmOldSnapshots(conn_state, vc_service)
	vm_old_snapshots.get_old_vm_snpshts()
	
	# 02 : Get the ESXi Server's NTP Settings Anomalies +
	get_esxi_ntp_set = GetNtpAnomEsxi(conn_state, vc_service)
	get_esxi_ntp_set.get_esxi_ntp_set_func()
	
	# 03 : Get the List of ESXi Servers with a Single Datastore +
	get_single_dsk_esxi = SingleDskESXi(conn_state, vc_service)
	get_single_dsk_esxi.get_esxi_w_sngl_dsk_func()

	# 04 : Get HA Disabled Clusters List +
	get_dis_ha_cl_instance = GetDisHAClu(conn_state, vc_service)
	get_dis_ha_cl_instance.get_ha_disabled_clusters_func()

	# 05 : Get the List of ESXi Servers with a Disabled HA Agent +
	get_dis_ha_esxi_instance = GetDisHAESXi(conn_state, vc_service)
	get_dis_ha_esxi_instance.get_esxi_ha_agent_dis_func()

	# 06 : Get VM Resource Limits (CPU/RAM) +
	get_vm_res_limit_instance = GetVmResLimit(conn_state, vc_service)
	get_vm_res_limit_instance.get_vms_res_limit_func()

	# 07 : Get VM Performance Stats : Swap, Balloon Data +
	get_vms_swap_bal_instance = GetVmsSwapBal(conn_state, vc_service)
	get_vms_swap_bal_instance.get_vm_swap_bal_func()

	# 08 : Get vSphere HA VM Monitoring Action +
	get_trig_alarms_instance = GetTrigAlarms(conn_state, vc_service)
	get_trig_alarms_instance.get_vm_trig_alrm_func()

	# 09 : Get VMs with Snapshot's Consolidation Needed Status +
	get_cnsldtn_ndd_instance = GetCnsldtnStatus(conn_state, vc_service)
	get_cnsldtn_ndd_instance.get_vm_w_cnsld_ndd_func()
	"""
	
  # 10 : Get VMware Tools Status +
	get_vm_tools_stat_instance = GetVmToolsStat(conn_state, vc_service)
	get_vm_tools_stat_instance.get_vm_tools_stat_func()
     
	"""
	# Get the ESXi Server's Multipath Round-Robin Limit
	get_esxi_mp_rr_set = GetMpRrLimitSet(conn_state, vc_service)
	get_esxi_mp_rr_set.get_esxi_mp_rr_lim_func()

	print("")
	"""

	"""
	# Get the ESXi Server's Boot Device Errors
	esxi_boot_dev_err_instance = GetEsxiBootErrors(conn_state, vc_service)
	esxi_boot_dev_err_instance.esxi_get_boot_dev_err_func()
	print("")
	"""
  
if __name__ == "__main__":
  main()
	