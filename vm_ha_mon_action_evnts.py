# Project Name  : 										Report-as-a-Service  
# ------------------------------------------------------- 
# Module Name   :         				vm_ha_mon_action_evt.py 
# Created by    :       									 Yuri P. Bodrov 
# Email         : 									 bodrovyp@hotmail.com 
# Phone Number  :         									 +79259929596 
# Telegram      :          										@YuriBodrov 
# LinkedIn Page : https://www.linkedin.com/in/yuribodrov/

"""
NOTE : This Module is used to Collect VMs with a Triggered Alarms about 
			 HA Monitoring Action Events.
"""

"""
TODO : CONSPECT THIS
-----------------------------------------------------------------------------------------
import ssl
from pyVim.connect import SmartConnectNoSSL, Disconnect
from pyVmomi import vim
import atexit

# Настройки подключения
VCENTER_IP = "vcenter.yourdomain.local"
USER = "administrator@vsphere.local"
PASSWORD = "YourPasswordHere"

def get_ha_monitoring_events():
    # Подключение к vCenter без проверки SSL-сертификата
    si = SmartConnectNoSSL(host=VCENTER_IP, user=USER, pwd=PASSWORD)
    atexit.register(Disconnect, si)

    content = si.RetrieveContent()
    event_manager = content.eventManager

    # Создаем фильтр для поиска событий
    filter_spec = vim.event.EventFilterSpec()
    
    # Интересующие типы событий vSphere HA (Fault Tolerance / VM Monitoring)
    # Например: VmBeingResetEvent (когда HA перезагружает зависшую ВМ)
    filter_spec.eventTypeId = [
        "VmBeingResetEvent", 
        "VmFailoverFailedEvent",
        "com.vmware.vc.ha.VmMonitoringStateChangedEvent"
    ]

    # Создаем коллектор событий (выбираем, например, последние 100 событий)
    event_collector = event_manager.CreateCollectorForEvents(filter_spec)
    atexit.register(event_collector.DestroyCollector)

    # Читаем события с конца
    events = event_collector.ReadNextEvents(maxCount=100)

    if not events:
        print("Событий по мониторингу vSphere HA VM не найдено.")
        return

    print(f"Найдено событий: {len(events)}\n")
    for event in events:
        print(f"Время: {event.createdTime}")
        print(f"ВМ: {event.vm.name if event.vm else 'Н/Д'}")
        print(f"Тип события: {type(event).__name__}")
        print(f"Сообщение: {event.fullFormattedMessage}")
        
        # Если это событие сброса ВМ, vSphere HA часто сохраняет путь к скриншоту BSOD/Panic
        if hasattr(event, 'screenshot') and event.screenshot:
            print(f"Скриншот экрана ВМ сохранен в: {event.screenshot}")
            
        print("-" * 50)

if __name__ == "__main__":
    get_ha_monitoring_events()
-----------------------------------------------------------------------------------------
from datetime import datetime
now = datetime.now()
# %f возвращает 6 цифр (микросекунды), срезом оставляем первые 3 цифры
time_str = now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
print(time_str)  # Например: '2026-09-29 09:44:15.123'
-----------------------------------------------------------------------------------------
from datetime import datetime
# Автоматически округляет/обрезает до миллисекунд
iso_str = datetime.now().isoformat(timespec='milliseconds') NOTE : Since Python 3.6!
print(iso_str)  # Например: '2026-09-29T09:44:15.123'
-----------------------------------------------------------------------------------------
from datetime import datetime
now = datetime.now()
# Оставляем только полную тысячу микросекунд (миллисекунды)
rounded_dt = now.replace(microsecond=(now.microsecond // 1000) * 1000)
print(rounded_dt.microsecond)  # Например: 123000 (что равно 123 мс)

"""
#########################################################################################
from dataclasses import dataclass  # Namespace for Operations with Data Classes
from time import sleep             # Getting a Sleep() Function for an Output Delay 
from pyVmomi import vim            # Namespace for a Core Operations with VI Objects
from sys import stdout						 # Function for an Output Print Options
import openpyxl                    # For an Operations with MS Excel File(s)

@dataclass # Main Class of this Module : 'vm_ha_mon_action_evt.py'
class GetVmsHaTrigAlrmClass(): ##########################################################
	pass
	connstate  : bool                # Is vCenter Server Connected : True/False
	vc_instance: vim.ServiceInstance # Passing SmartConnect Service Instance
	
	def get_vm_trig_alrm_func(self): ######################################################
		pass
		vi_content = self.vc_instance.RetrieveContent()
		if ((vi_content is not None) and (self.connstate)):
			
			container = vi_content.rootFolder          # Root/Parent Directory Starting Point
			trig_alrms = container.triggeredAlarmState # Searching for Triggered Alarms 
		
			# Open and Initialize the Existing(!) MS Excel File
			xlsx_wb = openpyxl.load_workbook("report.xlsx")

			# Create a Thematic Sheet and Set it as Active
			if (not ("vm_ha_monitor_errors" in xlsx_wb.sheetnames)):
				pass
				xlsx_wb.create_sheet(title = "vm_ha_monitor_errors")
				active_xlsx_sheet = xlsx_wb["vm_ha_monitor_errors"] # type: ignore

				# Set the Column Headers of Active XLSX Sheet
				column_header_list = ["VM Name", "Event Data", "Status/Color", "Event's Datetime"]
				active_xlsx_sheet.append(column_header_list)				# type: ignore
				xlsx_wb.save("report.xlsx")                  				# type: ignore

			else:
				active_xlsx_sheet = xlsx_wb["vm_ha_monitor_errors"] # type: ignore

			stdout.write(f"[INFO] : VM. Collecting HA Monitoring Action Triggered Alarms...")
			stdout.flush()
			for alrm_state in trig_alrms:     				 								# type: ignore
				pass
				# Get only Useful Information from all Alarms Data:
				alarm      = alrm_state.alarm
				alarm_info = alarm.info

				if ("virtual machine monitoring" in alarm_info.name.lower()):
					pass
					vm_entity_name = alrm_state.entity.name if hasattr(alrm_state.entity, "name") \
					else str(alrm_state.entity)
					xlsx_data = (vm_entity_name, alarm_info.name, alrm_state.overallStatus, \
									alrm_state.time.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3])
					active_xlsx_sheet.append(xlsx_data)
					xlsx_wb.save("report.xlsx")

			container.Destroy() # to Avoid Memory Accumulation in the vCenter Server
			print("Done.")
			sleep(2)
			print ("")
			