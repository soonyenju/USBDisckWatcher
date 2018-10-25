# coding: utf-8
import win32api, win32con, win32gui
from ctypes import *
from watchdog.observers import Observer
from watchdog.events import *
from datetime import datetime
import os, time, pickle
from pathlib import Path
import pandas as pd


def main():
	drivers = getdriveletter()
	print("Start monitoring! | 开始监测！")
	w = Notification(drivers) 
	win32gui.PumpMessages()

class Consts(object):
	# 
	# Device change events (WM_DEVICECHANGE wParam) 
	# 
	DBT_DEVICEARRIVAL = 0x8000 
	DBT_DEVICEQUERYREMOVE = 0x8001 
	DBT_DEVICEQUERYREMOVEFAILED = 0x8002 
	DBT_DEVICEMOVEPENDING = 0x8003 
	DBT_DEVICEREMOVECOMPLETE = 0x8004 
	DBT_DEVICETYPESSPECIFIC = 0x8005 
	DBT_CONFIGCHANGED = 0x0018 
	# 
	# type of device in DEV_BROADCAST_HDR 
	# 
	DBT_DEVTYP_OEM = 0x00000000 
	DBT_DEVTYP_DEVNODE = 0x00000001 
	DBT_DEVTYP_VOLUME = 0x00000002 
	DBT_DEVTYPE_PORT = 0x00000003 
	DBT_DEVTYPE_NET = 0x00000004 
	# 
	# media types in DBT_DEVTYP_VOLUME 
	# 
	DBTF_MEDIA = 0x0001 
	DBTF_NET = 0x0002 
	WORD = c_ushort
	DWORD = c_ulong


class DEV_BROADCAST_HDR (Structure, Consts): 
	_fields_ = [ 
		("dbch_size", Consts.DWORD), 
		("dbch_devicetype", Consts.DWORD), 
		("dbch_reserved", Consts.DWORD) 
	] 

class DEV_BROADCAST_VOLUME (Structure, Consts): 
	_fields_ = [ 
		("dbcv_size", Consts.DWORD), 
		("dbcv_devicetype", Consts.DWORD), 
		("dbcv_reserved", Consts.DWORD), 
		("dbcv_unitmask", Consts.DWORD), 
		("dbcv_flags", Consts.WORD) 
	]




class Notification: 
	def __init__(self, drivers):
		self.drivers = drivers
		message_map = { 
			win32con.WM_DEVICECHANGE : self.onDeviceChange 
		}
		wc = win32gui.WNDCLASS () 
		hinst = wc.hInstance = win32api.GetModuleHandle (None) 
		wc.lpszClassName = "DeviceChangeDemo" 
		wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW; 
		wc.hCursor = win32gui.LoadCursor (0, win32con.IDC_ARROW) 
		wc.hbrBackground = win32con.COLOR_WINDOW 
		wc.lpfnWndProc = message_map 
		classAtom = win32gui.RegisterClass (wc) 
		style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU 
		self.hwnd = win32gui.CreateWindow ( 
			classAtom, 
			"Device Change Demo", 
			style, 
			0, 0, 
			win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 
			0, 0, 
			hinst, None 
		) 

	def onDeviceChange (self, hwnd, msg, wparam, lparam): 
		# 
		# WM_DEVICECHANGE: 
		#  wParam - type of change: arrival, removal etc. 
		#  lParam - what's changed? 
		#    if it's a volume then... 
		#  lParam - what's changed more exactly 
		# 
		dev_broadcast_hdr = DEV_BROADCAST_HDR.from_address (lparam) 
		if wparam == Consts.DBT_DEVICEARRIVAL: 
			print("Detect new device!| 检测到新外接设备！") 
			if dev_broadcast_hdr.dbch_devicetype == Consts.DBT_DEVTYP_VOLUME: 
				print("Found a new storage device! | 发现新储存设备！")
				new_drivers = getdriveletter()
				ret_list = [item for item in new_drivers if item not in self.drivers]
				print(ret_list)
				arri_driver = ret_list[0]
				monitoring(arri_driver)
				savelog()

				dev_broadcast_volume = DEV_BROADCAST_VOLUME.from_address (lparam)
				if dev_broadcast_volume.dbcv_flags & Consts.DBTF_MEDIA: 
					print("with some media") 
					drive_letter = drive_from_mask (dev_broadcast_volume.dbcv_unitmask) 
					print("in drive", chr (ord ("A") + drive_letter))
		return 1 

def drive_from_mask (mask): 
	n_drive = 0 
	while 1: 
		if (mask & (2 ** n_drive)): return n_drive 
		else: n_drive += 1

def getdriveletter():
	data = os.popen("wmic VOLUME GET DriveLetter").read()
	# print(data.decode("gb2312"))
	data = [d.strip() for d in data.split("\n") if len(d.strip())][1::]
	return data

# 监测数据


def monitoring(listen_dir):
	tmpdir = "./tmp"
	if not os.path.exists(tmpdir): os.makedirs(tmpdir)
	with open("./tmp/tmp.pkl", "wb") as f:
		pickle.dump([], f)
	observer = Observer()
	event_handler = FileEventHandler()
	observer.schedule(event_handler, listen_dir, True)
	observer.start()
	try:
		logs = []; count = 0
		while True:
			isdir = os.path.isdir(listen_dir)
			if isdir:
				time.sleep(1)
				with open("./tmp/tmp.pkl", "rb") as f:
					data = pickle.load(f)
					logs.append(data)
				if data != []:
					if data != logs[count - 1]:
						timestamp = datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M")
						data = [[d, timestamp] for d in data]
						print("Data logged! | 记录完毕")
						with open("./tmp/tmp2.pkl", "wb") as f:
							pickle.dump(data, f)
				count += 1
			else:
				print("USB ejection | U盘拔出")
				break

	except KeyboardInterrupt:
		observer.stop()


class FileEventHandler(FileSystemEventHandler):
	def __init__(self):
		FileSystemEventHandler.__init__(self)
		self.copyfiles = []
	def on_created(self, event):
		if event.is_directory:
			path = Path(event.src_path)
			paths = path.rglob("*.*")
			for p in paths:
				if path.name != "copylog.csv": self.copyfiles.append(path.name)
		else:
			path = Path(event.src_path)
			if path.name != "copylog.csv": self.copyfiles.append(path.name)
		self.copyfiles = list(set(self.copyfiles))
		with open("./tmp/tmp.pkl", "wb") as f:
			pickle.dump(self.copyfiles, f)


def savelog():
	with open("./tmp/tmp2.pkl", "rb") as f:
		data = pickle.load(f)

		# 保存记录
		df = pd.DataFrame(data, columns = ["filename", "copy time"])
		savepath = "./tmp/copylog.csv"
		if not os.path.exists(savepath):
			df.to_csv(savepath)
			del(df)
		else:
			origin_df = pd.read_csv(savepath, encoding = "gbk")[["filename", "copy time"]]
			df = pd.concat([origin_df, df])
			df.to_csv(savepath)
			del(df, origin_df)
		print("Log data saved! | 记录保存完毕！")
		# 保存记录

"""
def monitoring(listen_dir):
	print("开始监测是否有数据拷入！")
	observer = Observer()
	event_handler = FileEventHandler()
	observer.schedule(event_handler, listen_dir, True)
	observer.start()
	try:
		while True:
			isdir = os.path.isdir(listen_dir)
			if isdir:
				time.sleep(1)
			else:
				print("U盘拔出")
				break
	except KeyboardInterrupt:
		observer.stop()

class FileEventHandler(FileSystemEventHandler):
	def __init__(self):
		FileSystemEventHandler.__init__(self)

	def on_moved(self, event):
		if event.is_directory:
			print("路径由 {0} 移动至 {1}".format(event.src_path,event.dest_path))
		else:
			print("文件由 {0} 移动至 {1}".format(event.src_path,event.dest_path))

	def on_created(self, event):
		if event.is_directory:
			print("已创建路径:{0}".format(event.src_path))
			pass
		else:
			print("已创建文件:{0}".format(event.src_path))

	def on_deleted(self, event):
		if event.is_directory:
			print("路径已被删除:{0}".format(event.src_path))
		else:
			print("文件已被删除:{0}".format(event.src_path))

	def on_modified(self, event):
		if event.is_directory:
			print("路径中内容被修改:{0}".format(event.src_path))
		else:
			print("文件被修改:{0}".format(event.src_path))
"""

if __name__ == '__main__':
	main()
