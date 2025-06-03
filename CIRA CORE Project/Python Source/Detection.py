import pandas as pd
from datetime import datetime, timedelta, time
import os
import time as time_sleep
from playsound import playsound

folder_path = r"C:/Users/Administrator/Desktop/CIRA CORE Project/Attendance/"
mp3_pass = r"C:/Users/Administrator/Desktop/CIRA CORE Project/Audio/PASS.mp3"
mp3_unknown = r"C:/Users/Administrator/Desktop/CIRA CORE Project/Audio/UNKNOWN.mp3"

today_date_str = datetime.now().strftime("%Y-%m-%d")
excel_file_path = os.path.join(folder_path, f"attendance_{today_date_str}.xlsx")
log_file_path = os.path.join(folder_path, f"log_{today_date_str}.txt")

if os.path.exists(excel_file_path):
	df = pd.read_excel(excel_file_path)
	attendance_data = df.to_dict(orient="records")
else:
	attendance_data = []
	df = pd.DataFrame(columns=["name", "entry_time", "exit_time", "duration", "status"])
	df.to_excel(excel_file_path, index=False)

last_recorded = {}

def write_log(message):
	with open(log_file_path, "a", encoding="utf-8") as log_file:
		log_file.write(f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")

def record_attendance(objects):
	if not attendance_data:
		print("ไม่มีข้อมูลใน attendance_data, เตรียมพร้อมสำหรับการบันทึกใหม่")

	for objs in objects:
		if "name" in objs:
			name = objs["name"]

			if name == "UNKNOWN":
				playsound(mp3_unknown)
				continue

			current_time = datetime.now()

			if current_time.time() > time(8, 0):
				status = "มาสาย"
			else:
				status = "ตรงเวลา"

			found = False
			for record in attendance_data:
				if record["name"] == name and pd.isna(record.get("exit_time")):
					print(f"บันทึกเวลาออก: {name} ออกเมื่อ {current_time.strftime('%H:%M:%S')}")
					record["exit_time"] = current_time

					entry_time = pd.to_datetime(record["entry_time"])
					duration = current_time - entry_time
					record["duration"] = str(duration).split('.')[0]

					found = True
					last_recorded[name] = current_time
					write_log(f"บันทึกเวลาออก: {name} ออกเมื่อ {current_time.strftime('%H:%M:%S')} (อยู่ {record['duration']})")

					playsound(mp3_pass)
					break

			if not found:
				print(f"บันทึกเวลาเข้า: {name} เข้าเมื่อ {current_time.strftime('%H:%M:%S')} ({status})")
				attendance_data.append({
					"name": name,
					"entry_time": current_time,
					"exit_time": None,
					"duration": None,
					"status": status
				})

				last_recorded[name] = current_time
				write_log(f"บันทึกเวลาเข้า: {name} เข้าเมื่อ {current_time.strftime('%H:%M:%S')} ({status})")

				playsound(mp3_pass)

			df = pd.DataFrame(attendance_data)
			df.to_excel(excel_file_path, index=False)

			time_sleep.sleep(3)

objects = payload["FaceRec"]["face_array"]

record_attendance(objects)

df = pd.DataFrame(attendance_data)
df.to_excel(excel_file_path, index=False)