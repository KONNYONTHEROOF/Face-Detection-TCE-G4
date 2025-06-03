import pandas as pd
from datetime import datetime
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

folder_path = r"C:/Users/Administrator/Desktop/CIRA CORE Project/Attendance/"
google_credentials_file = "C:/Users/Administrator/Desktop/CIRA CORE Project/PRIVATE.json"
google_sheet_name = "Attendance_Records"

def write_log(message, date_str=None):
	if date_str is None:
		date_str = datetime.now().strftime("%Y-%m-%d")
	log_file_path = os.path.join(folder_path, f"log_{date_str}.txt")
	with open(log_file_path, "a", encoding="utf-8") as log_file:
		log_file.write(f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")

def sync_rows(date_str=None):
	try:
		if date_str is None:
			date_str = datetime.now().strftime("%Y-%m-%d")
		
		scope = ["https://spreadsheets.google.com/feeds", 
				 "https://www.googleapis.com/auth/drive"]
		creds = ServiceAccountCredentials.from_json_keyfile_name(google_credentials_file, scope)
		client = gspread.authorize(creds)
		
		excel_path = os.path.join(folder_path, f"attendance_{date_str}.xlsx")
		if not os.path.exists(excel_path):
			write_log(f"No Excel file found for date {date_str}", date_str)
			return False
		
		df = pd.read_excel(excel_path)
		spreadsheet = client.open(google_sheet_name)
		
		try:
			sheet = spreadsheet.worksheet(date_str)
		except gspread.exceptions.WorksheetNotFound:
			sheet = spreadsheet.add_worksheet(title=date_str, rows=100, cols=5)
			sheet.append_row(["Name", "Entry Time", "Exit Time", "Duration", "Status"])

		sheet_data = sheet.get_all_values()
		current_row_count = len(sheet_data)  
		excel_row_count = len(df) + 1  

		for i in range(1, excel_row_count):  
			record = df.iloc[i - 1]
			row_data = [
				str(record["name"]),
				record["entry_time"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(record["entry_time"]) else "",
				record["exit_time"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(record["exit_time"]) else "",
				str(record["duration"]) if pd.notna(record["duration"]) else "",
				str(record["status"]) if pd.notna(record["status"]) else ""
			]
			
			if i < current_row_count:
				sheet.update(f"A{i+1}:E{i+1}", [row_data])
			else:
				sheet.append_row(row_data)
		
		write_log(f"Synced data with Google Sheet for {date_str}", date_str)
		return True

	except Exception as e:
		write_log(f"Error syncing to Google Sheet: {str(e)}", date_str)
		return False

if __name__ == "__main__":
	sync_rows()
