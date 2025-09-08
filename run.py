import csv
import gspread
import time
import warnings
import os
import sys
import json
import threading
import tempfile
import shutil
from collections import defaultdict
from datetime import datetime
from gspread_formatting import *
from gspread.utils import rowcol_to_a1
from google.oauth2 import service_account
from flask import Flask, request, render_template_string
from werkzeug.utils import secure_filename
warnings.filterwarnings('ignore', category=DeprecationWarning)
app = Flask(__name__)

LOCK_TIMEOUT = 300
DAILY_NORMS = {
        'Rent': 50.0,
        'Gym': 3.0,
        'Groceries': 3,
        'Transport': 0.27,
        'Entertainment': 0.17,
        'Utilities': 2.0,
        'Shopping': 3.33,
        'Dining': 10.00
    }

ALLOWED_EXTENSIONS = {'csv'}


def get_lock(lock_name):
    """File blocking"""
    # Используем временную директорию системы
    temp_dir = tempfile.gettempdir()
    lock_file = os.path.join(temp_dir, f"{lock_name}.lock")
    if os.path.exists(lock_file):
        # Проверяем время создания (если старше 5 минут - считаем устаревшей)
        if time.time() - os.path.getmtime(lock_file) < LOCK_TIMEOUT:
            return False
        else:
            try:
                os.remove(lock_file)
            except:
                pass  # Игнорируем ошибки удаления
    # Создаем новую блокировку
    try:
        with open(lock_file, 'w') as f:
            f.write(str(time.time()))
        return True
    except:
        return False  # Не удалось создать файл


def release_lock(lock_name):
    """Освободить блокировку"""
    temp_dir = tempfile.gettempdir()
    lock_file = os.path.join(temp_dir, f"{lock_name}.lock")
    if os.path.exists(lock_file):
        try:
            os.remove(lock_file)
        except:
            pass  # Игнорируем ошибки удаления


def check_global_lock(month_name):
    """Проверяет глобальную блокировку через Google Sheets"""
    try:
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials for global lock check")
            return False
            
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        
        # Пытаемся получить или создать worksheet для блокировок
        try:
            lock_sheet = sh.worksheet("SYSTEM_LOCKS")
        except gspread.WorksheetNotFound:
            try:
                lock_sheet = sh.add_worksheet(title="SYSTEM_LOCKS", rows="10", cols="3")
                lock_sheet.update('A1', [['Month', 'Locked', 'Timestamp']])
                time.sleep(2)
            except Exception as e:
                print(f"⚠️ Could not create lock sheet: {e}")
                return False
            
        # Нормализуем название месяца для сравнения
        normalized_month = month_name.capitalize()
        print(f"🔍 Checking lock for: {normalized_month}")
        
        # Проверяем есть ли блокировка для этого месяца
        try:
            locks = lock_sheet.get_all_records()
            for lock in locks:
                if lock['Month'] == normalized_month and lock['Locked'] == 'YES':
                    lock_time = datetime.fromisoformat(lock['Timestamp'])
                    if (datetime.now() - lock_time).total_seconds() < 600:  # 10 минут
                        print(f"🔒 Month {month_name} is locked by another process")
                        return False
                    else:
                        # Устаревшая блокировка - удаляем её
                        all_data = lock_sheet.get_all_values()
                        for i, row in enumerate(all_data[1:], start=2):  # Пропускаем заголовок
                            if row[0] == normalized_month:
                                lock_sheet.update_cell(i, 2, 'NO')
                                print(f"🔓 Removed stale lock for {normalized_month}")
                                break
                        break
        
        except Exception as e:
            print(f"⚠️ Error reading locks: {e}")
            return False
        
        # Создаем новую блокировку
        try:
            new_lock = [normalized_month, 'YES', datetime.now().isoformat()]
            lock_sheet.append_row(new_lock)
            print(f"🔒 Global lock acquired for {normalized_month}")
            return True
        except Exception as e:
            print(f"⚠️ Could not create lock: {e}")
            return False
        
    except Exception as e:
        print(f"⚠️ Global lock check error: {e}")
        return False

def release_global_lock(month_name):
    """Освобождает глобальную блокировку через Google Sheets"""
    try:
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials for global lock release")
            return
            
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        
        try:
            lock_sheet = sh.worksheet("SYSTEM_LOCKS")
            all_data = lock_sheet.get_all_values()

            normalized_month = month_name.capitalize()
            print(f"🔓 Releasing lock for: {normalized_month}")

            #search all rows with the month and status Yes
            rows_to_update = []
            for i, row in enumerate(all_data[1:], start=2):  # Пропускаем заголовок
                if row[0] == normalized_month and row[1] == 'YES':
                    rows_to_update.append(i) 
            
            for row_num in rows_to_update:
                lock_sheet.update_cell(row_num, 2, 'NO')  # Колонка B (Locked)
                print(f"🔓 Lock released for {normalized_month} in row {row_num}")
            
            if not rows_to_update:
                print(f"⚠️ No active lock found for {normalized_month} to release")
                    
        except gspread.WorksheetNotFound:
            print("⚠️ Lock sheet not found - nothing to release")
        except Exception as e:
            print(f"⚠️ Error accessing lock sheet: {e}")
            
    except Exception as e:
        print(f"⚠️ Global lock release error: {e}")
    finally:
        # Гарантируем, что функция всегда завершится
        print(f"🔓 Lock release process completed for {month_name}")
    


def allowed_file(filename):
    return ('.' in filename
            and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS)


# def sync_google_sheets_operation(month_name, table_data):
#     """Synchronic version of Google Sheets operation"""
#     lock_name = f"gsheets_lock_{month_name.lower()}"
    
#     # Пытаемся получить блокировку
#     if not get_lock(lock_name):
#         print(f"❌ Google Sheets is currently locked for {month_name}. Please try again later.")
#         return False
    
#     try:
#         print(f"📨 🔵 LOCAL MODE: Starting sync Google Sheets operation for {month_name}")
#         print(f"📊 Data to write: {len(table_data)} rows")
#         time.sleep(2)
#         # 1. Authentification
#         print("🔑 Getting credentials...")
#         creds = get_google_credentials()
#         if not creds:
#             print("❌ No credentials available")
#             return False
#         print("✅ Credentials obtained, authorizing...")
#         gc = gspread.authorize(creds)
#         print("✅ Authorized, opening spreadsheet...")
#         # 2. Open target table by ID
#         try:
#             spreadsheet_key = '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ'
#             target_spreadsheet = gc.open_by_key(spreadsheet_key)
#             print("✅ Spreadsheet opened successfully")
#         except Exception as e:
#             print(f"❌ Error opening spreadsheet: {e}")
#             return False
        
#         try:
#             summary_sheet = target_spreadsheet.worksheet('SUMMARY')
#             print("✅ SUMMARY worksheet accessed")
#         except Exception as e:
#             print(f"❌ Error accessing SUMMARY worksheet: {e}")
#             return False
#         print("📋 Getting headers...")
#         # 3. Get current headers
#         headers = summary_sheet.row_values(2)
#         print(f"📝 Current headers: {headers}")

#         # 4. Normalizing month name for comparison
#         normalized_month = month_name.capitalize()
#         print(f"🔍 Looking for column: {normalized_month}")

#         # 5. Find the month column
#         month_col = None
#         for i, header in enumerate(headers, 1):
#             if header == normalized_month:
#                 month_col = i
#                 print(f"✅ Found existing column for {normalized_month} "
#                       f"at position: {month_col}")
#                 break

#         if month_col is None:
#             print("🔍 No existing column found, looking for empty column...")
#             # Find first empty column
#             for i, header in enumerate(headers, 1):
#                 if not header.strip():  # Empty column
#                     month_col = i
#                     print(f"✅ Found empty column at position: {month_col}")
#                     print(f"📝 Creating new column for {normalized_month}...")
#                     summary_sheet.update_cell(2, month_col, normalized_month)
#                     summary_sheet.update_cell(
#                         3,
#                         month_col + 1,
#                         f"{normalized_month} %"
#                         )
#                     print(f"✅ Created new column for {normalized_month}"
#                           f" at position: {month_col}")
#                     break

#         if month_col is None:
#             print("🔍 No empty columns, adding at the end...")
#             # Add new columns at the end
#             month_col = len(headers) + 1
#             if month_col > 37:
#                 print("❌ Column limit reached (37)")
#                 return False
#             print(f"📝 Adding new column at position: {month_col}")
#             summary_sheet.update_cell(2, month_col, normalized_month)
#             summary_sheet.update_cell(
#                 3,
#                 month_col + 1,
#                 f"{normalized_month} %"
#                 )
#             print(f"✅ Added new column for {normalized_month}"
#                   f"at position: {month_col}")
#         print("📝 Preparing data for writing...")
#         # 6. Prepare data to be written
#         update_data = []
#         for i, row_data in enumerate(table_data, start=4):
#             if len(row_data) == 3:
#                 category, amount, percentage = row_data
#                 update_data.append({
#                     'range': f"{gspread.utils.rowcol_to_a1(i, month_col)}",
#                     'values': [[amount]]
#                 })
#                 update_data.append({
#                     'range': f"{gspread.utils.rowcol_to_a1(i, month_col + 1)}",
#                     'values': [[percentage]]
#                 })

#         print(f"📤 Ready to write {len(update_data)} cells")

#         # 7. batch-query
#         if update_data:
#             print("⏳ Writing data to Google Sheets...")
#             batch_size = 3
#             max_retries = 3
#             for i in range(0, len(update_data), batch_size):
#                 batch = update_data[i:i+batch_size]
#                 retry_count = 0
#                 success = False
            
#                 while not success and retry_count < max_retries:
#                     try:
#                         summary_sheet.batch_update(batch)
#                         print(f"✅ Batch {i//batch_size + 1} written")
#                         success = True
                        
#                     except Exception as e:
#                         if "429" in str(e) or "Quota exceeded" in str(e):
#                             retry_count += 1
#                             wait_time = 90 * retry_count  # Увеличиваем время ожидания
#                             print(f"⚠️ Rate limit exceeded. Retry {retry_count}/{max_retries} in {wait_time} seconds...")
#                             time.sleep(wait_time)
#                         else:
#                             print(f"❌ Error in batch update: {e}")
#                             raise e  # Другие ошибки прокидываем дальше
                
#                 if not success:
#                     print(f"❌ Failed to write batch {i//batch_size + 1} after {max_retries} retries")
#                     return False
                    
#                 if i + batch_size < len(update_data):
#                     time.sleep(15)  # Увеличиваем паузу между батчами

#             print("✅ All data written successfully!")

#             # Format percentage column
#             try:
#                 print("🎨 Formatting percentage column...")
#                 percent_col = month_col + 1
#                 start_row = 4
#                 end_row = start_row + len(table_data) - 1
               
#                 for row in range(start_row, end_row + 1):
#                     cell_address = f"{rowcol_to_a1(row, percent_col)}"
#                     summary_sheet.format(cell_address, {
#                         "numberFormat": {
#                             "type": "PERCENT",
#                             "pattern": "0.00%"
#                             },
#                         "horizontalAlignment": "CENTER"
#                     })
#                 print("✅ Percentage column formatted")
#             except Exception as format_error:
#                 print(f"⚠️ Formatting error: {format_error}")

#         print("✅ Google Sheets update completed successfully!")
#         return True

#     except Exception as e:
#         print(f"❌ Error in sync_google_sheets_operation: {e}")
#         import traceback
#         print(f"🔍 Traceback: {traceback.format_exc()}")
#         return False
#     finally:
#             # Всегда освобождаем блокировку
#             release_lock(lock_name)
#             print(f"🔓 Lock released for {month_name}")



# def sync_google_sheets_operation(month_name, table_data):
#     """Synchronic version of Google Sheets operation"""
#     lock_name = f"gsheets_lock_{month_name.lower()}"
    
#     # Пытаемся получить блокировку
#     if not get_lock(lock_name):
#         print(f"❌ Google Sheets is currently locked for {month_name}. Please try again later.")
#         return False
    
#     try:
#         print(f"📨 🔵 LOCAL MODE: Starting sync Google Sheets operation for {month_name}")
#         print(f"📊 Data to write: {len(table_data)} rows")
#         time.sleep(2)
#         # 1. Authentification
#         print("🔑 Getting credentials...")
#         creds = get_google_credentials()
#         if not creds:
#             print("❌ No credentials available")
#             return False
#         print("✅ Credentials obtained, authorizing...")
#         gc = gspread.authorize(creds)
#         print("✅ Authorized, opening spreadsheet...")
#         # 2. Open target table by ID
#         try:
#             spreadsheet_key = '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ'
#             target_spreadsheet = gc.open_by_key(spreadsheet_key)
#             print("✅ Spreadsheet opened successfully")
#         except Exception as e:
#             print(f"❌ Error opening spreadsheet: {e}")
#             return False
        
#         try:
#             summary_sheet = target_spreadsheet.worksheet('SUMMARY')
#             print("✅ SUMMARY worksheet accessed")
#         except Exception as e:
#             print(f"❌ Error accessing SUMMARY worksheet: {e}")
#             return False
#         print("📋 Getting headers...")
#         # 3. Get current headers
#         headers = summary_sheet.row_values(2)
#         print(f"📝 Current headers: {headers}")

#         # 4. Normalizing month name for comparison
#         normalized_month = month_name.capitalize()
#         print(f"🔍 Looking for column: {normalized_month}")

#         # 5. Find the month column
#         month_col = None
#         for i, header in enumerate(headers, 1):
#             if header == normalized_month:
#                 month_col = i
#                 print(f"✅ Found existing column for {normalized_month} "
#                       f"at position: {month_col}")
#                 break

#         if month_col is None:
#             print("🔍 No existing column found, looking for empty column...")
#             # Find first empty column
#             for i, header in enumerate(headers, 1):
#                 if not header.strip():  # Empty column
#                     month_col = i
#                     print(f"✅ Found empty column at position: {month_col}")
#                     print(f"📝 Creating new column for {normalized_month}...")
#                     summary_sheet.update_cell(2, month_col, normalized_month)
#                     summary_sheet.update_cell(
#                         3,
#                         month_col + 1,
#                         f"{normalized_month} %"
#                         )
#                     print(f"✅ Created new column for {normalized_month}"
#                           f" at position: {month_col}")
#                     break

#         if month_col is None:
#             print("🔍 No empty columns, adding at the end...")
#             # Add new columns at the end
#             month_col = len(headers) + 1
#             if month_col > 37:
#                 print("❌ Column limit reached (37)")
#                 return False
#             print(f"📝 Adding new column at position: {month_col}")
#             summary_sheet.update_cell(2, month_col, normalized_month)
#             summary_sheet.update_cell(
#                 3,
#                 month_col + 1,
#                 f"{normalized_month} %"
#                 )
#             print(f"✅ Added new column for {normalized_month}"
#                   f"at position: {month_col}")
#         print("📝 Preparing data for writing...")
#         # 6. Prepare data to be written - ИСПРАВЛЕННЫЙ ФОРМАТ ДИАПАЗОНА
#         update_data = []
#         for i, row_data in enumerate(table_data, start=4):
#             if len(row_data) == 3:
#                 category, amount, percentage = row_data
#                 # ИСПРАВЛЕННЫЙ ФОРМАТ - убрал лишние кавычки
#                 update_data.append({
#                     'range': f"{rowcol_to_a1(i, month_col)}",  # Без лишних кавычек!
#                     'values': [[amount]]
#                 })
#                 update_data.append({
#                     'range': f"{rowcol_to_a1(i, month_col + 1)}",  # Без лишних кавычек!
#                     'values': [[percentage]]
#                 })

#         print(f"📤 Ready to write {len(update_data)} cells")

#         # 7. batch-query
#         if update_data:
#             print("⏳ Writing data to Google Sheets...")
#             batch_size = 3  # Еще уменьшаем размер батча
#             max_retries = 5  # Увеличиваем количество попыток
            
#             for i in range(0, len(update_data), batch_size):
#                 batch = update_data[i:i+batch_size]
#                 retry_count = 0
#                 success = False
            
#                 while not success and retry_count < max_retries:
#                     try:
#                         summary_sheet.batch_update(batch)
#                         print(f"✅ Batch {i//batch_size + 1} written")
#                         success = True
                        
#                     except Exception as e:
#                         if "429" in str(e) or "Quota exceeded" in str(e):
#                             retry_count += 1
#                             wait_time = 90 * retry_count  # Увеличиваем время ожидания
#                             print(f"⚠️ Rate limit exceeded. Retry {retry_count}/{max_retries} in {wait_time} seconds...")
#                             time.sleep(wait_time)
#                         else:
#                             print(f"❌ Error in batch update: {e}")
#                             raise e  # Другие ошибки прокидываем дальше
                
#                 if not success:
#                     print(f"❌ Failed to write batch {i//batch_size + 1} after {max_retries} retries")
#                     return False
                    
#                 if i + batch_size < len(update_data):
#                     time.sleep(15)  # Увеличиваем паузу между батчами

#             print("✅ All data written successfully!")

#         print("✅ Google Sheets update completed successfully!")
#         return True

#     except Exception as e:
#         print(f"❌ Error in sync_google_sheets_operation: {e}")
#         import traceback
#         print(f"🔍 Traceback: {traceback.format_exc()}")
#         return False
#     finally:
#         # Всегда освобождаем блокировку
#         release_lock(lock_name)
#         print(f"🔓 Lock released for {month_name}")
def sync_google_sheets_operation(month_name, table_data):
    """Synchronic version of Google Sheets operation"""
    # Проверяем глобальную блокировку
    normalized_month = month_name.capitalize()
    print(f"🔐 Attempting to acquire lock for: {normalized_month}")
    if not check_global_lock(normalized_month):       
        print(f"❌ Google Sheets is currently locked for {normalized_month}. Please try again later.")
        return False
    
    try:
        print(f"📨 🔵 LOCAL MODE: Starting sync Google Sheets operation for {normalized_month}")
        print(f"📊 Data to write: {len(table_data)} rows")
        time.sleep(2)
        # 1. Authentification
        print("🔑 Getting credentials...")
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials available")
            return False
        print("✅ Credentials obtained, authorizing...")
        gc = gspread.authorize(creds)
        print("✅ Authorized, opening spreadsheet...")
        # 2. Open target table by ID
        try:
            spreadsheet_key = '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ'
            target_spreadsheet = gc.open_by_key(spreadsheet_key)
            print("✅ Spreadsheet opened successfully")
        except Exception as e:
            print(f"❌ Error opening spreadsheet: {e}")
            return False
        
        try:
            summary_sheet = target_spreadsheet.worksheet('SUMMARY')
            print("✅ SUMMARY worksheet accessed")
        except Exception as e:
            print(f"❌ Error accessing SUMMARY worksheet: {e}")
            return False
        print("📋 Getting headers...")
        # 3. Get current headers
        headers = summary_sheet.row_values(2)
        print(f"📝 Current headers: {headers}")

        # 4. Normalizing month name for comparison
        print(f"🔍 Looking for column: {normalized_month}")
        # if not check_global_lock(normalized_month):
        #     return False

        # 5. Find the month column
        month_col = None
        for i, header in enumerate(headers, 1):
            if header == normalized_month:
                month_col = i
                print(f"✅ Found existing column for {normalized_month} "
                      f"at position: {month_col}")
                break

        if month_col is None:
            print("🔍 No existing column found, looking for empty column...")
            # Find first empty column
            for i, header in enumerate(headers, 1):
                if not header.strip():  # Empty column
                    month_col = i
                    print(f"✅ Found empty column at position: {month_col}")
                    print(f"📝 Creating new column for {normalized_month}...")
                    summary_sheet.update_cell(2, month_col, normalized_month)
                    summary_sheet.update_cell(
                        3,
                        month_col + 1,
                        f"{normalized_month} %"
                        )
                    print(f"✅ Created new column for {normalized_month}"
                          f" at position: {month_col}")
                    break

        if month_col is None:
            print("🔍 No empty columns, adding at the end...")
            # Add new columns at the end
            month_col = len(headers) + 1
            if month_col > 37:
                print("❌ Column limit reached (37)")
                return False
            print(f"📝 Adding new column at position: {month_col}")
            summary_sheet.update_cell(2, month_col, normalized_month)
            summary_sheet.update_cell(
                3,
                month_col + 1,
                f"{normalized_month} %"
                )
            print(f"✅ Added new column for {normalized_month}"
                  f"at position: {month_col}")
        print("📝 Preparing data for writing...")
        # 6. Prepare data to be written - ИСПРАВЛЕННЫЙ ФОРМАТ ДИАПАЗОНА
        update_data = []
        for i, row_data in enumerate(table_data, start=4):
            if len(row_data) == 3:
                category, amount, percentage = row_data
                # ИСПРАВЛЕННЫЙ ФОРМАТ - убрал лишние кавычки
                update_data.append({
                    'range': f"{rowcol_to_a1(i, month_col)}",  # Без лишних кавычек!
                    'values': [[amount]]
                })
                update_data.append({
                    'range': f"{rowcol_to_a1(i, month_col + 1)}",  # Без лишних кавычек!
                    'values': [[percentage]]
                })

        print(f"📤 Ready to write {len(update_data)} cells")

        # 7. batch-query
        if update_data:
            print("⏳ Writing data to Google Sheets...")
            batch_size = 2  # Еще уменьшаем размер батча
            max_retries = 3  # Увеличиваем количество попыток
            
            for i in range(0, len(update_data), batch_size):
                batch = update_data[i:i+batch_size]
                retry_count = 0
                success = False
            
                while not success and retry_count < max_retries:
                    try:
                        summary_sheet.batch_update(batch)
                        print(f"✅ Batch {i//batch_size + 1} written")
                        success = True
                        
                    except Exception as e:
                        if "429" in str(e) or "Quota exceeded" in str(e):
                            retry_count += 1
                            wait_time = 60 * retry_count  # Decrease wait time
                            print(f"⚠️ Rate limit exceeded. Retry {retry_count}/{max_retries} in {wait_time} seconds...")
                            time.sleep(wait_time)
                        else:
                            print(f"❌ Error in batch update: {e}")
                            raise e  # Другие ошибки прокидываем дальше
                
                if not success:
                    print(f"❌ Failed to write batch {i//batch_size + 1} after {max_retries} retries")
                    return False
                    
                if i + batch_size < len(update_data):
                    time.sleep(15)  # Увеличиваем паузу между батчами

            print("✅ All data written successfully!")

        print("✅ Google Sheets update completed successfully!")
        return True

    except Exception as e:
        print(f"❌ Error in sync_google_sheets_operation: {e}")
        import traceback
        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False
    finally:
        # Всегда освобождаем глобальную блокировку
        release_global_lock(normalized_month)


def cleanup_stale_locks():
    """Очищает все устаревшие блокировки"""
    try:
        creds = get_google_credentials()
        if not creds:
            return
            
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        
        try:
            lock_sheet = sh.worksheet("SYSTEM_LOCKS")
            all_data = lock_sheet.get_all_values()
            
            for i, row in enumerate(all_data[1:], start=2):  # Пропускаем заголовок
                if len(row) >= 3 and row[1] == 'YES':  # Locked = YES
                    try:
                        lock_time = datetime.fromisoformat(row[2])
                        if (datetime.now() - lock_time).total_seconds() > 600:  # Старше 10 минут
                            lock_sheet.update_cell(i, 2, 'NO')
                            print(f"🔓 Cleaned up stale lock for {row[0]}")
                    except:
                        # Если timestamp некорректный, тоже снимаем блокировку
                        lock_sheet.update_cell(i, 2, 'NO')
                        print(f"🔓 Cleaned up invalid lock for {row[0]}")
                        
        except gspread.WorksheetNotFound:
            pass  # Если листа нет, нечего очищать
            
    except Exception as e:
        print(f"⚠️ Cleanup error: {e}")


def async_google_sheets_operation(month_name, table_data):
    """Asynchronous Google Sheets processing"""
    try:
        print(f"🚀 Starting async Google Sheets operation for {month_name}")
        time.sleep(5)
        
        normalized_month = month_name.capitalize()
        success = sync_google_sheets_operation(normalized_month, table_data)

        if success:
            print(f"✓ Asynchronous writing to Google Sheets completed "
                  f"for {month_name}")
        else:
            print(f"✗ Error during asynchronous writing to Google Sheets")

    except Exception as e:
        print(f"Async Google Sheets error: {e}")
        import traceback
        print(f"🔥 Traceback: {traceback.format_exc()}")


def get_google_credentials():
    """Get Google credentials with better error handling"""
    try:
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
                              'https://www.googleapis.com/auth/drive']
        if "DYNO" in os.environ:
            print("🔑 Using environment credentials from Heroku")
            service_account_json = os.environ.get(
                'GOOGLE_SERVICE_ACCOUNT_JSON'
                )
            if service_account_json:
                try:
                    creds_dict = json.loads(service_account_json)
                    
                    Credentials = service_account.Credentials
                    creator = Credentials.from_service_account_info
                    return creator(creds_dict, scopes=SCOPES)
                except json.JSONDecodeError:
                    print("❌ Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON")
                    return None
            else:
                print("❌ GOOGLE_SERVICE_ACCOUNT_JSON environment "
                      "variable not found")
                return None
        else:
            # Local development
            if os.path.exists('creds.json'):
                print("🔑 Using local creds.json file")
                Credentials = service_account.Credentials
                credentials = Credentials.from_service_account_file(
                 'creds.json',
                 scopes=SCOPES
                )
                return credentials
            else:
                print("❌ Local creds.json file not found")
                print("💡 Create creds.json with Google Service Account credentials")
                return None
    except Exception as e:
        print(f"❌ Error getting credentials: {e}")
        return None


def categorize(description):
    """Categorize transaction based on description."""
    desc = description.lower()

    categories = {
        'Salary': ['salary', 'wages', 'salary deposit'],
        'Bonus': ['bonus', 'tip', 'reward'],
        'Rent': ['rent', 'monthly rent', 'rent payment'],
        'Groceries': ['supermarket', 'grocery', 'food'],
        'Dining': ['restaurant', 'cafe', 'coffee'],
        'Transport': ['bus', 'train', 'taxi', 'uber', 'fuel', 'enoc'],
        'Entertainment': ['movie', 'netflix', 'concert', 'spotify'],
        'Utilities': ['electricity', 'water', 'gas', 'internet', 'phone'],
        'Gym': ['gym', 'fitness', 'yoga'],
        'Shopping': ['clothing', 'electronics', 'shopping'],
        'Health': ['pharmacy', 'doctor', 'health', 'dentist'],
        'Insurance': ['insurance', 'health insurance', 'car insurance'],
        'Travel': ['flight', 'hotel', 'travel', 'airline', 'hilton'],
        'Other': []
    }

    for cat, terms in categories.items():
        if any(term in desc for term in terms):
            return cat
    return 'Other'


def get_month_column_name(month_input):
    """Привести название месяца к стандартному формату"""
    month = month_input.strip().capitalize()
    month_mapping = {
        'Jan': 'January', 'Feb': 'February', 'Mar': 'March',
        'Apr': 'April', 'May': 'May', 'Jun': 'June',
        'Jul': 'July', 'Aug': 'August', 'Sep': 'September',
        'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
    }
    return month_mapping.get(month, month)


def analyze(transactions, daily_categories, month):
    """Perform financial analysis with daily tracking"""
    analysis = {
        'income': 0, 'expenses': 0, 'categories': defaultdict(float),
        'income_categories': defaultdict(float),
        'month': month, 'daily_categories': daily_categories,
        'days_count': 30,
        'daily_averages': defaultdict(float),
        'norms_violations': []
    }

    for t in transactions:
        if t['type'] == 'income':
            analysis['income'] += t['amount']
            analysis['income_categories'][t['category']] += t['amount']
        else:
            analysis['expenses'] += abs(t['amount'])
            analysis['categories'][t['category']] += abs(t['amount'])
    # Calculate daily averages
    for category, total in analysis['categories'].items():
        daily_avg = total / analysis['days_count']
        analysis['daily_averages'][category] = daily_avg
        if category in DAILY_NORMS:
            if daily_avg > DAILY_NORMS[category] * 1.1:
                analysis['norms_violations'].append(
                    f"Daily average for {category} "
                    f"overspent: {daily_avg:.2f}€ "
                    f"vs norm: {DAILY_NORMS[category]:.2f}€"
                )
    analysis['savings'] = analysis['income'] - analysis['expenses']
    return analysis


def format_terminal_output(data, month, transactions_count=0):
    """Форматирует вывод для терминала 80x24 символов как на скриншоте"""
    output = []
    output.append(' ')
    expense_rate = (
            (data['expenses'] / data['income'] * 100)
            if data['income'] > 0
            else 0
    )

    savings_rate = (
        (data['savings'] / data['income'] * 100)
        if data['income'] > 0
        else 0
    )

    centered_title = f"<u>FINANCIAL OVERVIEW: {month.upper()}</u>"
    output.append(centered_title)
    output.append(f"Income: {data['income']:8.2f}€ [{'■' * 20}] 100.0%")
    output.append(f"Expenses: {data['expenses']:8.2f}€ "
                  f"[{'■' * int(expense_rate/5)}] "
                  f"{expense_rate:.1f}%")
    output.append(f"Savings: {data['savings']:8.2f}€ "
                  f"[{'■' * int(savings_rate/5)}] "
                  f"{savings_rate:.1f}%")
    output.append("<u>EXPENSE CATEGORIES:</u>")
    top_categories = sorted(
     data['categories'].items(),
     key=lambda x: x[1],
     reverse=True
     )[:12]

    for category, amount in top_categories:
        if data['expenses'] > 0:
            percent = (amount / data['expenses'] * 100)
            bar_length = max(1, int(percent / 5))  # Масштабируем гистограмму
            output.append(f"{category[:15]:<15} {amount:8.2f}€ "
                          f"{'■' * bar_length} ({percent:.1f}%)")
        else:
            output.append(f"{category[:15]:<15} {amount:8.2f}€")
    output.append("<u>DAILY SPENDING and NORMS:</u>")
    sorted_categories = sorted(
        [
            (cat, avg)
            for cat, avg in data['daily_averages'].items()
            if cat in DAILY_NORMS
        ],
        key=lambda x: abs(x[1] - DAILY_NORMS.get(x[0], 0)),  # Добавить abs()
        reverse=True
    )[:3]
    for category, avg in sorted_categories:
        norm = DAILY_NORMS.get(category, 0)
        diff = avg - norm
        arrow = "▲" if diff > 0 else "▼"
        output.append(f"{category[:12]:<12} Avg: {avg:5.2f}€ "
                      f"Norm: {norm:5.2f}€ {arrow} {abs(diff):.2f}€")
    output.append("<u>DAILY SPENDING RECOMMENDATIONS:</u>")
    recommendations = generate_daily_recommendations(data)[:3]
    for i, rec in enumerate(recommendations, 1):
        if len(rec) > 70:
            rec = rec[:67] + "..."
        output.append(f"{i}. {rec}")
    while len(output) > 24:
        output.pop()
    return "\n".join(output)


def terminal_visualization(data):
    """Visualize financial data in terminal (80x24) как на скриншоте."""

    centered_title = f"FINANCIAL OVERVIEW {data['month'].upper()}"

    print(centered_title)
    print(f"Income:   {data['income']:8.2f}€ [{'■' * 20}] 100.0%")
    # Summary bars
    expense_rate = (
        (data['expenses'] / data['income'] * 100)
        if data['income'] > 0
        else 0
        )
    savings_rate = ((data['savings'] / data['income'] * 100)
                    if data['income'] > 0
                    else 0)
    print(f"Expenses: {data['expenses']:8.2f}€ "
          f"[{'■' * int(expense_rate/5)}] "
          f"{expense_rate:.1f}%")
    print(f"Savings:  {data['savings']:8.2f}€"
          f"[{'■' * int(savings_rate/5)}] "
          f"{savings_rate:.1f}%")
    # Categories breakdown в 3 колонки с гистограммами
    print("EXPENSE CATEGORIES: ")
    top_categories = (sorted(data['categories'].items(),
                      key=lambda x: x[1], reverse=True)[:8])
    categories_with_percent = []
    max_percent = (
        max((amount / data['expenses'] * 100)
            for category, amount in top_categories)
        if data['expenses'] > 0
        else 0
            )
    for category, amount in top_categories:
        percent = ((amount / data['expenses'] * 100)
                   if data['expenses'] > 0 else 0)
        if max_percent > 0:
            scaled_percent = max(1, int(percent / max_percent * 8))
        else:
            scaled_percent = 1
        categories_with_percent.append((category, amount, scaled_percent))
    col1 = categories_with_percent[0:3]
    col2 = categories_with_percent[3:6]
    col3 = categories_with_percent[6:9]
    for i in range(3):
        line = ""
        if i < len(col1):
            cat1, amt1, bar_len1 = col1[i]
            line += f"{cat1[:10]:<10} {amt1:6.2f}€ {'■' * bar_len1}"
        else:
            line += " " * 25
        line += " " * 2
        if i < len(col2):
            cat2, amt2, bar_len2 = col2[i]
            line += f"{cat2[:10]:<10} {amt2:6.2f}€ {'■' * bar_len2}"
        else:
            line += " " * 25
        line += " " * 2
        if i < len(col3):
            cat3, amt3, bar_len3 = col3[i]
            line += f"{cat3[:10]:<10} {amt3:6.2f}€ {'■' * bar_len3}"
        print(line)
    # Daily spending and norms
    print("DAILY SPENDING and NORMS: ")
    sorted_categories = sorted(
        [
            (cat, avg)
            for cat, avg in data['daily_averages'].items()
            if cat in DAILY_NORMS
        ],
        key=lambda x: x[1] - DAILY_NORMS.get(x[0], 0),
        reverse=True
        )[:3]
    for category, avg in sorted_categories:
        norm = DAILY_NORMS.get(category, 0)
        diff = avg - norm
        arrow = "▲" if diff > 0 else "▼"
        print(f"{category[:10]:<10}"
              f"Avg: {avg:5.2f}€ Norm: {norm:5.2f}€"
              f"{arrow} {abs(diff):.2f}€")
    # Recommendations
    print("DAILY SPENDING RECOMMENDATIONS: ")
    recommendations = generate_daily_recommendations(data)[:3]
    for i, rec in enumerate(recommendations, 1):
        if len(rec) > 70:
            rec = rec[:67] + "..."
        print(f"{i}. {rec}")


def generate_daily_recommendations(data):
    """Generate daily category-specific recommendations."""
    recs = []
    if not data or 'income' not in data:
        return ["No financial data available for recommendations."]
    if data['income'] <= 0:
        return ["No income data - cannot generate recommendations."]
    else:
        # 1. Savings rate recommendation
        expense_rate = (data['expenses'] / data['income'] * 100)
        savings_rate = (data['savings'] / data['income'] * 100)
        if savings_rate < 20:
            recs.append(f"Aim for 20% savings (current: {savings_rate:.1f}%)")
            # Add top 3 norms violations
            recs.extend(data['norms_violations'][:3])
        # Ensure minimum recommendations
        if len(recs) < 3:
            recs.extend([
                "Plan meals weekly to reduce grocery costs",
                "Use public transport more frequently",

            ])
        return recs[:3]  # Return only top 5 recommendations


def prepare_summary_data(data, transactions):
    """Prepare the data for the SUMMARY section - all categories and totals."""
    #
    all_categories = [
        'TOTAL INCOME',
        'TOTAL EXPENSES',
        'SAVINGS',
        '',
        'INCOME CATEGORIES:',
        'Salary',
        'Bonus',
        'Other Income',
        '',
        'EXPENSE CATEGORIES:',
        'Rent',
        'Groceries',
        'Dining',
        'Transport',
        'Entertainment',
        'Utilities',
        'Gym',
        'Shopping',
        'Health',
        'Insurance',
        'Education',
        'Travel',
        'Car',
        'Other'
    ]
    # Collecting data by income
    income_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'income':
            income_by_category[t['category']] += t['amount']

    # Colleciting data by expense
    expenses_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'expense':
            expenses_by_category[t['category']] += t['amount']

    # Preparing totals
    table_data = []
    for category in all_categories:
        if category == 'TOTAL INCOME':
            table_data.append([category, data['income'], 1.0])
        elif category == 'TOTAL EXPENSES':
            percentage = (data['expenses'] / data['income']
                          if data['income'] > 0 else 0)
            table_data.append([category, data['expenses'], percentage])

        elif category == 'SAVINGS':
            percentage = (data['savings'] / data['income']
                          if data['income'] > 0 else 0)
            table_data.append([category, data['savings'], percentage])

        elif category in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:

            table_data.append([category, '', ''])
        elif category in income_by_category:
            amount = income_by_category[category]
            percentage = (amount / data['income']
                          if data['income'] > 0 else 0)
            table_data.append([category, amount, percentage])
        elif category == 'Salary':
            matched = False
            for income_cat in income_by_category:
                amount = income_by_category[income_cat]
                percentage = (amount / data['income']
                              if data['income'] > 0 else 0)
                table_data.append([category, amount, percentage])
                matched = True
                break
            if not matched:
                if category == 'Salary':
                    for income_cat in income_by_category:
                        if ('salary' in income_cat.lower() or
                                'income' in income_cat.lower()):
                            amount = income_by_category[income_cat]
                            percentage = (amount / data['income']
                                          if data['income'] > 0 else 0)
                        table_data.append([category, amount, percentage])
                        matched = True
                        break
            if not matched:
                table_data.append([category, 0, 0])
        elif category in expenses_by_category:
            amount = expenses_by_category[category]
            percentage = amount / \
                data['expenses'] if data['expenses'] > 0 else 0
            table_data.append([category, amount, percentage])
        else:
            table_data.append([category, 0, 0])
    return table_data


def write_to_target_sheet(table_data, month_name):
    """Записать данные в целевую таблицу SUMMARY"""
    try:
        if not table_data:
            print("✗ No data to write to target sheet")
            return False
        if len(table_data) > 50:
            print(f"⚠️ Large dataset ({len(table_data)} rows), "
                  f"simplifying update")
            # Упрощаем данные для больших наборов
            simplified_data = []
            for row in table_data:
                if row[0] in ['TOTAL INCOME', 'TOTAL EXPENSES', 'SAVINGS']:
                    simplified_data.append(row)
                elif (row[0] and
                      not any(x in row[0] for x in ['CATEGORIES', ''])):
                    simplified_data.append([row[0], row[1], 0])
            table_data = simplified_data
        return sync_google_sheets_operation(month_name, table_data)

        # if "DYNO" in os.environ:
        #     thread = (threading.Thread(target=async_google_sheets_operation,
        #               args=(month_name, table_data)))
        #     thread.daemon = True
        #     thread.start()
        #     print("Google Sheets operation started in background")
        #     return True
        # else:
        #     return sync_google_sheets_operation(month_name, table_data)

    except Exception as e:
        print(f"✗ Error in writing into SUMMARY: {e}")
        return False


def check_system_lock():
    """Проверяет, не запущен ли уже другой экземпляр"""
    lock_file = os.path.join(tempfile.gettempdir(), "finance_system.lock")
    
    if os.path.exists(lock_file):
        # Проверяем время последней модификации
        if time.time() - os.path.getmtime(lock_file) < 600:  # 10 минут
            return False
        else:
            # Устаревшая блокировка
            try:
                os.remove(lock_file)
            except:
                pass
    
    # Создаем новую блокировку
    try:
        with open(lock_file, 'w') as f:
            f.write(str(time.time()))
        return True
    except:
        return False

def release_system_lock():
    """Освобождает системную блокировку"""
    lock_file = os.path.join(tempfile.gettempdir(), "finance_system.lock")
    if os.path.exists(lock_file):
        try:
            os.remove(lock_file)
        except:
            pass


# def main():
    
#     if "DYNO" in os.environ:
#         # Heroku mode
#         port = int(os.environ.get('PORT', 5000))
#         app.run(host='0.0.0.0', port=port)
#     else:
#         if not check_system_lock():
#             print("❌ Another instance is already running. Please wait.")
#             sys.exit(1)
            
#         try:
#             # Local mode
#             print(f"PERSONAL FINANCE ANALYZER ")
#             MONTH = (
#                     input("Enter the month (e.g. 'March, April, May'): ")
#                     .strip()
#                     .lower()
#                     )
#             FILE = f"hsbc_{MONTH}.csv"
#             print(f"Loading file: {FILE}")

#             transactions, daily_categories = load_transactions(FILE)
#             if not transactions:
#                 print(f"No transactions found")
#                 import sys
#                 sys.exit(1)
#             data = analyze(transactions, daily_categories, MONTH)
#             terminal_visualization(data)

#             print("\n" + "="*50)
#             print("📊 Preparing data for Google Sheets...")
#             monthly_success = write_to_month_sheet(MONTH, transactions, data)
#             if monthly_success:
#                 print("✅ Month worksheet updated successfully!")
#             else:
#                 print("❌ Failed to update Month worksheet")
            
#             time.sleep(10) 
#             print("⏳ Writing to Month worksheet...")
#             table_data = prepare_summary_data(data, transactions)
#             MONTH_NORMALIZED = get_month_column_name(MONTH)
            
#             print("⏳ Writing to Google Sheets SUMMARY...")
#             success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            
#             if success:
#                 print("✅ Google Sheets update completed successfully!")
#             else:
#                 print("❌ Failed to update Google Sheets")
#         finally:
#             # Всегда освобождаем системную блокировку
#             release_system_lock()
def main():
    if "DYNO" in os.environ:
        # Heroku mode
        port = int(os.environ.get('PORT', 5000))
        app.run(host='0.0.0.0', port=port)
    else:
        if not check_system_lock():
            print("❌ Another instance is already running. Please wait.")
            sys.exit(1)
            
        try:
            # Local mode
            print(f"PERSONAL FINANCE ANALYZER ")
            MONTH = (
                    input("Enter the month (e.g. 'March, April, May'): ")
                    .strip()
                    .lower()
                    )
            FILE = f"hsbc_{MONTH}.csv"
            print(f"Loading file: {FILE}")

            transactions, daily_categories = load_transactions(FILE)
            if not transactions:
                print(f"No transactions found")
                import sys
                sys.exit(1)
            data = analyze(transactions, daily_categories, MONTH)
            terminal_visualization(data)

            print("\n" + "="*50)
            print("📊 Preparing data for Google Sheets...")
            
            # Добавляем паузу перед началом операций с Google Sheets
            print("⏳ Waiting before Google Sheets operations...")
            time.sleep(10)
            
            monthly_success = write_to_month_sheet(MONTH, transactions, data)
            if monthly_success:
                print("✅ Month worksheet updated successfully!")
            else:
                print("❌ Failed to update Month worksheet")
            
            time.sleep(10)  # Увеличиваем паузу
            
            print("⏳ Writing to Month worksheet...")
            table_data = prepare_summary_data(data, transactions)
            MONTH_NORMALIZED = get_month_column_name(MONTH)
            
            print("⏳ Writing to Google Sheets SUMMARY...")
            success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            
            if success:
                print("✅ Google Sheets update completed successfully!")
            else:
                print("❌ Failed to update Google Sheets")
        
        finally:
            # Всегда освобождаем системную блокировку
            release_system_lock()
        # Recommendations
        # print(f"DAILY SPENDING RECOMMENDATIONS: ")
        # for i, rec in enumerate(generate_daily_recommendations(data), 1):
        #     print(f"{i}. {rec}")


HTML = '''
<!DOCTYPE html>
<html>
<head>
    <title>Finance Analyzer</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .main-container {
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
            overflow: hidden;
            width: 700px;
            max-width: 95%;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            padding: 25px;
            text-align: center;
            color: white;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 600;
            letter-spacing: 1px;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
        }
        .header p {
            margin: 10px 0 0 0;
            font-size: 14px;
            opacity: 0.9;
        }
        .content {
            padding: 30px;
        }
        .form-container {
            text-align: center;
            margin-bottom: 25px;
        }
        .input-group {
            display: flex;
            flex-direction: column;
            gap: 15px;
            justify-content: center;
            align-items: center;
            margin-bottom: 20px;
        }
        input[type="text"], input[type="file"] {
            padding: 14px 20px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            width: 300px;
            transition: all 0.3s ease;
            background: #f8f9fa;
        }
        input[type="text"]:focus, input[type="file"]:focus {
            outline: none;
            border-color: #667eea;
            background: white;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        input[type="file"] {
            padding: 12px;
            cursor: pointer;
        }
        button {
            padding: 14px 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }
        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
        }
        .terminal {
            background: #2d3748;
            color: #e2e8f0;
            border: 2px solid #667eea;
            padding: 10px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.4;
            overflow: auto;
            white-space: pre-wrap;
            max-height: 700px;
            margin-bottom: 20px;
        }
        .terminal:empty {
            display: none !important;
        }
        .status {
            text-align: center;
            padding: 15px;
            border-radius: 8px;
            font-weight: 500;
            margin: 10px 0;
        }
        .hidden { display: none !important; }
        .status-loading {
            background: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }
        .status-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .status-warning {
            background: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }
        .status-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        .file-info {
            margin-top: 10px;
            padding: 10px;
            background: #e8f4f8;
            border-radius: 6px;
            border-left: 4px solid #2196F3;
        }
    </style>
</head>
<body>
    <div class="main-container">
        <div class="header">
            <h1>💰 PERSONAL FINANCE ANALYZER</h1>
            <p>Upload your CSV file and analyze your finances</p>
        </div>
        <div class="content">
            <div class="form-container">
                <!-- Добавлены ID для формы и кнопки -->
                <form method="POST"
                      enctype="multipart/form-data"
                      id="uploadForm">
                    <div class="input-group">
                        <input type="text"
                        name="month"
                        placeholder="Enter month (e.g. March, April)"
                        required>
                        <input type="file" name="file" accept=".csv" required>
                        <button type="submit" id="submitBtn">Analyze</button>
                    </div>
                </form>
                {% if filename %}
                <div class="file-info">
                    📁 Using file: <strong>{{ filename }}</strong>
                </div>
                {% endif %}
            </div>

            

            {% if result %}
            <div class="terminal">
                {{ result|safe }}
            </div>
            {% endif %}

            <!-- Блок статуса -->
            <div class="status hidden"
                 id="statusMessage">
                Processing your financial data...
                Google Sheets update in progress
            </div>
        </div>
    </div>

    <script>
        // Показываем статус загрузки при нажатии кнопки
        document.getElementById('uploadForm').addEventListener(
                                                                'submit',
                                                                function(e) {
            const statusElement = document.getElementById('statusMessage');
            const submitBtn = document.getElementById('submitBtn');
            // Показываем статус загрузки
            statusElement.classList.remove('hidden');
            statusElement.classList.remove('status-success',
                                            'status-error',
                                            'status-warning');
            statusElement.classList.add('status-loading');
            // Отключаем кнопку и меняем текст
            submitBtn.disabled = true;
            submitBtn.textContent = 'Processing...';
            submitBtn.style.opacity = '0.7';
        });

        // Если есть статус сообщение с сервера - показываем его
        {% if status_message %}
        document.addEventListener('DOMContentLoaded', function() {
            const statusElement = document.getElementById('statusMessage');
            statusElement.classList.remove('hidden');
            statusElement.textContent = '{{ status_message }}';
            {% if 'success' in status_message %}
            statusElement.classList.add('status-success');
            {% elif 'failed' in status_message %}
            statusElement.classList.add('status-error');
            {% elif 'warning' in status_message %}
            statusElement.classList.add('status-warning');
            {% endif %}
        });
        {% endif %}
    </script>
    <script>
    // Smooth scroll to the terminal function
    function scrollToTerminal() {
        const terminal = document.querySelector('.terminal');
        if (terminal) {
            terminal.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    }

    {% if result %}
    document.addEventListener('DOMContentLoaded', function() {
        setTimeout(function() {
            scrollToTerminal();
        }, 100);
    });
    {% endif %}
</script>
<script>
        // Показываем статус загрузки при нажатии кнопки и очищаем предыдущие результаты
        document.getElementById('uploadForm').addEventListener('submit', function(e) {
            const statusElement = document.getElementById('statusMessage');
            const submitBtn = document.getElementById('submitBtn');
            const terminalElement = document.querySelector('.terminal');
            const fileInput = document.querySelector('input[type="file"]');
            
            // Очищаем предыдущий терминал
            if (terminalElement) {
                terminalElement.innerHTML = '';
                terminalElement.style.display = 'none';
            }
            
            // Обновляем информацию о файле сразу
            if (fileInput.files.length > 0) {
                const fileName = fileInput.files[0].name;
                let fileInfoElement = document.querySelector('.file-info');
                
                if (!fileInfoElement) {
                    fileInfoElement = document.createElement('div');
                    fileInfoElement.className = 'file-info';
                    document.querySelector('.input-group').after(fileInfoElement);
                }
                
                fileInfoElement.innerHTML = `📁 Using file: <strong>${fileName}</strong>`;
                fileInfoElement.style.display = 'block';
            }
            
            // Показываем статус загрузки
            statusElement.classList.remove('hidden');
            statusElement.classList.remove('status-success', 'status-error', 'status-warning');
            statusElement.classList.add('status-loading');
            statusElement.textContent = 'Processing your financial data... Google Sheets update in progress';
            
            // Отключаем кнопку и меняем текст
            submitBtn.disabled = true;
            submitBtn.textContent = 'Processing...';
            submitBtn.style.opacity = '0.7';
        });

        // Если есть статус сообщение с сервера - показываем его
        {% if status_message %}
        document.addEventListener('DOMContentLoaded', function() {
            const statusElement = document.getElementById('statusMessage');
            statusElement.classList.remove('hidden');
            statusElement.textContent = '{{ status_message }}';
            {% if 'success' in status_message %}
            statusElement.classList.add('status-success');
            {% elif 'failed' in status_message %}
            statusElement.classList.add('status-error');
            {% elif 'warning' in status_message %}
            statusElement.classList.add('status-warning');
            {% endif %}
        });
        {% endif %}
    </script>
</body>

</html>
'''


def load_transactions(file_path_or_object):
    """Load transactions from uploaded file with proper CSV parsing"""
    transactions = []
    daily_categories = defaultdict(lambda: defaultdict(float))
    try:
        # Handle both file objects and file paths
        if hasattr(file_path_or_object, 'read'):
            # File object - read content
            content = file_path_or_object.read()
            if isinstance(content, bytes):
                content = content.decode('utf-8')
            lines = content.split('\n')
        else:
            # File path
            with open(file_path_or_object, 'r', encoding='utf-8') as file:
                lines = file.readlines()
        # Parse CSV lines
        for line in lines:
            line = line.strip()
            if line and not line.startswith(('#', 'Date')):
                try:
                    parts = line.split(',')
                    if len(parts) >= 5:
                        # Parse date (assuming format: "31 Mar 2025")
                        date_str = parts[0].strip()
                        description = parts[1].strip()
                        amount = float(parts[2].strip())
                        currency = parts[3].strip()
                        transaction_type = parts[4].strip().lower()
                        # Convert date to standard format
                        try:
                            date_obj = datetime.strptime(date_str, "%d %b %Y")
                            # Check if the date exists
                            if date_obj.day != int(date_str.split()[0]):
                                print(f"Warning: Invalid date '{date_str}' - skipping")
                                continue
                        except ValueError:
                            print(f"Warning: Error parsing date '{date_str}' - skipping")
                            continue

                        date_formatted = date_obj.strftime("%Y-%m-%d")
                        # Categorize
                        category = categorize(description)
                        transaction = {
                            'date': date_formatted,
                            'desc': description[:30],
                            'amount': amount,
                            'type': ('income'
                                     if transaction_type == 'credit'
                                     else 'expense'),
                            'category': category
                        }
                        transactions.append(transaction)
                        # Track daily categories for expenses
                        if transaction_type != 'credit':
                            daily = daily_categories[date_formatted]
                            daily[category] += amount
                except (ValueError, IndexError) as e:
                    print(f"Warning: Error parsing line '{line}' - {e}")
                    continue
    except Exception as e:
        print(f"Error loading transactions: {e}")
        return [], defaultdict(lambda: defaultdict(float))
    return transactions, daily_categories


def get_operation_status(analysis_success, sheets_success):
    """Возвращает сообщение о статусе операций"""
    if analysis_success and sheets_success:
        return (f"✅ Analysis completed successfully "
                f"and data written to Google Sheets")
    elif analysis_success and not sheets_success:
        return (f"⚠️ Analysis completed"
                f"but failed to write data to Google Sheets")
    elif not analysis_success and sheets_success:
        return "⚠️ Analysis failed but Google Sheets operation completed"
    else:
        return "❌ Both analysis and Google Sheets operations failed"


def run_full_analysis_with_file(month, file_path, temp_dir):
    """Полная обработка в фоновом режиме с использованием загруженного файла"""
    analysis_success = False
    sheets_success = False
    try:
        print(f"🚀 Starting FULL background analysis"
              f"for {month} with uploaded file")
        transactions, daily_categories = load_transactions(file_path)
        if not transactions:
            print("No transactions found in uploaded file")
            return
        data = analyze(transactions, daily_categories, month)
        analysis_success = True

        print(f"{month.upper()} ANALYSIS COMPLETED")
        print(f"Income: {data['income']:.2f}€")
        print(f"Expenses: {data['expenses']:.2f}€")
        print(f"Savings: {data['savings']:.2f}€")
        # 1. Writing into Month sheet
        print(f"📝 Writing to {month} worksheet...")
        monthly_success = write_to_month_sheet(month, transactions, data)
        if monthly_success:
            print(f"✅ Successfully updated {month} worksheet")
        else:
            print(f"❌ Failed to update {month} worksheet")
        time.sleep(10)
        # 2. Writing into Summary sheet
        print("⏳ Starting Google Sheets SUMMARY update...")
        table_data = prepare_summary_data(data, transactions)
        MONTH_NORMALIZED = get_month_column_name(month)
        summary_success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
        sheets_success = summary_success
        if summary_success:
            print("✅ Successfully updated Google Sheets SUMMARY")
        else:
            print("❌ Failed to update Google Sheets SUMMARY")
        # Printing status message
        status_message = get_operation_status(analysis_success, sheets_success)
        print(f"🎉 {status_message}")
    except Exception as e:
        print(f"Background analysis error: {e}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
    finally:
        # Clearing temporary data
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                print(f"Cleaned up temporary directory: {temp_dir}")
        except Exception as cleanup_error:
            print(f"Error cleaning up temporary files: {cleanup_error}")
        # Returning operations status
        return analysis_success, sheets_success


@app.route('/', methods=['GET', 'POST'])
def index():
    print("🔍 Index function called")
    if not check_system_lock():
        print("⚠️ System is locked")
        return render_template_string(HTML,
                                    result="⚠️ System is busy. Please try again in a few minutes.",
                                    status_message="❌ System busy - please wait")


    result = None
    month = None
    filename = None
    status_message = None

    try:
        if request.method == 'POST':
            print("📨 POST request received")
            month = request.form['month'].strip().lower()
            print(f"📅 Month: {month}")
            if 'file' not in request.files:
                return render_template_string(HTML,
                                            result="No file uploaded",
                                            month=month)
            file = request.files['file']
            if file.filename == '':
                return render_template_string(HTML,
                                            result="No file selected",
                                            month=month)
            if file and allowed_file(file.filename):
                try:
                    filename = secure_filename(file.filename)
                    # Create temporary file for processing
                    temp_dir = tempfile.mkdtemp()
                    temp_file_path = os.path.join(temp_dir, f"hsbc_{month}.csv")
                    # Save uploaded file
                    file.save(temp_file_path)
                    # Load transactions for immediate display
                    transactions, daily_categories = (
                        load_transactions(temp_file_path)
                        )
                    if transactions:
                        data = analyze(transactions, daily_categories, month)
                        result = format_terminal_output(data,
                                                        month,
                                                        len(transactions))
                        # Start background processing and get status
                        analysis_success, sheets_success = (
                            run_full_analysis_with_file(
                                                            month,
                                                            temp_file_path,
                                                            temp_dir))
                        status_message = get_operation_status(analysis_success,
                                                            sheets_success)
                    else:
                        result = f"No valid transactions found in {filename}"
                        status_message = (f"❌ Analysis failed "
                                        f"no transactions found")
                except Exception as e:
                    result = f"Error processing file: {str(e)}"
                    status_message = "❌ Analysis failed due to error"
            else:
                result = "Invalid file type. Please upload a CSV file."
                status_message = "❌ Invalid file type"
        return render_template_string(HTML,
                                    result=result,
                                    month=month,
                                    filename=filename,
                                    status_message=status_message)
    finally:
        # Освобождаем системную блокировку
        release_system_lock()

def set_column_width(worksheet, column_letter, width):
    """Set column width for worksheet"""
    try:
        # Convert column letter to index (A=1, B=2, etc.)
        col_index = gspread.utils.a1_to_rowcol(column_letter + '1')[1]
        
        body = {
            "requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_index - 1,
                        "endIndex": col_index
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize"
                }
            }]
        }
        worksheet.spreadsheet.batch_update(body)
        print(f"✅ Column {column_letter} width set to {width}px")
    except Exception as e:
        print(f"⚠️ Error setting column width for {column_letter}: {e}")


# def write_to_month_sheet(month_name, transactions, data):
#     """Запись данных в лист месяца в формате как на скриншоте"""
#     lock_name = f"month_sheet_lock_{month_name.lower()}"
#     # Пытаемся получить блокировку
#     if not get_lock(lock_name):
#         print(f"❌ Month sheet is currently locked for {month_name}. Please try again later.")
#         return False
#     try:
#         print(f"📊 Writing to {month_name} worksheet with lock...")
#         print(f"📊 Writing to {month_name} worksheet...")
#         # 1. Authentification
#         creds = get_google_credentials()
#         if not creds:
#             print("❌ No credentials for month sheet")
#             return False
#         gc = gspread.authorize(creds)
#         sh = gc.open("Personal Finances")
#         # 2. Get or create worksheet
#         try:
#             worksheet = sh.worksheet(month_name)
#             print(f"✅ Worksheet '{month_name}' found")
#         except gspread.WorksheetNotFound:
#             print(f"📝 Creating new worksheet '{month_name}'...")
#             worksheet = sh.add_worksheet(title=month_name,
#                                          rows="100", cols="20")
#             print(f"✅ Worksheet '{month_name}' created")
#         # 3. Clear existing data
#         worksheet.clear()
#         time.sleep(5)
#         # 4. Create the layout as shown in the screenshot
#         # Financial Overview header
#         worksheet.update('A6', [['FINANCIAL OVERVIEW']])
#         worksheet.merge_cells('A6:E6')
#         worksheet.format('A6', {
#             "textFormat": {"bold": True, "fontSize": 14},
#             "horizontalAlignment": "CENTER"
#         })
#         # Table headers
#         headers = ["Date", "Description", "Amount", "Type", "Category"]
#         worksheet.update('A7', [headers])
#         worksheet.format('A7:E7', {
#             "textFormat": {"bold": True, "fontSize": 12},
#             "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Write transactions
#         all_data = []
#         for t in transactions:
#             all_data.append([t['date'],
#                              t['desc'][:30],
#                              t['amount'],
#                              t['type'],
#                              t['category']])
#         if all_data:
#             worksheet.update('A8', all_data)
#         # Format transaction table with alternating column backgrounds
#         last_transaction_row = 7 + len(transactions)
#         # Column A (Date) - light-grey background
#         worksheet.format(f'A8:A{last_transaction_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Column B (Description) - white background
#         worksheet.format(f'B8:B{last_transaction_row}', {
#             "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Column C (Amount) - light-grey background
#         worksheet.format(f'C8:C{last_transaction_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#         })
#         # Column D (Type) - white background
#         worksheet.format(f'D8:D{last_transaction_row}', {
#             "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Column E (Category) - light-grey background
#         worksheet.format(f'E8:E{last_transaction_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Transaction Categories header
#         worksheet.update('G6', [['TRANSACTION CATEGORIES']])
#         worksheet.merge_cells('G6:I6')
#         worksheet.format('G6', {
#             "textFormat": {"bold": True, "fontSize": 14},
#             "horizontalAlignment": "CENTER"
#         })
#         # Categories table headers
#         category_headers = ["Category", "Amount", "Percentage"]
#         worksheet.update('G7', [category_headers])
#         worksheet.format('G7:I7', {
#             "textFormat": {"bold": True, "fontSize": 12},
#             "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Prepare and write category data
#         table_data = prepare_summary_data(data, transactions)
#         category_data = []
#         for row in table_data:
#             if row[0] and row[0] not in ['', 'INCOME CATEGORIES:',
#                                          'EXPENSE CATEGORIES:']:
#                 category_data.append([row[0], row[1], row[2]])
#         if category_data:
#             worksheet.update('G8', category_data)
#         # Format category table with alternating column backgrounds
#         last_category_row = 7 + len(category_data)
#         # Column G (Category) - light-grey background
#         worksheet.format(f'G8:G{last_category_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Column H (Amount) - white background
#         worksheet.format(f'H8:H{last_category_row}', {
#             "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#         })
#         # Column I (Percentage) - light-grey background
#         worksheet.format(f'I8:I{last_category_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
#         })
#         # Daily Recommendations header
#         worksheet.update('K6', [['DAILY RECOMMENDATIONS']])
#         worksheet.merge_cells('K6:L6')
#         worksheet.format('K6', {
#             "textFormat": {"bold": True, "fontSize": 14},
#             "horizontalAlignment": "CENTER"
#         })
#         # Recommendations headers
#         rec_headers = ["Priority", "Recommendation"]
#         worksheet.update('K7', [rec_headers])
#         worksheet.format('K7:L7', {
#             "textFormat": {"bold": True, "fontSize": 12},
#             "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Write recommendations
#         recommendations = generate_daily_recommendations(data)
#         rec_data = []
#         for i, rec in enumerate(recommendations, 1):
#             rec_data.append([f"{i}", rec])
#         if rec_data:
#             worksheet.update('K8', rec_data)
#         # Format recommendations table with alternating column backgrounds
#         last_rec_row = 7 + len(rec_data)
#         # Column K (Priority) - light-grey background
#         worksheet.format(f'K8:K{last_rec_row}', {
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         # Column L (Recommendation) - white background
#         worksheet.format(f'L8:L{last_rec_row}', {
#             "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "wrapStrategy": "WRAP"
#         })
#         # Calculate percentages
#         expense_percentage = ((data['expenses'] / data['income'])
#                               if data['income'] > 0
#                               else 0)
#         savings_percentage = ((data['savings'] / data['income'])
#                               if data['income'] > 0
#                               else 0)
#         # Summary section at the top
#         summary_data = [
#             ["Total Income:", data['income'], 1.0],
#             ["Total Expenses:", data['expenses'], expense_percentage],
#             ["Savings:", data['savings'], savings_percentage]
#         ]
#         worksheet.update('A2', summary_data)
#         # Format summary section with alternating backgrounds
#         worksheet.format('A2:A4', {
#             "textFormat": {"bold": True},
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             }
#         })
#         worksheet.format('B2:B4', {
#             "textFormat": {"bold": False},
#             "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#         })
#         worksheet.format('C2:C4', {
#             "textFormat": {"bold": False},
#             "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#             "borders": {
#                 "top": {"style": "SOLID", "width": 1},
#                 "bottom": {"style": "SOLID", "width": 1},
#                 "left": {"style": "SOLID", "width": 1},
#                 "right": {"style": "SOLID", "width": 1}
#             },
#             "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
#         })
#         # Set column widths to match the screenshot
#         set_column_width(worksheet, 'A', 100)  # Date
#         set_column_width(worksheet, 'B', 200)  # Description
#         set_column_width(worksheet, 'C', 80)   # Amount
#         set_column_width(worksheet, 'D', 80)   # Type
#         set_column_width(worksheet, 'E', 100)  # Category
#         set_column_width(worksheet, 'G', 150)  # Category name
#         set_column_width(worksheet, 'H', 80)   # Amount
#         set_column_width(worksheet, 'I', 100)  # Percentage
#         set_column_width(worksheet, 'K', 90)   # Priority
#         set_column_width(worksheet, 'L', 300)  # Recommendation
#         print(f"✅ Successfully formatted {month_name} "
#               f"worksheet to match screenshot")
#         return True
#     except Exception as e:
#         print(f"❌ Error writing to {month_name} worksheet: {e}")
#         import traceback
#         print(f"🔍 Traceback: {traceback.format_exc()}")
#         return False
#     finally:
#         release_lock(lock_name)
#         print(f"🔓 Month sheet lock released for {month_name}")
# def write_to_month_sheet(month_name, transactions, data):
#     """Запись данных в лист месяца в формате как на скриншоте"""
#     lock_name = f"month_sheet_lock_{month_name.lower()}"
#     # Пытаемся получить блокировку
#     if not get_lock(lock_name):
#         print(f"❌ Month sheet is currently locked for {month_name}. Please try again later.")
#         return False
    
#     try:
#         print(f"📊 Writing to {month_name} worksheet...")
#         # 1. Authentification
#         creds = get_google_credentials()
#         if not creds:
#             print("❌ No credentials for month sheet")
#             return False
        
#         gc = gspread.authorize(creds)
#         sh = gc.open("Personal Finances")
        
#         # 2. Get or create worksheet
#         try:
#             worksheet = sh.worksheet(month_name)
#             print(f"✅ Worksheet '{month_name}' found")
#         except gspread.WorksheetNotFound:
#             print(f"📝 Creating new worksheet '{month_name}'...")
#             worksheet = sh.add_worksheet(title=month_name, rows="100", cols="20")
#             print(f"✅ Worksheet '{month_name}' created")
#             time.sleep(3)
        
#         # 3. Clear existing data с обработкой ошибок
#         try:
#             worksheet.clear()
#             time.sleep(3)
#         except Exception as e:
#             print(f"⚠️ Warning: Could not clear worksheet: {e}")
        
#         # 4. Основные данные с повторными попытками
#         max_retries = 3
#         retry_count = 0
        
#         while retry_count < max_retries:
#             try:
#                 # Financial Overview header
#                 worksheet.update('A6', [['FINANCIAL OVERVIEW']])
#                 worksheet.merge_cells('A6:E6')
                
#                 # Table headers
#                 headers = ["Date", "Description", "Amount", "Type", "Category"]
#                 worksheet.update('A7', [headers])
                
#                 # Write transactions
#                 all_data = []
#                 for t in transactions:
#                     all_data.append([
#                         t['date'],
#                         t['desc'][:30],
#                         t['amount'],
#                         t['type'],
#                         t['category']
#                     ])
                
#                 if all_data:
#                     worksheet.update('A8', all_data)
                
#                 # Transaction Categories header
#                 worksheet.update('G6', [['TRANSACTION CATEGORIES']])
#                 worksheet.merge_cells('G6:I6')
                
#                 # Categories table headers
#                 category_headers = ["Category", "Amount", "Percentage"]
#                 worksheet.update('G7', [category_headers])
                
#                 # Prepare and write category data
#                 table_data = prepare_summary_data(data, transactions)
#                 category_data = []
#                 for row in table_data:
#                     if row[0] and row[0] not in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:
#                         category_data.append([row[0], row[1], row[2]])
                
#                 if category_data:
#                     worksheet.update('G8', category_data)
                
#                 # Daily Recommendations header
#                 worksheet.update('K6', [['DAILY RECOMMENDATIONS']])
#                 worksheet.merge_cells('K6:L6')
                
#                 # Recommendations headers
#                 rec_headers = ["Priority", "Recommendation"]
#                 worksheet.update('K7', [rec_headers])
                
#                 # Write recommendations
#                 recommendations = generate_daily_recommendations(data)
#                 rec_data = []
#                 for i, rec in enumerate(recommendations, 1):
#                     rec_data.append([f"{i}", rec[:100]])  # Ограничиваем длину рекомендации
                
#                 if rec_data:
#                     worksheet.update('K8', rec_data)
                
#                 # Summary section at the top
#                 expense_percentage = (data['expenses'] / data['income']) if data['income'] > 0 else 0
#                 savings_percentage = (data['savings'] / data['income']) if data['income'] > 0 else 0
                
#                 summary_data = [
#                     ["Total Income:", data['income'], 1.0],
#                     ["Total Expenses:", data['expenses'], expense_percentage],
#                     ["Savings:", data['savings'], savings_percentage]
#                 ]
#                 worksheet.update('A2', summary_data)
                
#                 print(f"✅ Basic data written to {month_name} worksheet")
#                 break  # Успешно завершили
                
#             except Exception as e:
#                 retry_count += 1
#                 if "429" in str(e) or "Quota exceeded" in str(e):
#                     wait_time = 60 * retry_count
#                     print(f"⚠️ Rate limit exceeded. Retry {retry_count}/{max_retries} in {wait_time} seconds...")
#                     time.sleep(wait_time)
#                 else:
#                     print(f"❌ Error writing data: {e}")
#                     raise e
        
#         if retry_count >= max_retries:
#             print(f"❌ Failed to write data after {max_retries} retries")
#             return False
        
#         # 5. ФОРМАТИРОВАНИЕ С ЦВЕТОМ ЯЧЕЕК
#         try:
#             last_transaction_row = 7 + len(transactions)
#             last_category_row = 7 + len(category_data) if category_data else 7
#             last_rec_row = 7 + len(rec_data) if rec_data else 7
            
#             # Financial Overview header
#             worksheet.format('A6', {
#                 "textFormat": {"bold": True, "fontSize": 14},
#                 "horizontalAlignment": "CENTER"
#             })
            
#             # Table headers
#             worksheet.format('A7:E7', {
#                 "textFormat": {"bold": True, "fontSize": 12},
#                 "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 }
#             })
            
#             # Transaction table with alternating colors
#             if last_transaction_row > 7:
#                 # Column A (Date) - light-grey background
#                 worksheet.format(f'A8:A{last_transaction_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
                
#                 # Column B (Description) - white background
#                 worksheet.format(f'B8:B{last_transaction_row}', {
#                     "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
                
#                 # Column C (Amount) - light-grey background
#                 worksheet.format(f'C8:C{last_transaction_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     },
#                     "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#                 })
                
#                 # Column D (Type) - white background
#                 worksheet.format(f'D8:D{last_transaction_row}', {
#                     "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
                
#                 # Column E (Category) - light-grey background
#                 worksheet.format(f'E8:E{last_transaction_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
            
#             # Transaction Categories header
#             worksheet.format('G6', {
#                 "textFormat": {"bold": True, "fontSize": 14},
#                 "horizontalAlignment": "CENTER"
#             })
            
#             # Categories table headers
#             worksheet.format('G7:I7', {
#                 "textFormat": {"bold": True, "fontSize": 12},
#                 "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 }
#             })
            
#             # Category table with alternating colors
#             if last_category_row > 7:
#                 # Column G (Category) - light-grey background
#                 worksheet.format(f'G8:G{last_category_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
                
#                 # Column H (Amount) - white background
#                 worksheet.format(f'H8:H{last_category_row}', {
#                     "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     },
#                     "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#                 })
                
#                 # Column I (Percentage) - light-grey background
#                 worksheet.format(f'I8:I{last_category_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     },
#                     "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
#                 })
            
#             # Daily Recommendations header
#             worksheet.format('K6', {
#                 "textFormat": {"bold": True, "fontSize": 14},
#                 "horizontalAlignment": "CENTER"
#             })
            
#             # Recommendations headers
#             worksheet.format('K7:L7', {
#                 "textFormat": {"bold": True, "fontSize": 12},
#                 "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 }
#             })
            
#             # Recommendations table with alternating colors
#             if last_rec_row > 7:
#                 # Column K (Priority) - light-grey background
#                 worksheet.format(f'K8:K{last_rec_row}', {
#                     "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     }
#                 })
                
#                 # Column L (Recommendation) - white background
#                 worksheet.format(f'L8:L{last_rec_row}', {
#                     "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#                     "borders": {
#                         "top": {"style": "SOLID", "width": 1},
#                         "bottom": {"style": "SOLID", "width": 1},
#                         "left": {"style": "SOLID", "width": 1},
#                         "right": {"style": "SOLID", "width": 1}
#                     },
#                     "wrapStrategy": "WRAP"
#                 })
            
#             # Summary section formatting
#             worksheet.format('A2:A4', {
#                 "textFormat": {"bold": True},
#                 "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 }
#             })
            
#             worksheet.format('B2:B4', {
#                 "textFormat": {"bold": False},
#                 "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 },
#                 "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
#             })
            
#             worksheet.format('C2:C4', {
#                 "textFormat": {"bold": False},
#                 "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
#                 "borders": {
#                     "top": {"style": "SOLID", "width": 1},
#                     "bottom": {"style": "SOLID", "width": 1},
#                     "left": {"style": "SOLID", "width": 1},
#                     "right": {"style": "SOLID", "width": 1}
#                 },
#                 "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
#             })
            
#             print("✅ Full formatting with colors applied")
            
#         except Exception as format_error:
#             print(f"⚠️ Formatting error: {format_error}")
        
#         # 6. Установка ширины колонок
#         try:
#             set_column_width(worksheet, 'A', 100)   # Date
#             set_column_width(worksheet, 'B', 200)   # Description
#             set_column_width(worksheet, 'C', 80)    # Amount
#             set_column_width(worksheet, 'D', 80)    # Type
#             set_column_width(worksheet, 'E', 100)   # Category
#             set_column_width(worksheet, 'G', 150)   # Category name
#             set_column_width(worksheet, 'H', 80)    # Amount
#             set_column_width(worksheet, 'I', 100)   # Percentage
#             set_column_width(worksheet, 'K', 90)    # Priority
#             set_column_width(worksheet, 'L', 300)   # Recommendation
            
#             print("✅ Column widths set")
            
#         except Exception as width_error:
#             print(f"⚠️ Column width error: {width_error}")
        
#         print(f"✅ Successfully formatted {month_name} worksheet to match screenshot")
#         return True
        
#     except Exception as e:
#         print(f"❌ Error writing to {month_name} worksheet: {e}")
#         import traceback
#         print(f"🔍 Traceback: {traceback.format_exc()}")
#         return False
        
#     finally:
#         release_lock(lock_name)
#         print(f"🔓 Month sheet lock released for {month_name}")

def write_to_month_sheet(month_name, transactions, data):
    """Запись данных в лист месяца в формате как на скриншоте"""
    # Проверяем глобальную блокировку
    normalized_month = month_name.capitalize()
    if not check_global_lock(normalized_month):
        print(f"❌ Month sheet is currently locked for {month_name}. Please try again later.")
        return False
    
    try:
        print(f"📊 Writing to {month_name} worksheet...")
        # 1. Authentification
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials for month sheet")
            return False
        
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        
        # 2. Get or create worksheet
        try:
            worksheet = sh.worksheet(month_name)
            print(f"✅ Worksheet '{month_name}' found")
        except gspread.WorksheetNotFound:
            print(f"📝 Creating new worksheet '{month_name}'...")
            worksheet = sh.add_worksheet(title=month_name, rows="100", cols="20")
            print(f"✅ Worksheet '{month_name}' created")
            time.sleep(3)
        
        # 3. Clear existing data с обработкой ошибок
        try:
            worksheet.clear()
            time.sleep(3)
        except Exception as e:
            print(f"⚠️ Warning: Could not clear worksheet: {e}")
        
        # 4. Основные данные с повторными попытками
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # Financial Overview header
                worksheet.update('A6', [['FINANCIAL OVERVIEW']])
                worksheet.merge_cells('A6:E6')
                
                # Table headers
                headers = ["Date", "Description", "Amount", "Type", "Category"]
                worksheet.update('A7', [headers])
                
                # Write transactions
                all_data = []
                for t in transactions:
                    all_data.append([
                        t['date'],
                        t['desc'][:30],
                        t['amount'],
                        t['type'],
                        t['category']
                    ])
                
                if all_data:
                    worksheet.update('A8', all_data)
                
                # Transaction Categories header
                worksheet.update('G6', [['TRANSACTION CATEGORIES']])
                worksheet.merge_cells('G6:I6')
                
                # Categories table headers
                category_headers = ["Category", "Amount", "Percentage"]
                worksheet.update('G7', [category_headers])
                
                # Prepare and write category data
                table_data = prepare_summary_data(data, transactions)
                category_data = []
                for row in table_data:
                    if row[0] and row[0] not in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:
                        category_data.append([row[0], row[1], row[2]])
                
                if category_data:
                    worksheet.update('G8', category_data)
                
                # Daily Recommendations header
                worksheet.update('K6', [['DAILY RECOMMENDATIONS']])
                worksheet.merge_cells('K6:L6')
                
                # Recommendations headers
                rec_headers = ["Priority", "Recommendation"]
                worksheet.update('K7', [rec_headers])
                
                # Write recommendations
                recommendations = generate_daily_recommendations(data)
                rec_data = []
                for i, rec in enumerate(recommendations, 1):
                    rec_data.append([f"{i}", rec[:100]])  # Ограничиваем длину рекомендации
                
                if rec_data:
                    worksheet.update('K8', rec_data)
                
                # Summary section at the top
                expense_percentage = (data['expenses'] / data['income']) if data['income'] > 0 else 0
                savings_percentage = (data['savings'] / data['income']) if data['income'] > 0 else 0
                
                summary_data = [
                    ["Total Income:", data['income'], 1.0],
                    ["Total Expenses:", data['expenses'], expense_percentage],
                    ["Savings:", data['savings'], savings_percentage]
                ]
                worksheet.update('A2', summary_data)
                
                print(f"✅ Basic data written to {month_name} worksheet")
                break  # Успешно завершили
                
            except Exception as e:
                retry_count += 1
                if "429" in str(e) or "Quota exceeded" in str(e):
                    wait_time = 60 * retry_count
                    print(f"⚠️ Rate limit exceeded. Retry {retry_count}/{max_retries} in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    print(f"❌ Error writing data: {e}")
                    raise e
        
        if retry_count >= max_retries:
            print(f"❌ Failed to write data after {max_retries} retries")
            return False
        
        # 5. ФОРМАТИРОВАНИЕ С ЦВЕТОМ ЯЧЕЕК
        try:
            last_transaction_row = 7 + len(transactions)
            last_category_row = 7 + len(category_data) if category_data else 7
            last_rec_row = 7 + len(rec_data) if rec_data else 7
            
            # Financial Overview header
            worksheet.format('A6', {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            
            # Table headers
            worksheet.format('A7:E7', {
                "textFormat": {"bold": True, "fontSize": 12},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                }
            })
            
            # Transaction table with alternating colors
            if last_transaction_row > 7:
                # Column A (Date) - light-grey background
                worksheet.format(f'A8:A{last_transaction_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
                
                # Column B (Description) - white background
                worksheet.format(f'B8:B{last_transaction_row}', {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
                
                # Column C (Amount) - light-grey background
                worksheet.format(f'C8:C{last_transaction_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    },
                    "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
                })
                
                # Column D (Type) - white background
                worksheet.format(f'D8:D{last_transaction_row}', {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
                
                # Column E (Category) - light-grey background
                worksheet.format(f'E8:E{last_transaction_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
            
            # Transaction Categories header
            worksheet.format('G6', {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            
            # Categories table headers
            worksheet.format('G7:I7', {
                "textFormat": {"bold": True, "fontSize": 12},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                }
            })
            
            # Category table with alternating colors
            if last_category_row > 7:
                # Column G (Category) - light-grey background
                worksheet.format(f'G8:G{last_category_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
                
                # Column H (Amount) - white background
                worksheet.format(f'H8:H{last_category_row}', {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    },
                    "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
                })
                
                # Column I (Percentage) - light-grey background
                worksheet.format(f'I8:I{last_category_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    },
                    "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
                })
            
            # Daily Recommendations header
            worksheet.format('K6', {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            
            # Recommendations headers
            worksheet.format('K7:L7', {
                "textFormat": {"bold": True, "fontSize": 12},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                }
            })
            
            # Recommendations table with alternating colors
            if last_rec_row > 7:
                # Column K (Priority) - light-grey background
                worksheet.format(f'K8:K{last_rec_row}', {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                })
                
                # Column L (Recommendation) - white background
                worksheet.format(f'L8:L{last_rec_row}', {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    },
                    "wrapStrategy": "WRAP"
                })
            
            # Summary section formatting
            worksheet.format('A2:A4', {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                }
            })
            
            worksheet.format('B2:B4', {
                "textFormat": {"bold": False},
                "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                },
                "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
            })
            
            worksheet.format('C2:C4', {
                "textFormat": {"bold": False},
                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                "borders": {
                    "top": {"style": "SOLID", "width": 1},
                    "bottom": {"style": "SOLID", "width": 1},
                    "left": {"style": "SOLID", "width": 1},
                    "right": {"style": "SOLID", "width": 1}
                },
                "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
            })
            
            print("✅ Full formatting with colors applied")
            
        except Exception as format_error:
            print(f"⚠️ Formatting error: {format_error}")
        
        # 6. Установка ширины колонок
        try:
            set_column_width(worksheet, 'A', 100)   # Date
            set_column_width(worksheet, 'B', 200)   # Description
            set_column_width(worksheet, 'C', 80)    # Amount
            set_column_width(worksheet, 'D', 80)    # Type
            set_column_width(worksheet, 'E', 100)   # Category
            set_column_width(worksheet, 'G', 150)   # Category name
            set_column_width(worksheet, 'H', 80)    # Amount
            set_column_width(worksheet, 'I', 100)   # Percentage
            set_column_width(worksheet, 'K', 90)    # Priority
            set_column_width(worksheet, 'L', 300)   # Recommendation
            
            print("✅ Column widths set")
            
        except Exception as width_error:
            print(f"⚠️ Column width error: {width_error}")
        
        print(f"✅ Successfully formatted {month_name} worksheet to match screenshot")
        return True
        
    except Exception as e:
        print(f"❌ Error writing to {month_name} worksheet: {e}")
        import traceback
        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False
        
    finally:
        # Всегда освобождаем глобальную блокировку
        release_global_lock(normalized_month)


def run_full_analysis(month):
    """FULL background analysis"""
    try:
        import sys
        sys.stdout = sys.__stdout__
        print(f"🚀 Starting FULL background analysis for {month}")
        old_stdout = sys.stdout
        sys.stdout = sys.__stdout__
        print(f"Starting FULL background analysis for {month}")
        FILE = f"hsbc_{month}.csv"
        print(f"Loading file: {FILE}")

        transactions, daily_categories = load_transactions(FILE)
        if not transactions:
            print("No transactions found")
            return
        data = analyze(transactions, daily_categories, month)
        print(f"{month.upper()} ANALYSIS COMPLETED")
        print(f"Income: {data['income']:.2f}€")
        print(f"Expenses: {data['expenses']:.2f}€")
        print(f"Savings: {data['savings']:.2f}€")
        # 1. Writing into month sheet
        print(f"📝 Writing to {month} worksheet...")
        # 1. Writing into month sheet (new format)
        print(f"📝 Writing to {month} worksheet in screenshot format...")
        write_to_month_sheet(month, transactions, data)
        time.sleep(5)
        print("⏳ Starting Google Sheets update...")
        # starting Google Sheets
        table_data = prepare_summary_data(data, transactions)
        MONTH_NORMALIZED = get_month_column_name(month)
        write_to_target_sheet(table_data, MONTH_NORMALIZED)
        print("🎉 All background tasks completed!")
    except Exception as e:
        print(f"Background analysis error: {e}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
    finally:
        if 'old_stdout' in locals():
            sys.stdout = old_stdout


if __name__ == '__main__':
    cleanup_stale_locks()
    if "DYNO" in os.environ:
        # Heroku mode
        port = int(os.environ.get('PORT', 5000))
        app.run(host='0.0.0.0', port=port)
    else:
        # Local mode
        main()
