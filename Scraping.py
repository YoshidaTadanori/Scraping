import os
import threading
# GUI and Selenium imports are deferred to reduce optional dependencies when
# running unit tests that only need transformation logic.
from collections import defaultdict
import re
import csv
from datetime import datetime, timedelta
import time
import math
import traceback  # 追加


def extract_hotel_name(hotel_element, index):
    """Return a hotel name even when normal selectors fail."""
    from selenium.webdriver.common.by import By

    try:
        name = hotel_element.find_element(By.XPATH, ".//h2").text.strip()
        if name:
            return name
    except Exception:
        pass

    try:
        img = hotel_element.find_element(By.CSS_SELECTOR, "img[alt]")
        alt = img.get_attribute("alt").strip()
        if alt:
            return alt
    except Exception:
        pass

    alt_attr = hotel_element.get_attribute("data-hotel-name")
    if alt_attr:
        return alt_attr.strip()

    return f"不明なホテル{index}"

def scrape_data_for_date(driver, current_date):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    all_data = []
    base_url = "https://www.jalan.net/010000/LRG_012000/SML_012005/"
    params = f"?stayYear={current_date.year}&stayMonth={current_date.month}&stayDay={current_date.day}&stayCount=1&roomCount=1&adultNum=1"
    driver.get(base_url + params)

    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*/div[2]/a/div/div/div[1]/div[1]/h2')))

    page_number = 1  # ページ番号のカウント
    while True:
        print(f"Scraping page {page_number} for date {current_date.strftime('%Y/%m/%d')}")
        hotel_list = driver.find_elements(By.XPATH, '//*/div[2]/a/div/div/div[1]')
        if not hotel_list:
            print(f"No hotels found on page {page_number}. Ending scraping.")
            break
        for idx, hotel in enumerate(hotel_list, 1):
            try:
                hotel_name = extract_hotel_name(hotel, idx)
            except Exception as e:
                print(f"Error finding hotel name: {e}")
                hotel_name = f"不明なホテル{idx}"

            price = 0
            try:
                price_text = hotel.find_element(By.XPATH, "div[2]/dl/dd/span[1]").text
                price_match = re.search(r'\d+', price_text.replace(',', ''))
                if price_match:
                    price = int(price_match.group())
            except Exception:
                pass

            all_data.append([current_date.strftime("%Y/%m/%d"), hotel_name, price])

        # 「次へ」ボタンの探索
        next_buttons = driver.find_elements(By.XPATH, '//a[contains(text(), "次") or contains(text(), "Next")]')
        if next_buttons:
            try:
                next_button = next_buttons[0]
                # 「次へ」ボタンがクリック可能か確認
                if next_button.is_enabled():
                    next_button.click()
                    page_number += 1
                    # 新しいページが読み込まれるのを待つ
                    wait.until(EC.presence_of_element_located((By.XPATH, '//*/div[2]/a/div/div/div[1]/div[1]/h2')))
                else:
                    print("Next button is disabled. Ending scraping.")
                    break
            except Exception as e:
                print(f"Error clicking next button: {e}")
                break
        else:
            print("No next button found. Ending scraping.")
            break

    return all_data

def transform_and_save_data(data, csv_output_filename):
    transformed_data = defaultdict(dict)
    dates = set()
    hotels = set()

    for row in data:
        date, hotel_name, price = row
        transformed_data[date][hotel_name] = price
        dates.add(date)
        hotels.add(hotel_name)

    sorted_dates = sorted(dates)
    sorted_hotels = sorted(hotels)

    header_row = ['日付'] + sorted_hotels + ['平均']
    rows = [header_row]

    for date in sorted_dates:
        row_prices = []
        for hotel in sorted_hotels:
            price = transformed_data[date].get(hotel, 0)
            row_prices.append(price)
        nonzero_prices = [p for p in row_prices if p != 0]
        average = math.ceil(sum(nonzero_prices)/len(nonzero_prices)) if nonzero_prices else 0
        rows.append([date] + row_prices + [average])

    with open(csv_output_filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(rows)

def collect_data(start_date, end_date, root, progress_var, status_label, start_button):
    from tkinter import messagebox
    import tkinter as tk
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager
    try:
        status_label.config(text="データ収集中…")
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        today_str = datetime.now().strftime("%Y%m%d")
        csv_output_filename = os.path.join(desktop_path, f"hotel_price_{today_str}.csv")

        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')  # headlessモードでの互換性向上
        options.add_argument('--no-sandbox')
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        temp_data = []
        total_days = (end_date - start_date).days + 1
        current_day = 0

        try:
            current_date = start_date
            while current_date <= end_date:
                status_label.config(text=f"{current_date.strftime('%Y/%m/%d')} のデータ収集中…")
                root.update_idletasks()

                day_data = scrape_data_for_date(driver, current_date)
                temp_data.extend(day_data)

                current_day += 1
                progress = int((current_day / total_days) * 100)
                progress_var.set(progress)
                root.update_idletasks()

                current_date += timedelta(days=1)

            status_label.config(text="データ変換中…")
            root.update_idletasks()
            transform_and_save_data(temp_data, csv_output_filename)

            status_label.config(text="完了！")
            messagebox.showinfo("完了", f"データ収集が完了しました。\n\n{csv_output_filename}\nがデスクトップに保存されました。")
        except Exception as e:
            # エラー内容をポップアップで表示
            messagebox.showerror("エラー", f"データ収集中にエラーが発生しました:\n{e}")
            # コンソール上にも表示（スタックトレースを含めて表示）
            print("エラーが発生しました。詳細:")
            traceback.print_exc()
        finally:
            driver.quit()

    except Exception as e:
        # スレッド内エラーも表示
        messagebox.showerror("エラー", f"スレッド内でエラーが発生しました:\n{e}")
        # コンソール上にも表示
        print("スレッド内でエラーが発生しました。詳細:")
        traceback.print_exc()
    finally:
        start_button.config(state=tk.NORMAL)
        status_label.config(text="")
        progress_var.set(0)

def start_collection_thread(start_date, end_date, root, progress_var, status_label, start_button):
    import tkinter as tk
    start_button.config(state=tk.DISABLED)
    threading.Thread(target=collect_data, args=(start_date, end_date, root, progress_var, status_label, start_button), daemon=True).start()


def run_schedule(root, start_cal, end_cal, progress_var, status_label, start_button, time_var, stop_event):
    from tkinter import messagebox
    while not stop_event.is_set():
        try:
            t = datetime.strptime(time_var.get(), "%H:%M").time()
        except ValueError:
            time.sleep(60)
            continue
        now = datetime.now()
        next_run = datetime.combine(now.date(), t)
        if next_run <= now:
            next_run += timedelta(days=1)
        wait_seconds = (next_run - now).total_seconds()
        if stop_event.wait(wait_seconds):
            break
        try:
            start_date = datetime.strptime(start_cal.get_date(), "%Y/%m/%d")
            end_date = datetime.strptime(end_cal.get_date(), "%Y/%m/%d")
            if start_date > end_date:
                root.after(0, lambda: messagebox.showerror("エラー", "開始日は終了日より前にしてください。"))
                continue
            root.after(0, lambda sd=start_date, ed=end_date: start_collection_thread(sd, ed, root, progress_var, status_label, start_button))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("エラー", f"スケジュール実行でエラーが発生しました:\n{e}"))

def create_app():
    import tkinter as tk
    from tkinter import messagebox, ttk
    from tkcalendar import Calendar

    root = tk.Tk()
    root.title("ホテル料金収集ツール")
    root.geometry("720x850")
    root.configure(bg="#F8F9FA")

    accent_color = "#4285F4"
    text_color = "#202124"
    font_family = "Helvetica"

    header_frame = tk.Frame(root, bg="#FFFFFF", bd=0, highlightthickness=0)
    header_frame.pack(fill="x")

    title_label = tk.Label(header_frame, text="ホテル料金データ収集", bg="#FFFFFF", fg=text_color, 
                           font=(font_family, 20, "bold"), pady=20)
    title_label.pack()

    subtitle_label = tk.Label(header_frame, text="Google AnalyticsやLooker Studio風のシンプルUI", 
                              bg="#FFFFFF", fg="#5F6368", font=(font_family, 12))
    subtitle_label.pack(pady=(0, 20))

    content_frame = tk.Frame(root, bg="#FFFFFF", bd=1, relief="solid", padx=20, pady=20)
    content_frame.pack(padx=20, pady=20, fill="both", expand=True)

    info_label = tk.Label(content_frame, text="使い方ガイド\n\n1. 開始日と終了日を選択\n2. 「データ収集開始」をクリック\n3. 完了後、デスクトップにCSVファイルが生成\n4. 任意で毎日実行する時間を設定\n\n期間が長い場合は処理に時間がかかります。",
                          bg="#FFFFFF", fg="#5F6368", font=(font_family, 11), justify="left", wraplength=300)
    info_label.pack(pady=(0, 20))

    date_label = tk.Label(content_frame, text="収集期間を選択", bg="#FFFFFF", fg=text_color, font=(font_family, 14, "bold"))
    date_label.pack(pady=(0,10))

    date_frame = tk.Frame(content_frame, bg="#FFFFFF")
    date_frame.pack()

    start_frame = tk.Frame(date_frame, bg="#FFFFFF")
    start_frame.grid(row=0, column=0, padx=10, pady=10)

    start_label = tk.Label(start_frame, text="開始日", bg="#FFFFFF", fg=text_color, font=(font_family, 12, "bold"))
    start_label.pack(pady=(0,5))

    start_cal = Calendar(start_frame, selectmode='day', date_pattern='yyyy/mm/dd', 
                         background="#FFFFFF", foreground=text_color, 
                         headersbackground=accent_color, normalbackground="#FFFFFF", 
                         weekendbackground="#F0F0F0", selectbackground=accent_color)
    start_cal.pack()

    end_frame = tk.Frame(date_frame, bg="#FFFFFF")
    end_frame.grid(row=0, column=1, padx=10, pady=10)

    end_label = tk.Label(end_frame, text="終了日", bg="#FFFFFF", fg=text_color, font=(font_family, 12, "bold"))
    end_label.pack(pady=(0,5))

    end_cal = Calendar(end_frame, selectmode='day', date_pattern='yyyy/mm/dd', 
                       background="#FFFFFF", foreground=text_color, 
                       headersbackground=accent_color, normalbackground="#FFFFFF", 
                       weekendbackground="#F0F0F0", selectbackground=accent_color)
    end_cal.pack()

    status_label = tk.Label(content_frame, text="", bg="#FFFFFF", fg=accent_color, font=(font_family, 12))
    status_label.pack(pady=(20,5))

    progress_var = tk.IntVar()
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TProgressbar", troughcolor="#E0E0E0", background=accent_color, bordercolor="#E0E0E0", 
                    lightcolor=accent_color, darkcolor=accent_color)

    progress_bar = ttk.Progressbar(content_frame, orient="horizontal", length=280, mode="determinate", variable=progress_var, style="TProgressbar")
    progress_bar.pack(pady=(0,20))

    time_var = tk.StringVar(value="06:00")
    schedule_stop_event = threading.Event()
    schedule_thread = None

    time_frame = tk.Frame(content_frame, bg="#FFFFFF")
    time_frame.pack(pady=(0,10))

    time_label = tk.Label(time_frame, text="自動実行時間(HH:MM)", bg="#FFFFFF", fg=text_color, font=(font_family, 12))
    time_label.pack(side="left")

    time_entry = tk.Entry(time_frame, textvariable=time_var, width=6)
    time_entry.pack(side="left", padx=(5,10))

    def toggle_schedule():
        nonlocal schedule_thread
        if schedule_button.config('text')[-1] == 'スケジュール ON':
            schedule_button.config(text='スケジュール OFF')
            schedule_stop_event.set()
        else:
            try:
                datetime.strptime(time_var.get(), "%H:%M")
            except ValueError:
                messagebox.showerror("エラー", "時刻はHH:MM形式で入力してください。")
                return
            schedule_stop_event.clear()
            schedule_button.config(text='スケジュール ON')
            if schedule_thread is None or not schedule_thread.is_alive():
                schedule_thread = threading.Thread(target=run_schedule, args=(root, start_cal, end_cal, progress_var, status_label, start_button, time_var, schedule_stop_event), daemon=True)
                schedule_thread.start()

    schedule_button = tk.Button(time_frame, text='スケジュール OFF', bg=accent_color, fg='#FFFFFF', bd=0, font=(font_family, 11, 'bold'), command=toggle_schedule, cursor='hand2', activebackground="#3367D6", activeforeground='#FFFFFF')
    schedule_button.pack(side='left')

    def start_collection():
        try:
            start_date = datetime.strptime(start_cal.get_date(), "%Y/%m/%d")
            end_date = datetime.strptime(end_cal.get_date(), "%Y/%m/%d")
            if start_date > end_date:
                messagebox.showerror("エラー", "開始日は終了日より前にしてください。")
            else:
                response = messagebox.askokcancel("確認", 
                    f"以下の日程でデータ収集を開始します。\n\n開始日: {start_date.strftime('%Y/%m/%d')}\n終了日: {end_date.strftime('%Y/%m/%d')}\n\nよろしいですか？")
                if response:
                    start_collection_thread(start_date, end_date, root, progress_var, status_label, start_button)
        except Exception as e:
            messagebox.showerror("エラー", f"日付の取得に失敗しました:\n{e}")
            print("日付の取得に失敗しました。詳細:")
            traceback.print_exc()

    start_button = tk.Button(content_frame, text="データ収集開始", bg=accent_color, fg="#FFFFFF", bd=0, 
                             font=(font_family, 13, "bold"), activebackground="#3367D6", activeforeground="#FFFFFF", 
                             padx=20, pady=10, command=start_collection, cursor="hand2")
    start_button.pack(pady=(0,20))

    bottom_space = tk.Label(root, bg="#F8F9FA")
    bottom_space.pack(fill="x", pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_app()
