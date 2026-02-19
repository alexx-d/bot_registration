import pandas as pd
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import re
import winreg

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

SETTINGS_FILE = "settings.json"

class QuizBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoReg Quiz v1")
        self.root.geometry("800x700")
        
        self.style = ttk.Style()
        self.style.theme_use('clam') 
        self.style.configure("Treeview", rowheight=35, font=('Segoe UI', 10))
        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        self.style.configure("Treeview", padding=(5, 5)) 

        self.file_path = tk.StringVar()
        self.start_from = tk.StringVar(value="1")
        self.is_running = False

        self.t_fio = tk.StringVar()
        self.t_phone = tk.StringVar()
        self.t_email = tk.StringVar()
        self.q_url = tk.StringVar()
        self.q_name = tk.StringVar()

        self.load_settings()
        self.setup_ui()

        self.root.bind_class("Entry", "<Control-KeyPress>", self.handle_control_hotkeys)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill="both", expand=True, padx=15, pady=10)

        # 1. Файл
        file_frame = tk.LabelFrame(main_container, text=" 1. База данных Excel ", padx=10, pady=5)
        file_frame.pack(fill="x", pady=5)
        tk.Entry(file_frame, textvariable=self.file_path, state='readonly', width=85).pack(side="left", padx=5)
        tk.Button(file_frame, text="Обзор", command=self.browse_file).pack(side="left")

        # Настройки
        top_settings = tk.Frame(main_container)
        top_settings.pack(fill="x")
        
        t_frame = tk.LabelFrame(top_settings, text=" 2. Данные педагога ", padx=10, pady=5)
        t_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.create_entry_row(t_frame, "ФИО:", self.t_fio, 0)
        self.create_entry_row(t_frame, "Тел:", self.t_phone, 1)
        self.create_entry_row(t_frame, "Email:", self.t_email, 2)

        q_frame = tk.LabelFrame(top_settings, text=" 3. Викторина ", padx=10, pady=5)
        q_frame.pack(side="left", fill="both", expand=True)
        self.create_entry_row(q_frame, "Название:", self.q_name, 0)
        self.create_entry_row(q_frame, "URL:", self.q_url, 1)

        # 4. Таблица
        table_frame = tk.LabelFrame(main_container, text=" 4. Предпросмотр списка ", padx=5, pady=5)
        table_frame.pack(fill="both", expand=True, pady=10)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.cols = ("№", "ФИО обучающегося", "Дата рождения", "СНИЛС", "Группа/Класс", 
                     "Образовательная организация", "ФИО заявителя", "Контактный телефон", "e-mail")
        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", height=3)
        
        c_conf = {
            "№": (45, "center"), "ФИО обучающегося": (250, "w"), "Дата рождения": (120, "center"),
            "СНИЛС": (170, "center"), "Группа/Класс": (110, "center"), "Образовательная организация": (250, "w"),
            "ФИО заявителя": (180, "w"), "Контактный телефон": (140, "center"), "e-mail": (150, "w")
        }

        for col in self.cols:
            w, anc = c_conf[col]
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=anc, stretch=False)

        sy = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns"); sx.grid(row=1, column=0, sticky="ew")

        # Блок прогресса
        progress_frame = tk.Frame(main_container)
        progress_frame.pack(fill="x", pady=5)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", side="top", pady=2)

        p_status_frame = tk.Frame(progress_frame)
        p_status_frame.pack(fill="x")
        
        self.progress_label = tk.Label(p_status_frame, text="Прогресс: 0%", font=("Segoe UI", 9))
        self.progress_label.pack(side="left")

        self.eta_label = tk.Label(p_status_frame, text="Осталось: --:--", font=("Segoe UI", 9, "bold"), fg="#1976D2")
        self.eta_label.pack(side="right")
        
        ctrl_frame = tk.Frame(main_container)
        ctrl_frame.pack(pady=10)
        tk.Label(ctrl_frame, text="Начать с №:").pack(side="left")
        tk.Entry(ctrl_frame, textvariable=self.start_from, width=5).pack(side="left", padx=10)
        self.btn_run = tk.Button(ctrl_frame, text="ЗАПУСТИТЬ РЕГИСТРАЦИЮ", bg="#2E7D32", fg="white", 
                                 font=("Arial", 11, "bold"), padx=40, pady=10, command=self.start_thread)
        self.btn_run.pack(side="left")

        self.status_var = tk.StringVar(value="Готов")
        tk.Label(self.root, textvariable=self.status_var, bd=1, relief="sunken", anchor="w").pack(side="bottom", fill="x")

    def create_entry_row(self, frame, label, var, row):
        tk.Label(frame, text=label, width=8, anchor="e").grid(row=row, column=0, padx=5, pady=2)
        tk.Entry(frame, textvariable=var, width=38).grid(row=row, column=1, padx=5, pady=2)

    def load_settings(self):
        """Загрузка настроек из реестра Windows"""
        try:
            # Путь в реестре: HKEY_CURRENT_USER\Software\QuizBot
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\QuizBot")
            self.t_fio.set(winreg.QueryValueEx(key, "t_fio")[0])
            self.t_phone.set(winreg.QueryValueEx(key, "t_phone")[0])
            self.t_email.set(winreg.QueryValueEx(key, "t_email")[0])
            self.q_url.set(winreg.QueryValueEx(key, "q_url")[0])
            self.q_name.set(winreg.QueryValueEx(key, "q_name")[0])
            winreg.CloseKey(key)
        except WindowsError:
            # Если ключа еще нет (первый запуск), оставляем пустые поля
            pass

    def save_settings(self):
        """Сохранение настроек в реестр Windows"""
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\QuizBot")
            winreg.SetValueEx(key, "t_fio", 0, winreg.REG_SZ, self.t_fio.get())
            winreg.SetValueEx(key, "t_phone", 0, winreg.REG_SZ, self.t_phone.get())
            winreg.SetValueEx(key, "t_email", 0, winreg.REG_SZ, self.t_email.get())
            winreg.SetValueEx(key, "q_url", 0, winreg.REG_SZ, self.q_url.get())
            winreg.SetValueEx(key, "q_name", 0, winreg.REG_SZ, self.q_name.get())
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Ошибка сохранения в реестр: {e}")

    def on_closing(self):
        self.save_settings(); self.root.destroy()

    def browse_file(self):
        fn = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if fn: self.file_path.set(fn); self.load_table_data(fn)

    def load_table_data(self, path):
        try:
            for item in self.tree.get_children(): self.tree.delete(item)
            
            df = pd.read_excel(path, skiprows=1)
            
            req = ["ФИО обучающегося", "Дата рождения", "СНИЛС", "ФИО заявителя", 
                   "Контактный телефон", "e-mail", "Образовательная организация", "Группа/Класс"]
            
            if not all(c in df.columns for c in req):
                raise ValueError("Неверные заголовки")

            df = df.dropna(subset=['ФИО обучающегося'])
            for index, row in df.iterrows():
                snils = str(row.get('СНИЛС', '-')).strip()
                row_data = (index + 1, row.get('ФИО обучающегося', '-'), row.get('Дата рождения', '-'), 
                            snils, row.get('Группа/Класс', '-'), row.get('Образовательная организация', '-'), 
                            row.get('ФИО заявителя', '-'), row.get('Контактный телефон', '-'), row.get('e-mail', '-'))
                self.tree.insert("", "end", values=row_data)
            
            self.status_var.set(f"Загружено записей: {len(df)}")
            
        except Exception as e:
            msg = ("Внимание! В файле с данными заголовки таблицы должны начинаться со второй строки. "
                   "Проверьте, что в таблице имеются столбцы:\n\n"
                   "\"ФИО обучающегося\", \"Дата рождения\", \"СНИЛС\", \"ФИО заявителя\", "
                   "\"Контактный телефон\", \"e-mail\", \"Образовательная организация\", \"Группа/Класс\"")
            messagebox.showerror("Внимание", msg)

    def start_thread(self):
        if not self.file_path.get(): return
        self.save_settings()
        threading.Thread(target=self.run_bot, daemon=True).start()

    # --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
    def safe_fill(self, wait, label_text, value):
        xpath = f"//div[contains(text(), '{label_text}')]/following::input[1]"
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        field.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
        field.send_keys(str(value))
        return field

    def get_field_data(self, driver, label_name):
        try:
            container_xpath = f"//div[contains(text(), '{label_name}')]/following::div[contains(@class, 'v-input__control')][1]"
            container = driver.find_element(By.XPATH, container_xpath)
            
            input_el = container.find_element(By.TAG_NAME, "input")
            val = driver.execute_script("return arguments[0].value;", input_el)
            
            if not val:
                val = container.text
                
            return str(val).replace('\n', ' ').strip().lower()
        except:
            return ""

    def handle_control_hotkeys(self, event):
        char = event.char.lower()
        
        # Проверяем именно символ (char), так как keycode может прыгать от раскладки
        if char == '\x03': # Ctrl + C
            event.widget.event_generate("<<Copy>>")
            return "break"
        elif char == '\x16': # Ctrl + V
            event.widget.event_generate("<<Paste>>")
            return "break"
        elif char == '\x18': # Ctrl + X
            event.widget.event_generate("<<Cut>>")
            return "break"
        elif char == '\x01': # Ctrl + A
            event.widget.selection_range(0, 'end')
            event.widget.icursor('end')
            return "break"

    # --- ГЛАВНЫЙ БОТ ---
    def run_bot(self):
        if not self.file_path.get():
            return messagebox.showwarning("Внимание", "Выберите файл Excel!")
        
        self.btn_run.config(state="disabled", text="БОТ РАБОТАЕТ")
        driver = None
        
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            wait = WebDriverWait(driver, 25)

            df_full = pd.read_excel(self.file_path.get(), skiprows=1).dropna(subset=['ФИО обучающегося'])
            start_idx = int(self.start_from.get()) - 1
            work_df = df_full.iloc[start_idx:]
            
            total = len(work_df)
            self.progress["maximum"] = total
            self.progress["value"] = 0
            start_time_all = time.time()
            processed = 0
            
            for index, row in work_df.iterrows():
                processed += 1
                fio_current = str(row['ФИО обучающегося']).strip()
                
                # UI Обновление
                self.progress["value"] = processed
                percent = int((processed / total) * 100)
                self.progress_label.config(text=f"Обработано: {processed} из {total} ({percent}%)")
                
                # Визуальное выделение строки в таблице
                for item in self.tree.get_children():
                    if self.tree.item(item)['values'][0] == index + 1:
                        self.tree.selection_set(item)
                        self.tree.see(item) # Прокрутка к строке
                        break
                
                fio_current = str(row['ФИО обучающегося']).strip()
                self.status_var.set(f"Регистрируем: {fio_current}")
                driver.get(self.q_url.get())
                
                # 1. Выбор викторины
                btn_xp = f"//div[contains(., '{self.q_name.get()}')]/ancestor::div[contains(@class, 'white')]//button"
                quiz_btn = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xp)))
                driver.execute_script("arguments[0].click();", quiz_btn)

                # 2. ФИО
                wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Фамилия участника')]")))
                fio_parts = fio_current.split()
                self.safe_fill(wait, 'Фамилия участника', fio_parts[0])
                self.safe_fill(wait, 'Имя участника', fio_parts[1] if len(fio_parts)>1 else "-")
                self.safe_fill(wait, 'Отчество участника', fio_parts[2] if len(fio_parts)>2 else "-")

                # 3. Дата
                if not pd.isna(row['Дата рождения']):
                    try:
                        d_input = driver.find_element(By.XPATH, "//div[contains(text(), 'Дата рождения')]/following::input[1]")
                        d_val = pd.to_datetime(row['Дата рождения'], dayfirst=True).strftime('%Y-%m-%d')
                        driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'))", d_input, d_val)
                    except: pass

                # 4. Школа
                try:
                    org_in = self.safe_fill(wait, 'ОБУЧАЕТСЯ', row['Образовательная организация'])
                    time.sleep(0.5)
                    org_in.send_keys(Keys.ARROW_DOWN, Keys.ENTER)
                    
                    repr_in = self.safe_fill(wait, 'ПРЕДСТАВЛЯЕТ', 'ГБОУ "Воробьевы горы"')
                    time.sleep(0.5)
                    repr_in.send_keys(Keys.ARROW_DOWN, Keys.ENTER)
                except: pass

                # 5. Класс
                raw_class = str(row['Группа/Класс']).strip().lower()
                digit = "".join(filter(str.isdigit, raw_class))
                text_cl = f"{digit} курс" if 'курс' in raw_class else f"{digit} класс"
                
                try:
                    cl_in = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Класс/ курс')]/following::input[1]")))
                    driver.execute_script("arguments[0].click();", cl_in)
                    cl_in.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                    for char in text_cl: cl_in.send_keys(char)
                    time.sleep(0.5)
                    cl_in.send_keys(Keys.ENTER)
                except: pass

                # 6. СНИЛС
                try:
                    consent = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Согласен на внесение')]")))
                    driver.execute_script("arguments[0].click();", consent)
                    snils_xpath = "//div[contains(translate(text(), 'СНИЛС', 'снилс'), 'снилс')]/following::input[1]"
                    snils_field = wait.until(EC.visibility_of_element_located((By.XPATH, snils_xpath)))
                    snils_field.send_keys(str(row['СНИЛС']))
                except: pass

                # 7. Контакты
                self.safe_fill(wait, 'Телефон участника', self.t_phone.get())
                self.safe_fill(wait, 'Email участника', self.t_email.get())
                self.safe_fill(wait, 'Телефон родителя', row['Контактный телефон'])
                self.safe_fill(wait, 'Email родителя', row['e-mail'])

                # --- БЛОК СТРОГОЙ ВАЛИДАЦИИ ---
                while True:
                    time.sleep(1.5)
                    errors = []
                    
                    # 1. УНИВЕРСАЛЬНАЯ ПРОВЕРКА ОШИБОК САЙТА (Только активные ошибки)
                    # Ищем сообщения об ошибках, которые видны (обычно родитель имеет класс error--text)
                    all_messages = driver.find_elements(By.XPATH, "//div[contains(@class, 'v-messages__message')]")
                    for msg_el in all_messages:
                        err_text = msg_el.text.strip()
                        if not err_text:
                            continue
                            
                        is_visible = driver.execute_script(
                            "return window.getComputedStyle(arguments[0]).display !== 'none' && "
                            "arguments[0].offsetHeight > 0;", msg_el
                        )
                        
                        try:
                            parent_container = msg_el.find_element(By.XPATH, "./ancestor::div[contains(@class, 'v-input')]")
                            is_error_state = "error--text" in parent_container.get_attribute("class")
                        except:
                            is_error_state = True 

                        if is_visible and is_error_state:
                            errors.append(f"САЙТ: '{err_text}'")

                    # 2. Сверка данных: Фамилия
                    fact_fio = self.get_field_data(driver, 'Фамилия участника')
                    if not fact_fio or fio_parts[0].lower() not in fact_fio:
                        errors.append(f"Данные: ждали фамилию '{fio_parts[0]}', на сайте '{fact_fio}'")

                    # 3. Сверка данных: СНИЛС (сверка цифр)
                    fact_snils = "".join(filter(str.isdigit, self.get_field_data(driver, 'СНИЛС')))
                    excel_snils = "".join(filter(str.isdigit, str(row['СНИЛС'])))
                    if fact_snils != excel_snils:
                        errors.append(f"Данные: СНИЛС в Excel({excel_snils}) не совпадает с сайтом({fact_snils})")

                    # 4. Проверка школы (по номеру)
                    fact_org = self.get_field_data(driver, 'ОБУЧАЕТСЯ')
                    org_num = "".join(filter(str.isdigit, str(row['Образовательная организация'])))
                    if org_num and org_num not in fact_org:
                        errors.append(f"Данные: В поле школы нет номера '{org_num}'")

                    # Убираем дубликаты сообщений
                    errors = list(dict.fromkeys(errors))
                    
                    if not errors:
                        break 
                    else:
                        # Бот выведет ВСЕ ошибки: и несоответствие данных, и ругань самого сайта
                        msg = f"ОШИБКИ у {fio_current}:\n\n" + "\n".join(errors) + \
                              "\n\n1. Исправьте ошибки прямо в браузере.\n2. Нажмите 'Да', чтобы перепроверить и продолжить.\n3. 'Нет' - пропустить.\n4. 'Отмена' - стоп."
                        res = messagebox.askyesnocancel("Обнаружены ошибки", msg)
                        
                        if res is True: continue   # Рестарт цикла валидации
                        elif res is False: break   # Игнор
                        else: raise Exception("Прервано пользователем")

                # 8. Финал (Окна 1 -> 2 -> 3)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//label[contains(., 'персональных данных')]"))
                driver.execute_script("arguments[0].click();", driver.find_elements(By.TAG_NAME, "button")[-1])

                # Окно 2
                wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "Данные о подающем"))
                driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Педагог')]"))))
                self.safe_fill(wait, 'Фамилия Имя Отчество', self.t_fio.get())
                self.safe_fill(wait, 'Контактный телефон', self.t_phone.get())
                self.safe_fill(wait, 'Электронная почта', self.t_email.get())
                driver.execute_script("arguments[0].click();", driver.find_elements(By.TAG_NAME, "button")[-1])

                # Окно 3
                wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "ФИО педагога"))
                self.safe_fill(wait, 'ФИО педагога', self.t_fio.get())
                self.safe_fill(wait, 'Контактный телефон педагога', self.t_phone.get())
                self.safe_fill(wait, 'Email педагога', self.t_email.get())
                driver.execute_script("arguments[0].click();", driver.find_elements(By.TAG_NAME, "button")[-1])
                
                time.sleep(3)

                # Таймер
                elapsed = time.time() - start_time_all
                avg = elapsed / processed
                rem = int(avg * (total - processed))
                self.eta_label.config(text=f"Осталось: ~{rem // 60:02d}:{rem % 60:02d}")
                
                driver.refresh()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button")))

            messagebox.showinfo("Готово", "Все ученики зарегистрированы!")
        except Exception as e:
            if "Прервано пользователем" not in str(e):
                messagebox.showerror("Ошибка", f"Бот остановлен: {e}")
        finally:
            if driver: driver.quit()
            self.btn_run.config(state="normal", text="ЗАПУСТИТЬ РЕГИСТРАЦИЮ")

if __name__ == "__main__":
    root = tk.Tk(); app = QuizBotGUI(root); root.mainloop()
