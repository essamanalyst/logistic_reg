import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog
import pandas as pd
import os
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from scipy.stats import chi2
from statsmodels.discrete.discrete_model import Logit
from statsmodels.tools import add_constant
from docx import Document
import tempfile
import platform
import subprocess
import threading
import sqlalchemy

class LogisticPredictorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Logistic Regression Predictor")
        self.root.geometry("950x820")
        self.root.resizable(True, True)

        self.language = "en"
        self.df = None
        self.selected_features = []
        self.target_column = None

        self.create_menu()
        self.create_widgets()
        self.update_texts()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.get_text("File"), menu=file_menu)
        file_menu.add_command(label=self.get_text("Open File"), command=self.open_file)
        file_menu.add_command(label=self.get_text("Load from DB"), command=self.load_from_db_dialog)
        file_menu.add_separator()
        file_menu.add_command(label=self.get_text("Save as Excel"), command=self.save_as_excel)
        file_menu.add_command(label=self.get_text("Save as Word"), command=self.save_as_word)
        file_menu.add_command(label=self.get_text("Print Result"), command=self.print_result)
        file_menu.add_separator()
        file_menu.add_command(label=self.get_text("Exit"), command=self.root.quit)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.get_text("Edit"), menu=edit_menu)
        edit_menu.add_command(label=self.get_text("Undo"), command=self.undo_action)
        edit_menu.add_command(label=self.get_text("Copy"), command=self.copy_action)
        edit_menu.add_command(label=self.get_text("Paste"), command=self.paste_action)

        options_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.get_text("Options"), menu=options_menu)
        options_menu.add_command(label=self.get_text("Settings"), command=self.settings)

        about_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.get_text("About"), menu=about_menu)
        about_menu.add_command(label=self.get_text("About this app"), command=self.show_about)

        lang_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.get_text("Language"), menu=lang_menu)
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        lang_menu.add_command(label="عربي", command=lambda: self.change_language("ar"))

    def create_widgets(self):
        self.lbl_file = tk.Label(self.root, text="", font=("Arial", 10))
        self.lbl_file.pack(pady=5)

        tk.Label(self.root, text=self.get_text("Select Columns (Ctrl+Click for multi):")).pack()
        self.list_columns = tk.Listbox(self.root, selectmode=tk.MULTIPLE, height=15, exportselection=0)
        self.list_columns.pack(fill=tk.X, padx=10)

        self.btn_set_target = tk.Button(self.root, text=self.get_text("Set Target Column"), command=self.set_target_column)
        self.btn_set_target.pack(pady=8)

        self.lbl_target = tk.Label(self.root, text=self.get_text("Target Column: None"), font=("Arial", 10, "bold"))
        self.lbl_target.pack()

        self.missing_label = tk.Label(self.root, text=self.get_text("Handle Missing Data:"))
        self.missing_label.pack(pady=5)

        self.missing_var = tk.StringVar()
        self.combo_missing = ttk.Combobox(self.root, textvariable=self.missing_var, state="readonly",
                                          values=[self.get_text("Drop rows with missing data"), self.get_text("Fill missing with mean")])
        self.combo_missing.current(0)
        self.combo_missing.pack(pady=5)

        frm_buttons = tk.Frame(self.root)
        frm_buttons.pack(pady=10)

        self.btn_run_stats = tk.Button(frm_buttons, text=self.get_text("Run Statistical Tests"), command=self.run_statistics_threaded)
        self.btn_run_stats.grid(row=0, column=0, padx=10)

        self.btn_run_predict = tk.Button(frm_buttons, text=self.get_text("Run Prediction"), command=self.run_prediction_threaded)
        self.btn_run_predict.grid(row=0, column=1, padx=10)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 9), fg="blue")
        self.status_label.pack()

        self.txt_output = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, height=20, font=("Consolas", 10))
        self.txt_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.txt_output.bind("<KeyPress>", self.on_keypress)

    # ====================== ترجمة النصوص ======================
    def get_text(self, key):
        translations = {
            "en": {
                "Language": "Language",
                "File": "File",
                "Open File": "Open File",
                "Load from DB": "Load from DB",
                "Exit": "Exit",
                "Edit": "Edit",
                "Undo": "Undo",
                "Copy": "Copy",
                "Paste": "Paste",
                "Options": "Options",
                "Settings": "Settings",
                "About": "About",
                "About this app": "About this app",
                "Select Columns (Ctrl+Click for multi):": "Select Columns (Ctrl+Click for multi):",
                "Set Target Column": "Set Target Column",
                "Target Column: None": "Target Column: None",
                "Handle Missing Data:": "Handle Missing Data:",
                "Drop rows with missing data": "Drop rows with missing data",
                "Fill missing with mean": "Fill missing with mean",
                "Run Statistical Tests": "Run Statistical Tests",
                "Run Prediction": "Run Prediction",
                "Save as Excel": "Save as Excel",
                "Save as Word": "Save as Word",
                "Print Result": "Print Result",
                "No results to save": "No results to save",
                "No results to print": "No results to print",
                "File Saved": "File Saved",
                "Welcome to the Logistic Regression App": "Welcome to the Logistic Regression App",
                "File Selected": "File Selected",
                "You selected:": "You selected:",
                "Error": "Error",
                "Cannot undo.": "Cannot undo.",
                "Settings": "Settings",
                "Settings window or options here.": "Settings window or options here.",
                "This app performs Logistic Regression": "This app performs Logistic Regression",
                "Developed by: Essam Sabbah": "Developed by: Essam Sabbah",
                "This app uses logistic regression for data analysis.": "This app uses logistic regression for data analysis.",
                "For inquiries, contact via email: essam.fathi.sabbah@gmail.com": "For inquiries, contact via email: essam.fathi.sabbah@gmail.com",
                "Invalid feature columns": "Invalid feature columns",
                "Target column must be binary (0 and 1)": "Target column must be binary (0 and 1)",
                "Please select at least one column first": "Please select at least one column first",
                "Please select target column from selected columns": "Please select target column from selected columns",
                "Unsupported file format": "Unsupported file format",
                "Database Connection": "Database Connection",
                "Database Type:": "Database Type:",
                "Host / File Path:": "Host / File Path:",
                "Port (optional):": "Port (optional):",
                "Database Name:": "Database Name:",
                "Username:": "Username:",
                "Password:": "Password:",
                "SQL Query (optional):": "SQL Query (optional):",
                "Or select Table:": "Or select Table:",
                "Connect and Load": "Connect and Load",
                "Info": "Info",
                "Please enter SQL query or select a table to fetch data.": "Please enter SQL query or select a table to fetch data.",
            },
            "ar": {
                "Language": "اللغة",
                "File": "ملف",
                "Open File": "فتح ملف",
                "Load from DB": "تحميل من قاعدة بيانات",
                "Exit": "خروج",
                "Edit": "تحرير",
                "Undo": "تراجع",
                "Copy": "نسخ",
                "Paste": "لصق",
                "Options": "خيارات",
                "Settings": "الإعدادات",
                "About": "حول",
                "About this app": "حول هذا التطبيق",
                "Select Columns (Ctrl+Click for multi):": "اختر الأعمدة (اضغط Ctrl لاختيار أكثر):",
                "Set Target Column": "حدد العمود الهدف",
                "Target Column: None": "العمود الهدف: لا يوجد",
                "Handle Missing Data:": "التعامل مع القيم المفقودة:",
                "Drop rows with missing data": "حذف الصفوف التي تحتوي على قيم مفقودة",
                "Fill missing with mean": "تعويض القيم المفقودة بالوسط الحسابي",
                "Run Statistical Tests": "تشغيل الاختبارات الإحصائية",
                "Run Prediction": "تشغيل التنبؤ",
                "Save as Excel": "حفظ كملف إكسل",
                "Save as Word": "حفظ كملف وورد",
                "Print Result": "طباعة النتائج",
                "No results to save": "لا توجد نتائج للحفظ",
                "No results to print": "لا توجد نتائج للطباعة",
                "File Saved": "تم حفظ الملف",
                "Welcome to the Logistic Regression App": "مرحبًا بك في تطبيق الانحدار اللوجستي",
                "File Selected": "تم اختيار الملف",
                "You selected:": "لقد اخترت:",
                "Error": "خطأ",
                "Cannot undo.": "لا يمكن التراجع.",
                "Settings": "الإعدادات",
                "Settings window or options here.": "نافذة الإعدادات أو الخيارات هنا.",
                "This app performs Logistic Regression": "هذا التطبيق يقوم بتحليل الانحدار اللوجستي",
                "Developed by: Essam Sabbah": "تم التطوير بواسطة: عصام صباح",
                "This app uses logistic regression for data analysis.": "يستخدم هذا التطبيق الانحدار اللوجستي لتحليل البيانات.",
                "For inquiries, contact via email: essam.fathi.sabbah@gmail.com": "للاستفسارات، تواصل عبر البريد الإلكتروني: essam.fathi.sabbah@gmail.com",
                "Invalid feature columns": "أعمدة الميزات غير صحيحة",
                "Target column must be binary (0 and 1)": "العمود الهدف يجب أن يكون ثنائي (0 و 1)",
                "Please select at least one column first": "الرجاء اختيار عمود واحد على الأقل أولاً",
                "Please select target column from selected columns": "الرجاء تحديد العمود الهدف من الأعمدة المختارة",
                "Unsupported file format": "صيغة الملف غير مدعومة",
                "Database Connection": "اتصال بقاعدة البيانات",
                "Database Type:": "نوع قاعدة البيانات:",
                "Host / File Path:": "المضيف / مسار الملف:",
                "Port (optional):": "المنفذ (اختياري):",
                "Database Name:": "اسم القاعدة:",
                "Username:": "اسم المستخدم:",
                "Password:": "كلمة السر:",
                "SQL Query (optional):": "استعلام SQL (اختياري):",
                "Or select Table:": "أو اختر جدول:",
                "Connect and Load": "اتصال وتحميل",
                "Info": "معلومة",
                "Please enter SQL query or select a table to fetch data.": "الرجاء إدخال استعلام SQL أو اختيار جدول لجلب البيانات.",
            }
        }
        return translations[self.language].get(key, key)

    # ====================== تحديث النصوص ======================
    def update_texts(self):
        self.root.title(self.get_text("Logistic Regression Predictor"))
        self.lbl_file.config(text="")
        self.missing_label.config(text=self.get_text("Handle Missing Data:"))
        self.btn_run_stats.config(text=self.get_text("Run Statistical Tests"))
        self.btn_run_predict.config(text=self.get_text("Run Prediction"))
        self.btn_set_target.config(text=self.get_text("Set Target Column"))
        if self.target_column:
            self.lbl_target.config(text=f"{self.get_text('Target Column: None').replace('None', self.target_column)}")
        else:
            self.lbl_target.config(text=self.get_text("Target Column: None"))

    def change_language(self, lang):
        self.language = lang
        self.update_texts()

    # ====================== فتح ملفات البيانات ======================
    def open_file(self):
        filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("JSON files", "*.json"),
            ("Parquet files", "*.parquet"),
        ]
        filename = filedialog.askopenfilename(title=self.get_text("Open File"), filetypes=filetypes)
        if filename:
            try:
                ext = os.path.splitext(filename)[1].lower()
                if ext in ['.xlsx', '.xls']:
                    self.df = pd.read_excel(filename)
                elif ext == '.csv':
                    self.df = pd.read_csv(filename)
                elif ext == '.json':
                    self.df = pd.read_json(filename)
                elif ext == '.parquet':
                    self.df = pd.read_parquet(filename)
                else:
                    messagebox.showerror(self.get_text("Error"), self.get_text("Unsupported file format"))
                    return

                self.lbl_file.config(text=f"{self.get_text('File Selected')}: {filename}")

                cols = list(self.df.columns)
                self.list_columns.delete(0, tk.END)
                for c in cols:
                    self.list_columns.insert(tk.END, c)

                self.selected_features = []
                self.target_column = None
                self.lbl_target.config(text=self.get_text("Target Column: None"))

            except Exception as e:
                messagebox.showerror(self.get_text("Error"), str(e))

    # ====================== نافذة اتصال قواعد البيانات ======================
    def load_from_db_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.get_text("Database Connection"))
        dialog.geometry("480x460")
        dialog.grab_set()

        db_types = ["mysql", "postgresql", "sqlite", "mssql"]

        tk.Label(dialog, text=self.get_text("Database Type:")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        combo_db_type = ttk.Combobox(dialog, values=db_types, state="readonly")
        combo_db_type.current(0)
        combo_db_type.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Host / File Path:")).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        entry_host = tk.Entry(dialog)
        entry_host.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Port (optional):")).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        entry_port = tk.Entry(dialog)
        entry_port.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Database Name:")).grid(row=3, column=0, sticky="w", padx=5, pady=5)
        entry_dbname = tk.Entry(dialog)
        entry_dbname.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Username:")).grid(row=4, column=0, sticky="w", padx=5, pady=5)
        entry_user = tk.Entry(dialog)
        entry_user.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Password:")).grid(row=5, column=0, sticky="w", padx=5, pady=5)
        entry_pass = tk.Entry(dialog, show="*")
        entry_pass.grid(row=5, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("SQL Query (optional):")).grid(row=6, column=0, sticky="nw", padx=5, pady=5)
        text_query = tk.Text(dialog, height=4, width=40)
        text_query.grid(row=6, column=1, padx=5, pady=5)

        tk.Label(dialog, text=self.get_text("Or select Table:")).grid(row=7, column=0, sticky="w", padx=5, pady=5)
        combo_tables = ttk.Combobox(dialog, values=[], state="readonly")
        combo_tables.grid(row=7, column=1, padx=5, pady=5)

        def fetch_tables(event=None):
            dbtype = combo_db_type.get()
            host = entry_host.get()
            port = entry_port.get()
            dbname = entry_dbname.get()
            user = entry_user.get()
            password = entry_pass.get()

            try:
                if dbtype == "sqlite":
                    conn_str = f"sqlite:///{host}"
                else:
                    port_part = f":{port}" if port else ""
                    conn_str = f"{dbtype}://{user}:{password}@{host}{port_part}/{dbname}"

                engine = sqlalchemy.create_engine(conn_str)
                inspector = sqlalchemy.inspect(engine)
                tables = inspector.get_table_names()
                combo_tables['values'] = tables
                if tables:
                    combo_tables.current(0)

            except Exception as e:
                messagebox.showerror(self.get_text("Error"), str(e))

        # ربط تحديث الجداول بتغيير بيانات الاتصال
        combo_db_type.bind("<<ComboboxSelected>>", fetch_tables)
        entry_host.bind("<FocusOut>", fetch_tables)
        entry_port.bind("<FocusOut>", fetch_tables)
        entry_dbname.bind("<FocusOut>", fetch_tables)
        entry_user.bind("<FocusOut>", fetch_tables)
        entry_pass.bind("<FocusOut>", fetch_tables)

        def connect_and_load():
            dbtype = combo_db_type.get()
            host = entry_host.get()
            port = entry_port.get()
            dbname = entry_dbname.get()
            user = entry_user.get()
            password = entry_pass.get()
            query = text_query.get("1.0", tk.END).strip()
            table = combo_tables.get()

            try:
                if dbtype == "sqlite":
                    conn_str = f"sqlite:///{host}"
                else:
                    port_part = f":{port}" if port else ""
                    conn_str = f"{dbtype}://{user}:{password}@{host}{port_part}/{dbname}"

                engine = sqlalchemy.create_engine(conn_str)

                if query:
                    df = pd.read_sql(query, engine)
                elif table:
                    df = pd.read_sql_table(table, engine)
                else:
                    messagebox.showinfo(self.get_text("Info"), self.get_text("Please enter SQL query or select a table to fetch data."))
                    return

                self.df = df
                cols = list(df.columns)
                self.list_columns.delete(0, tk.END)
                for c in cols:
                    self.list_columns.insert(tk.END, c)
                self.selected_features = []
                self.target_column = None
                self.lbl_target.config(text=self.get_text("Target Column: None"))
                self.lbl_file.config(text=f"{self.get_text('File Selected')}: {dbtype} database")

                dialog.destroy()

            except Exception as e:
                messagebox.showerror(self.get_text("Error"), str(e))

        btn_connect = tk.Button(dialog, text=self.get_text("Connect and Load"), command=connect_and_load)
        btn_connect.grid(row=8, column=0, columnspan=2, pady=10)

    # ====================== تحديد العمود الهدف ======================
    def set_target_column(self):
        selected_indices = self.list_columns.curselection()
        if not selected_indices:
            messagebox.showerror(self.get_text("Error"), self.get_text("Please select at least one column first"))
            return

        selected_cols = [self.list_columns.get(i) for i in selected_indices]

        target = simpledialog.askstring(self.get_text("Set Target Column"),
                                        f"{self.get_text('Please select target column from selected columns')}:\n{selected_cols}")

        if target is None:
            return

        if target not in selected_cols:
            messagebox.showerror(self.get_text("Error"), self.get_text("Please select target column from selected columns"))
            return

        self.target_column = target
        self.selected_features = [col for col in selected_cols if col != target]

        self.lbl_target.config(text=f"{self.get_text('Target Column: None').replace('None', target)}")

    # ====================== تحقق صحة الأعمدة ======================
    def validate_columns(self, features, target):
        for col in features:
            if not pd.api.types.is_numeric_dtype(self.df[col]):
                raise ValueError(self.get_text("Invalid feature columns") + f": {col}")

        if not pd.api.types.is_numeric_dtype(self.df[target]):
            raise ValueError(self.get_text("Target column must be binary (0 and 1)"))

        unique_vals = self.df[target].dropna().unique()
        if set(unique_vals) != {0, 1} and set(unique_vals) != {1, 0}:
            raise ValueError(self.get_text("Target column must be binary (0 and 1)"))

    # ====================== اختبار هوسمر-ليمشو ======================
    def hosmer_lemeshow_test(self, model, X, y, groups=10):
        probs = model.predict(X)
        data = pd.DataFrame({'y': y, 'prob': probs})
        data['decile'] = pd.qcut(data['prob'], groups, duplicates='drop')
        obs = data.groupby('decile')['y'].sum()
        exp = data.groupby('decile')['prob'].sum()
        n = data.groupby('decile').size()
        hl_stat = ((obs - exp) ** 2 / (exp * (1 - exp / n))).sum()
        dof = groups - 2
        p_value = 1 - chi2.cdf(hl_stat, dof)
        return hl_stat, p_value

    # ====================== تشغيل التحليل في Thread ======================
    def run_statistics_threaded(self):
        threading.Thread(target=self.run_statistics, daemon=True).start()

    def run_prediction_threaded(self):
        threading.Thread(target=self.run_prediction, daemon=True).start()

    # ====================== التحليل الاحصائي ======================
    def run_statistics(self):
        if self.df is None:
            messagebox.showerror(self.get_text("Error"), self.get_text("Open File"))
            return

        if not self.selected_features or not self.target_column:
            messagebox.showerror(self.get_text("Error"), self.get_text("Please select at least one column first"))
            return

        try:
            self.set_status(self.get_text("Running Statistical Tests..."))
            self.validate_columns(self.selected_features, self.target_column)

            df_selected = self.df[self.selected_features + [self.target_column]].copy()

            if self.combo_missing.get() == self.get_text("Drop rows with missing data"):
                df_clean = df_selected.dropna()
            else:
                df_clean = df_selected.fillna(df_selected.mean(numeric_only=True))

            X = df_clean[self.selected_features]
            y = df_clean[self.target_column]
            X_const = add_constant(X, has_constant='add')

            model = Logit(y, X_const).fit(disp=0)

            hl_stat, p_value = self.hosmer_lemeshow_test(model, X_const, y)

            result_text = f"{self.get_text('Run Statistical Tests')}:\n"
            result_text += f"Hosmer-Lemeshow test statistic: {hl_stat:.4f}\n"
            result_text += f"Hosmer-Lemeshow p-value: {p_value:.4f}\n\n"
            result_text += "Model Summary:\n" + str(model.summary()) + "\n\n"
            result_text += "Statistical summary:\n" + str(df_clean.describe()) + "\n"

            self.txt_output.after(0, lambda: self.update_output(result_text))
            self.set_status("")

        except Exception as e:
            self.set_status("")
            messagebox.showerror(self.get_text("Error"), str(e))

    # ====================== التنبؤ ======================
    def run_prediction(self):
        if self.df is None:
            messagebox.showerror(self.get_text("Error"), self.get_text("Open File"))
            return

        if not self.selected_features or not self.target_column:
            messagebox.showerror(self.get_text("Error"), self.get_text("Please select at least one column first"))
            return

        try:
            self.set_status(self.get_text("Running Prediction..."))
            self.validate_columns(self.selected_features, self.target_column)

            df_selected = self.df[self.selected_features + [self.target_column]].copy()

            if self.combo_missing.get() == self.get_text("Drop rows with missing data"):
                df_clean = df_selected.dropna()
            else:
                df_clean = df_selected.fillna(df_selected.mean(numeric_only=True))

            X = df_clean[self.selected_features]
            y = df_clean[self.target_column]

            X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

            model = LogisticRegression(max_iter=1000)
            model.fit(X_train, y_train)

            y_pred = model.predict(X_test)

            acc = accuracy_score(y_test, y_pred)
            report = classification_report(y_test, y_pred)
            conf_matrix = confusion_matrix(y_test, y_pred)

            result_text = f"{self.get_text('Run Prediction')}:\n"
            result_text += f"Accuracy: {acc:.4f}\n\n"
            result_text += "Classification Report:\n" + report + "\n"
            result_text += "Confusion Matrix:\n" + str(conf_matrix)

            self.txt_output.after(0, lambda: self.update_output(result_text))
            self.set_status("")

        except Exception as e:
            self.set_status("")
            messagebox.showerror(self.get_text("Error"), str(e))

    # ====================== تحديث مخرجات النص ======================
    def update_output(self, text):
        self.txt_output.delete('1.0', tk.END)
        self.txt_output.insert(tk.END, text)

    # ====================== حفظ النتائج ======================
    def save_as_excel(self):
        text = self.txt_output.get('1.0', tk.END).strip()
        if not text:
            messagebox.showerror(self.get_text("Error"), self.get_text("No results to save"))
            return

        filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title=self.get_text("Save as Excel"))
        if not filename:
            return

        try:
            df = pd.DataFrame(text.splitlines(), columns=["Result"])
            df.to_excel(filename, index=False)
            messagebox.showinfo(self.get_text("File Saved"), f"{self.get_text('File Saved')}: {filename}")
        except Exception as e:
            messagebox.showerror(self.get_text("Error"), str(e))

    def save_as_word(self):
        text = self.txt_output.get('1.0', tk.END).strip()
        if not text:
            messagebox.showerror(self.get_text("Error"), self.get_text("No results to save"))
            return

        filename = filedialog.asksaveasfilename(defaultextension=".docx",
                                                filetypes=[("Word files", "*.docx")],
                                                title=self.get_text("Save as Word"))
        if not filename:
            return

        try:
            doc = Document()
            for line in text.splitlines():
                doc.add_paragraph(line)
            doc.save(filename)
            messagebox.showinfo(self.get_text("File Saved"), f"{self.get_text('File Saved')}: {filename}")
        except Exception as e:
            messagebox.showerror(self.get_text("Error"), str(e))

    # ====================== طباعة النتائج ======================
    def print_result(self):
        text = self.txt_output.get('1.0', tk.END).strip()
        if not text:
            messagebox.showerror(self.get_text("Error"), self.get_text("No results to print"))
            return

        try:
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
            tmp_file.write(text.encode('utf-8'))
            tmp_file.close()

            if platform.system() == "Windows":
                os.startfile(tmp_file.name, "print")
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["lp", tmp_file.name])
            else:  # Linux / Unix
                subprocess.run(["lp", tmp_file.name])
        except Exception as e:
            messagebox.showerror(self.get_text("Error"), str(e))

    # ====================== تحرير النص ======================
    def undo_action(self):
        try:
            self.txt_output.edit_undo()
        except Exception:
            messagebox.showerror(self.get_text("Error"), self.get_text("Cannot undo."))

    def copy_action(self):
        self.txt_output.event_generate("<<Copy>>")

    def paste_action(self):
        self.txt_output.event_generate("<<Paste>>")

    def on_keypress(self, event):
        self.txt_output.edit_modified(False)

    def settings(self):
        messagebox.showinfo(self.get_text("Settings"), self.get_text("Settings window or options here."))

    def show_about(self):
        about_text = f"""
{self.get_text("This app performs Logistic Regression")}
{self.get_text("Developed by: Essam Sabbah")}

{self.get_text("This app uses logistic regression for data analysis.")}
{self.get_text("For inquiries, contact via email: essam.fathi.sabbah@gmail.com")}
"""
        messagebox.showinfo(self.get_text("About this app"), about_text)

    # ====================== تحديث حالة البرنامج ======================
    def set_status(self, message):
        def clear():
            self.status_label.config(text="")
        self.status_label.config(text=message)
        self.root.after(5000, clear)

if __name__ == "__main__":
    root = tk.Tk()
    app = LogisticPredictorApp(root)
    root.mainloop()
