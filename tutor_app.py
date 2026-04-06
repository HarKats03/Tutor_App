import customtkinter as ctk
import sqlite3
from datetime import datetime
import pandas as pd
import os
import platform
import calendar
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image, ImageDraw, ImageTk

# --- ΕΙΔΙΚΗ ΡΥΘΜΙΣΗ ΓΙΑ ΤΗΝ TASKBAR ΤΩΝ WINDOWS ---
if platform.system() == 'Windows':
    import ctypes

    try:
        app_id = 'tutor.manager.pro.final.v5'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


# --- 1. ΑΥΤΟΜΑΤΗ ΔΗΜΙΟΥΡΓΙΑ ΛΟΓΟΤΥΠΟΥ ---
def ensure_logo_exists():
    if not os.path.exists("logo.png") or not os.path.exists("logo.ico"):
        size = (512, 512)
        img = Image.new('RGBA', size, (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        x0, y0, x1, y1 = 32, 32, 480, 480
        draw.rounded_rectangle((x0, y0, x1, y1), radius=110, fill="#0A84FF")
        draw.rectangle([216, 140, 296, 400], fill="white")
        draw.rectangle([130, 140, 382, 210], fill="white")
        img.save("logo.png")
        img.save("logo.ico", format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32)])


def apply_window_icon(window):
    try:
        if platform.system() == "Windows" and os.path.exists("logo.ico"):
            window.iconbitmap("logo.ico")
        else:
            icon_img = ImageTk.PhotoImage(Image.open("logo.png"))
            window.iconphoto(True, icon_img)
    except Exception:
        pass


# --- 2. CUSTOM ALERT DIALOG ---
def show_custom_alert(parent, title, message, is_error=False, callback=None):
    dialog = ctk.CTkToplevel(parent)
    dialog.title(title)
    dialog.geometry("450x250")
    dialog.transient(parent)
    dialog.grab_set()

    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
    y = parent.winfo_y() + (parent.winfo_height() // 2) - (250 // 2)
    dialog.geometry(f"+{x}+{y}")

    logo_img = ctk.CTkImage(Image.open("logo.png"), size=(60, 60))
    ctk.CTkLabel(dialog, text="", image=logo_img).pack(pady=(20, 10))

    txt_color = "#e74c3c" if is_error else "#ffffff"
    ctk.CTkLabel(dialog, text=title, font=("Arial", 18, "bold"), text_color=txt_color).pack(pady=5)
    ctk.CTkLabel(dialog, text=message, font=("Arial", 14), wraplength=400).pack(pady=5)

    def on_click():
        dialog.destroy()
        if callback: callback()

    btn_color = "#e74c3c" if is_error else "#27ae60"
    ctk.CTkButton(dialog, text="ΟΚ", fg_color=btn_color, width=120, command=on_click).pack(pady=15)


# --- 3. ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ ---
def init_db():
    conn = sqlite3.connect("tutor_manager.db")
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS students (
                        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, 
                        group_name TEXT, rate_per_hour REAL NOT NULL, hours_per_session REAL NOT NULL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS schedule (
                        id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, 
                        day_of_week TEXT NOT NULL, FOREIGN KEY(student_id) REFERENCES students(id))''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS session_logs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, 
                        date TEXT NOT NULL, hours_done REAL NOT NULL, earned_amount REAL NOT NULL, 
                        FOREIGN KEY(student_id) REFERENCES students(id))''')
    conn.commit()
    conn.close()


# --- 4. Η ΚΥΡΙΑ ΕΦΑΡΜΟΓΗ ---
class TutorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ====== ΤΟ ΜΥΣΤΙΚΟ ΓΙΑ ΝΑ ΜΗΝ ΕΧΟΥΜΕ FLICKERING ΟΥΤΕ ERRORS ======
        # Κάνουμε την εφαρμογή 100% αόρατη κατά τη φόρτωση του UI!
        self.attributes("-alpha", 0.0)
        self.withdraw()

        self.title("Tutor Manager Pro")
        self.geometry("1300x850")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        apply_window_icon(self)

        self.day_finalized = False
        self.protocol("WM_DELETE_WINDOW", self.on_closing_app)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.logo_img = ctk.CTkImage(Image.open("logo.png"), size=(80, 80))

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color="#1a1c23")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="", image=self.logo_img).pack(pady=(30, 10))
        ctk.CTkLabel(self.sidebar, text="TutorManager", font=("Montserrat", 20, "bold"), text_color="#0A84FF").pack(
            pady=(0, 30))

        self.btn_calendar = ctk.CTkButton(self.sidebar, text="📅 Έλεγχος & Ημερολόγιο", font=("Arial", 15), height=45,
                                          fg_color="#2980b9", command=self.show_calendar_view)
        self.btn_calendar.pack(pady=10, padx=20, fill="x")

        self.btn_add_student = ctk.CTkButton(self.sidebar, text="👥 Προσθήκη / Πρόγραμμα", font=("Arial", 15), height=45,
                                             fg_color="#34495e", command=self.show_add_student_ui)
        self.btn_add_student.pack(pady=10, padx=20, fill="x")

        self.btn_export = ctk.CTkButton(self.sidebar, text="📊 Εξαγωγή Excel", font=("Arial", 15), height=45,
                                        fg_color="#27ae60", hover_color="#219150", command=self.export_excel)
        self.btn_export.pack(pady=10, padx=20, fill="x")

        self.btn_summary = ctk.CTkButton(self.sidebar, text="✅ Κλείσιμο Ημέρας", font=("Arial", 15, "bold"), height=45,
                                         fg_color="#d35400", hover_color="#e67e22", command=self.show_daily_summary)
        self.btn_summary.pack(pady=40, padx=20, fill="x", side="bottom")

        self.calendar_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.add_student_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")

        self.setup_add_student_ui()
        self.show_calendar_view()

        # --- Εκκίνηση Splash Screen ---
        self.show_splash()

    # ==================== SPLASH SCREEN LOGIC ====================
    def show_splash(self):
        self.splash = ctk.CTkToplevel(self)
        self.splash.title("Φόρτωση...")
        self.splash.geometry("450x300")
        self.splash.overrideredirect(True)  # Χωρίς περίγραμμα παραθύρου

        # Κεντράρισμα Splash Screen
        self.splash.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.winfo_screenheight() // 2) - (300 // 2)
        self.splash.geometry(f"+{x}+{y}")

        self.splash.configure(fg_color="#1a1c23")
        logo = ctk.CTkImage(Image.open("logo.png"), size=(100, 100))
        ctk.CTkLabel(self.splash, text="", image=logo).pack(pady=(40, 10))
        ctk.CTkLabel(self.splash, text="TUTOR MANAGER PRO", font=("Montserrat", 22, "bold"),
                     text_color="#0A84FF").pack()

        self.lbl_status = ctk.CTkLabel(self.splash, text="Εκκίνηση συστημάτων...", font=("Arial", 14),
                                       text_color="gray")
        self.lbl_status.pack(pady=20)

        # Εξέλιξη φόρτωσης (δεν προκαλεί errors, γιατί ανήκει στο main loop του app)
        self.after(5000, lambda: self.lbl_status.configure(text="Σύνδεση στη βάση δεδομένων..."))
        self.after(8000, lambda: self.lbl_status.configure(text="Φόρτωση γραφικών..."))
        self.after(12000, self.finish_splash)

    def finish_splash(self):
        self.splash.destroy()  # Σκοτώνουμε το Loading screen

        # Κεντράρουμε το κεντρικό παράθυρο πριν το εμφανίσουμε
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (1300 // 2)
        y = (self.winfo_screenheight() // 2) - (850 // 2)
        self.geometry(f"+{x}+{y}")

        self.deiconify()  # Εμφανίζει το παράθυρο
        self.attributes("-alpha", 1.0)  # Το κάνει 100% ορατό (Σπάει την αορατότητα)

    # ==============================================================

    def hide_all_frames(self):
        self.calendar_frame.grid_forget()
        self.add_student_frame.grid_forget()

    # ==================== ΟΘΟΝΗ ΗΜΕΡΟΛΟΓΙΟΥ & ΗΜΕΡΗΣΙΟΥ ΕΛΕΓΧΟΥ ====================
    def show_calendar_view(self):
        self.hide_all_frames()
        self.calendar_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        for w in self.calendar_frame.winfo_children(): w.destroy()

        self.calendar_frame.grid_columnconfigure(0, weight=2)
        self.calendar_frame.grid_columnconfigure(1, weight=3)

        left_panel = ctk.CTkFrame(self.calendar_frame, fg_color="#232731", corner_radius=15)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        cal_header = ctk.CTkFrame(left_panel, fg_color="transparent")
        cal_header.pack(fill="x", pady=15, padx=20)

        current_month = datetime.today().month
        current_year = datetime.today().year

        self.combo_month = ctk.CTkOptionMenu(cal_header, values=[f"{i:02d}" for i in range(1, 13)], width=70,
                                             command=lambda _: self.build_calendar_grid())
        self.combo_month.set(f"{current_month:02d}")
        self.combo_month.pack(side="left", padx=5)

        self.combo_year = ctk.CTkOptionMenu(cal_header, values=[str(y) for y in range(2024, 2030)], width=90,
                                            command=lambda _: self.build_calendar_grid())
        self.combo_year.set(str(current_year))
        self.combo_year.pack(side="left", padx=5)

        self.cal_grid_container = ctk.CTkFrame(left_panel, fg_color="transparent")
        self.cal_grid_container.pack(fill="both", expand=True, padx=15, pady=10)

        self.right_panel = ctk.CTkFrame(self.calendar_frame, fg_color="#1a1c23", corner_radius=15)
        self.right_panel.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self.lbl_day_header = ctk.CTkLabel(self.right_panel, text="", font=("Arial", 22, "bold"), text_color="#f1c40f")
        self.lbl_day_header.pack(pady=15)

        self.scroll_area = ctk.CTkScrollableFrame(self.right_panel, fg_color="transparent")
        self.scroll_area.pack(fill="both", expand=True, padx=10, pady=5)

        self.frame_quick_add = ctk.CTkFrame(self.right_panel, fg_color="#1a1c23", border_width=1,
                                            border_color="#34495e")
        self.frame_quick_add.pack(fill="x", pady=(10, 10), padx=10, ipady=10)

        self.build_calendar_grid()
        self.select_day(datetime.today().year, datetime.today().month, datetime.today().day)

    def build_calendar_grid(self):
        for w in self.cal_grid_container.winfo_children(): w.destroy()

        year = int(self.combo_year.get())
        month = int(self.combo_month.get())

        days_titles = ["Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ", "Κυρ"]
        for i, d in enumerate(days_titles):
            ctk.CTkLabel(self.cal_grid_container, text=d, font=("Arial", 14, "bold"), text_color="#0A84FF").grid(row=0,
                                                                                                                 column=i,
                                                                                                                 pady=5)

        cal_matrix = calendar.monthcalendar(year, month)
        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()

        for r, week in enumerate(cal_matrix):
            for c, day in enumerate(week):
                if day != 0:
                    date_str = f"{year}-{month:02d}-{day:02d}"
                    cursor.execute("SELECT SUM(hours_done) FROM session_logs WHERE date=?", (date_str,))
                    hours_sum = cursor.fetchone()[0]

                    bg_color = "#27ae60" if hours_sum and hours_sum > 0 else "#2c3e50"
                    is_today = (
                                day == datetime.today().day and month == datetime.today().month and year == datetime.today().year)
                    border = 2 if is_today else 0
                    border_color = "#f1c40f" if is_today else None

                    text_display = f"{day}\n({hours_sum}h)" if hours_sum else str(day)

                    btn = ctk.CTkButton(self.cal_grid_container, text=text_display, width=50, height=50,
                                        fg_color=bg_color, font=("Arial", 13, "bold"),
                                        border_width=border, border_color=border_color,
                                        command=lambda d=day, m=month, y=year: self.select_day(y, m, d))
                    btn.grid(row=r + 1, column=c, padx=3, pady=3)
        conn.close()

    def select_day(self, year, month, day):
        self.selected_date = f"{year}-{month:02d}-{day:02d}"
        date_obj = datetime(year, month, day)
        days_gr = ["Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο", "Κυριακή"]
        self.selected_day_name = days_gr[date_obj.weekday()]

        self.lbl_day_header.configure(text=f"Ημερήσιος Έλεγχος: {self.selected_day_name} {day}/{month}/{year}")
        self.refresh_day_lists()
        self.setup_quick_add()

    def refresh_day_lists(self):
        for w in self.scroll_area.winfo_children(): w.destroy()

        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()

        ctk.CTkLabel(self.scroll_area, text="📋 Προγραμματισμένα για σήμερα:", font=("Arial", 16, "bold")).pack(
            anchor="w", pady=(0, 5))

        cursor.execute('''SELECT s.id, s.name, s.group_name, s.rate_per_hour, s.hours_per_session 
                          FROM students s JOIN schedule sch ON s.id = sch.student_id WHERE sch.day_of_week = ?''',
                       (self.selected_day_name,))
        scheduled_students = cursor.fetchall()

        scheduled_ids = []

        if not scheduled_students:
            ctk.CTkLabel(self.scroll_area, text="Κανένα προγραμματισμένο μάθημα.", text_color="gray").pack(anchor="w",
                                                                                                           padx=10)
        else:
            for sid, name, gname, rate, hours in scheduled_students:
                scheduled_ids.append(sid)
                cursor.execute("SELECT id, hours_done FROM session_logs WHERE student_id=? AND date=?",
                               (sid, self.selected_date))
                exists = cursor.fetchone()

                f = ctk.CTkFrame(self.scroll_area, fg_color="#2b323d", corner_radius=10)
                f.pack(fill="x", pady=5, ipady=5)

                title = f"🎓 {name} {'[' + gname + ']' if gname else '[Ατομικό]'}"
                ctk.CTkLabel(f, text=title, font=("Arial", 15, "bold")).pack(side="left", padx=15)

                if exists:
                    log_id, hours_logged = exists
                    ctk.CTkButton(f, text="↺ Αναίρεση", fg_color="#e74c3c", width=80,
                                  command=lambda l=log_id: self.delete_log(l)).pack(side="right", padx=15)
                    ctk.CTkLabel(f, text=f"✔️ {hours_logged}h", text_color="#2ecc71", font=("Arial", 14, "bold")).pack(
                        side="right", padx=10)
                else:
                    e_h = ctk.CTkEntry(f, width=50)
                    e_h.insert(0, str(hours))

                    btn = ctk.CTkButton(f, text="Επιβεβαίωση", fg_color="#27ae60", width=100,
                                        command=lambda s=sid, r=rate, e=e_h: self.save_lesson(s, r, e))
                    btn.pack(side="right", padx=15)
                    e_h.pack(side="right", padx=5)
                    ctk.CTkLabel(f, text="Ώρες:").pack(side="right")

        cursor.execute('''SELECT l.id, s.name, s.group_name, l.hours_done, s.id 
                          FROM session_logs l JOIN students s ON l.student_id = s.id 
                          WHERE l.date = ?''', (self.selected_date,))
        all_logs_today = cursor.fetchall()
        conn.close()

        extra_logs = [log for log in all_logs_today if log[4] not in scheduled_ids]

        if extra_logs:
            ctk.CTkLabel(self.scroll_area, text="⚠️ Έκτακτες Καταχωρήσεις / Αναπληρώσεις:", font=("Arial", 16, "bold"),
                         text_color="#e67e22").pack(anchor="w", pady=(20, 5))
            for lid, name, gname, hrs, _ in extra_logs:
                f = ctk.CTkFrame(self.scroll_area, fg_color="#161b22")
                f.pack(fill="x", pady=3)
                ctk.CTkLabel(f, text=f"{name} {'[' + gname + ']' if gname else ''} | {hrs}h", font=("Arial", 13)).pack(
                    side="left", padx=10)
                ctk.CTkButton(f, text="🗑️", width=30, fg_color="#c0392b",
                              command=lambda log_id=lid: self.delete_log(log_id)).pack(side="right", padx=15)

    def setup_quick_add(self):
        for w in self.frame_quick_add.winfo_children(): w.destroy()

        ctk.CTkLabel(self.frame_quick_add, text="Προσθήκη Εκτάκτου Μαθήματος:", text_color="gray").pack(pady=(0, 5))

        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, group_name, rate_per_hour FROM students")
        all_students = cursor.fetchall()
        conn.close()

        if all_students:
            self.student_options = {f"{s[1]} ({s[2] if s[2] else 'Ατομικό'})": s for s in all_students}
            dd_frame = ctk.CTkFrame(self.frame_quick_add, fg_color="transparent")
            dd_frame.pack()

            self.student_dropdown = ctk.CTkOptionMenu(dd_frame, values=list(self.student_options.keys()), width=200)
            self.student_dropdown.pack(side="left", padx=5)
            self.entry_quick_hours = ctk.CTkEntry(dd_frame, width=50, placeholder_text="Ώρες")
            self.entry_quick_hours.pack(side="left", padx=5)
            ctk.CTkButton(dd_frame, text="➕", width=40, fg_color="#2980b9", command=self.add_extra_session).pack(
                side="left", padx=5)

    def save_lesson(self, sid, rate, entry_widget):
        try:
            hours = float(entry_widget.get())
            conn = sqlite3.connect("tutor_manager.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO session_logs (student_id, date, hours_done, earned_amount) VALUES (?,?,?,?)",
                           (sid, self.selected_date, hours, hours * rate))
            conn.commit()
            conn.close()
            self.day_finalized = False
            self.refresh_day_lists()
            self.build_calendar_grid()
        except ValueError:
            show_custom_alert(self, "Σφάλμα", "Εισάγετε έγκυρο αριθμό ωρών.", is_error=True)

    def add_extra_session(self):
        selected = self.student_dropdown.get()
        s_data = self.student_options[selected]
        sid, _, _, rate = s_data
        try:
            hours = float(self.entry_quick_hours.get())
            conn = sqlite3.connect("tutor_manager.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO session_logs (student_id, date, hours_done, earned_amount) VALUES (?,?,?,?)",
                           (sid, self.selected_date, hours, hours * rate))
            conn.commit()
            conn.close()
            self.day_finalized = False
            self.refresh_day_lists()
            self.build_calendar_grid()
        except ValueError:
            show_custom_alert(self, "Σφάλμα", "Εισάγετε αριθμό ωρών.", is_error=True)

    def delete_log(self, log_id):
        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM session_logs WHERE id=?", (log_id,))
        conn.commit()
        conn.close()
        self.day_finalized = False
        self.refresh_day_lists()
        self.build_calendar_grid()

    # ==================== POP-UP ΚΛΕΙΣΙΜΑΤΟΣ ΗΜΕΡΑΣ ====================
    def show_daily_summary(self, is_closing=False):
        today_date = datetime.today().strftime('%Y-%m-%d')

        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()
        cursor.execute('''SELECT s.name, s.group_name, l.hours_done 
                          FROM session_logs l JOIN students s ON l.student_id = s.id 
                          WHERE l.date = ?''', (today_date,))
        completed = cursor.fetchall()
        conn.close()

        if not completed:
            if is_closing:
                self.destroy()
            else:
                show_custom_alert(self, "Πληροφορία", "Δεν υπάρχουν καθόλου καταχωρήσεις για τη σημερινή μέρα.")
            return

        popup = ctk.CTkToplevel(self)
        popup.title("Τελικός Έλεγχος Ημέρας")
        popup.geometry("550x550")
        popup.transient(self)
        popup.grab_set()

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (550 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (550 // 2)
        popup.geometry(f"+{x}+{y}")

        if is_closing:
            ctk.CTkLabel(popup, text="⚠️ ΠΡΟΣΟΧΗ!", font=("Arial", 20, "bold"), text_color="#e74c3c").pack(pady=(20, 0))
            ctk.CTkLabel(popup, text="Πας να κλείσεις το πρόγραμμα χωρίς να έχεις επιβεβαιώσει τη μέρα σου.",
                         font=("Arial", 14)).pack(pady=5)
        else:
            ctk.CTkLabel(popup, text="📋 Σύνοψη Σημερινής Ημέρας", font=("Arial", 22, "bold"),
                         text_color="#0A84FF").pack(pady=20)

        scroll_popup = ctk.CTkScrollableFrame(popup, fg_color="#1f252e")
        scroll_popup.pack(fill="both", expand=True, padx=20, pady=10)

        for name, gname, hours in completed:
            g_text = f"[{gname}]" if gname else "[Ατομικό]"
            row = ctk.CTkFrame(scroll_popup, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text=f"{name} {g_text}", font=("Arial", 15)).pack(side="left")
            ctk.CTkLabel(row, text=f"{hours} ώρες", font=("Arial", 15, "bold"), text_color="#f1c40f").pack(side="right")

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="🔙 Επιστροφή", fg_color="#7f8c8d", hover_color="#95a5a6", height=40,
                      command=popup.destroy).pack(side="left", padx=10)

        ctk.CTkButton(btn_frame, text="✅ Οριστική Αποθήκευση", fg_color="#27ae60", hover_color="#219150", height=40,
                      command=lambda: self.finalize_day(popup, is_closing)).pack(side="left", padx=10)

    def finalize_day(self, popup, is_closing):
        self.day_finalized = True
        popup.destroy()
        if is_closing:
            self.destroy()
        else:
            show_custom_alert(self, "Μπράβο!", "Η σημερινή ημέρα επιβεβαιώθηκε και αποθηκεύτηκε επιτυχώς.")

    def on_closing_app(self):
        today_date = datetime.today().strftime('%Y-%m-%d')
        conn = sqlite3.connect("tutor_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM session_logs WHERE date=?", (today_date,))
        logs_today = cursor.fetchone()[0]
        conn.close()

        if logs_today > 0 and not self.day_finalized:
            self.show_daily_summary(is_closing=True)
        else:
            self.destroy()

            # ==================== ΠΡΟΣΘΗΚΗ ΜΑΘΗΤΩΝ ====================

    def setup_add_student_ui(self):
        card = ctk.CTkFrame(self.add_student_frame, corner_radius=15, fg_color="#232731")
        card.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(card, text="Προσθήκη Μαθητή / Γκρουπ & Προγράμματος", font=("Arial", 22, "bold")).pack(pady=20)

        self.group_var = ctk.BooleanVar(value=False)
        self.cb_is_group = ctk.CTkCheckBox(card, text="Το μάθημα είναι σε Γκρουπ;", font=("Arial", 14),
                                           variable=self.group_var, command=self.toggle_group_fields)
        self.cb_is_group.pack(pady=10)

        self.entry_group = ctk.CTkEntry(card, placeholder_text="Όνομα Γκρουπ (π.χ. Β' Λυκείου)", width=400, height=40)

        self.students_container = ctk.CTkFrame(card, fg_color="transparent")
        self.students_container.pack(pady=10)
        self.student_entries = []
        self.add_student_field()

        self.btn_add_more = ctk.CTkButton(card, text="+ Προσθήκη Μαθητή στο Γκρουπ", fg_color="#555",
                                          hover_color="#444", command=self.add_student_field)

        self.entry_rate = ctk.CTkEntry(card, placeholder_text="Χρέωση ανά ώρα ανά μαθητή (€)", width=400, height=40)
        self.entry_rate.pack(pady=10)
        self.entry_def_hours = ctk.CTkEntry(card, placeholder_text="Συνήθης διάρκεια (π.χ. 1.5)", width=400, height=40)
        self.entry_def_hours.pack(pady=10)

        ctk.CTkLabel(card, text="Επιλογή Ημερών Μαθήματος:", font=("Arial", 16, "bold")).pack(pady=(20, 5))
        self.days_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.days_frame.pack(pady=5)
        self.days_dict = {}
        for d in ["Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο", "Κυριακή"]:
            v = ctk.BooleanVar()
            cb = ctk.CTkCheckBox(self.days_frame, text=d, variable=v)
            cb.pack(side="left", padx=10)
            self.days_dict[d] = v

        ctk.CTkButton(card, text="Αποθήκευση στο Σύστημα", fg_color="#27ae60", height=45,
                      command=self.save_student_to_db).pack(pady=30)

    def toggle_group_fields(self):
        if self.group_var.get():
            self.entry_group.pack(after=self.cb_is_group, pady=10)
            self.btn_add_more.pack(after=self.students_container, pady=5)
        else:
            self.entry_group.pack_forget()
            self.btn_add_more.pack_forget()
            while len(self.student_entries) > 1:
                self.student_entries.pop().destroy()

    def add_student_field(self):
        en = ctk.CTkEntry(self.students_container, placeholder_text=f"Όνομα Μαθητή {len(self.student_entries) + 1}",
                          width=400, height=40)
        en.pack(pady=5)
        self.student_entries.append(en)

    def save_student_to_db(self):
        group_name = self.entry_group.get() if self.group_var.get() else ""
        rate = self.entry_rate.get()
        hours = self.entry_def_hours.get()
        sel_days = [d for d, v in self.days_dict.items() if v.get()]

        names = [e.get().strip() for e in self.student_entries if e.get().strip()]

        if not names or not rate or not sel_days:
            show_custom_alert(self, "Προσοχή", "Συμπληρώστε Όνομα, Χρέωση και τουλάχιστον ΜΙΑ Ημέρα.", is_error=True)
            return

        try:
            conn = sqlite3.connect("tutor_manager.db")
            cursor = conn.cursor()
            for n in names:
                cursor.execute(
                    "INSERT INTO students (name, group_name, rate_per_hour, hours_per_session) VALUES (?,?,?,?)",
                    (n, group_name, float(rate), float(hours)))
                sid = cursor.lastrowid
                for day in sel_days:
                    cursor.execute("INSERT INTO schedule (student_id, day_of_week) VALUES (?,?)", (sid, day))
            conn.commit()
            conn.close()

            show_custom_alert(self, "Επιτυχία", "Οι καταχωρήσεις αποθηκεύτηκαν!")

            self.entry_group.delete(0, 'end')
            self.entry_rate.delete(0, 'end')
            self.entry_def_hours.delete(0, 'end')
            for var in self.days_dict.values(): var.set(False)
            for e in self.student_entries: e.delete(0, 'end')

        except ValueError:
            show_custom_alert(self, "Σφάλμα", "Η χρέωση και οι ώρες πρέπει να είναι αριθμοί.", is_error=True)

    def show_add_student_ui(self):
        self.hide_all_frames()
        self.add_student_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

    # ==================== EXCEL EXPORT ====================
    def export_excel(self):
        conn = sqlite3.connect("tutor_manager.db")
        query = '''SELECT s.name AS 'Μαθητής', IFNULL(s.group_name, 'Ατομικό') AS 'Γκρουπ',
                   strftime('%Y-%m', l.date) AS 'Μήνας', SUM(l.hours_done) AS 'Σύνολο Ωρών', 
                   SUM(l.earned_amount) AS 'Οφειλόμενο Ποσό'
                   FROM session_logs l JOIN students s ON l.student_id = s.id 
                   GROUP BY s.id, strftime('%Y-%m', l.date) ORDER BY 'Μήνας' DESC, 'Γκρουπ', 'Μαθητής' '''
        df = pd.read_sql_query(query, conn)
        conn.close()

        if df.empty:
            show_custom_alert(self, "Άδειο", "Δεν υπάρχουν δεδομένα μαθημάτων για εξαγωγή.", is_error=True)
            return

        filename = "Αναφορά_Εσόδων_Μαθητών.xlsx"

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Αναλυτικά')
            ws = writer.sheets['Αναλυτικά']

            header_fill = PatternFill(start_color="0A84FF", end_color="0A84FF", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=12)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                                 bottom=Side(style='thin'))

            for cell in ws["1:1"]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            fill_light = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
            fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

            for row_idx, row in enumerate(
                    ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
                row_fill = fill_light if row_idx % 2 == 0 else fill_white
                for cell in row:
                    cell.fill = row_fill
                    cell.border = thin_border
                    if cell.column_letter == 'E':
                        cell.number_format = '#,##0.00 €'
                    elif cell.column_letter == 'D':
                        cell.alignment = Alignment(horizontal="center")

            max_row = ws.max_row
            ws.cell(row=max_row + 1, column=4, value="ΓΕΝΙΚΟ ΣΥΝΟΛΟ:").font = Font(bold=True)

            sum_cell = ws.cell(row=max_row + 1, column=5)
            sum_cell.value = f"=SUM(E2:E{max_row})"
            sum_cell.font = Font(bold=True, size=12, color="D32F2F")
            sum_cell.number_format = '#,##0.00 €'
            sum_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            sum_cell.border = Border(top=Side(style='thick'), bottom=Side(style='thick'))

            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[column].width = max_length + 5

        show_custom_alert(self, "Ολοκληρώθηκε", f"Το Excel δημιουργήθηκε επιτυχώς!")

        filepath = os.path.abspath(filename)
        try:
            if platform.system() == 'Windows':
                os.startfile(filepath)
            elif platform.system() == 'Darwin':
                os.system(f"open '{filepath}'")
            else:
                os.system(f"xdg-open '{filepath}'")
        except Exception:
            pass


# --- ΕΚΚΙΝΗΣΗ ---
if __name__ == "__main__":
    ensure_logo_exists()
    init_db()

    app = TutorApp()
    app.mainloop()