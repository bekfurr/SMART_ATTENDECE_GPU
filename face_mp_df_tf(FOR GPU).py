import cv2
import mediapipe as mp
from deepface import DeepFace
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import pandas as pd
from datetime import datetime, time, timedelta
import json
import schedule
import time as tm
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import tensorflow as tf
from PIL import Image, ImageTk

# GPU sozlamalarini tekshirish
physical_devices = tf.config.list_physical_devices('GPU')
if physical_devices:
    tf.config.experimental.set_memory_growth(physical_devices[0], True)

class AttendanceApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1000x700")
        self.root.minsize(1000, 700)
        self.root.title("Davomat Tizimi")
        self.root.configure(bg="#f0f2f5")
        
        self.current_frame = None
        self.database_path = None
        self.running = False
        self.status_text = None
        self.cap = None
        self.deadline = None
        self.late_deadline = None
        self.mode_choice = None
        self.schedule_file = None
        self.schedule_data = None
        self.attendance_data = None
        self.contacts_file = "contacts.json"
        self.smtp_settings_file = "smtp_settings.json"
        self.contacts = self.load_contacts()
        self.smtp_settings = self.load_smtp_settings()
        
        # Mediapipe sozlamalari
        self.mp_face_detection = mp.solutions.face_detection
        self.face_detection = self.mp_face_detection.FaceDetection(model_selection=1, min_detection_confidence=0.5)
        
        # Pre-load ArcFace model
        try:
            self.arcface_model = DeepFace.build_model("ArcFace")
        except Exception as e:
            messagebox.showerror("Xato", f"ArcFace modelini yuklashda xato: {e}")
            self.arcface_model = None
        
        self.create_main_interface()
        
    def load_contacts(self):
        """Kontaktlar faylini yuklash"""
        if os.path.exists(self.contacts_file):
            try:
                with open(self.contacts_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Kontakts faylini o'qishda xato: {e}")
                return {}
        return {}

    def save_contacts(self):
        """Kontaktlarni faylga saqlash"""
        try:
            with open(self.contacts_file, 'w') as f:
                json.dump(self.contacts, f, indent=4)
        except Exception as e:
            print(f"Kontakts faylini saqlashda xato: {e}")

    def load_smtp_settings(self):
        """SMTP sozlamalarini yuklash"""
        if os.path.exists(self.smtp_settings_file):
            try:
                with open(self.smtp_settings_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"SMTP sozlamalarini o'qishda xato: {e}")
                return {}
        return {}

    def save_smtp_settings(self):
        """SMTP sozlamalarini faylga saqlash"""
        try:
            with open(self.smtp_settings_file, 'w') as f:
                json.dump(self.smtp_settings, f, indent=4)
        except Exception as e:
            print(f"SMTP sozlamalarini saqlashda xato: {e}")

    def create_main_interface(self):
        """Asosiy interfeysni yaratish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Davomat Tizimi", 
                 font=("Helvetica", 24, "bold")).pack(pady=20)
        
        btn_frame = ttk.Frame(self.current_frame)
        btn_frame.pack(fill="x", pady=20)
        
        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 12), padding=10)
        
        ttk.Button(btn_frame, text="Davomat",
                  command=self.show_attendance_section).pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Email Jo'natish",
                  command=self.show_email_section).pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Kontaktlar",
                  command=self.show_contacts_section).pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="SMTP Sozlamalari",
                  command=self.show_smtp_settings_section).pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Chiqish",
                  command=self.quit_application).pack(fill="x", pady=5)
        
        ttk.Label(self.current_frame, 
                 text="BEKFURR INC 2025",
                 font=("Arial", 8)).place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def quit_application(self):
        """Ilovadan chiqish"""
        self.running = False
        if self.cap:
            self.cap.release()
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass

    def show_attendance_section(self):
        """Davomat bo'limini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Davomat Bo'limi",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        ttk.Button(self.current_frame, text="Baza Yaratish",
                  command=self.show_create_database).pack(fill="x", pady=5)
        ttk.Button(self.current_frame, text="Davomat Boshlash",
                  command=self.show_attendance_setup).pack(fill="x", pady=5)
        ttk.Button(self.current_frame, text="Jadval Yaratish",
                  command=self.show_create_schedule).pack(fill="x", pady=5)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.create_main_interface).pack(fill="x", pady=5)

    def show_create_database(self):
        """Yangi o'quvchi qo'shish interfeysini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Yangi O'quvchi Qo'shish",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        fields = ["Ism", "Familiya", "Otasining ismi", "Fakultet", "Yo'nalish", "Guruh"]
        self.entries = {}
        
        for field in fields:
            frame = ttk.Frame(self.current_frame)
            frame.pack(fill="x", pady=5)
            ttk.Label(frame, text=f"{field}:", width=15).pack(side="left")
            entry = ttk.Entry(frame)
            entry.pack(fill="x", expand=True, padx=5)
            self.entries[field.lower().replace("otasining ismi", "father_name")] = entry
        
        db_frame = ttk.Frame(self.current_frame)
        db_frame.pack(fill="x", pady=5)
        ttk.Label(db_frame, text="Baza papkasi:", width=15).pack(side="left")
        self.db_name_entry = ttk.Entry(db_frame)
        self.db_name_entry.pack(fill="x", expand=True, padx=5)
        self.db_name_entry.insert(0, "face_database")
        
        file_frame = ttk.Frame(self.current_frame)
        file_frame.pack(fill="x", pady=5)
        ttk.Label(file_frame, text="Suratlar (kamida 4 ta):", width=15).pack(side="left")
        self.file_paths_var = tk.StringVar()
        ttk.Button(file_frame, text="Tanlash",
                  command=self.select_files).pack(side="left", padx=5)
        ttk.Entry(file_frame, textvariable=self.file_paths_var,
                 state="readonly").pack(fill="x", expand=True, padx=5)
        
        ttk.Button(self.current_frame, text="Saqlash",
                  command=self.save_to_database).pack(pady=10)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.show_attendance_section).pack(pady=5)

    def select_files(self):
        """Suratlarni tanlash"""
        file_paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg;*.jpeg;*.png"), ("Barcha fayllar", "*.*")])
        if len(file_paths) < 4:
            messagebox.showwarning("Ogohlantirish", "Kamida 4 ta surat tanlashingiz kerak!")
            return
        if file_paths:
            self.file_paths_var.set(";".join(file_paths))

    def save_to_database(self):
        """Ma'lumotlar bazasiga o'quvchini saqlash"""
        name = self.entries["ism"].get().strip()
        surname = self.entries["familiya"].get().strip()
        father_name = self.entries["father_name"].get().strip()
        faculty = self.entries["fakultet"].get().strip()
        direction = self.entries["yo'nalish"].get().strip()
        group = self.entries["guruh"].get().strip()
        file_paths = self.file_paths_var.get().split(";")
        db_name = self.db_name_entry.get().strip() or "face_database"
        
        if not all([name, surname, father_name, faculty, direction, group, len(file_paths) >= 4]):
            messagebox.showwarning("Ogohlantirish", "Barcha maydonlarni to'ldiring va kamida 4 ta surat tanlang!")
            return
            
        valid_images = 0
        db_path = os.path.join(os.getcwd(), db_name)
        person_path = os.path.join(db_path, f"{name}_{surname}")
        os.makedirs(person_path, exist_ok=True)
        
        metadata = {}
        metadata_path = os.path.join(db_path, "metadata.json")
        if os.path.exists(metadata_path):
            try:
                with open(metadata_path, 'r') as f:
                    metadata = json.load(f)
            except:
                pass
        
        for i, file_path in enumerate(file_paths):
            try:
                img = cv2.imread(file_path)
                if img is None:
                    continue
                rgb_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                results = self.face_detection.process(rgb_img)
                
                if results.detections:
                    valid_images += 1
                    new_path = os.path.join(person_path, f"{name}_{i}.jpg")
                    cv2.imwrite(new_path, img)
            except Exception as e:
                print(f"Suratni qayta ishlashda xato: {e}")
                continue
        
        if valid_images < 4:
            messagebox.showerror("Xato", "Kamida 4 ta suratda yuz aniqlanishi kerak!")
            return
        
        metadata[name] = {
            "surname": surname,
            "father_name": father_name,
            "faculty": faculty,
            "direction": direction,
            "group": group,
            "image_folder": person_path
        }
        
        try:
            with open(metadata_path, 'w') as f:
                json.dump(metadata, f, indent=4)
        except Exception as e:
            messagebox.showerror("Xato", f"Metadata faylini saqlashda xato: {e}")
            return
        
        messagebox.showinfo("Muvaffaqiyat", f"{name} {db_name} bazasiga qo'shildi!")
        self.show_attendance_section()

    def show_create_schedule(self):
        """Jadval yaratish interfeysini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Jadval Yaratish",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        self.schedule_entries = {}
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        for day in days:
            ttk.Label(self.current_frame, text=day, font=("Helvetica", 12, "bold")).pack(pady=5)
            
            frame = ttk.Frame(self.current_frame)
            frame.pack(fill="x", pady=2)
            
            ttk.Label(frame, text="Boshlanish vaqti (HH:MM):", width=20).pack(side="left")
            start_entry = ttk.Entry(frame, width=10)
            start_entry.pack(side="left", padx=5)
            
            ttk.Label(frame, text="Kech qolish (HH:MM):", width=20).pack(side="left")
            late_entry = ttk.Entry(frame, width=10)
            late_entry.pack(side="left", padx=5)
            
            ttk.Label(frame, text="Tugash vaqti (HH:MM):", width=20).pack(side="left")
            end_entry = ttk.Entry(frame, width=10)
            end_entry.pack(side="left", padx=5)
            
            self.schedule_entries[day] = {
                "start": start_entry,
                "late": late_entry,
                "end": end_entry
            }
        
        ttk.Button(self.current_frame, text="Jadvalni Saqlash",
                  command=self.save_schedule).pack(pady=10)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.show_attendance_section).pack(pady=5)

    def save_schedule(self):
        """Jadvalni saqlash"""
        schedule_data = {}
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        for day in days:
            start_time = self.schedule_entries[day]["start"].get().strip()
            late_time = self.schedule_entries[day]["late"].get().strip()
            end_time = self.schedule_entries[day]["end"].get().strip()
            
            try:
                if start_time:
                    datetime.strptime(start_time, "%H:%M")
                if late_time:
                    datetime.strptime(late_time, "%H:%M")
                if end_time:
                    datetime.strptime(end_time, "%H:%M")
            except ValueError:
                messagebox.showwarning("Ogohlantirish", f"{day} kuni uchun vaqt formati noto'g'ri! (HH:MM, masalan, 09:00)")
                return
            
            if start_time and late_time and end_time:
                schedule_data[day] = {
                    "start": start_time,
                    "late": late_time,
                    "end": end_time
                }
        
        if not schedule_data:
            messagebox.showwarning("Ogohlantirish", "Kamida bir kun uchun jadval kiritilishi kerak!")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(schedule_data, f, indent=4)
                messagebox.showinfo("Muvaffaqiyat", "Jadval saqlandi!")
                self.show_attendance_section()
            except Exception as e:
                messagebox.showerror("Xato", f"Jadval faylini saqlashda xato: {e}")

    def show_attendance_setup(self):
        """Davomat sozlamalari interfeysini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        canvas = tk.Canvas(self.current_frame)
        scrollbar = ttk.Scrollbar(self.current_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        ttk.Label(scrollable_frame, text="Davomat Sozlamalari",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        mode_frame = ttk.Frame(scrollable_frame)
        mode_frame.pack(fill="x", pady=5)
        self.mode_choice = tk.StringVar(value="manual")
        ttk.Radiobutton(mode_frame, text="Qo'lda (Manual)", variable=self.mode_choice, value="manual").pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="Jadval bo'yicha", variable=self.mode_choice, value="schedule").pack(side="left", padx=5)
        
        db_frame = ttk.Frame(scrollable_frame)
        db_frame.pack(fill="x", pady=5)
        ttk.Label(db_frame, text="Baza papkasi:", width=15).pack(side="left")
        self.db_select_var = tk.StringVar()
        ttk.Button(db_frame, text="Papka tanlash",
                  command=self.select_database).pack(side="left", padx=5)
        ttk.Entry(db_frame, textvariable=self.db_select_var,
                 state="readonly").pack(fill="x", expand=True, padx=5)
        
        self.schedule_frame = ttk.Frame(scrollable_frame)
        self.schedule_frame.pack(fill="x", pady=5)
        ttk.Label(self.schedule_frame, text="Jadval fayli:", width=15).pack(side="left")
        self.schedule_var = tk.StringVar()
        ttk.Button(self.schedule_frame, text="Fayl tanlash",
                  command=self.select_schedule_file).pack(side="left", padx=5)
        ttk.Entry(self.schedule_frame, textvariable=self.schedule_var,
                 state="readonly").pack(fill="x", expand=True, padx=5)
        
        cam_frame = ttk.Frame(scrollable_frame)
        cam_frame.pack(fill="x", pady=5)
        ttk.Label(cam_frame, text="Kamera:", width=15).pack(side="left")
        self.camera_choice = tk.StringVar()
        cameras = self.detect_available_cameras()
        camera_options = [cam[1] for cam in cameras] + ["IP Camera"]
        self.camera_menu = ttk.OptionMenu(cam_frame, self.camera_choice,
                                         cameras[0][1] if cameras else "Kamera topilmadi",
                                         *camera_options)
        self.camera_menu.pack(fill="x", expand=True, padx=5)
        
        ip_frame = ttk.Frame(scrollable_frame)
        ip_frame.pack(fill="x", pady=5)
        ttk.Label(ip_frame, text="IP URL:", width=15).pack(side="left")
        self.ip_entry = ttk.Entry(ip_frame)
        self.ip_entry.pack(fill="x", expand=True, padx=5)
        self.ip_entry.config(state="disabled")
        
        ttk.Label(scrollable_frame, text="Kech qolish chegarasi:", font=("Helvetica", 12, "bold")).pack(pady=10)
        late_deadline_frame = ttk.Frame(scrollable_frame)
        late_deadline_frame.pack(fill="x", pady=5)
        
        self.late_deadline_choice = tk.StringVar(value="time")
        ttk.Radiobutton(late_deadline_frame, text="Soat bo'yicha", variable=self.late_deadline_choice, value="time").pack(side="left", padx=5)
        ttk.Radiobutton(late_deadline_frame, text="Taymer bo'yicha", variable=self.late_deadline_choice, value="timer").pack(side="left", padx=5)
        
        self.late_time_frame = ttk.Frame(scrollable_frame)
        self.late_time_frame.pack(fill="x", pady=5)
        ttk.Label(self.late_time_frame, text="Soat (0-23):", width=15).pack(side="left")
        self.late_hour_entry = ttk.Entry(self.late_time_frame, width=5)
        self.late_hour_entry.pack(side="left", padx=5)
        ttk.Label(self.late_time_frame, text="Daqiqa (0-59):", width=15).pack(side="left")
        self.late_minute_entry = ttk.Entry(self.late_time_frame, width=5)
        self.late_minute_entry.pack(side="left", padx=5)
        
        self.late_timer_frame = ttk.Frame(scrollable_frame)
        self.late_timer_frame.pack(fill="x", pady=5)
        ttk.Label(self.late_timer_frame, text="Daqiqa:", width=15).pack(side="left")
        self.late_deadline_var = tk.StringVar(value="5")
        ttk.OptionMenu(self.late_timer_frame, self.late_deadline_var, "5", "10", "15", "30").pack(side="left", padx=5)
        
        ttk.Label(scrollable_frame, text="Davomat tugash vaqti:", font=("Helvetica", 12, "bold")).pack(pady=10)
        deadline_frame = ttk.Frame(scrollable_frame)
        deadline_frame.pack(fill="x", pady=5)
        
        self.deadline_choice = tk.StringVar(value="time")
        ttk.Radiobutton(deadline_frame, text="Soat bo'yicha", variable=self.deadline_choice, value="time").pack(side="left", padx=5)
        ttk.Radiobutton(deadline_frame, text="Taymer bo'yicha", variable=self.deadline_choice, value="timer").pack(side="left", padx=5)
        
        self.time_frame = ttk.Frame(scrollable_frame)
        self.time_frame.pack(fill="x", pady=5)
        ttk.Label(self.time_frame, text="Soat (0-23):", width=15).pack(side="left")
        self.hour_entry = ttk.Entry(self.time_frame, width=5)
        self.hour_entry.pack(side="left", padx=5)
        ttk.Label(self.time_frame, text="Daqiqa (0-59):", width=15).pack(side="left")
        self.minute_entry = ttk.Entry(self.time_frame, width=5)
        self.minute_entry.pack(side="left", padx=5)
        
        self.timer_frame = ttk.Frame(scrollable_frame)
        self.timer_frame.pack(fill="x", pady=5)
        ttk.Label(self.timer_frame, text="Daqiqa:", width=15).pack(side="left")
        self.deadline_var = tk.StringVar(value="5")
        ttk.OptionMenu(self.timer_frame, self.deadline_var, "5", "10", "15", "30").pack(side="left", padx=5)
        
        ttk.Button(scrollable_frame, text="Davomatni Boshlash",
                  command=self.start_attendance).pack(pady=10)
        ttk.Button(scrollable_frame, text="Orqaga",
                  command=self.show_attendance_section).pack(pady=5)
        
        self.camera_choice.trace("w", self.toggle_ip_entry)
        self.late_deadline_choice.trace("w", self.toggle_late_deadline_input)
        self.deadline_choice.trace("w", self.toggle_deadline_input)
        self.mode_choice.trace("w", self.toggle_mode_input)
        
        self.toggle_late_deadline_input()
        self.toggle_deadline_input()
        self.toggle_mode_input()

    def toggle_ip_entry(self, *args):
        """IP kamera kiritish maydonini yoqish/o'chirish"""
        self.ip_entry.config(state="normal" if self.camera_choice.get() == "IP Camera" else "disabled")

    def toggle_late_deadline_input(self, *args):
        """Kech qolish chegarasi kiritish maydonini boshqarish"""
        if self.late_deadline_choice.get() == "time":
            self.late_time_frame.pack(fill="x", pady=5)
            self.late_timer_frame.pack_forget()
        else:
            self.late_time_frame.pack_forget()
            self.late_timer_frame.pack(fill="x", pady=5)

    def toggle_deadline_input(self, *args):
        """Davomat tugash vaqti kiritish maydonini boshqarish"""
        if self.deadline_choice.get() == "time":
            self.time_frame.pack(fill="x", pady=5)
            self.timer_frame.pack_forget()
        else:
            self.time_frame.pack_forget()
            self.timer_frame.pack(fill="x", pady=5)

    def toggle_mode_input(self, *args):
        """Rejim tanlashga qarab kiritish maydonlarini boshqarish"""
        if self.mode_choice.get() == "manual":
            self.schedule_frame.pack_forget()
            self.late_time_frame.pack(fill="x", pady=5)
            self.late_timer_frame.pack(fill="x", pady=5)
            self.time_frame.pack(fill="x", pady=5)
            self.timer_frame.pack(fill="x", pady=5)
        else:
            self.late_time_frame.pack_forget()
            self.late_timer_frame.pack_forget()
            self.time_frame.pack_forget()
            self.timer_frame.pack_forget()
            self.schedule_frame.pack(fill="x", pady=5)

    def select_database(self):
        """Ma'lumotlar bazasi papkasini tanlash"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.db_select_var.set(folder_path)
            self.database_path = folder_path

    def select_schedule_file(self):
        """Jadval faylini tanlash"""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("Barcha fayllar", "*.*")])
        if file_path:
            self.schedule_var.set(file_path)
            self.schedule_file = file_path
            self.load_schedule()

    def load_schedule(self):
        """Jadval faylini yuklash"""
        try:
            with open(self.schedule_file, 'r') as f:
                self.schedule_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Xato", f"Jadval faylini o'qishda xato: {e}")
            self.schedule_data = None

    def detect_available_cameras(self, max_index=3):
        """Mavjud kameralarni aniqlash"""
        available_cameras = []
        for i in range(max_index):
            try:
                cap = cv2.VideoCapture(i)
                if cap.isOpened():
                    available_cameras.append((i, f"Kamera {i} (Indeks: {i})"))
                    cap.release()
            except:
                pass
        return available_cameras

    def calculate_statistics(self, distances):
        """Masofalar statistikasini hisoblash"""
        if not distances:
            return 0.0, 0.0, 0.0
        
        mean = np.mean(distances)
        variance = np.var(distances)
        std_dev = np.sqrt(variance)
        
        return mean, variance, std_dev

    def start_attendance(self):
        """Davomatni boshlash"""
        if not self.db_select_var.get():
            messagebox.showwarning("Xato", "Baza tanlanmadi!")
            return
            
        try:
            with open(os.path.join(self.db_select_var.get(), "metadata.json"), 'r') as f:
                self.database = json.load(f)
        except Exception as e:
            messagebox.showerror("Xato", f"Baza faylini yuklashda xato: {e}")
            return
            
        camera_source = self.camera_choice.get()
        if camera_source == "IP Camera":
            camera_source = self.ip_entry.get().strip()
            if not camera_source:
                messagebox.showwarning("Xato", "IP kamera URL kiritilmadi!")
                return
        else:
            try:
                camera_source = int(camera_source.split(" (Indeks: ")[1].rstrip(")"))
            except:
                messagebox.showerror("Xato", "Kamera tanlashda xato!")
                return

        if self.mode_choice.get() == "manual":
            if self.late_deadline_choice.get() == "time":
                try:
                    hour = int(self.late_hour_entry.get())
                    minute = int(self.late_minute_entry.get())
                    if 0 <= hour <= 23 and 0 <= minute <= 59:
                        self.late_deadline = datetime.combine(datetime.now().date(), time(hour, minute))
                        if self.late_deadline < datetime.now():
                            self.late_deadline += timedelta(days=1)
                    else:
                        messagebox.showwarning("Xato", "Soat 0-23, daqiqa 0-59 oralig'ida bo'lishi kerak!")
                        return
                except ValueError:
                    messagebox.showwarning("Xato", "Soat va daqiqa raqam bo'lishi kerak!")
                    return
            else:
                minutes = int(self.late_deadline_var.get())
                self.late_deadline = datetime.now() + timedelta(minutes=minutes)

            if self.deadline_choice.get() == "time":
                try:
                    hour = int(self.hour_entry.get())
                    minute = int(self.minute_entry.get())
                    if 0 <= hour <= 23 and 0 <= minute <= 59:
                        self.deadline = datetime.combine(datetime.now().date(), time(hour, minute))
                        if self.deadline < datetime.now():
                            self.deadline += timedelta(days=1)
                    else:
                        messagebox.showwarning("Xato", "Soat 0-23, daqiqa 0-59 oralig'ida bo'lishi kerak!")
                        return
                except ValueError:
                    messagebox.showwarning("Xato", "Soat va daqiqa raqam bo'lishi kerak!")
                    return
            else:
                minutes = int(self.deadline_var.get())
                self.deadline = datetime.now() + timedelta(minutes=minutes)

            if self.late_deadline >= self.deadline:
                messagebox.showwarning("Xato", "Kech qolish chegarasi umumiy tugash vaqtidan oldin bo'lishi kerak!")
                return

            self.attendance_system(camera_source)
        else:
            if not self.schedule_file or not self.schedule_data:
                messagebox.showwarning("Xato", "Jadval fayli tanlanmadi yoki noto'g'ri!")
                return
            self.running = True
            threading.Thread(target=self.run_scheduled_attendance, args=(camera_source,), daemon=True).start()

    def run_scheduled_attendance(self, camera_source):
        """Jadval bo'yicha davomatni boshqarish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Jadval bo'yicha kutish rejimi",
                 font=("Helvetica", 16, "bold")).pack(pady=10)
        
        ttk.Label(self.current_frame, text="Joriy kun:", font=("Helvetica", 12)).pack(pady=2)
        self.day_label = ttk.Label(self.current_frame, text="", font=("Helvetica", 12))
        self.day_label.pack(pady=2)
        
        ttk.Label(self.current_frame, text="Jadval bo'yicha boshlanish vaqti:", font=("Helvetica", 12)).pack(pady=2)
        self.start_time_label = ttk.Label(self.current_frame, text="", font=("Helvetica", 12))
        self.start_time_label.pack(pady=2)
        
        ttk.Label(self.current_frame, text="Joriy vaqt:", font=("Helvetica", 12)).pack(pady=2)
        self.current_time_label = ttk.Label(self.current_frame, text="", font=("Helvetica", 12))
        self.current_time_label.pack(pady=2)
        
        ttk.Label(self.current_frame, text="Qolgan vaqt:", font=("Helvetica", 12)).pack(pady=2)
        self.remaining_time_label = ttk.Label(self.current_frame, text="", font=("Helvetica", 12))
        self.remaining_time_label.pack(pady=2)
        
        ttk.Button(self.current_frame, text="To'xtatish",
                  command=self.stop_attendance).pack(pady=10)
        
        def update_timer():
            while self.running:
                current_time = datetime.now()
                current_day = current_time.strftime("%A")
                try:
                    self.day_label.config(text=f"{current_day}")
                except tk.TclError:
                    self.running = False
                    break
                
                if current_day in self.schedule_data:
                    schedule_entry = self.schedule_data[current_day]
                    start_time = datetime.strptime(schedule_entry["start"], "%H:%M").time()
                    start_datetime = datetime.combine(current_time.date(), start_time)
                    if current_time > start_datetime:
                        start_datetime += timedelta(days=1)
                    try:
                        self.start_time_label.config(text=f"{start_datetime.strftime('%H:%M')}")
                    except tk.TclError:
                        self.running = False
                        break
                
                try:
                    self.current_time_label.config(text=f"{current_time.strftime('%H:%M:%S')}")
                except tk.TclError:
                    self.running = False
                    break
                
                if current_day in self.schedule_data:
                    schedule_entry = self.schedule_data[current_day]
                    start_time = datetime.strptime(schedule_entry["start"], "%H:%M").time()
                    start_datetime = datetime.combine(current_time.date(), start_time)
                    if current_time > start_datetime:
                        start_datetime += timedelta(days=1)
                    seconds_to_wait = max(0, (start_datetime - current_time).total_seconds())
                    minutes, seconds = divmod(int(seconds_to_wait), 60)
                    hours, minutes = divmod(minutes, 60)
                    try:
                        self.remaining_time_label.config(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
                    except tk.TclError:
                        self.running = False
                        break
                
                self.root.after(1000, update_timer)
                return

        self.root.after(0, update_timer)

        while self.running:
            current_day = datetime.now().strftime("%A")
            current_time = datetime.now()
            
            if current_day in self.schedule_data:
                schedule_entry = self.schedule_data[current_day]
                start_time = datetime.strptime(schedule_entry["start"], "%H:%M").time()
                late_time = datetime.strptime(schedule_entry["late"], "%H:%M").time()
                end_time = datetime.strptime(schedule_entry["end"], "%H:%M").time()
                
                start_datetime = datetime.combine(current_time.date(), start_time)
                late_datetime = datetime.combine(current_time.date(), late_time)
                end_datetime = datetime.combine(current_time.date(), end_time)
                
                if current_time > end_datetime:
                    start_datetime += timedelta(days=1)
                    late_datetime += timedelta(days=1)
                    end_datetime += timedelta(days=1)
                
                if current_time < start_datetime:
                    seconds_to_wait = (start_datetime - current_time).total_seconds()
                    tm.sleep(seconds_to_wait)
                
                if start_datetime <= datetime.now() <= end_datetime:
                    self.late_deadline = late_datetime
                    self.deadline = end_datetime
                    self.attendance_system(camera_source)
                elif datetime.now() > end_datetime:
                    self.start_surveillance(camera_source)
            
            tm.sleep(60)

    def attendance_system(self, camera_source):
        """Davomat jarayonini boshqarish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.running = True
        self.cap = cv2.VideoCapture(camera_source)
        
        if not self.cap.isOpened():
            messagebox.showerror("Xato", "Kamera ochilmadi!")
            self.show_attendance_section()
            return
            
        self.attendance = {
            name: {
                "surname": data["surname"],
                "father_name": data["father_name"],
                "faculty": data["faculty"],
                "direction": data["direction"],
                "group": data["group"],
                "status": "Kelmagan",
                "arrival_time": None,
                "late_time": None,
                "recorded": False,
                "distances": [],
                "probability": 0.0,
                "mean_distance": 0.0,
                "variance": 0.0,
                "std_dev": 0.0
            } 
            for name, data in self.database.items()
        }
        
        ttk.Label(self.current_frame, text="Davomat Davom Etmoqda...",
                 font=("Helvetica", 16, "bold")).pack(pady=10)
        
        self.time_label = ttk.Label(self.current_frame, text="", font=("Helvetica", 12))
        self.time_label.pack(pady=5)
        
        # Video display in Tkinter
        self.video_label = tk.Label(self.current_frame)
        self.video_label.pack(pady=10)
        
        self.status_text = tk.Text(self.current_frame, height=10, width=80)
        self.status_text.pack(pady=10)
        
        ttk.Button(self.current_frame, text="Yakunlash",
                  command=self.stop_attendance).pack(pady=10)
        
        self.timer_event = threading.Event()
        self.video_event = threading.Event()

        def update_timer():
            while self.running and not self.timer_event.is_set():
                current_time = datetime.now()
                if current_time >= self.deadline:
                    self.stop_attendance()
                    break
                remaining_time = self.deadline - current_time
                minutes, seconds = divmod(remaining_time.seconds, 60)
                try:
                    self.time_label.config(text=f"Davomat tugashiga qolgan vaqt: {minutes:02d}:{seconds:02d}")
                except tk.TclError:
                    self.running = False
                    break
                self.timer_event.wait(1)

        def video_loop():
            while self.running and not self.video_event.is_set():
                if not self.cap or not self.cap.isOpened():
                    self.root.after(0, lambda: messagebox.showerror("Xato", "Kamera uzildi! Davomat to'xtatildi."))
                    self.stop_attendance()
                    break
                    
                ret, frame = self.cap.read()
                if not ret:
                    self.root.after(0, lambda: messagebox.showerror("Xato", "Kamera o'qishda xato! Davomat to'xtatildi."))
                    self.stop_attendance()
                    break
                    
                rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                results = self.face_detection.process(rgb_frame)
                
                current_time = datetime.now()
                if results.detections:
                    for detection in results.detections:
                        bbox = detection.location_data.relative_bounding_box
                        h, w, _ = frame.shape
                        x, y = int(bbox.xmin * w), int(bbox.ymin * h)
                        width, height = int(bbox.width * w), int(bbox.height * h)
                        
                        face_img = frame[y:y+height, x:x+width]
                        if face_img.size == 0:
                            continue
                            
                        name = "Nomalum"
                        color = (0, 0, 255)
                        status = ""
                        max_probability = 0.0
                        best_match_name = None
                        best_distance = float('inf')
                        
                        if self.arcface_model:
                            for person_name, person_data in self.database.items():
                                image_folder = person_data["image_folder"]
                                for img_file in os.listdir(image_folder):
                                    ref_img_path = os.path.join(image_folder, img_file)
                                    try:
                                        result = DeepFace.verify(face_img, ref_img_path, 
                                                                model=self.arcface_model, 
                                                                enforce_detection=False)
                                        distance = result["distance"]
                                        probability = 1 - distance
                                        if probability > max_probability and probability > 0.5:
                                            max_probability = probability
                                            best_match_name = person_name
                                            best_distance = distance
                                            self.attendance[person_name]["distances"].append(distance)
                                            mean, variance, std_dev = self.calculate_statistics(
                                                self.attendance[person_name]["distances"])
                                            self.attendance[person_name]["probability"] = probability
                                            self.attendance[person_name]["mean_distance"] = mean
                                            self.attendance[person_name]["variance"] = variance
                                            self.attendance[person_name]["std_dev"] = std_dev
                                    except Exception as e:
                                        print(f"Yuz taqqoslashda xato: {e}")
                                        continue
                        
                        if best_match_name:
                            name = best_match_name
                            if not self.attendance[name]["recorded"]:
                                if current_time <= self.late_deadline:
                                    self.attendance[name]["status"] = "Kelgan"
                                    self.attendance[name]["arrival_time"] = current_time.strftime("%H:%M:%S")
                                else:
                                    self.attendance[name]["status"] = "Kech qolgan"
                                    self.attendance[name]["late_time"] = current_time.strftime("%H:%M:%S")
                                self.attendance[name]["recorded"] = True
                                self.update_status_text(
                                    name, 
                                    self.attendance[name]["status"],
                                    self.attendance[name]["probability"],
                                    self.attendance[name]["mean_distance"],
                                    self.attendance[name]["variance"],
                                    self.attendance[name]["std_dev"]
                                )
                            status = self.attendance[name]["status"]
                            probability = self.attendance[name]["probability"]
                            if status == "Kelgan":
                                color = (0, 255, 0)
                            elif status == "Kech qolgan":
                                color = (255, 165, 0)
                            
                            cv2.rectangle(frame, (x, y), (x + width, y + height), color, 2)
                            cv2.putText(frame, f"{name} ({status}, {probability:.2%})", 
                                       (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
                
                # Convert frame to Tkinter-compatible format
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                img = img.resize((640, 480), Image.LANCZOS)  # Resize for display
                imgtk = ImageTk.PhotoImage(image=img)
                
                # Update video label in main thread
                self.root.after(0, lambda: self.video_label.config(image=imgtk))
                self.video_label.imgtk = imgtk  # Keep reference to avoid garbage collection
                
                try:
                    self.root.update()
                except tk.TclError:
                    self.running = False
                    break
                
                # Small delay to avoid overloading
                self.video_event.wait(0.03)  # ~30 FPS
        
            if self.cap:
                self.cap.release()
            self.cap = None
        
        threading.Thread(target=video_loop, daemon=True).start()
        threading.Thread(target=update_timer, daemon=True).start()

    def start_surveillance(self, camera_source):
        """Video kuzatuv rejimini boshqarish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Video Kuzatuv Rejimi",
                 font=("Helvetica", 16, "bold")).pack(pady=10)
        
        self.cap = cv2.VideoCapture(camera_source)
        if not self.cap.isOpened():
            messagebox.showerror("Xato", "Kamera ochilmadi!")
            self.show_attendance_section()
            return
        
        self.video_label = tk.Label(self.current_frame)
        self.video_label.pack(pady=10)
        
        ttk.Button(self.current_frame, text="To'xtatish",
                  command=self.stop_attendance).pack(pady=10)
        
        self.surveillance_event = threading.Event()

        def surveillance_loop():
            while self.running and not self.surveillance_event.is_set():
                if not self.cap or not self.cap.isOpened():
                    self.root.after(0, lambda: messagebox.showerror("Xato", "Kamera uzildi! Kuzatuv to'xtatildi."))
                    self.stop_attendance()
                    break
                    
                ret, frame = self.cap.read()
                if not ret:
                    self.root.after(0, lambda: messagebox.showerror("Xato", "Kamera o'qishda xato! Kuzatuv to'xtatildi."))
                    self.stop_attendance()
                    break
                    
                # Convert frame to Tkinter-compatible format
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                img = img.resize((640, 480), Image.LANCZOS)
                imgtk = ImageTk.PhotoImage(image=img)
                
                # Update video label in main thread
                self.root.after(0, lambda: self.video_label.config(image=imgtk))
                self.video_label.imgtk = imgtk
                
                try:
                    self.root.update()
                except tk.TclError:
                    self.running = False
                    break
                
                self.surveillance_event.wait(0.03)  # ~30 FPS
            
            if self.cap:
                self.cap.release()
            self.cap = None
        
        threading.Thread(target=surveillance_loop, daemon=True).start()

    def update_status_text(self, name, status, probability, mean_distance, variance, std_dev):
        """Holatni matn maydonida yangilash"""
        try:
            if self.status_text and self.running:
                self.status_text.config(state="normal")
                self.status_text.insert(tk.END, f"{'='*50}\n")
                self.status_text.insert(tk.END, 
                    f"{name}: {status} ({datetime.now().strftime('%H:%M:%S')})\n"
                    f"  Ehtimollik: {probability:.2%}\n"
                    f"  O'rtacha masofa: {mean_distance:.4f}\n"
                    f"  Dispersiya: {variance:.4f}\n"
                    f"  Kvadrat chetlanish: {std_dev:.4f}\n\n"
                    f"{'-'*30}\n"
                )
                self.status_text.config(state="disabled")
                self.status_text.see(tk.END)
        except tk.TclError:
            pass

    def stop_attendance(self):
        """Davomatni to'xtatish"""
        self.running = False
        if hasattr(self, 'timer_event'):
            self.timer_event.set()
        if hasattr(self, 'video_event'):
            self.video_event.set()
        if hasattr(self, 'surveillance_event'):
            self.surveillance_event.set()
        if self.cap:
            self.cap.release()
            self.cap = None
        if self.current_frame:
            try:
                self.current_frame.destroy()
            except tk.TclError:
                pass
        if hasattr(self, 'attendance'):
            self.save_attendance()
            self.show_email_contact_selection()
        else:
            self.create_main_interface()

    def save_attendance(self):
        """Davomat ma'lumotlarini saqlash"""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.filename = f"attendance_{timestamp}.xlsx"
        df = pd.DataFrame([
            {
                "Ism": name,
                "Familiya": data["surname"],
                "Otasining ismi": data["father_name"],
                "Fakultet": data["faculty"],
                "Yo'nalish": data["direction"],
                "Guruh": data["group"],
                "Holati": data["status"],
                "Kelgan vaqti": data["arrival_time"],
                "Kech qolgan vaqti": data["late_time"],
                "Ehtimollik": f"{data['probability']:.2%}" if data["probability"] > 0 else None,
                "O'rtacha masofa": f"{data['mean_distance']:.4f}" if data["mean_distance"] > 0 else None,
                "Dispersiya": f"{data['variance']:.4f}" if data["variance"] > 0 else None,
                "Kvadrat chetlanish": f"{data['std_dev']:.4f}" if data["std_dev"] > 0 else None
            }
            for name, data in self.attendance.items()
        ])
        try:
            df.to_excel(self.filename, index=False)
            self.attendance_data = self.attendance
        except Exception as e:
            messagebox.showerror("Xato", f"Davomat faylini saqlashda xato: {e}")

    def show_email_contact_selection(self):
        """Email yuborish uchun kontakt tanlash interfeysini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Hisobotni Email orqali Yuborish",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        if not self.contacts:
            ttk.Label(self.current_frame, text="Kontaktlar mavjud emas! Iltimos, Kontaktlar bo'limida kontakt qo'shing.",
                     font=("Helvetica", 12, "italic"), foreground="red").pack(pady=10)
            ttk.Button(self.current_frame, text="Bosh menyuga",
                      command=self.create_main_interface).pack(pady=10)
            return
        
        if not self.smtp_settings:
            ttk.Label(self.current_frame, text="SMTP sozlamalari kiritilmagan! Iltimos, SMTP Sozlamalari bo'limida sozlamalarni kiriting.",
                     font=("Helvetica", 12, "italic"), foreground="red").pack(pady=10)
            ttk.Button(self.current_frame, text="Bosh menyuga",
                      command=self.create_main_interface).pack(pady=10)
            return
        
        ttk.Label(self.current_frame, text="Qabul Qiluvchi Tanlash",
                 font=("Helvetica", 14, "bold")).pack(pady=5)
        
        contact_frame = ttk.Frame(self.current_frame)
        contact_frame.pack(fill="x", pady=5)
        ttk.Label(contact_frame, text="Kontakt:", width=15).pack(side="left")
        self.contact_choice = tk.StringVar()
        contact_options = list(self.contacts.keys())
        self.contact_menu = ttk.OptionMenu(contact_frame, self.contact_choice,
                                          contact_options[0] if contact_options else "Kontaktlar mavjud emas",
                                          *contact_options)
        self.contact_menu.pack(fill="x", expand=True, padx=5)
        
        ttk.Button(self.current_frame, text="Email Jo'natish",
                  command=self.send_email_after_attendance).pack(pady=10)
        ttk.Button(self.current_frame, text="O'tkazib Yuborish",
                  command=self.show_summary).pack(pady=5)

    def send_email_after_attendance(self):
        """Davomat hisobotini email orqali yuborish"""
        contact_name = self.contact_choice.get()
        if contact_name not in self.contacts:
            messagebox.showwarning("Ogohlantirish", "Kontaktni tanlang!")
            return
        
        sender = self.smtp_settings["email"]
        receiver = self.contacts[contact_name]["email"]
        smtp_server = self.smtp_settings["smtp_server"]
        smtp_port = self.smtp_settings["smtp_port"]
        password = self.smtp_settings["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = f"Davomat Hisoboti {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        body = "Davomat hisoboti ilova sifatida yuborildi."
        msg.attach(MIMEText(body, 'plain'))
        
        try:
            with open(self.filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(self.filename)}"
            )
            msg.attach(part)
            
            progress_label = ttk.Label(self.current_frame, text="Email yuborilmoqda...", font=("Helvetica", 12))
            progress_label.pack(pady=5)
            self.root.update()
            
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender, password)
            text = msg.as_string()
            server.sendmail(sender, receiver, text)
            server.quit()
            
            progress_label.destroy()
            messagebox.showinfo("Muvaffaqiyat", "Hisobot email orqali yuborildi!")
            self.show_summary()
        except Exception as e:
            progress_label.destroy()
            messagebox.showerror("Xato", f"Email yuborishda xato: {e}")
            self.show_summary()

    def show_summary(self):
        """Davomat yakuniy hisobotini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Davomat Yakuni",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        summary_text = tk.Text(self.current_frame, height=20, width=80)
        summary_text.pack(pady=10)
        
        summary_text.insert(tk.END, "=== Davomat Yakuni ===\n\n")
        summary_text.insert(tk.END, "Kelganlar:\n")
        for name, data in self.attendance.items():
            if data["status"] == "Kelgan":
                summary_text.insert(tk.END, 
                    f"- {name} {data['surname']} {data['father_name']} "
                    f"({data['faculty']}, {data['direction']}, {data['group']}) - "
                    f"{data['arrival_time']}\n"
                    f"  Ehtimollik: {data['probability']:.2%}\n"
                    f"  O'rtacha masofa: {data['mean_distance']:.4f}\n"
                    f"  Dispersiya: {data['variance']:.4f}\n"
                    f"  Kvadrat chetlanish: {data['std_dev']:.4f}\n"
                )
        
        summary_text.insert(tk.END, "\nKech qolganlar:\n")
        for name, data in self.attendance.items():
            if data["status"] == "Kech qolgan":
                summary_text.insert(tk.END, 
                    f"- {name} {data['surname']} {data['father_name']} "
                    f"({data['faculty']}, {data['direction']}, {data['group']}) - "
                    f"{data['late_time']}\n"
                    f"  Ehtimollik: {data['probability']:.2%}\n"
                    f"  O'rtacha masofa: {data['mean_distance']:.4f}\n"
                    f"  Dispersiya: {data['variance']:.4f}\n"
                    f"  Kvadrat chetlanish: {data['std_dev']:.4f}\n"
                )
        
        summary_text.insert(tk.END, "\nKelmaganlar:\n")
        for name, data in self.attendance.items():
            if data["status"] == "Kelmagan":
                summary_text.insert(tk.END, 
                    f"- {name} {data['surname']} {data['father_name']} "
                    f"({data['faculty']}, {data['direction']}, {data['group']})\n"
                )
        
        summary_text.config(state="disabled")
        
        ttk.Button(self.current_frame, text="Bosh menyuga",
                  command=self.create_main_interface).pack(pady=10)

    def show_contacts_section(self):
        """Kontaktlar bo'limini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Kontaktlar Bo'limi",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        ttk.Label(self.current_frame, text="Yangi Kontakt Qo'shish",
                 font=("Helvetica", 14, "bold")).pack(pady=5)
        
        name_frame = ttk.Frame(self.current_frame)
        name_frame.pack(fill="x", pady=5)
        ttk.Label(name_frame, text="Ism:", width=15).pack(side="left")
        self.contact_name_entry = ttk.Entry(name_frame)
        self.contact_name_entry.pack(fill="x", expand=True, padx=5)
        
        email_frame = ttk.Frame(self.current_frame)
        email_frame.pack(fill="x", pady=5)
        ttk.Label(email_frame, text="Email:", width=15).pack(side="left")
        self.contact_email_entry = ttk.Entry(email_frame)
        self.contact_email_entry.pack(fill="x", expand=True, padx=5)
        
        ttk.Button(self.current_frame, text="Kontaktni Saqlash",
                  command=self.save_contact).pack(pady=10)
        
        ttk.Label(self.current_frame, text="Saqlangan Kontaktlar",
                 font=("Helvetica", 14, "bold")).pack(pady=5)
        
        self.contacts_listbox = tk.Listbox(self.current_frame, height=10, width=60)
        self.contacts_listbox.pack(pady=5)
        self.update_contacts_listbox()
        
        ttk.Button(self.current_frame, text="Kontaktni O'chirish",
                  command=self.delete_contact).pack(pady=5)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.create_main_interface).pack(pady=5)

    def save_contact(self):
        """Yangi kontaktni saqlash"""
        name = self.contact_name_entry.get().strip()
        email = self.contact_email_entry.get().strip()
        
        if not all([name, email]):
            messagebox.showwarning("Ogohlantirish", "Ism va Email maydonlarini to'ldiring!")
            return
        
        self.contacts[name] = {
            "email": email
        }
        self.save_contacts()
        self.update_contacts_listbox()
        messagebox.showinfo("Muvaffaqiyat", f"{name} kontakt sifatida saqlandi!")
        
        self.contact_name_entry.delete(0, tk.END)
        self.contact_email_entry.delete(0, tk.END)

    def update_contacts_listbox(self):
        """Kontaktlar ro'yxatini yangilash"""
        self.contacts_listbox.delete(0, tk.END)
        for name, data in self.contacts.items():
            self.contacts_listbox.insert(tk.END, f"{name}: {data['email']}")

    def delete_contact(self):
        """Kontaktni o'chirish"""
        selection = self.contacts_listbox.curselection()
        if not selection:
            messagebox.showwarning("Ogohlantirish", "Kontaktni tanlang!")
            return
        
        name = self.contacts_listbox.get(selection[0]).split(":")[0].strip()
        del self.contacts[name]
        self.save_contacts()
        self.update_contacts_listbox()
        messagebox.showinfo("Muvaffaqiyat", f"{name} kontakt o'chirildi!")

    def show_smtp_settings_section(self):
        """SMTP sozlamalari bo'limini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="SMTP Sozlamalari",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        email_frame = ttk.Frame(self.current_frame)
        email_frame.pack(fill="x", pady=5)
        ttk.Label(email_frame, text="Yuboruvchi Email:", width=15).pack(side="left")
        self.smtp_email_entry = ttk.Entry(email_frame)
        self.smtp_email_entry.pack(fill="x", expand=True, padx=5)
        self.smtp_email_entry.insert(0, self.smtp_settings.get("email", ""))
        
        smtp_frame = ttk.Frame(self.current_frame)
        smtp_frame.pack(fill="x", pady=5)
        ttk.Label(smtp_frame, text="SMTP server:", width=15).pack(side="left")
        self.smtp_server_entry = ttk.Entry(smtp_frame)
        self.smtp_server_entry.pack(fill="x", expand=True, padx=5)
        self.smtp_server_entry.insert(0, self.smtp_settings.get("smtp_server", ""))
        
        port_frame = ttk.Frame(self.current_frame)
        port_frame.pack(fill="x", pady=5)
        ttk.Label(port_frame, text="SMTP port:", width=15).pack(side="left")
        self.smtp_port_entry = ttk.Entry(port_frame)
        self.smtp_port_entry.pack(fill="x", expand=True, padx=5)
        self.smtp_port_entry.insert(0, self.smtp_settings.get("smtp_port", "587"))
        
        password_frame = ttk.Frame(self.current_frame)
        password_frame.pack(fill="x", pady=5)
        ttk.Label(password_frame, text="Parol:", width=15).pack(side="left")
        self.smtp_password_entry = ttk.Entry(password_frame, show="*")
        self.smtp_password_entry.pack(fill="x", expand=True, padx=5)
        self.smtp_password_entry.insert(0, self.smtp_settings.get("password", ""))
        
        ttk.Button(self.current_frame, text="Sozlamalarni Saqlash",
                  command=self.save_smtp_settings_ui).pack(pady=10)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.create_main_interface).pack(pady=5)

    def save_smtp_settings_ui(self):
        """SMTP sozlamalarini saqlash"""
        email = self.smtp_email_entry.get().strip()
        smtp_server = self.smtp_server_entry.get().strip()
        try:
            smtp_port = int(self.smtp_port_entry.get() or 587)
        except ValueError:
            smtp_port = 587
        password = self.smtp_password_entry.get().strip()
        
        if not all([email, smtp_server, password]):
            messagebox.showwarning("Ogohlantirish", "Barcha maydonlarni to'ldiring!")
            return
        
        self.smtp_settings = {
            "email": email,
            "smtp_server": smtp_server,
            "smtp_port": smtp_port,
            "password": password
        }
        self.save_smtp_settings()
        messagebox.showinfo("Muvaffaqiyat", "SMTP sozlamalari saqlandi!")

    def show_email_section(self):
        """Email jo'natish bo'limini ko'rsatish"""
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = ttk.Frame(self.root)
        self.current_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(self.current_frame, text="Email Jo'natish Bo'limi",
                 font=("Helvetica", 18, "bold")).pack(pady=10)
        
        ttk.Label(self.current_frame, text="Hisobot Faylini Tanlash",
                 font=("Helvetica", 14, "bold")).pack(pady=5)
        
        file_frame = ttk.Frame(self.current_frame)
        file_frame.pack(fill="x", pady=5)
        ttk.Label(file_frame, text="Fayl:", width=15).pack(side="left")
        self.email_file_var = tk.StringVar()
        ttk.Button(file_frame, text="Tanlash",
                  command=self.select_email_file).pack(side="left", padx=5)
        ttk.Entry(file_frame, textvariable=self.email_file_var,
                 state="readonly").pack(fill="x", expand=True, padx=5)
        
        ttk.Label(self.current_frame, text="Qabul Qiluvchi Tanlash",
                 font=("Helvetica", 14, "bold")).pack(pady=5)
        
        contact_frame = ttk.Frame(self.current_frame)
        contact_frame.pack(fill="x", pady=5)
        ttk.Label(contact_frame, text="Kontakt:", width=15).pack(side="left")
        self.contact_choice = tk.StringVar()
        contact_options = list(self.contacts.keys())
        self.contact_menu = ttk.OptionMenu(contact_frame, self.contact_choice,
                                          contact_options[0] if contact_options else "Kontaktlar mavjud emas",
                                          *contact_options)
        self.contact_menu.pack(fill="x", expand=True, padx=5)
        
        ttk.Button(self.current_frame, text="Email Jo'natish",
                  command=self.send_email).pack(pady=10)
        ttk.Button(self.current_frame, text="Orqaga",
                  command=self.create_main_interface).pack(pady=5)

    def select_email_file(self):
        """Email uchun fayl tanlash"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("Barcha fayllar", "*.*")])
        if file_path:
            self.email_file_var.set(file_path)

    def send_email(self):
        """Hisobot faylini email orqali yuborish"""
        if not self.email_file_var.get():
            messagebox.showwarning("Ogohlantirish", "Hisobot faylini tanlang!")
            return
        
        contact_name = self.contact_choice.get()
        if contact_name not in self.contacts:
            messagebox.showwarning("Ogohlantirish", "Kontaktni tanlang!")
            return
        
        if not self.smtp_settings:
            messagebox.showwarning("Ogohlantirish", "SMTP sozlamalari kiritilmagan! Iltimos, SMTP Sozlamalari bo'limida sozlamalarni kiriting.")
            return
        
        sender = self.smtp_settings["email"]
        receiver = self.contacts[contact_name]["email"]
        smtp_server = self.smtp_settings["smtp_server"]
        smtp_port = self.smtp_settings["smtp_port"]
        password = self.smtp_settings["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = f"Davomat Hisoboti {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        body = "Davomat hisoboti ilova sifatida yuborildi."
        msg.attach(MIMEText(body, 'plain'))
        
        try:
            with open(self.email_file_var.get(), "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(self.email_file_var.get())}"
            )
            msg.attach(part)
            
            progress_label = ttk.Label(self.current_frame, text="Email yuborilmoqda...", font=("Helvetica", 12))
            progress_label.pack(pady=5)
            self.root.update()
            
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender, password)
            text = msg.as_string()
            server.sendmail(sender, receiver, text)
            server.quit()
            
            progress_label.destroy()
            messagebox.showinfo("Muvaffaqiyat", "Hisobot email orqali yuborildi!")
            self.create_main_interface()
        except Exception as e:
            progress_label.destroy()
            messagebox.showerror("Xato", f"Email yuborishda xato: {e}")
            self.create_main_interface()

    def run(self):
        """Ilovani ishga tushirish"""
        try:
            self.root.mainloop()
        except Exception as e:
            print(f"Ilova ishga tushirishda xato: {e}")

if __name__ == "__main__":
    app = AttendanceApp()
    app.run()
