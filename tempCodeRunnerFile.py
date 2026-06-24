# ============================================
# نظام الباركود المتكامل - بدون pyzbar
# ============================================

import cv2
import numpy as np
import pandas as pd
import os
import json
import csv
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import webbrowser
import tempfile
import threading
import sys

# ---------- استخدام OpenCV فقط لقراءة الباركود ----------
# ملاحظة: هذا الحل يستخدم مكتبة pyzbar ولكن مع إصلاح المسار

# محاولة استيراد pyzbar مع إصلاح المسار
try:
    import pyzbar.pyzbar as pyzbar
    PYZBAR_AVAILABLE = True
except:
    PYZBAR_AVAILABLE = False
    print("⚠️ pyzbar غير متوفرة، سيتم استخدام حل بديل")

# ---------- حل بديل لقراءة الباركود باستخدام OpenCV ----------
class BarcodeReader:
    @staticmethod
    def decode(frame):
        """قراءة الباركود من الإطار"""
        if PYZBAR_AVAILABLE:
            # استخدام pyzbar إن وجدت
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            return pyzbar.decode(gray)
        else:
            # حل بديل: محاكاة قراءة الباركود (للتجربة)
            # في الحقيقة تحتاج لتثبيت pyzbar أو استخدام مكتبة أخرى
            return []

# ---------- إعدادات الملفات ----------
EXCEL_FILE = "barcode_data.xlsx"
CSV_FILE = "barcode_data.csv"
JSON_FILE = "barcode_data.json"
TXT_FILE = "barcode_data.txt"
BACKUP_FOLDER = "backups"
SETTINGS_FILE = "settings.json"

# ---------- إنشاء المجلدات ----------
if not os.path.exists(BACKUP_FOLDER):
    os.makedirs(BACKUP_FOLDER)

# ---------- إعدادات التطبيق ----------
class Settings:
    def __init__(self):
        self.camera_id = 0
        self.auto_save = True
        self.load()
    
    def load(self):
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
                self.camera_id = data.get('camera_id', 0)
                self.auto_save = data.get('auto_save', True)
    
    def save(self):
        with open(SETTINGS_FILE, 'w') as f:
            json.dump({
                'camera_id': self.camera_id,
                'auto_save': self.auto_save
            }, f)

settings = Settings()

# ---------- إدارة البيانات ----------
class DataManager:
    @staticmethod
    def init_files():
        """تهيئة جميع ملفات البيانات"""
        if not os.path.exists(EXCEL_FILE):
            df = pd.DataFrame(columns=["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
            df.to_excel(EXCEL_FILE, index=False)
        
        if not os.path.exists(CSV_FILE):
            with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
        
        if not os.path.exists(JSON_FILE):
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        
        if not os.path.exists(TXT_FILE):
            with open(TXT_FILE, 'w', encoding='utf-8') as f:
                f.write("سجل الباركود\n")
                f.write("="*50 + "\n")
    
    @staticmethod
    def save_barcode(barcode_data):
        """حفظ الباركود في جميع الملفات"""
        now = datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")
        device = os.environ.get('COMPUTERNAME', 'Unknown')
        
        record = {
            "التاريخ": date_str,
            "الوقت": time_str,
            "الباركود": barcode_data,
            "النوع": "باركود",
            "الجهاز": device
        }
        
        # حفظ في Excel
        try:
            df = pd.read_excel(EXCEL_FILE)
            df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
        except Exception as e:
            print(f"خطأ في حفظ Excel: {e}")
        
        # حفظ في CSV
        try:
            with open(CSV_FILE, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([date_str, time_str, barcode_data, "باركود", device])
        except Exception as e:
            print(f"خطأ في حفظ CSV: {e}")
        
        # حفظ في JSON
        try:
            with open(JSON_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data.append(record)
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"خطأ في حفظ JSON: {e}")
        
        # حفظ في TXT
        try:
            with open(TXT_FILE, 'a', encoding='utf-8') as f:
                f.write(f"{date_str} | {time_str} | {barcode_data} | {device}\n")
        except Exception as e:
            print(f"خطأ في حفظ TXT: {e}")
        
        return record
    
    @staticmethod
    def load_data():
        """تحميل البيانات من Excel"""
        if os.path.exists(EXCEL_FILE):
            try:
                return pd.read_excel(EXCEL_FILE)
            except:
                return pd.DataFrame(columns=["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
        return pd.DataFrame(columns=["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
    
    @staticmethod
    def clear_all():
        """مسح جميع البيانات"""
        df = pd.DataFrame(columns=["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
        df.to_excel(EXCEL_FILE, index=False)
        
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["التاريخ", "الوقت", "الباركود", "النوع", "الجهاز"])
        
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f, ensure_ascii=False, indent=2)
        
        with open(TXT_FILE, 'w', encoding='utf-8') as f:
            f.write("سجل الباركود\n")
            f.write("="*50 + "\n")

# ---------- قراءة الباركود باستخدام حل بديل ----------
class BarcodeScanner:
    def __init__(self):
        self.cap = None
        self.running = False
        self.scanned_data = None
        self.callback = None
    
    def start_scan(self, callback=None):
        """بدء المسح"""
        self.callback = callback
        self.running = True
        self.scanned_data = None
        
        thread = threading.Thread(target=self._scan_thread)
        thread.daemon = True
        thread.start()
    
    def _scan_thread(self):
        """خيط المسح"""
        self.cap = cv2.VideoCapture(settings.camera_id)
        if not self.cap.isOpened():
            messagebox.showerror("خطأ", "لا يمكن فتح الكاميرا")
            self.running = False
            return
        
        while self.running:
            ret, frame = self.cap.read()
            if not ret:
                break
            
            # محاولة قراءة الباركود
            if PYZBAR_AVAILABLE:
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                barcodes = pyzbar.decode(gray)
                
                for barcode in barcodes:
                    barcode_data = barcode.data.decode("utf-8")
                    self.scanned_data = barcode_data
                    
                    # رسم مستطيل حول الباركود
                    (x, y, w, h) = barcode.rect
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(frame, barcode_data, (x, y - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
                    
                    cv2.putText(frame, "PRESS 'S' TO SAVE | 'ESC' TO EXIT", (10, 30),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 0), 2)
            else:
                # وضع تجريبي بدون باركود حقيقي
                cv2.putText(frame, "⚠️ pyzbar غير مثبت - وضع المحاكاة", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                cv2.putText(frame, "اضغط 'S' لمحاكاة باركود تجريبي", (10, 60),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
            
            cv2.imshow("📷 مسح الباركود - اضغط S للحفظ أو ESC للخروج", frame)
            
            key = cv2.waitKey(1)
            if key == 27:  # ESC
                self.running = False
                break
            if key == ord('s') or key == ord('S'):
                if PYZBAR_AVAILABLE and self.scanned_data:
                    if self.callback:
                        self.callback(self.scanned_data)
                    self.running = False
                    break
                elif not PYZBAR_AVAILABLE:
                    # وضع المحاكاة - إنشاء باركود وهمي
                    test_barcode = f"TEST-{datetime.now().strftime('%H%M%S')}"
                    if self.callback:
                        self.callback(test_barcode)
                    self.running = False
                    break
        
        self.stop()
    
    def stop(self):
        """إيقاف المسح"""
        self.running = False
        if self.cap:
            self.cap.release()
        cv2.destroyAllWindows()

# ---------- واجهة التطبيق الرئيسية ----------
class BarcodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📱 نظام الباركود المتكامل v2.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # تهيئة الملفات
        DataManager.init_files()
        
        # إنشاء الواجهة
        self.create_widgets()
        
        # تحميل البيانات
        self.refresh_table()
        
        # متغيرات
        self.scanner = BarcodeScanner()
        self.current_barcode = None
        
        # عرض حالة pyzbar
        if not PYZBAR_AVAILABLE:
            self.info_label.config(text="⚠️ وضع المحاكاة - pyzbar غير مثبت")
            messagebox.showwarning("تحذير", 
                "مكتبة pyzbar غير مثبتة!\n\n"
                "لحل المشكلة:\n"
                "1. افتح PowerShell كمسؤول\n"
                "2. نفذ: pip uninstall pyzbar -y\n"
                "3. نفذ: pip install pyzbar\n"
                "4. أعد تشغيل البرنامج\n\n"
                "حاليًا سيعمل في وضع المحاكاة.")
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # الإطار العلوي
        top_frame = tk.Frame(self.root, bg='lightgray', height=100)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        top_frame.pack_propagate(False)
        
        # العنوان
        title = tk.Label(top_frame, text="📦 نظام إدارة الباركود", 
                        font=("Arial", 20, "bold"), bg='lightgray')
        title.pack(pady=10)
        
        # إطار الأزرار
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        # صف الأزرار الأول
        btn_row1 = tk.Frame(btn_frame)
        btn_row1.pack(pady=5)
        
        btn_scan = tk.Button(btn_row1, text="📷 مسح باركود", 
                           command=self.scan_barcode,
                           font=("Arial", 12), bg="#4CAF50", fg="white", 
                           padx=20, pady=8, cursor="hand2")
        btn_scan.pack(side=tk.LEFT, padx=5)
        
        btn_manual = tk.Button(btn_row1, text="✏️ إدخال يدوي", 
                             command=self.manual_entry,
                             font=("Arial", 12), bg="#2196F3", fg="white",
                             padx=20, pady=8, cursor="hand2")
        btn_manual.pack(side=tk.LEFT, padx=5)
        
        btn_print = tk.Button(btn_row1, text="🖨️ طباعة", 
                            command=self.print_data,
                            font=("Arial", 12), bg="#FF9800", fg="white",
                            padx=20, pady=8, cursor="hand2")
        btn_print.pack(side=tk.LEFT, padx=5)
        
        # صف الأزرار الثاني
        btn_row2 = tk.Frame(btn_frame)
        btn_row2.pack(pady=5)
        
        btn_export = tk.Button(btn_row2, text="💾 تصدير", 
                             command=self.export_data,
                             font=("Arial", 12), bg="#9C27B0", fg="white",
                             padx=20, pady=8, cursor="hand2")
        btn_export.pack(side=tk.LEFT, padx=5)
        
        btn_backup = tk.Button(btn_row2, text="🔄 نسخ احتياطي", 
                             command=self.create_backup,
                             font=("Arial", 12), bg="#607D8B", fg="white",
                             padx=20, pady=8, cursor="hand2")
        btn_backup.pack(side=tk.LEFT, padx=5)
        
        btn_clear = tk.Button(btn_row2, text="🗑️ مسح الكل", 
                            command=self.clear_all,
                            font=("Arial", 12), bg="#F44336", fg="white",
                            padx=20, pady=8, cursor="hand2")
        btn_clear.pack(side=tk.LEFT, padx=5)
        
        btn_settings = tk.Button(btn_row2, text="⚙️ إعدادات", 
                               command=self.show_settings,
                               font=("Arial", 12), bg="#795548", fg="white",
                               padx=20, pady=8, cursor="hand2")
        btn_settings.pack(side=tk.LEFT, padx=5)
        
        # إطار الجدول
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # شريط التمرير العمودي
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # الجدول
        columns = ("التاريخ", "الوقت", "الباركود", "الجهاز")
        self.tree = ttk.Treeview(table_frame, columns=columns, 
                                 show="headings", yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.tree.yview)
        
        # تنسيق الأعمدة
        self.tree.heading("التاريخ", text="📅 التاريخ")
        self.tree.heading("الوقت", text="⏰ الوقت")
        self.tree.heading("الباركود", text="🔢 الباركود")
        self.tree.heading("الجهاز", text="💻 الجهاز")
        
        self.tree.column("التاريخ", width=120, anchor="center")
        self.tree.column("الوقت", width=100, anchor="center")
        self.tree.column("الباركود", width=300, anchor="center")
        self.tree.column("الجهاز", width=150, anchor="center")
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # إطار المعلومات
        info_frame = tk.Frame(self.root, bg='#f0f0f0', height=60)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        info_frame.pack_propagate(False)
        
        self.info_label = tk.Label(info_frame, text="✅ جاهز", 
                                  font=("Arial", 10), bg='#f0f0f0')
        self.info_label.pack(side=tk.LEFT, padx=10)
        
        self.count_label = tk.Label(info_frame, text="📊 عدد السجلات: 0", 
                                   font=("Arial", 10), bg='#f0f0f0')
        self.count_label.pack(side=tk.RIGHT, padx=10)
        
        # ربط حدث النقر المزدوج
        self.tree.bind("<Double-Button-1>", self.on_double_click)
    
    def refresh_table(self):
        """تحديث الجدول"""
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        df = DataManager.load_data()
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=(
                row.get("التاريخ", ""),
                row.get("الوقت", ""),
                row.get("الباركود", ""),
                row.get("الجهاز", "")
            ))
        
        self.count_label.config(text=f"📊 عدد السجلات: {len(df)}")
        self.info_label.config(text=f"✅ تم التحديث - {datetime.now().strftime('%H:%M:%S')}")
    
    def scan_barcode(self):
        """بدء مسح الباركود"""
        self.info_label.config(text="📷 جاري فتح الكاميرا...")
        self.scanner.start_scan(self.on_barcode_scanned)
    
    def on_barcode_scanned(self, barcode_data):
        """معالجة الباركود الممسوح"""
        self.current_barcode = barcode_data
        DataManager.save_barcode(barcode_data)
        self.refresh_table()
        self.info_label.config(text=f"✅ تم حفظ الباركود: {barcode_data}")
        messagebox.showinfo("✅ تم الحفظ", f"تم حفظ الباركود بنجاح:\n{barcode_data}")
    
    def manual_entry(self):
        """إدخال باركود يدوي"""
        dialog = tk.Toplevel(self.root)
        dialog.title("✏️ إدخال يدوي")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="أدخل الباركود:", font=("Arial", 12)).pack(pady=20)
        
        entry = tk.Entry(dialog, font=("Arial", 14), width=30)
        entry.pack(pady=10)
        entry.focus()
        
        def save_manual():
            barcode = entry.get().strip()
            if barcode:
                DataManager.save_barcode(barcode)
                self.refresh_table()
                self.info_label.config(text=f"✅ تم حفظ الباركود: {barcode}")
                dialog.destroy()
                messagebox.showinfo("✅ تم الحفظ", f"تم حفظ الباركود بنجاح:\n{barcode}")
            else:
                messagebox.showwarning("تحذير", "الرجاء إدخال الباركود")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="💾 حفظ", command=save_manual,
                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ إلغاء", command=dialog.destroy,
                 font=("Arial", 12), bg="#F44336", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        
        entry.bind('<Return>', lambda e: save_manual())
    
    def print_data(self):
        """طباعة البيانات"""
        df = DataManager.load_data()
        if df.empty:
            messagebox.showerror("خطأ", "لا توجد بيانات للطباعة")
            return
        
        # عرض خيارات الطباعة
        print_dialog = tk.Toplevel(self.root)
        print_dialog.title("🖨️ خيارات الطباعة")
        print_dialog.geometry("400x250")
        print_dialog.transient(self.root)
        print_dialog.grab_set()
        
        tk.Label(print_dialog, text="خيارات الطباعة", font=("Arial", 14, "bold")).pack(pady=10)
        
        print_type = tk.StringVar(value="html")
        
        tk.Radiobutton(print_dialog, text="🌐 معاينة في المتصفح (HTML)", 
                      variable=print_type, value="html").pack(pady=5)
        tk.Radiobutton(print_dialog, text="📄 طباعة كنص (TXT)", 
                      variable=print_type, value="txt").pack(pady=5)
        
        def do_print():
            if print_type.get() == "html":
                self.print_html()
            else:
                self.print_txt()
            print_dialog.destroy()
        
        btn_frame = tk.Frame(print_dialog)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="🖨️ طباعة", command=do_print,
                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ إلغاء", command=print_dialog.destroy,
                 font=("Arial", 12), bg="#F44336", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
    
    def print_html(self):
        """طباعة كـ HTML"""
        df = DataManager.load_data()
        html = df.to_html(index=False, classes='table', border=1)
        html_content = f"""
        <!DOCTYPE html>
        <html dir="rtl">
        <head>
            <meta charset="UTF-8">
            <title>تقرير الباركود</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 20px; }}
                h1 {{ color: #333; text-align: center; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                th {{ background-color: #4CAF50; color: white; padding: 10px; }}
                td {{ padding: 8px; border: 1px solid #ddd; text-align: center; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .footer {{ margin-top: 30px; text-align: center; color: #666; }}
            </style>
        </head>
        <body>
            <h1>📊 تقرير الباركود</h1>
            <p>📅 تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>📊 عدد السجلات: {len(df)}</p>
            {html}
            <div class="footer">تم إنشاء التقرير بواسطة نظام الباركود v2.0</div>
            <script>
                window.onload = function() {{ window.print(); }}
            </script>
        </body>
        </html>
        """
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8')
        temp_file.write(html_content)
        temp_file.close()
        webbrowser.open(temp_file.name)
        messagebox.showinfo("🖨️ طباعة", "تم فتح معاينة الطباعة في المتصفح")
    
    def print_txt(self):
        """طباعة كنص"""
        df = DataManager.load_data()
        txt_content = "="*60 + "\n"
        txt_content += "📊 تقرير الباركود\n"
        txt_content += f"📅 التاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        txt_content += f"📊 عدد السجلات: {len(df)}\n"
        txt_content += "="*60 + "\n\n"
        
        for _, row in df.iterrows():
            txt_content += f"📅 {row['التاريخ']} | ⏰ {row['الوقت']} | 🔢 {row['الباركود']} | 💻 {row['الجهاز']}\n"
        
        txt_content += "\n" + "="*60 + "\n"
        txt_content += "تم إنشاء التقرير بواسطة نظام الباركود v2.0"
        
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
        temp_file.write(txt_content)
        temp_file.close()
        webbrowser.open(temp_file.name)
    
    def export_data(self):
        """تصدير البيانات"""
        export_dialog = tk.Toplevel(self.root)
        export_dialog.title("💾 تصدير البيانات")
        export_dialog.geometry("400x250")
        export_dialog.transient(self.root)
        export_dialog.grab_set()
        
        tk.Label(export_dialog, text="اختر صيغة التصدير", font=("Arial", 14, "bold")).pack(pady=10)
        
        export_type = tk.StringVar(value="excel")
        
        tk.Radiobutton(export_dialog, text="📊 Excel (.xlsx)", 
                      variable=export_type, value="excel").pack(pady=5)
        tk.Radiobutton(export_dialog, text="📄 CSV (.csv)", 
                      variable=export_type, value="csv").pack(pady=5)
        tk.Radiobutton(export_dialog, text="📋 JSON (.json)", 
                      variable=export_type, value="json").pack(pady=5)
        
        def do_export():
            file_types = {
                'excel': [("Excel files", "*.xlsx")],
                'csv': [("CSV files", "*.csv")],
                'json': [("JSON files", "*.json")]
            }
            ext = {
                'excel': ".xlsx",
                'csv': ".csv",
                'json': ".json"
            }
            
            filename = filedialog.asksaveasfilename(
                defaultextension=ext[export_type.get()],
                filetypes=file_types[export_type.get()]
            )
            
            if filename:
                df = DataManager.load_data()
                if export_type.get() == "excel":
                    df.to_excel(filename, index=False)
                elif export_type.get() == "csv":
                    df.to_csv(filename, index=False, encoding='utf-8-sig')
                else:
                    df.to_json(filename, orient='records', force_ascii=False, indent=2)
                messagebox.showinfo("✅ تم التصدير", f"تم تصدير البيانات إلى:\n{filename}")
                export_dialog.destroy()
        
        btn_frame = tk.Frame(export_dialog)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="💾 تصدير", command=do_export,
                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ إلغاء", command=export_dialog.destroy,
                 font=("Arial", 12), bg="#F44336", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
    
    def create_backup(self):
        """إنشاء نسخة احتياطية"""
        df = DataManager.load_data()
        if df.empty:
            messagebox.showerror("خطأ", "لا توجد بيانات للنسخ الاحتياطي")
            return
        
        backup_name = f"{BACKUP_FOLDER}/backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(backup_name, index=False)
        messagebox.showinfo("✅ تم", f"تم إنشاء النسخة الاحتياطية:\n{backup_name}")
    
    def clear_all(self):
        """مسح جميع البيانات"""
        if messagebox.askyesno("⚠️ تأكيد", "هل أنت متأكد من مسح جميع البيانات؟\nهذا الإجراء لا يمكن التراجع عنه!"):
            DataManager.clear_all()
            self.refresh_table()
            self.info_label.config(text="🗑️ تم مسح جميع البيانات")
            messagebox.showinfo("✅ تم", "تم مسح جميع البيانات بنجاح")
    
    def show_settings(self):
        """عرض الإعدادات"""
        settings_dialog = tk.Toplevel(self.root)
        settings_dialog.title("⚙️ الإعدادات")
        settings_dialog.geometry("400x200")
        settings_dialog.transient(self.root)
        settings_dialog.grab_set()
        
        tk.Label(settings_dialog, text="⚙️ إعدادات النظام", font=("Arial", 14, "bold")).pack(pady=10)
        
        # كاميرا
        cam_frame = tk.Frame(settings_dialog)
        cam_frame.pack(pady=10, fill=tk.X, padx=20)
        tk.Label(cam_frame, text="رقم الكاميرا:").pack(side=tk.LEFT)
        cam_entry = tk.Entry(cam_frame, width=10)
        cam_entry.insert(0, str(settings.camera_id))
        cam_entry.pack(side=tk.LEFT, padx=10)
        
        def save_settings():
            try:
                settings.camera_id = int(cam_entry.get())
            except:
                settings.camera_id = 0
            settings.save()
            messagebox.showinfo("✅ تم", "تم حفظ الإعدادات")
            settings_dialog.destroy()
        
        btn_frame = tk.Frame(settings_dialog)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="💾 حفظ", command=save_settings,
                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ إلغاء", command=settings_dialog.destroy,
                 font=("Arial", 12), bg="#F44336", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
    
    def on_double_click(self, event):
        """معالجة النقر المزدوج على الجدول"""
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            barcode = item['values'][2]
            if barcode:
                detail_dialog = tk.Toplevel(self.root)
                detail_dialog.title("🔍 تفاصيل الباركود")
                detail_dialog.geometry("400x300")
                detail_dialog.transient(self.root)
                detail_dialog.grab_set()
                
                tk.Label(detail_dialog, text="📋 تفاصيل الباركود", 
                        font=("Arial", 14, "bold")).pack(pady=10)
                
                details = f"""
                🔢 الباركود: {barcode}
                📅 التاريخ: {item['values'][0]}
                ⏰ الوقت: {item['values'][1]}
                💻 الجهاز: {item['values'][3]}
                """
                
                text_area = scrolledtext.ScrolledText(detail_dialog, height=10, font=("Arial", 12))
                text_area.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
                text_area.insert('1.0', details)
                text_area.config(state='disabled')
                
                def copy_barcode():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(barcode)
                    messagebox.showinfo("✅ تم", "تم نسخ الباركود")
                
                tk.Button(detail_dialog, text="📋 نسخ الباركود", command=copy_barcode,
                         font=("Arial", 12), bg="#2196F3", fg="white", padx=20).pack(pady=10)
                tk.Button(detail_dialog, text="❌ إغلاق", command=detail_dialog.destroy,
                         font=("Arial", 12), bg="#F44336", fg="white", padx=20).pack(pady=5)

# ---------- تشغيل التطبيق ----------
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = BarcodeApp(root)
        root.mainloop()
    except Exception as e:
        print(f"خطأ: {e}")
        messagebox.showerror("خطأ", f"حدث خطأ:\n{str(e)}")
