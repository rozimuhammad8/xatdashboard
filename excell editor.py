import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl import load_workbook
import threading
import os
import re

class ExcelMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Ism-Sharif Moslashtiruvchi (PNFL)")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # O'zgaruvchilar
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.status_text = ""
        self.is_processing = False
        
        # Asosiy interfeys
        self.create_widgets()
        
    def create_widgets(self):
        # Asosiy ramka
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sarlavha
        title_label = ttk.Label(
            main_frame, 
            text="Excel Ism-Sharif Moslashtiruvchi (PNFL)", 
            font=("Helvetica", 16, "bold")
        )
        title_label.pack(pady=10)
        
        # Kirish fayli
        file_frame = ttk.LabelFrame(main_frame, text="Fayl tanlash", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="Excel fayl:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(file_frame, textvariable=self.input_file_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Tanlash", command=self.select_file).grid(row=0, column=2, padx=5)
        
        # Chiqish fayli
        ttk.Label(file_frame, text="Saqlash joyi:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_file_path, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Tanlash", command=self.select_output_file).grid(row=1, column=2, padx=5)
        
        # Tugmalar
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        self.process_btn = ttk.Button(
            btn_frame, 
            text="Qayta ishlashni boshlash", 
            command=self.start_processing,
            width=30
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="Natijani ko'rish", 
            command=self.show_result,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        # Holat paneli
        status_frame = ttk.LabelFrame(main_frame, text="Holat", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Tayyor", foreground="blue")
        self.status_label.pack(side=tk.LEFT)
        
        self.progress_bar = ttk.Progressbar(
            status_frame, 
            mode='indeterminate',
            length=200
        )
        self.progress_bar.pack(side=tk.RIGHT)
        
        # Natija matni
        result_frame = ttk.LabelFrame(main_frame, text="Natija", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(
            result_frame, 
            wrap=tk.WORD, 
            height=15,
            font=("Consolas", 10)
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # Qo'shimcha ma'lumot
        info_frame = ttk.LabelFrame(main_frame, text="Ma'lumot", padding="10")
        info_frame.pack(fill=tk.X, pady=5)
        
        info_text = """
        📌 Qanday ishlaydi:
        1. '2-жадвал' varag'idagi F ustunidagi ismlarni oladi
        2. 'royxat' varag'idagi K, L, M (yoki N) ustunlaridan mos keladigan qatorni topadi
        3. 'royxat2' varag'idagi E, F, G (yoki D) ustunlaridan ham qidiradi
        4. Topilgan qatordagi PinFL ma'lumotini G ustuniga yozadi
        """
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(anchor="w")
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Excel faylini tanlang",
            filetypes=[("Excel fayllar", "*.xlsx *.xls"), ("Barcha fayllar", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            # Avtomatik chiqish fayli nomini yaratish
            base_name = os.path.splitext(file_path)[0]
            self.output_file_path.set(f"{base_name}_natija.xlsx")
    
    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Natijani saqlash joyi",
            defaultextension=".xlsx",
            filetypes=[("Excel fayllar", "*.xlsx"), ("Barcha fayllar", "*.*")]
        )
        if file_path:
            self.output_file_path.set(file_path)
    
    def log_message(self, message, is_error=False):
        """Xabarni natija oynasiga va holat paneliga chiqarish"""
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)
        self.status_label.config(
            text=message[:50] + "..." if len(message) > 50 else message,
            foreground="red" if is_error else "blue"
        )
        self.root.update()
    
    def normalize_name(self, text):
        """Ismni normalize qilish"""
        if not text:
            return ""
        text = str(text).strip()
        # Turli apostroflarni bir xilga keltirish
        text = text.replace("'", "‘").replace("'", "‘")
        text = text.upper()
        # Qo'shimcha bo'shliqlarni tozalash
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def extract_parts(self, full_name):
        """Ismdan qismlarni ajratish"""
        if not full_name:
            return "", "", ""
        
        parts = full_name.split()
        if len(parts) == 1:
            return parts[0], "", ""
        elif len(parts) == 2:
            return parts[0], parts[1], ""
        else:
            # O'g'li, Qizi kabi qo'shimchalarni hisobga olish
            patronymic = " ".join(parts[2:])
            return parts[0], parts[1], patronymic
    
    def build_royxat_index(self, sheet_royxat):
        """royxat sheetidan indeks yaratish"""
        index = {}
        
        for row in sheet_royxat.iter_rows(min_row=2, values_only=False):
            try:
                # J ustuni (index 9) - PinFL
                pinfl = row[9].value if len(row) > 9 and row[9] and row[9].value else ""
                
                # K, L, M ustunlari (index 10, 11, 12)
                first_name = row[10].value if len(row) > 10 and row[10] and row[10].value else ""
                surname = row[11].value if len(row) > 11 and row[11] and row[11].value else ""
                patronymic = row[12].value if len(row) > 12 and row[12] and row[12].value else ""
                
                # N ustuni (index 13) - to'liq ism
                full_name_col = row[13].value if len(row) > 13 and row[13] and row[13].value else ""
                
                if pinfl:
                    # Turli formatlarda kalit yaratish
                    variants = []
                    
                    if full_name_col:
                        variants.append(self.normalize_name(full_name_col))
                    
                    if first_name or surname or patronymic:
                        fn = self.normalize_name(first_name)
                        sn = self.normalize_name(surname)
                        pn = self.normalize_name(patronymic)
                        
                        if fn and sn:
                            variants.append(f"{sn} {fn} {pn}".strip())
                            variants.append(f"{fn} {sn} {pn}".strip())
                            variants.append(f"{sn} {fn}".strip())
                            variants.append(f"{fn} {sn}".strip())
                        
                        if fn:
                            variants.append(fn)
                        if sn:
                            variants.append(sn)
                    
                    # Har bir variantni indeksga qo'shish
                    for variant in variants:
                        if variant and len(variant) > 1:
                            # Agar variant mavjud bo'lmasa yoki uzunroq bo'lsa
                            if variant not in index or len(str(pinfl)) > 0:
                                index[variant] = pinfl
            
            except Exception:
                continue
        
        return index
    
    def build_royxat2_index(self, sheet_royxat2):
        """royxat2 sheetidan indeks yaratish"""
        index = {}
        
        for row in sheet_royxat2.iter_rows(min_row=2, values_only=False):
            try:
                # D ustuni (index 3) - PinFL
                pinfl = row[3].value if len(row) > 3 and row[3] and row[3].value else ""
                
                # E, F, G ustunlari (index 4, 5, 6)
                surname = row[4].value if len(row) > 4 and row[4] and row[4].value else ""
                first_name = row[5].value if len(row) > 5 and row[5] and row[5].value else ""
                patronymic = row[6].value if len(row) > 6 and row[6] and row[6].value else ""
                
                # I ustuni (index 8) - to'liq ism
                full_name_col = row[8].value if len(row) > 8 and row[8] and row[8].value else ""
                
                if pinfl:
                    variants = []
                    
                    if full_name_col:
                        variants.append(self.normalize_name(full_name_col))
                    
                    if first_name or surname or patronymic:
                        fn = self.normalize_name(first_name)
                        sn = self.normalize_name(surname)
                        pn = self.normalize_name(patronymic)
                        
                        if fn and sn:
                            variants.append(f"{sn} {fn} {pn}".strip())
                            variants.append(f"{fn} {sn} {pn}".strip())
                            variants.append(f"{sn} {fn}".strip())
                            variants.append(f"{fn} {sn}".strip())
                        
                        if fn:
                            variants.append(fn)
                        if sn:
                            variants.append(sn)
                    
                    for variant in variants:
                        if variant and len(variant) > 1:
                            if variant not in index or len(str(pinfl)) > 0:
                                index[variant] = pinfl
            
            except Exception:
                continue
        
        return index
    
    def find_pinfl(self, name, index1, index2):
        """Ism bo'yicha PinFL ni topish"""
        if not name:
            return None
        
        normalized = self.normalize_name(name)
        
        # 1. To'g'ridan-to'g'ri qidirish
        if normalized in index1:
            return index1[normalized]
        if normalized in index2:
            return index2[normalized]
        
        # 2. Qismlarga ajratib qidirish
        surname, first_name, patronymic = self.extract_parts(normalized)
        
        # Turli kombinatsiyalarni tekshirish
        variants = []
        if surname and first_name and patronymic:
            variants.append(f"{surname} {first_name} {patronymic}")
            variants.append(f"{first_name} {surname} {patronymic}")
        if surname and first_name:
            variants.append(f"{surname} {first_name}")
            variants.append(f"{first_name} {surname}")
        if surname:
            variants.append(surname)
        if first_name:
            variants.append(first_name)
        
        # O'g'li/Qizi qo'shimchalarini tekshirish
        if surname and first_name and 'O‘G‘LI' in normalized:
            variants.append(f"{surname} {first_name}")
        if surname and first_name and 'QIZI' in normalized:
            variants.append(f"{surname} {first_name}")
        
        for variant in variants:
            if variant in index1:
                return index1[variant]
            if variant in index2:
                return index2[variant]
        
        # 3. Qisman moslik (eng yaxshi moslikni topish)
        best_match = None
        best_score = 0
        
        # royxat indeksida qidirish
        for key, value in index1.items():
            score = self.calculate_similarity(normalized, key)
            if score > best_score and score >= 60:  # 60% dan yuqori
                best_score = score
                best_match = value
        
        # royxat2 indeksida qidirish
        for key, value in index2.items():
            score = self.calculate_similarity(normalized, key)
            if score > best_score and score >= 60:
                best_score = score
                best_match = value
        
        return best_match
    
    def calculate_similarity(self, name1, name2):
        """Ikki ism o'rtasidagi o'xshashlikni hisoblash"""
        if not name1 or not name2:
            return 0
        
        name1_parts = set(name1.split())
        name2_parts = set(name2.split())
        
        if not name1_parts or not name2_parts:
            return 0
        
        common = len(name1_parts & name2_parts)
        total = len(name1_parts | name2_parts)
        
        if total == 0:
            return 0
        
        return (common / total) * 100
    
    def start_processing(self):
        if self.is_processing:
            messagebox.showwarning("Ogohlantirish", "Jarayon allaqachon ishlamoqda!")
            return
        
        if not self.input_file_path.get():
            messagebox.showerror("Xato", "Iltimos, Excel faylini tanlang!")
            return
        
        if not os.path.exists(self.input_file_path.get()):
            messagebox.showerror("Xato", "Tanlangan fayl mavjud emas!")
            return
        
        # Natija oynasini tozalash
        self.result_text.delete(1.0, tk.END)
        
        # Holatni yangilash
        self.is_processing = True
        self.process_btn.config(state=tk.DISABLED)
        self.progress_bar.start()
        
        # Alohida threadda ishlatish
        thread = threading.Thread(target=self.process_file, daemon=True)
        thread.start()
    
    def process_file(self):
        try:
            file_path = self.input_file_path.get()
            output_path = self.output_file_path.get()
            
            if not output_path:
                output_path = os.path.splitext(file_path)[0] + "_natija.xlsx"
            
            self.log_message(f"📂 Fayl ochilmoqda: {file_path}")
            
            # Excel faylini yuklash
            wb = load_workbook(file_path, data_only=True)
            
            # Varaqlarni tekshirish
            if '2-жадвал' not in wb.sheetnames:
                self.log_message("❌ Xato: '2-жадвал' varag'i topilmadi!", True)
                self.finish_processing()
                return
            
            sheet_main = wb['2-жадвал']
            
            # royxat indeksini yaratish
            self.log_message("📊 'royxat' ma'lumotlari indekslanmoqda...")
            royxat_index = {}
            if 'royxat' in wb.sheetnames:
                sheet_royxat = wb['royxat']
                royxat_index = self.build_royxat_index(sheet_royxat)
                self.log_message(f"✅ 'royxat' dan {len(royxat_index)} ta yozuv indekslandi")
            else:
                self.log_message("⚠️ 'royxat' varag'i topilmadi, faqat 'royxat2' dan qidiriladi", True)
            
            # royxat2 indeksini yaratish
            self.log_message("📊 'royxat2' ma'lumotlari indekslanmoqda...")
            royxat2_index = {}
            if 'royxat2' in wb.sheetnames:
                sheet_royxat2 = wb['royxat2']
                royxat2_index = self.build_royxat2_index(sheet_royxat2)
                self.log_message(f"✅ 'royxat2' dan {len(royxat2_index)} ta yozuv indekslandi")
            else:
                self.log_message("⚠️ 'royxat2' varag'i topilmadi, faqat 'royxat' dan qidiriladi", True)
            
            if not royxat_index and not royxat2_index:
                self.log_message("❌ Xato: Hech qanday indeks yaratilmadi!", True)
                self.finish_processing()
                return
            
            # `2-жадвал` dagi ma'lumotlarni qayta ishlash
            updated_count = 0
            not_found_count = 0
            already_had_pinfl = 0
            total_rows = 0
            
            self.log_message("🔄 '2-жадвал' dagi ma'lumotlar qayta ishlanmoqda...")
            
            for row in sheet_main.iter_rows(min_row=8, values_only=False):
                try:
                    # F ustunidagi ma'lumotni olish (index 5)
                    f_value = row[5].value if len(row) > 5 and row[5] else None
                    
                    if not f_value or not str(f_value).strip():
                        continue
                    
                    full_name = str(f_value).strip()
                    total_rows += 1
                    
                    # G ustunidagi mavjud PNFL ni tekshirish (index 6)
                    existing_pinfl = row[6].value if len(row) > 6 and row[6] else None
                    
                    if existing_pinfl and str(existing_pinfl).strip():
                        already_had_pinfl += 1
                        continue
                    
                    # PinFL ni topish
                    pinfl = self.find_pinfl(full_name, royxat_index, royxat2_index)
                    
                    if pinfl:
                        # G ustuniga yozish (index 6)
                        if len(row) <= 6:
                            # G ustuni mavjud emas, kengaytirish
                            pass
                        row[6].value = pinfl
                        updated_count += 1
                    else:
                        not_found_count += 1
                        
                        # Debug uchun topilmagan ismlarni ko'rsatish
                        if total_rows <= 10:
                            self.log_message(f"🔍 Topilmadi: '{full_name}'")
                        
                except Exception as e:
                    self.log_message(f"⚠️ Qatorni qayta ishlashda xatolik: {str(e)}", True)
                    continue
            
            self.log_message("💾 Natija saqlanmoqda...")
            
            # Natijalarni saqlash
            wb.save(output_path)
            
            # Yakuniy hisobot
            self.log_message("="*50)
            self.log_message(f"✅ Jarayon muvaffaqiyatli yakunlandi!")
            self.log_message(f"📊 Umumiy qatorlar: {total_rows}")
            self.log_message(f"✅ Topilgan va yangilangan: {updated_count}")
            self.log_message(f"⚠️ Oldin PNFL bo'lgan: {already_had_pinfl}")
            self.log_message(f"❌ Topilmagan: {not_found_count}")
            self.log_message(f"📁 Natija saqlandi: {output_path}")
            self.log_message("="*50)
            
            messagebox.showinfo(
                "Muvaffaqiyat", 
                f"Jarayon yakunlandi!\n\n"
                f"Umumiy qatorlar: {total_rows}\n"
                f"Topilgan va yangilangan: {updated_count}\n"
                f"Oldin PNFL bo'lgan: {already_had_pinfl}\n"
                f"Topilmagan: {not_found_count}\n\n"
                f"Natija: {output_path}"
            )
            
        except Exception as e:
            self.log_message(f"❌ Xatolik yuz berdi: {str(e)}", True)
            messagebox.showerror("Xatolik", f"Jarayonda xatolik yuz berdi:\n{str(e)}")
        
        finally:
            self.finish_processing()
    
    def finish_processing(self):
        self.is_processing = False
        self.process_btn.config(state=tk.NORMAL)
        self.progress_bar.stop()
        self.status_label.config(text="Tayyor", foreground="blue")
    
    def show_result(self):
        """Natijani alohida oynada ko'rsatish"""
        if not self.result_text.get(1.0, tk.END).strip():
            messagebox.showinfo("Ma'lumot", "Hali natija mavjud emas. Avval 'Qayta ishlashni boshlash' tugmasini bosing.")
            return
        
        result_window = tk.Toplevel(self.root)
        result_window.title("Natija")
        result_window.geometry("800x500")
        
        text_widget = scrolledtext.ScrolledText(
            result_window, 
            wrap=tk.WORD, 
            font=("Consolas", 10)
        )
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(1.0, self.result_text.get(1.0, tk.END))
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(result_window, text="Yopish", command=result_window.destroy).pack(pady=5)

def main():
    root = tk.Tk()
    app = ExcelMatcherApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()