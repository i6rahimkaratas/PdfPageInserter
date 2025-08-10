import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pypdf
import os



def duplicate_and_insert_pages(input_path, output_path, pages_to_duplicate_str, insert_at_page, before_or_after, num_copies):
    
    try:
        reader = pypdf.PdfReader(input_path)
        writer = pypdf.PdfWriter()

        total_pages = len(reader.pages)

        
        page_indices_to_duplicate = set()
        parts = pages_to_duplicate_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                if start < 1 or end > total_pages or start > end:
                    raise ValueError(f"Geçersiz sayfa aralığı: {part}")
                for i in range(start, end + 1):
                    page_indices_to_duplicate.add(i - 1)
            else:
                page_num = int(part)
                if page_num < 1 or page_num > total_pages:
                    raise ValueError(f"Geçersiz sayfa numarası: {page_num}")
                page_indices_to_duplicate.add(page_num - 1)

        
        duplicated_pages = [reader.pages[i] for i in sorted(list(page_indices_to_duplicate))]

        if not duplicated_pages:
            raise ValueError("Çoğaltılacak sayfa seçilmedi.")
            
        insert_at_index = insert_at_page - 1
        if insert_at_index < 0 or insert_at_index >= total_pages:
             raise ValueError(f"Ekleme konumu geçersiz. 1 ile {total_pages} arasında olmalı.")

        
        for i in range(total_pages):
            
            if before_or_after == "öncesine" and i == insert_at_index:
                
                for _ in range(num_copies):
                    for page in duplicated_pages:
                        writer.add_page(page)
            
            
            writer.add_page(reader.pages[i])

            
            if before_or_after == "sonrasına" and i == insert_at_index:
                
                for _ in range(num_copies):
                    for page in duplicated_pages:
                        writer.add_page(page)

        
        with open(output_path, "wb") as f_out:
            writer.write(f_out)

        return True, f"İşlem tamamlandı! Yeni dosya '{os.path.basename(output_path)}' adıyla kaydedildi."

    except Exception as e:
        return False, f"Hata oluştu: {e}"



class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Gelişmiş PDF Sayfa Çoğaltıcı")
        self.geometry("550x350") 
        
        self.input_file_path = tk.StringVar()
        self.insert_option = tk.StringVar(value="sonrasına")

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        
        ttk.Label(main_frame, text="1. PDF Dosyasını Seçin:").grid(row=0, column=0, sticky="w", pady=2)
        self.file_entry = ttk.Entry(main_frame, textvariable=self.input_file_path, width=50, state="readonly")
        self.file_entry.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        self.browse_button = ttk.Button(main_frame, text="Gözat...", command=self.select_file)
        self.browse_button.grid(row=1, column=1, sticky="w")

        
        ttk.Label(main_frame, text="2. Çoğaltılacak Sayfalar: (Örn: 5 veya 3-7 veya 1,4,8-10)").grid(row=2, column=0, sticky="w", pady=(10, 2))
        self.pages_to_copy_entry = ttk.Entry(main_frame, width=50)
        self.pages_to_copy_entry.grid(row=3, column=0, columnspan=2, sticky="ew")

        
        ttk.Label(main_frame, text="3. Eklenecek Sayfa Konumu:").grid(row=4, column=0, sticky="w", pady=(10, 2))
        insert_frame = ttk.Frame(main_frame)
        insert_frame.grid(row=5, column=0, columnspan=2, sticky="w")
        self.insert_at_page_entry = ttk.Entry(insert_frame, width=10)
        self.insert_at_page_entry.pack(side="left", padx=(0, 5))
        ttk.Label(insert_frame, text="numaralı sayfanın").pack(side="left", padx=(0, 5))
        self.before_radio = ttk.Radiobutton(insert_frame, text="öncesine", variable=self.insert_option, value="öncesine")
        self.before_radio.pack(side="left")
        self.after_radio = ttk.Radiobutton(insert_frame, text="sonrasına", variable=self.insert_option, value="sonrasına")
        self.after_radio.pack(side="left", padx=(5, 0))

        
        ttk.Label(main_frame, text="4. Kaç kopya eklensin:").grid(row=6, column=0, sticky="w", pady=(10, 2))
        self.num_copies_entry = ttk.Entry(main_frame, width=10)
        self.num_copies_entry.grid(row=7, column=0, sticky="w")
        self.num_copies_entry.insert(0, "1") 

        
        self.process_button = ttk.Button(main_frame, text="5. PDF'i İşle ve Kaydet", command=self.process_pdf)
        self.process_button.grid(row=8, column=0, columnspan=2, pady=(20, 0))

        main_frame.columnconfigure(0, weight=1)

    def select_file(self):
        filepath = filedialog.askopenfilename(title="PDF Dosyası Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if filepath:
            self.input_file_path.set(filepath)

    def process_pdf(self):
        
        input_pdf = self.input_file_path.get()
        pages_str = self.pages_to_copy_entry.get()
        insert_at_str = self.insert_at_page_entry.get()
        num_copies_str = self.num_copies_entry.get() 
        insert_mode = self.insert_option.get()

        
        if not all([input_pdf, pages_str, insert_at_str, num_copies_str]):
            messagebox.showerror("Hata", "Lütfen tüm alanları doldurun.")
            return
        
        try:
            insert_at_page = int(insert_at_str)
            num_copies = int(num_copies_str) 
            if num_copies < 1:
                raise ValueError("Kopya sayısı 1'den küçük olamaz.")
        except ValueError as e:
            messagebox.showerror("Hata", f"Lütfen geçerli sayılar girin.\nDetay: {e}")
            return

        
        dir_name = os.path.dirname(input_pdf)
        base_name = os.path.basename(input_pdf)
        file_name, file_ext = os.path.splitext(base_name)
        output_pdf = os.path.join(dir_name, f"{file_name}_duzenlenmis{file_ext}")

        
        success, message = duplicate_and_insert_pages(
            input_pdf,
            output_pdf,
            pages_str,
            insert_at_page,
            insert_mode,
            num_copies 
        )

        if success:
            messagebox.showinfo("Başarılı", message)
        else:
            messagebox.showerror("Hata", message)


if __name__ == "__main__":
    app = App()
    app.mainloop()