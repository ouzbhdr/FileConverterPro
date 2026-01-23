import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import os
import threading
from docx2pdf import convert as docx_to_pdf_convert
from pdf2docx import Converter

# Theme Settings
ctk.set_appearance_mode("Dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

class ProConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Settings
        self.title("FileConverterPro v2.0")
        self.geometry("600x500")
        self.resizable(False, False)

        # Variables
        self.selected_file = ""
        
        # --- UI LAYOUT ---
        
        # 1. Title
        self.lbl_title = ctk.CTkLabel(self, text="File Converter Pro", font=("Roboto", 24, "bold"))
        self.lbl_title.pack(pady=20)

        # 2. Tabs
        self.tabview = ctk.CTkTabview(self, width=500, height=300)
        self.tabview.pack(pady=10)

        # Add Tabs
        self.tab_image = self.tabview.add("Images")
        self.tab_doc = self.tabview.add("Documents")

        # --- IMAGE TAB SETUP ---
        self.setup_image_tab()

        # --- DOCUMENT TAB SETUP ---
        self.setup_doc_tab()

        # Status Bar
        self.lbl_status = ctk.CTkLabel(self, text="Ready", text_color="gray")
        self.lbl_status.pack(side="bottom", pady=10)

    def setup_image_tab(self):
        # Select Button
        btn_select = ctk.CTkButton(self.tab_image, text="Select Image File", command=lambda: self.select_file("image"))
        btn_select.pack(pady=20)

        self.lbl_image_file = ctk.CTkLabel(self.tab_image, text="No file selected...", text_color="gray")
        self.lbl_image_file.pack(pady=5)

        # Format Selection
        self.image_formats = [".png", ".jpg", ".pdf", ".webp", ".ico", ".bmp", ".tiff"]
        self.combo_image = ctk.CTkComboBox(self.tab_image, values=self.image_formats, state="readonly")
        self.combo_image.set(".png")
        self.combo_image.pack(pady=20)

        # Convert Button
        btn_convert = ctk.CTkButton(self.tab_image, text="Start Conversion", fg_color="green", hover_color="darkgreen", command=lambda: self.start_thread("image"))
        btn_convert.pack(pady=10)

    def setup_doc_tab(self):
        # Select Button
        btn_select = ctk.CTkButton(self.tab_doc, text="Select Word or PDF", command=lambda: self.select_file("doc"))
        btn_select.pack(pady=20)

        self.lbl_doc_file = ctk.CTkLabel(self.tab_doc, text="No file selected...", text_color="gray")
        self.lbl_doc_file.pack(pady=5)

        # Format Selection
        self.doc_formats = ["Convert to PDF", "Convert to Word (.docx)"]
        self.combo_doc = ctk.CTkComboBox(self.tab_doc, values=self.doc_formats, state="readonly")
        self.combo_doc.set("Convert to PDF")
        self.combo_doc.pack(pady=20)

        # Convert Button
        btn_convert = ctk.CTkButton(self.tab_doc, text="Start Conversion", fg_color="blue", hover_color="darkblue", command=lambda: self.start_thread("doc"))
        btn_convert.pack(pady=10)

    def select_file(self, file_type):
        if file_type == "image":
            filetypes = [("Image Files", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff")]
            label = self.lbl_image_file
        else:
            filetypes = [("Documents", "*.docx *.pdf")]
            label = self.lbl_doc_file

        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.selected_file = filename
            display_name = os.path.basename(filename)
            label.configure(text=display_name, text_color="white")
            
            # Auto-detect format suggestion for documents
            if file_type == "doc":
                if filename.endswith(".docx"):
                    self.combo_doc.set("Convert to PDF")
                elif filename.endswith(".pdf"):
                    self.combo_doc.set("Convert to Word (.docx)")

    def start_thread(self, file_type):
        # Run conversion in a separate thread to prevent UI freezing
        if not self.selected_file:
            messagebox.showwarning("Warning", "Please select a file first!")
            return
        
        self.lbl_status.configure(text="Converting... Please wait.", text_color="orange")
        threading.Thread(target=self.convert_process, args=(file_type,)).start()

    def convert_process(self, file_type):
        try:
            folder = os.path.dirname(self.selected_file)
            name = os.path.splitext(os.path.basename(self.selected_file))[0]
            
            if file_type == "image":
                target_ext = self.combo_image.get()
                new_path = os.path.join(folder, name + target_ext)
                
                img = Image.open(self.selected_file)
                # JPG and PDF do not support transparency (Alpha channel)
                if target_ext in [".jpg", ".jpeg", ".bmp", ".pdf"]:
                    img = img.convert("RGB")
                img.save(new_path)
            
            elif file_type == "doc":
                operation = self.combo_doc.get()
                
                if operation == "Convert to PDF":
                    if not self.selected_file.endswith(".docx"):
                        raise ValueError("You must select a .docx file to convert to PDF.")
                    new_path = os.path.join(folder, name + ".pdf")
                    docx_to_pdf_convert(self.selected_file, new_path)
                    
                elif operation == "Convert to Word (.docx)":
                    if not self.selected_file.endswith(".pdf"):
                        raise ValueError("You must select a .pdf file to convert to Word.")
                    new_path = os.path.join(folder, name + ".docx")
                    cv = Converter(self.selected_file)
                    cv.convert(new_path, start=0, end=None)
                    cv.close()

            self.lbl_status.configure(text=f"Success! Saved: {os.path.basename(new_path)}", text_color="green")
            messagebox.showinfo("Success", "Conversion completed successfully!")

        except Exception as e:
            self.lbl_status.configure(text="Error Occurred!", text_color="red")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    app = ProConverterApp()
    app.mainloop()
