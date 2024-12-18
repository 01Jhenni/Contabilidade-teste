import tkinter as tk
from tkinter import filedialog, simpledialog, Canvas, Toplevel, messagebox
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import json
import os


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Data Extractor")
        
        # Botão para carregar PDFs
        self.upload_button = tk.Button(root, text="Upload PDFs", command=self.upload_pdfs)
        self.upload_button.pack(pady=10)
        
        # Botão para selecionar campos
        self.select_fields_button = tk.Button(root, text="Selecionar Campos", command=self.select_fields)
        self.select_fields_button.pack(pady=10)

        # Botão para carregar um modelo salvo
        self.load_model_button = tk.Button(root, text="Carregar Modelo", command=self.load_model)
        self.load_model_button.pack(pady=10)

        # Botão para salvar o modelo
        self.save_model_button = tk.Button(root, text="Salvar Modelo", command=self.save_model)
        self.save_model_button.pack(pady=10)

        # Botão para processar os PDFs
        self.process_button = tk.Button(root, text="Processar PDFs", command=self.process_pdfs)
        self.process_button.pack(pady=10)

        # Área para mostrar os PDFs carregados
        self.pdf_listbox = tk.Listbox(root, height=10, width=50)
        self.pdf_listbox.pack(pady=10)

        self.pdf_files = []
        self.current_pdf_path = None
        self.pdf_doc = None
        self.selected_areas = {}  # Armazena as caixas desenhadas (agora por página)
        self.extraction_model = {}  # Modelo que mapeia áreas para campos
        self.current_page = 0  # Página inicial

    def upload_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        self.pdf_files.extend(files)
        for file in files:
            self.pdf_listbox.insert(tk.END, file)

    def select_fields(self):
        if len(self.pdf_files) > 0:
            self.current_pdf_path = self.pdf_files[0]
            self.open_pdf_page(self.current_page)

    def open_pdf_page(self, page_num):
        self.pdf_doc = fitz.open(self.current_pdf_path)
        page = self.pdf_doc.load_page(page_num)
        pix = page.get_pixmap()

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.img_tk = ImageTk.PhotoImage(img)

        self.window = Toplevel(self.root)
        self.window.title(f"Selecione os Campos")

        self.canvas = Canvas(self.window, width=pix.width, height=pix.height)
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img_tk)

        self.prev_page_button = tk.Button(self.window, text="Página Anterior", command=self.prev_page)
        self.prev_page_button.pack(side=tk.LEFT, padx=10)

        self.next_page_button = tk.Button(self.window, text="Próxima Página", command=self.next_page)
        self.next_page_button.pack(side=tk.LEFT, padx=10)

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

        self.start_x = None
        self.start_y = None
        self.rect = None

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.window.destroy()
            self.open_pdf_page(self.current_page)

    def next_page(self):
        if self.current_page < len(self.pdf_doc) - 1:
            self.current_page += 1
            self.window.destroy()
            self.open_pdf_page(self.current_page)

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)

        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_mouse_drag(self, event):
        cur_x, cur_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        field_name = simpledialog.askstring("Campo", "Digite o nome do campo (ex: CNPJ, Valor, etc.)")
        
        if field_name:
            if self.current_page not in self.extraction_model:
                self.extraction_model[self.current_page] = {}

            self.extraction_model[self.current_page][field_name] = (self.start_x, self.start_y, end_x, end_y)
            print(f"Campo '{field_name}'")

    def save_model(self):
        if not self.extraction_model:
            messagebox.showwarning("Atenção", "Nenhum campo foi selecionado para salvar.")
            return
       
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w') as json_file:
                json.dump(self.extraction_model, json_file)
            messagebox.showinfo("Sucesso", "Modelo salvo com sucesso!")

    def load_model(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path and os.path.exists(file_path):
            with open(file_path, 'r') as json_file:
                self.extraction_model = json.load(json_file)
            messagebox.showinfo("Sucesso", "Modelo carregado com sucesso!")
        else:
            messagebox.showerror("Erro", "Falha ao carregar o modelo.")

    def process_pdfs(self):
        if not self.extraction_model:
            messagebox.showwarning("Atenção", "Nenhum modelo foi carregado ou criado.")
            return
       
        extracted_data = []
        
        for pdf_file in self.pdf_files:
            pdf_data = {}
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, fields in self.extraction_model.items():
                    try:
                        page_num = int(page_num)
                        if page_num >= len(pdf.pages):
                            print(f"Página {page_num + 1} não encontrada no PDF '{pdf_file}'. Pulando.")
                            continue
                        
                        page = pdf.pages[page_num]
                        for field, coords in fields.items():
                            # Extrair a área do PDF com uma margem para garantir captura completa
                            adjusted_coords = (
                                coords[0] - 2, coords[1] - 2,
                                coords[2] + 2, coords[3] + 2
                            )
                            cropped = page.within_bbox(adjusted_coords).extract_text() or " "
                            pdf_data[f"{field} "] = cropped
                            print(f"Extraindo campo '{field}' na página {page_num + 1} com coords {adjusted_coords}: {cropped}")
                    except Exception as e:
                        print(f"Erro ao processar a página {page_num + 1} do PDF '{pdf_file}': {str(e)}")
            extracted_data.append(pdf_data)
        
        # Exportar para Excel
        df = pd.DataFrame(extracted_data)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Sucesso", "Dados extraídos e salvos em Excel com sucesso!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

