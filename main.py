import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import ImageTk, Image
import os

POPPLER_PATH = r"C:\poppler-24.08.0\Library\bin"

class PDFDragPreviewApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 미리보기 + 드래그 정렬 병합기")
        self.pdf_files = []
        self.page_inputs = []
        self.preview_widgets = []
        self.preview_data = []

        self.select_btn = tk.Button(root, text="PDF 파일 선택", command=self.select_files)
        self.select_btn.pack(pady=5)

        self.input_frame = tk.Frame(root)
        self.input_frame.pack()

        self.preview_btn = tk.Button(root, text="미리보기 생성", command=self.show_preview)
        self.preview_btn.pack(pady=5)

        # 캔버스 + 스크롤바 구성
        self.canvas_frame = tk.Frame(root)
        self.canvas_frame.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(self.canvas_frame)
        self.scrollbar = Scrollbar(self.canvas_frame, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollable_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')

        # 마우스 휠 스크롤
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self.merge_btn = tk.Button(root, text="PDF 병합 실행", command=self.merge_pdf)
        self.merge_btn.pack(pady=5)

        self.drag_start_index = None

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not files:
            return
        self.pdf_files = list(files)
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.page_inputs = []
        for i, file in enumerate(self.pdf_files):
            name = os.path.basename(file)
            tk.Label(self.input_frame, text=f"{name} (예: 1,3-5):").grid(row=i, column=0, padx=5, pady=2)
            entry = tk.Entry(self.input_frame, width=30)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.page_inputs.append(entry)

    def parse_pages(self, text):
        result = []
        for part in text.split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    result.extend(range(start - 1, end))
                except:
                    continue
            else:
                try:
                    result.append(int(part) - 1)
                except:
                    continue
        return result

    def show_preview(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.preview_widgets.clear()
        self.preview_data.clear()

        self.root.update_idletasks()
        self.canvas.update_idletasks()
        frame_width = self.canvas.winfo_width()
        target_width = int(frame_width * 0.9)  # 여백 고려

        for file, entry in zip(self.pdf_files, self.page_inputs):
            pages = self.parse_pages(entry.get())
            if not pages:
                continue
            try:
                images = convert_from_path(
                    file,
                    dpi=100,
                    first_page=min(pages) + 1,
                    last_page=max(pages) + 1,
                    poppler_path=POPPLER_PATH
                )
                for page_num in pages:
                    index = page_num - min(pages)
                    if 0 <= index < len(images):
                        img = images[index]
                        orig_w, orig_h = img.size
                        ratio = target_width / orig_w
                        resized = img.resize((int(orig_w * ratio), int(orig_h * ratio)))
                        tk_img = ImageTk.PhotoImage(resized)

                        frame = tk.Frame(self.scrollable_frame, bd=2, relief='ridge')
                        label = tk.Label(frame, image=tk_img)
                        label.image = tk_img
                        label.pack()
                        info = tk.Label(frame, text=f"{os.path.basename(file)} - {page_num + 1}")
                        info.pack()

                        frame.pack(padx=10, pady=10)
                        frame.bind("<Button-1>", self.start_drag)
                        frame.bind("<B1-Motion>", self.perform_drag)
                        frame.bind("<ButtonRelease-1>", self.end_drag)

                        self.preview_widgets.append(frame)
                        self.preview_data.append((file, page_num))
            except Exception as e:
                messagebox.showerror("미리보기 오류", str(e))

    def start_drag(self, event):
        widget = event.widget
        while not isinstance(widget, tk.Frame) and widget.master:
            widget = widget.master
        self.drag_start_index = self.preview_widgets.index(widget)

    def perform_drag(self, event):
        widget = event.widget
        while not isinstance(widget, tk.Frame) and widget.master:
            widget = widget.master

        y = widget.winfo_y() + event.y
        for i, w in enumerate(self.preview_widgets):
            if w == widget:
                continue
            if abs(w.winfo_y() - y) < 50:
                i1 = self.drag_start_index
                i2 = i
                if i1 != i2:
                    self.preview_widgets[i1], self.preview_widgets[i2] = self.preview_widgets[i2], self.preview_widgets[i1]
                    self.preview_data[i1], self.preview_data[i2] = self.preview_data[i2], self.preview_data[i1]
                    self.refresh_preview_order()
                    self.drag_start_index = i2
                break

    def end_drag(self, event):
        self.drag_start_index = None

    def refresh_preview_order(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.pack_forget()
        for widget in self.preview_widgets:
            widget.pack(padx=10, pady=10)

    def merge_pdf(self):
        writer = PdfWriter()
        try:
            for file, page_num in self.preview_data:
                reader = PdfReader(file)
                if 0 <= page_num < len(reader.pages):
                    writer.add_page(reader.pages[page_num])
            out_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if out_path:
                with open(out_path, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("완료", f"병합 완료: {out_path}")
        except Exception as e:
            messagebox.showerror("병합 오류", str(e))

root = tk.Tk()
root.geometry("900x800")
app = PDFDragPreviewApp(root)
root.mainloop()
