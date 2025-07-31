import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import speech_recognition as sr
from docx import Document
from docx.shared import Pt
import pyperclip

class DictationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hindi Dictation Recorder")
        self.is_recording = False
        self.transcribed_text = ""

        # Text preview area
        self.text_area = scrolledtext.ScrolledText(
            root, wrap=tk.WORD, font=("Arial", 15), width=60, height=15)
        self.text_area.pack(pady=10)

        # Control buttons
        btn_frame = tk.Frame(root)
        btn_frame.pack()

        self.start_btn = tk.Button(btn_frame, text="Start", command=self.start_recording)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_recording, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        self.copy_btn = tk.Button(btn_frame, text="Copy", command=self.copy_text)
        self.copy_btn.pack(side=tk.LEFT, padx=5)
        self.save_btn = tk.Button(btn_frame, text="Save as DOCX", command=self.save_docx)
        self.save_btn.pack(side=tk.LEFT, padx=5)

        self.recognizer = sr.Recognizer()
        self.mic = sr.Microphone()

    def start_recording(self):
        self.is_recording = True
        self.text_area.delete("1.0", tk.END)
        self.transcribed_text = ""
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        threading.Thread(target=self.record).start()

    def stop_recording(self):
        self.is_recording = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def record(self):
        with self.mic as source:
            self.recognizer.adjust_for_ambient_noise(source)
            while self.is_recording:
                try:
                    audio = self.recognizer.listen(source, timeout=5)
                    text = self.recognizer.recognize_google(audio, language="hi-IN")
                    self.transcribed_text += text + " "
                    self.text_area.insert(tk.END, text + " ")
                    self.text_area.see(tk.END)
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    continue
                except sr.RequestError as e:
                    messagebox.showerror("Error", f"Speech Recognition error:\n{e}")
                    break

    def copy_text(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if text:
            pyperclip.copy(text)
            messagebox.showinfo("Copied", "Text copied to clipboard!")

    def save_docx(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Empty", "No text to save!")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if not file_path:
            return
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.font.name = "Mangal"
        run.font.size = Pt(16)
        # DOCX trick for font: set eastAsia as well
        import docx.oxml.ns
        rFonts = run._element.rPr.rFonts
        rFonts.set(docx.oxml.ns.qn('w:eastAsia'), 'Mangal')
        doc.save(file_path)
        messagebox.showinfo("Saved", f"Saved at:\n{file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DictationApp(root)
    root.mainloop()
