import customtkinter as ctk
import nltk
import pyperclip
import re
from deep_translator import GoogleTranslator
import pystray
from pystray import MenuItem as item
from PIL import Image
import threading
import sys
import os
from win32com.client import Dispatch

# --- SUMMARIZER ---
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer


# --- СИСТЕМНЫЕ ФУНКЦИИ ---
def resource_path(relative_path):
    """Позволяет EXE видеть иконку внутри себя"""
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def create_shortcut():
    """Создает ярлык в меню Пуск для поиска Windows"""
    try:
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
            start_menu = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
            shortcut_path = os.path.join(start_menu, "Smart AI Assistant.lnk")
            if not os.path.exists(shortcut_path):
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                shortcut.save()
    except:
        pass


try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')


class SmartApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.placeholders = {
            "russian": "📝 Вставьте текст любой длины...",
            "english": "📝 Paste text of any length..."
        }

        self.translations = {
            "russian": {
                "paste_btn": "ВСТАВИТЬ",
                "clear_btn": "ОЧИСТИТЬ",
                "run_btn": "СОКРАТИТЬ ТЕКСТ",
                "copy_all": "СКОПИРОВАТЬ ВСЁ",
                "copied": "✅ ГОТОВО!",
                "loading": "⚙️ Анализирую текст..."
            },
            "english": {
                "paste_btn": "PASTE",
                "clear_btn": "CLEAR",
                "run_btn": "SUMMARIZE TEXT",
                "copy_all": "COPY ALL",
                "copied": "✅ DONE!",
                "loading": "⚙️ Analyzing text..."
            }
        }

        self.title("Smart AI Assistant")
        self.geometry("750x850")
        ctk.set_appearance_mode("dark")

        # При закрытии — скрываем в трей
        self.protocol('WM_DELETE_WINDOW', self.withdraw)

        # --- UI ---
        self.top_bar = ctk.CTkFrame(self, fg_color="transparent")
        self.top_bar.pack(fill="x", padx=20, pady=(10, 0))

        self.lang_switch = ctk.CTkSegmentedButton(self.top_bar, values=["RU", "EN"], command=self.change_lang,
                                                  width=100)
        self.lang_switch.set("RU")
        self.lang_switch.pack(side="left")

        self.main_label = ctk.CTkLabel(self, text="Smart AI Assistant", font=("Arial", 22, "bold"))
        self.main_label.pack(pady=10)

        self.textbox = ctk.CTkTextbox(self, width=650, height=250)
        self.textbox.pack(pady=10)
        self.textbox.bind("<FocusIn>", self.clear_placeholder)

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=10)

        self.paste_btn = ctk.CTkButton(self.btn_frame, text="", command=self.do_paste, width=140)
        self.paste_btn.grid(row=0, column=0, padx=5)

        # КРАСНАЯ КНОПКА ОЧИСТКИ
        self.clear_btn = ctk.CTkButton(self.btn_frame, text="", command=self.do_clear_all, width=140,
                                       fg_color="#A13333", hover_color="#C0392B")
        self.clear_btn.grid(row=0, column=1, padx=5)

        self.run_btn = ctk.CTkButton(self.btn_frame, text="", command=self.do_run, width=220, fg_color="#1f6aa5")
        self.run_btn.grid(row=0, column=2, padx=5)

        self.result_text = ctk.CTkTextbox(self, width=650, height=300, fg_color="#1a1a1a")
        self.result_text.pack(pady=10)

        self.copy_btn = ctk.CTkButton(self, text="", command=self.do_copy, width=300, fg_color="#2da94f")
        self.copy_btn.pack(pady=10)

        self.current_lang = "russian"
        self.do_clear_all()
        self.update_ui_texts()

        create_shortcut()
        self.create_tray_icon()

    def create_tray_icon(self):
        try:
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
            else:
                img = Image.new('RGB', (64, 64), color='#1f6aa5')

            menu = (item('Открыть', self.deiconify), item('Выход', self.quit_app))
            self.tray_icon = pystray.Icon("SmartAI", img, "Smart AI Assistant", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except:
            pass

    def quit_app(self):
        try:
            self.tray_icon.stop()
        except:
            pass
        self.destroy()
        sys.exit()

    def update_ui_texts(self):
        t = self.translations[self.current_lang]
        self.paste_btn.configure(text=t["paste_btn"])
        self.clear_btn.configure(text=t["clear_btn"])
        self.run_btn.configure(text=t["run_btn"])
        self.copy_btn.configure(text=t["copy_all"])

    def change_lang(self, val):
        self.current_lang = "russian" if val == "RU" else "english"
        self.update_ui_texts()
        self.do_clear_all()

    def do_clear_all(self):
        self.textbox.delete("1.0", "end")
        self.result_text.delete("1.0", "end")
        self.textbox.insert("1.0", self.placeholders[self.current_lang])

    def clear_placeholder(self, event):
        if self.textbox.get("1.0", "end-1c").strip() in self.placeholders.values():
            self.textbox.delete("1.0", "end")

    def do_paste(self):
        txt = pyperclip.paste()
        if txt:
            if self.textbox.get("1.0", "end-1c").strip() in self.placeholders.values():
                self.textbox.delete("1.0", "end")
            self.textbox.insert("insert", txt)

    def do_copy(self):
        res = self.result_text.get("1.0", "end-1c").strip()
        if res:
            pyperclip.copy(res)
            old = self.copy_btn.cget("text")
            self.copy_btn.configure(text=self.translations[self.current_lang]["copied"])
            self.after(1000, lambda: self.copy_btn.configure(text=old))

    def do_run(self):
        raw = self.textbox.get("1.0", "end-1c").strip()
        if not raw or raw in self.placeholders.values():
            return

        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", self.translations[self.current_lang]["loading"])

        def run_ai():
            try:
                clean = re.sub(r'http\S+', '', raw)
                clean = re.sub(r'\s+', ' ', clean)

                is_ru = bool(re.search(r'[а-яё]', clean.lower()))
                src_lang = "russian" if is_ru else "english"

                # Разбиваем на чанки по 5000 символов
                chunk_size = 5000
                chunks = [clean[i:i + chunk_size] for i in range(0, len(clean), chunk_size)]

                summarizer = LexRankSummarizer()
                final_sentences = []

                for chunk in chunks:
                    parser = PlaintextParser.from_string(chunk, Tokenizer(src_lang))
                    # Берем 35% предложений из чанка (больше текста на выходе)
                    sentences_count = max(3, int(len(parser.document.sentences) * 0.35))
                    summary = summarizer(parser.document, sentences_count)

                    for s in summary:
                        final_sentences.append(str(s).strip())

                final_sentences = list(dict.fromkeys(final_sentences))

                target = 'ru' if self.current_lang == 'russian' else 'en'
                trans = GoogleTranslator(source='auto', target=target)

                res_txt = ""
                for line in final_sentences:
                    # Перевод только если нужно
                    if (src_lang == "russian" and self.current_lang == "english") or \
                            (src_lang == "english" and self.current_lang == "russian"):
                        try:
                            line = trans.translate(line)
                        except:
                            pass
                    res_txt += f"• {line}\n\n"

                self.after(0, lambda: self.update_result(res_txt.strip()))

            except Exception as e:
                self.after(0, lambda: self.update_result(f"Error: {e}"))

        threading.Thread(target=run_ai, daemon=True).start()

    def update_result(self, text):
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", text)


if __name__ == "__main__":
    app = SmartApp()
    app.mainloop()