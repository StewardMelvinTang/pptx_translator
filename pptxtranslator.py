import customtkinter as ctk
from tkinter import filedialog, messagebox
from pptx import Presentation
import os
import json
import time
import threading
import re
from openai import OpenAI

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pptxtranslator_config.json")

MODELS = {
    "gpt-4.1-nano": {"label": "4.1 Nano", "input_cost": 0.10, "output_cost": 0.40},
    "gpt-4.1-mini": {"label": "4.1 Mini", "input_cost": 0.40, "output_cost": 1.60},
    "gpt-4.1": {"label": "4.1", "input_cost": 2.00, "output_cost": 8.00},
    "gpt-4o-mini": {"label": "4o Mini", "input_cost": 0.15, "output_cost": 0.60},
    "gpt-4o": {"label": "4o", "input_cost": 2.50, "output_cost": 10.00},
}

TARGET_BATCH_CHARS = 2000
SEP_TOKEN = "\n[SEP]\n"


def load_config():
    try:
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_config(config):
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f)


def contains_chinese(text):
    return bool(re.search(r'[\u4e00-\u9fff]', text))


def translate_text_batch(texts, model, client):
    joined = SEP_TOKEN.join(texts)
    messages = [
        {
            "role": "system",
            "content": (
                "You are a professional translator. Translate the following text "
                "from Traditional Chinese to English. Each segment is separated by [SEP]. "
                "Return your translations in the same order, separated by [SEP]. "
                "Keep the same format and style. Do not add any additional text."
            ),
        },
        {"role": "user", "content": joined},
    ]

    try:
        response = client.chat.completions.create(
            model=model, messages=messages, temperature=0.1
        )
        translated_text = response.choices[0].message.content
        input_tokens = response.usage.prompt_tokens
        output_tokens = response.usage.completion_tokens

        translations = [t.strip() for t in translated_text.split("[SEP]")]

        if len(translations) < len(texts):
            translations.extend(texts[len(translations):])
        elif len(translations) > len(texts):
            translations = translations[: len(texts)]

        return translations, input_tokens, output_tokens
    except Exception as e:
        print(f"Error during translation: {e}")
        return texts, 0, 0


def build_batches(text_items):
    batches = []
    current_batch = []
    current_chars = 0
    for text in text_items:
        if current_chars + len(text) > TARGET_BATCH_CHARS and current_batch:
            batches.append(current_batch)
            current_batch = []
            current_chars = 0
        current_batch.append(text)
        current_chars += len(text)
    if current_batch:
        batches.append(current_batch)
    return batches


def scan_pptx_paragraphs(file_path):
    prs = Presentation(file_path)
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in para.runs)
                    if contains_chinese(full_text):
                        count += 1
    return count


def process_pptx(file_path, output_path, model, client, progress_callback=None, cancel_event=None, para_offset=0):
    prs = Presentation(file_path)
    text_items = []
    para_map = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in para.runs)
                    if contains_chinese(full_text):
                        text_items.append(full_text)
                        para_map.append(para)

    total = len(text_items)
    translated_results = []
    input_tokens = 0
    output_tokens = 0
    cancelled = False

    batches = build_batches(text_items)
    paras_done = 0

    for batch in batches:
        if cancel_event and cancel_event.is_set():
            cancelled = True
            break

        translations, batch_in, batch_out = translate_text_batch(batch, model, client)
        input_tokens += batch_in
        output_tokens += batch_out
        translated_results.extend(translations)
        paras_done += len(batch)

        if progress_callback:
            progress_callback(para_offset + paras_done, input_tokens, output_tokens)

    if not cancelled:
        for para, translated_text in zip(para_map, translated_results):
            if para.runs:
                para.runs[0].text = translated_text
                for run in para.runs[1:]:
                    run.text = ""
        prs.save(output_path)

    return input_tokens, output_tokens, not cancelled, total


class FileEntry(ctk.CTkFrame):
    def __init__(self, master, filename, on_remove, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(fg_color=("gray88", "gray20"), corner_radius=8)

        self.label = ctk.CTkLabel(
            self, text=filename, anchor="w", font=ctk.CTkFont(size=13)
        )
        self.label.pack(side="left", fill="x", expand=True, padx=(12, 4), pady=6)

        self.remove_btn = ctk.CTkButton(
            self,
            text="\u2715",
            width=28,
            height=28,
            corner_radius=6,
            fg_color="transparent",
            hover_color=("gray75", "gray35"),
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=14),
            command=on_remove,
        )
        self.remove_btn.pack(side="right", padx=(0, 8), pady=6)


class PPTTranslatorApp:
    def __init__(self, master):
        self.master = master
        master.title("PPTX Translator")
        master.geometry("540x620")
        master.minsize(480, 580)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.file_queue = []
        self.file_widgets = []
        self.cancel_event = threading.Event()
        self.translating = False

        self.setup_ui()
        self.load_api_key()

    def setup_ui(self):
        main = ctk.CTkFrame(self.master, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=16)

        # Title
        ctk.CTkLabel(
            main,
            text="PPTX Translator",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            main,
            text="Traditional Chinese \u2192 English",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", pady=(0, 16))

        # API Key
        key_frame = ctk.CTkFrame(main, fg_color="transparent")
        key_frame.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(key_frame, text="API Key", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")

        key_row = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_row.pack(fill="x", pady=(4, 0))

        self.api_key_var = ctk.StringVar()
        self.api_key_entry = ctk.CTkEntry(
            key_row, textvariable=self.api_key_var, show="\u2022", placeholder_text="sk-...", height=36
        )
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.show_key = False
        self.toggle_key_btn = ctk.CTkButton(
            key_row, text="Show", width=56, height=36, command=self.toggle_key_visibility
        )
        self.toggle_key_btn.pack(side="right")

        # Files section
        files_header = ctk.CTkFrame(main, fg_color="transparent")
        files_header.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(files_header, text="Files", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        self.file_count_label = ctk.CTkLabel(
            files_header, text="0 files", font=ctk.CTkFont(size=12), text_color=("gray50", "gray60")
        )
        self.file_count_label.pack(side="right")

        self.files_frame = ctk.CTkScrollableFrame(main, height=140, corner_radius=10)
        self.files_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.empty_label = ctk.CTkLabel(
            self.files_frame,
            text="No files added yet. Click 'Add Files' below.",
            text_color=("gray55", "gray55"),
            font=ctk.CTkFont(size=12),
        )
        self.empty_label.pack(pady=20)

        self.add_files_btn = ctk.CTkButton(
            main, text="+ Add Files", height=34, command=self.add_files
        )
        self.add_files_btn.pack(fill="x", pady=(0, 12))

        # Model selection
        model_frame = ctk.CTkFrame(main, fg_color="transparent")
        model_frame.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(model_frame, text="Model", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")

        model_labels = [MODELS[m]["label"] for m in MODELS]
        self.model_ids = list(MODELS.keys())
        self.model_var = ctk.StringVar(value=MODELS[self.model_ids[1]]["label"])
        self.model_menu = ctk.CTkOptionMenu(
            model_frame, variable=self.model_var, values=model_labels, height=34
        )
        self.model_menu.pack(fill="x", pady=(4, 0))

        # Translate button
        self.translate_btn = ctk.CTkButton(
            main,
            text="Start Translation",
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.start_translation,
            state="disabled",
        )
        self.translate_btn.pack(fill="x", pady=(0, 12))

        # Progress section
        self.progress_frame = ctk.CTkFrame(main, corner_radius=10)
        self.progress_frame.pack(fill="x", pady=(0, 4))
        self.progress_frame.pack_forget()

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, height=14, corner_radius=7)
        self.progress_bar.pack(fill="x", padx=16, pady=(16, 8))
        self.progress_bar.set(0)

        self.progress_percent = ctk.CTkLabel(
            self.progress_frame,
            text="0%",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.progress_percent.pack(pady=(0, 4))

        self.progress_detail = ctk.CTkLabel(
            self.progress_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self.progress_detail.pack(pady=(0, 4))

        self.progress_stats = ctk.CTkLabel(
            self.progress_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray55", "gray55"),
        )
        self.progress_stats.pack(pady=(0, 12))

        self.cancel_btn = ctk.CTkButton(
            self.progress_frame,
            text="Cancel",
            height=32,
            fg_color=("gray70", "gray35"),
            hover_color=("gray60", "gray45"),
            command=self.cancel_translation,
        )
        self.cancel_btn.pack(pady=(0, 16))

        # Status
        self.status_label = ctk.CTkLabel(
            main,
            text="Ready",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self.status_label.pack(pady=(4, 0))

    def toggle_key_visibility(self):
        self.show_key = not self.show_key
        self.api_key_entry.configure(show="" if self.show_key else "\u2022")
        self.toggle_key_btn.configure(text="Hide" if self.show_key else "Show")

    def load_api_key(self):
        config = load_config()
        key = config.get("api_key", "")
        if key:
            self.api_key_var.set(key)

    def save_api_key(self):
        config = load_config()
        config["api_key"] = self.api_key_var.get()
        save_config(config)

    def get_selected_model_id(self):
        label = self.model_var.get()
        for mid, info in MODELS.items():
            if info["label"] == label:
                return mid
        return self.model_ids[1]

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PowerPoint Files", "*.pptx")])
        if files:
            for f in files:
                if f not in self.file_queue:
                    self.file_queue.append(f)
                    self.refresh_file_list()
            self.translate_btn.configure(state="normal")

    def remove_file(self, filepath):
        if filepath in self.file_queue:
            self.file_queue.remove(filepath)
            self.refresh_file_list()
        if not self.file_queue:
            self.translate_btn.configure(state="disabled")

    def refresh_file_list(self):
        for w in self.file_widgets:
            w.destroy()
        self.file_widgets.clear()

        if not self.file_queue:
            self.empty_label.pack(pady=20)
            self.file_count_label.configure(text="0 files")
            return

        self.empty_label.pack_forget()
        self.file_count_label.configure(
            text=f"{len(self.file_queue)} file{'s' if len(self.file_queue) != 1 else ''}"
        )

        for fp in self.file_queue:
            entry = FileEntry(
                self.files_frame,
                os.path.basename(fp),
                on_remove=lambda p=fp: self.remove_file(p),
            )
            entry.pack(fill="x", pady=(0, 4))
            self.file_widgets.append(entry)

    def start_translation(self):
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("API Key Required", "Please enter your OpenAI API key.")
            return
        if not self.file_queue:
            return

        self.save_api_key()

        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
            return

        self.translating = True
        self.cancel_event.clear()
        self.translate_btn.configure(state="disabled")
        self.add_files_btn.configure(state="disabled")
        self.progress_frame.pack(fill="x", pady=(0, 4))
        self.progress_bar.set(0)
        self.progress_percent.configure(text="0%")
        self.progress_detail.configure(text="Scanning files...")
        self.progress_stats.configure(text="")
        self.status_label.configure(text="Translating...")

        model_id = self.get_selected_model_id()
        files = self.file_queue.copy()

        threading.Thread(
            target=self.run_translation,
            args=(files, output_dir, model_id, api_key),
            daemon=True,
        ).start()

    def run_translation(self, files, output_dir, model_id, api_key):
        client = OpenAI(api_key=api_key)

        total_paras = 0
        for f in files:
            try:
                total_paras += scan_pptx_paragraphs(f)
            except Exception:
                pass

        if total_paras == 0:
            self.master.after(0, lambda: self.translation_done(0, 0, True, "No Chinese text found in files."))
            return

        paras_done = 0
        total_input_tokens = 0
        total_output_tokens = 0
        start_time = time.time()
        completed_files = 0

        for f in files:
            if self.cancel_event.is_set():
                break

            filename = os.path.basename(f)
            default_name = os.path.splitext(filename)[0] + "_Translated.pptx"
            output_path = os.path.join(output_dir, default_name)

            self.master.after(
                0, lambda fn=filename: self.progress_detail.configure(text=f"Translating: {fn}")
            )

            def progress_cb(global_done, in_tok, out_tok, _st=start_time, _tp=total_paras):
                nonlocal total_input_tokens, total_output_tokens
                total_input_tokens = in_tok
                total_output_tokens = out_tok
                elapsed = time.time() - _st
                pct = global_done / _tp if _tp > 0 else 0
                eta = (elapsed / global_done * (_tp - global_done)) if global_done > 0 else 0
                model_info = MODELS.get(model_id, {})
                cost = (in_tok / 1_000_000) * model_info.get("input_cost", 0) + \
                       (out_tok / 1_000_000) * model_info.get("output_cost", 0)

                self.master.after(0, lambda p=pct, e=eta, gd=global_done, tp=_tp, c=cost, it=in_tok + out_tok: (
                    self.progress_bar.set(p),
                    self.progress_percent.configure(text=f"{int(p * 100)}%"),
                    self.progress_detail.configure(text=f"Paragraph {gd}/{tp}"),
                    self.progress_stats.configure(
                        text=f"ETA: {int(e)}s  |  Tokens: {it:,}  |  ~${c:.4f}"
                    ),
                ))

            try:
                in_tok, out_tok, success, file_paras = process_pptx(
                    f, output_path, model_id, client,
                    progress_callback=progress_cb,
                    cancel_event=self.cancel_event,
                    para_offset=paras_done,
                )
                total_input_tokens += in_tok
                total_output_tokens += out_tok
                if success:
                    paras_done += file_paras
                    completed_files += 1
            except Exception as e:
                self.master.after(0, lambda err=str(e): self.status_label.configure(text=f"Error: {err}"))

        if self.cancel_event.is_set():
            self.master.after(0, lambda: self.translation_done(
                total_input_tokens, total_output_tokens, False, "Translation cancelled."
            ))
        else:
            model_info = MODELS.get(model_id, {})
            cost = (total_input_tokens / 1_000_000) * model_info.get("input_cost", 0) + \
                   (total_output_tokens / 1_000_000) * model_info.get("output_cost", 0)
            elapsed = time.time() - start_time
            msg = (
                f"Translation complete!\n\n"
                f"Files: {completed_files}/{len(files)}\n"
                f"Paragraphs: {paras_done}\n"
                f"Tokens: {total_input_tokens + total_output_tokens:,}\n"
                f"Cost: ${cost:.4f}\n"
                f"Time: {int(elapsed)}s\n"
                f"Output: {output_dir}"
            )
            self.master.after(0, lambda m=msg: self.translation_done(
                total_input_tokens, total_output_tokens, True, m
            ))

    def translation_done(self, input_tokens, output_tokens, success, message):
        self.translating = False
        self.translate_btn.configure(state="normal" if self.file_queue else "disabled")
        self.add_files_btn.configure(state="normal")

        if success and "complete" in message.lower():
            self.progress_bar.set(1.0)
            self.progress_percent.configure(text="100%")
            messagebox.showinfo("Done", message)
            self.file_queue.clear()
            self.refresh_file_list()
            self.progress_frame.pack_forget()
            self.status_label.configure(text="Complete!")
        elif not success:
            messagebox.showwarning("Cancelled", message)
            self.progress_frame.pack_forget()
            self.status_label.configure(text="Cancelled")
        else:
            messagebox.showinfo("Info", message)
            self.progress_frame.pack_forget()
            self.status_label.configure(text=message.split("\n")[0])

    def cancel_translation(self):
        self.cancel_event.set()
        self.cancel_btn.configure(state="disabled")
        self.progress_detail.configure(text="Cancelling...")


if __name__ == "__main__":
    app = ctk.CTk()
    try:
        app.iconbitmap("pptx_icon.ico")
    except Exception:
        pass
    PPTTranslatorApp(app)
    app.mainloop()
