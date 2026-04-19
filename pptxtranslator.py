import os
import ssl

# Fix conda SSL_CERT_FILE pointing to nonexistent path
if "SSL_CERT_FILE" in os.environ and not os.path.exists(os.environ["SSL_CERT_FILE"]):
    del os.environ["SSL_CERT_FILE"]
if "SSL_CERT_DIR" in os.environ and not os.path.exists(os.environ["SSL_CERT_DIR"]):
    del os.environ["SSL_CERT_DIR"]

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
import json
import time
import threading
import re
import base64
import colorsys
import webbrowser
import fitz  # PyMuPDF
from openai import OpenAI

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pptxtranslator_config.json")

MODELS = {
    "gpt-4.1-nano": {"label": "4.1 Nano", "input_cost": 0.10, "output_cost": 0.40, "vision": False},
    "gpt-4.1-mini": {"label": "4.1 Mini", "input_cost": 0.40, "output_cost": 1.60, "vision": True},
    "gpt-4.1": {"label": "4.1", "input_cost": 2.00, "output_cost": 8.00, "vision": True},
    "gpt-4o-mini": {"label": "4o Mini", "input_cost": 0.15, "output_cost": 0.60, "vision": True},
    "gpt-4o": {"label": "4o", "input_cost": 2.50, "output_cost": 10.00, "vision": True},
    "gpt-5.4": {"label": "5.4", "input_cost": 5.00, "output_cost": 20.00, "vision": True},
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

    # Report initial progress so progress bar shows 0% even for single-batch files
    if progress_callback and total > 0:
        progress_callback(para_offset, input_tokens, output_tokens)

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


def extract_pptx_content(file_path):
    """Extract text and images from a PPTX file, organized by slide."""
    prs = Presentation(file_path)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        slide_data = {"number": i, "texts": [], "images": []}
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    slide_data["texts"].append(text)
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                rows_text = []
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    rows_text.append(" | ".join(cells))
                if rows_text:
                    slide_data["texts"].append("[Table]\n" + "\n".join(rows_text))
            try:
                img = shape.image
                b64 = base64.b64encode(img.blob).decode()
                slide_data["images"].append({
                    "base64": b64,
                    "content_type": img.content_type or "image/png",
                })
            except (AttributeError, Exception):
                pass
        slides.append(slide_data)
    return slides


def scan_pdf_paragraphs(file_path):
    """Count Chinese text spans in a PDF file."""
    doc = fitz.open(file_path)
    count = 0
    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] != 0:  # skip non-text blocks
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    if contains_chinese(span["text"]):
                        count += 1
    doc.close()
    return count


def process_pdf(file_path, output_path, model, client, progress_callback=None, cancel_event=None, para_offset=0):
    """Translate Chinese text in a PDF, replacing in-place."""
    doc = fitz.open(file_path)
    text_items = []
    span_info = []  # (page_idx, rect, font_size, font_name, color)

    for page_idx, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    if contains_chinese(span["text"]):
                        text_items.append(span["text"])
                        span_info.append({
                            "page": page_idx,
                            "bbox": fitz.Rect(span["bbox"]),
                            "size": span["size"],
                            "font": span["font"],
                            "color": span["color"],
                        })

    total = len(text_items)
    translated_results = []
    input_tokens = 0
    output_tokens = 0
    cancelled = False

    batches = build_batches(text_items)
    paras_done = 0

    if progress_callback and total > 0:
        progress_callback(para_offset, input_tokens, output_tokens)

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
        # Group spans by page for efficient redaction
        pages_to_redact = {}
        for info, translated in zip(span_info, translated_results):
            pg = info["page"]
            if pg not in pages_to_redact:
                pages_to_redact[pg] = []
            pages_to_redact[pg].append((info, translated))

        for pg_idx, items in pages_to_redact.items():
            page = doc[pg_idx]
            # Add redaction annotations for all spans on this page
            for info, _ in items:
                page.add_redact_annot(info["bbox"])
            # Apply all redactions at once (removes original text)
            page.apply_redactions()
            # Insert translated text
            for info, translated in items:
                rect = info["bbox"]
                fontsize = info["size"]
                # Adjust font size to fit if needed
                text_width = fitz.get_text_length(translated, fontsize=fontsize)
                rect_width = rect.width
                if text_width > rect_width and rect_width > 0:
                    fontsize = fontsize * rect_width / text_width
                    fontsize = max(fontsize, 5)  # minimum readable size
                # Convert int color to RGB tuple
                c = info["color"]
                rgb = ((c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF)
                color = (rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)
                page.insert_text(
                    fitz.Point(rect.x0, rect.y1 - 2),
                    translated,
                    fontsize=fontsize,
                    color=color,
                )

        doc.save(output_path)

    doc.close()
    return input_tokens, output_tokens, not cancelled, total


def extract_pdf_content(file_path):
    """Extract text and images from a PDF file, organized by page."""
    doc = fitz.open(file_path)
    pages = []
    for i, page in enumerate(doc, 1):
        page_data = {"number": i, "texts": [], "images": []}
        text = page.get_text().strip()
        if text:
            page_data["texts"].append(text)
        # Extract images
        for img_info in page.get_images(full=True):
            try:
                xref = img_info[0]
                img_data = doc.extract_image(xref)
                if img_data:
                    b64 = base64.b64encode(img_data["image"]).decode()
                    ext = img_data.get("ext", "png")
                    content_type = f"image/{ext}" if ext != "jpg" else "image/jpeg"
                    page_data["images"].append({
                        "base64": b64,
                        "content_type": content_type,
                    })
            except Exception:
                pass
        pages.append(page_data)
    doc.close()
    return pages


def extract_file_content(file_path):
    """Dispatch to PPTX or PDF content extractor based on file extension."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extract_pdf_content(file_path)
    else:
        return extract_pptx_content(file_path)


class FileEntry(ctk.CTkFrame):
    def __init__(self, master, filepath, on_remove, **kwargs):
        super().__init__(master, **kwargs)
        self.filepath = filepath
        self.configure(fg_color=("gray88", "gray20"), corner_radius=8)

        self.label = ctk.CTkLabel(
            self, text=os.path.basename(filepath), anchor="w", font=ctk.CTkFont(size=13)
        )
        self.label.pack(side="left", fill="x", expand=True, padx=(12, 4), pady=6)

        # Remove button (rightmost)
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

        # Open file button
        self.file_btn = ctk.CTkButton(
            self,
            text="\U0001f4c4",
            width=28,
            height=28,
            corner_radius=6,
            fg_color="transparent",
            hover_color=("gray75", "gray35"),
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=14),
            command=lambda: os.startfile(self.filepath),
        )
        self.file_btn.pack(side="right", padx=(0, 2), pady=6)

        # Open folder button
        self.folder_btn = ctk.CTkButton(
            self,
            text="\U0001f4c2",
            width=28,
            height=28,
            corner_radius=6,
            fg_color="transparent",
            hover_color=("gray75", "gray35"),
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=14),
            command=lambda: os.startfile(os.path.dirname(self.filepath)),
        )
        self.folder_btn.pack(side="right", padx=(0, 2), pady=6)


class ChatWindow:
    def __init__(self, master, pptx_path, model_id, api_key):
        self.pptx_path = pptx_path
        self.model_id = model_id
        self.api_key = api_key
        self.client = OpenAI(api_key=api_key)
        self.messages = []
        self.streaming = False
        self._current_ai_label = None

        self.win = ctk.CTkToplevel(master)
        self.win.title(f"AI Chat \u2014 {os.path.basename(pptx_path)}")
        self.win.geometry("720x660")
        self.win.minsize(520, 450)
        self.win.transient(master)

        self.setup_ui()
        self.load_presentation()

    def setup_ui(self):
        # Header
        header = ctk.CTkFrame(self.win, height=52, corner_radius=0, fg_color=("gray90", "gray17"))
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="\U0001f4ac  AI Chat",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(side="left", padx=16, pady=10)

        ctk.CTkLabel(
            header, text=os.path.basename(self.pptx_path),
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        ).pack(side="left", padx=(0, 16), pady=10)

        model_label = MODELS.get(self.model_id, {}).get("label", self.model_id)
        ctk.CTkLabel(
            header, text=f"Model: {model_label}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=16, pady=10)

        # Chat area
        self.chat_frame = ctk.CTkScrollableFrame(self.win, corner_radius=0, fg_color=("gray95", "gray14"))
        self.chat_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # Input area
        input_frame = ctk.CTkFrame(self.win, corner_radius=0, height=64, fg_color=("gray90", "gray17"))
        input_frame.pack(fill="x", side="bottom")
        input_frame.pack_propagate(False)

        self.input_var = ctk.StringVar()
        self.input_entry = ctk.CTkEntry(
            input_frame, textvariable=self.input_var,
            placeholder_text="Ask about the presentation...",
            height=42, font=ctk.CTkFont(size=13),
        )
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(12, 8), pady=11)
        self.input_entry.bind("<Return>", lambda e: self.send_message())

        self.send_btn = ctk.CTkButton(
            input_frame, text="Send", width=74, height=42,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.send_message,
        )
        self.send_btn.pack(side="right", padx=(0, 12), pady=11)

    def load_presentation(self):
        self.add_system_bubble("Loading presentation...")
        threading.Thread(target=self._load_pptx, daemon=True).start()

    def _load_pptx(self):
        try:
            self.slides_content = extract_file_content(self.pptx_path)

            # Build system message with text content
            text_summary = ""
            for slide in self.slides_content:
                text_summary += f"\n--- Slide {slide['number']} ---\n"
                for t in slide["texts"]:
                    text_summary += t + "\n"
                if slide["images"]:
                    text_summary += f"[{len(slide['images'])} image(s) on this slide]\n"

            system_msg = (
                "You are a helpful AI assistant analyzing a PowerPoint presentation. "
                "Here is the text content of each slide:\n" + text_summary + "\n"
                "Answer the user's questions about this presentation. Be specific about "
                "slide numbers when relevant. If the user asks about visual content, "
                "describe what you can see in the images provided."
            )
            self.messages = [{"role": "system", "content": system_msg}]

            # Include images for vision-capable models
            model_info = MODELS.get(self.model_id, {})
            has_vision = model_info.get("vision", False)

            if has_vision:
                content_parts = [{"type": "text", "text": "Here are the images from the presentation for your reference. Please acknowledge."}]
                img_count = 0
                for slide in self.slides_content:
                    for img in slide["images"]:
                        content_parts.append({
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{img['content_type']};base64,{img['base64']}",
                                "detail": "low",
                            }
                        })
                        img_count += 1
                if img_count > 0:
                    self.messages.append({"role": "user", "content": content_parts})
                    self.messages.append({
                        "role": "assistant",
                        "content": f"I've received and analyzed all {img_count} image(s) from the presentation. I'm ready to answer your questions about both the text and visual content."
                    })

            total_slides = len(self.slides_content)
            total_images = sum(len(s["images"]) for s in self.slides_content)

            self.win.after(0, lambda: self._on_pptx_loaded(total_slides, total_images))
        except Exception as e:
            self.win.after(0, lambda err=str(e): self._on_pptx_error(err))

    def _on_pptx_loaded(self, num_slides, num_images):
        for w in self.chat_frame.winfo_children():
            w.destroy()

        img_text = f" and {num_images} image(s)" if num_images > 0 else ""
        self.add_ai_bubble(
            f"I've analyzed your presentation ({num_slides} slide(s){img_text}). "
            "Ask me anything! For example:\n\n"
            "\u2022  \"What is slide 3 about?\"\n"
            "\u2022  \"Summarize the entire presentation\"\n"
            "\u2022  \"What are the key points?\"\n"
            "\u2022  \"Explain the diagram on slide 5\""
        )
        self.input_entry.focus()

    def _on_pptx_error(self, error):
        for w in self.chat_frame.winfo_children():
            w.destroy()
        self.add_system_bubble(f"Error loading presentation: {error}")

    def add_user_bubble(self, text):
        outer = ctk.CTkFrame(self.chat_frame, fg_color="transparent")
        outer.pack(fill="x", padx=12, pady=(6, 2))

        inner = ctk.CTkFrame(outer, fg_color=("#2563eb", "#1e40af"), corner_radius=16)
        inner.pack(side="right")

        label = ctk.CTkLabel(
            inner, text=text, wraplength=420, justify="left",
            text_color="white", font=ctk.CTkFont(size=13),
        )
        label.pack(padx=14, pady=10)

        self._scroll_to_bottom()

    def add_ai_bubble(self, text="", formatted=True):
        outer = ctk.CTkFrame(self.chat_frame, fg_color="transparent")
        outer.pack(fill="x", padx=12, pady=(6, 2))

        inner = ctk.CTkFrame(outer, fg_color=("gray82", "#2b2b2b"), corner_radius=16)
        inner.pack(side="left", anchor="nw")

        textbox = tk.Text(
            inner, wrap="word", borderwidth=0, highlightthickness=0,
            bg="#2b2b2b", fg="#dcdcdc", font=("Segoe UI", 11),
            cursor="arrow", padx=14, pady=10,
            selectbackground="#3b82f6", selectforeground="white",
            width=52, height=1, relief="flat",
            insertbackground="#2b2b2b",
        )
        textbox.pack(fill="x")

        # Configure tags
        textbox.tag_configure("bold", font=("Segoe UI", 11, "bold"))
        textbox.tag_configure("link", foreground="#60a5fa", underline=True)
        textbox.tag_configure("error", foreground="#ef4444")

        if text:
            if formatted:
                self._insert_formatted(textbox, text)
            else:
                textbox.insert("end", text)
            self._autosize_textbox(textbox)

        textbox.configure(state="disabled")
        self._scroll_to_bottom()
        return textbox

    def _insert_formatted(self, textbox, text):
        """Parse markdown bold (**text**) and URLs, insert with tags."""
        # Pattern matches **bold** or URLs
        pattern = r'(\*\*.*?\*\*|https?://[^\s\)\]\},]+)'
        parts = re.split(pattern, text)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                textbox.insert("end", part[2:-2], "bold")
            elif re.match(r'https?://', part):
                tag_name = f"link_{textbox.index('end')}"
                textbox.tag_configure(tag_name, foreground="#60a5fa", underline=True)
                url = part
                textbox.tag_bind(tag_name, "<Button-1>", lambda e, u=url: webbrowser.open(u))
                textbox.tag_bind(tag_name, "<Enter>", lambda e, tb=textbox: tb.configure(cursor="hand2"))
                textbox.tag_bind(tag_name, "<Leave>", lambda e, tb=textbox: tb.configure(cursor="arrow"))
                textbox.insert("end", part, tag_name)
            else:
                textbox.insert("end", part)

    def _autosize_textbox(self, textbox):
        """Auto-resize a Text widget to fit its content."""
        textbox.update_idletasks()
        # Count display lines (accounts for word wrap)
        try:
            count = textbox.count("1.0", "end", "displaylines")
            if count and count[0] > 0:
                textbox.configure(height=count[0])
            else:
                # Fallback: count newlines
                num_lines = int(textbox.index("end-1c").split(".")[0])
                textbox.configure(height=max(1, num_lines))
        except Exception:
            num_lines = int(textbox.index("end-1c").split(".")[0])
            textbox.configure(height=max(1, num_lines))

    def add_system_bubble(self, text):
        outer = ctk.CTkFrame(self.chat_frame, fg_color="transparent")
        outer.pack(fill="x", padx=12, pady=(8, 4))

        label = ctk.CTkLabel(
            outer, text=text, wraplength=500, justify="center",
            text_color=("gray50", "gray55"), font=ctk.CTkFont(size=12),
        )
        label.pack(pady=4)

    def _scroll_to_bottom(self):
        self.chat_frame.after(80, lambda: self.chat_frame._parent_canvas.yview_moveto(1.0))

    def send_message(self):
        text = self.input_var.get().strip()
        if not text or self.streaming:
            return

        self.input_var.set("")
        self.add_user_bubble(text)
        self.messages.append({"role": "user", "content": text})

        self.streaming = True
        self.send_btn.configure(state="disabled")
        self._current_ai_label = self.add_ai_bubble("\u2026")

        threading.Thread(target=self._stream_response, daemon=True).start()

    def _stream_response(self):
        try:
            response = self.client.chat.completions.create(
                model=self.model_id,
                messages=self.messages,
                stream=True,
                temperature=0.7,
            )

            full_text = ""
            for chunk in response:
                if chunk.choices and chunk.choices[0].delta.content:
                    full_text += chunk.choices[0].delta.content
                    text_snapshot = full_text
                    self.win.after(0, lambda t=text_snapshot: self._update_ai_label(t))

            self.messages.append({"role": "assistant", "content": full_text})
            self.win.after(0, self._stream_done)
        except Exception as e:
            error_msg = str(e)
            self.win.after(0, lambda err=error_msg: self._stream_error(err))

    def _update_ai_label(self, text):
        tb = self._current_ai_label
        if tb and tb.winfo_exists():
            tb.configure(state="normal")
            tb.delete("1.0", "end")
            tb.insert("1.0", text)
            self._autosize_textbox(tb)
            tb.configure(state="disabled")
            self._scroll_to_bottom()

    def _stream_done(self):
        self.streaming = False
        self.send_btn.configure(state="normal")
        # Re-format the completed response with bold/links
        tb = self._current_ai_label
        if tb and tb.winfo_exists():
            tb.configure(state="normal")
            raw_text = tb.get("1.0", "end-1c")
            tb.delete("1.0", "end")
            self._insert_formatted(tb, raw_text)
            self._autosize_textbox(tb)
            tb.configure(state="disabled")
        self.input_entry.focus()

    def _stream_error(self, error):
        self.streaming = False
        self.send_btn.configure(state="normal")
        tb = self._current_ai_label
        if tb and tb.winfo_exists():
            tb.configure(state="normal")
            tb.delete("1.0", "end")
            tb.insert("1.0", f"Error: {error}", "error")
            self._autosize_textbox(tb)
            tb.configure(state="disabled")


class PPTTranslatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Document Translator")
        master.geometry("600x760")
        master.minsize(520, 700)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.file_queue = []
        self.file_widgets = []
        self.cancel_event = threading.Event()
        self.translating = False
        self._rainbow_job = None

        self.setup_ui()
        self.load_api_key()

    def setup_ui(self):
        main = ctk.CTkFrame(self.master, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=16)

        # Title
        ctk.CTkLabel(
            main,
            text="Document Translator",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            main,
            text="PPTX & PDF  \u2022  Traditional Chinese \u2192 English",
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

        self.files_frame = ctk.CTkScrollableFrame(main, height=150, corner_radius=10)
        self.files_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.empty_label = ctk.CTkLabel(
            self.files_frame,
            text="No files added yet. Click 'Add Files' below.",
            text_color=("gray55", "gray55"),
            font=ctk.CTkFont(size=12),
        )
        self.empty_label.pack(pady=20)

        self.add_files_btn = ctk.CTkButton(
            main, text="+ Add Files", height=36, command=self.add_files
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
            height=44,
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
            font=ctk.CTkFont(size=22, weight="bold"),
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
            height=34,
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
        files = filedialog.askopenfilenames(filetypes=[("Supported Files", "*.pptx *.pdf"), ("PowerPoint", "*.pptx"), ("PDF", "*.pdf")])
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
                fp,
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
        self.status_label.configure(text="Translating... 0%")
        self.cancel_btn.configure(state="normal")

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
                ext = os.path.splitext(f)[1].lower()
                if ext == ".pdf":
                    total_paras += scan_pdf_paragraphs(f)
                else:
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
        output_paths = []

        for f in files:
            if self.cancel_event.is_set():
                break

            filename = os.path.basename(f)
            file_ext = os.path.splitext(filename)[1]
            default_name = os.path.splitext(filename)[0] + "_Translated" + file_ext
            output_path = os.path.join(output_dir, default_name)

            self.master.after(
                0, lambda fn=filename: self.progress_detail.configure(text=f"Translating: {fn}")
            )

            tokens_before_file_in = total_input_tokens
            tokens_before_file_out = total_output_tokens

            def progress_cb(global_done, in_tok, out_tok, _st=start_time, _tp=total_paras,
                            _tbi=tokens_before_file_in, _tbo=tokens_before_file_out):
                nonlocal total_input_tokens, total_output_tokens
                total_input_tokens = _tbi + in_tok
                total_output_tokens = _tbo + out_tok
                elapsed = time.time() - _st
                pct = global_done / _tp if _tp > 0 else 0
                total_tok = total_input_tokens + total_output_tokens
                eta = (elapsed / global_done * (_tp - global_done)) if global_done > 0 else 0
                model_info = MODELS.get(model_id, {})
                cost = (total_input_tokens / 1_000_000) * model_info.get("input_cost", 0) + \
                       (total_output_tokens / 1_000_000) * model_info.get("output_cost", 0)
                pct_int = int(pct * 100)

                self.master.after(0, lambda p=pct, e=eta, gd=global_done, tp=_tp, c=cost, it=total_tok, pi=pct_int: (
                    self.progress_bar.set(p),
                    self.progress_percent.configure(text=f"{pi}%"),
                    self.progress_detail.configure(text=f"Paragraph {gd}/{tp}"),
                    self.progress_stats.configure(
                        text=f"ETA: {int(e)}s  |  Tokens: {it:,}  |  ~${c:.4f}"
                    ),
                    self.status_label.configure(text=f"Translating... {pi}%"),
                ))

            try:
                ext = os.path.splitext(f)[1].lower()
                if ext == ".pdf":
                    process_fn = process_pdf
                else:
                    process_fn = process_pptx
                in_tok, out_tok, success, file_paras = process_fn(
                    f, output_path, model_id, client,
                    progress_callback=progress_cb,
                    cancel_event=self.cancel_event,
                    para_offset=paras_done,
                )
                total_input_tokens = tokens_before_file_in + in_tok
                total_output_tokens = tokens_before_file_out + out_tok
                if success:
                    paras_done += file_paras
                    completed_files += 1
                    output_paths.append(output_path)
            except Exception as e:
                self.master.after(0, lambda err=str(e): self.status_label.configure(text=f"Error: {err}"))

        if self.cancel_event.is_set():
            self.master.after(0, lambda: self.translation_done(
                total_input_tokens, total_output_tokens, False, "Translation cancelled.",
                None, None
            ))
        else:
            model_info = MODELS.get(model_id, {})
            cost = (total_input_tokens / 1_000_000) * model_info.get("input_cost", 0) + \
                   (total_output_tokens / 1_000_000) * model_info.get("output_cost", 0)
            elapsed = time.time() - start_time
            result_info = {
                "completed_files": completed_files,
                "total_files": len(files),
                "paragraphs": paras_done,
                "tokens": total_input_tokens + total_output_tokens,
                "cost": cost,
                "elapsed": int(elapsed),
                "output_dir": output_dir,
                "output_paths": output_paths,
                "model_id": model_id,
                "api_key": api_key,
            }
            # Show 100% briefly before showing completion window
            self.master.after(0, lambda: (
                self.progress_bar.set(1.0),
                self.progress_percent.configure(text="100%"),
                self.progress_detail.configure(text="Done!"),
                self.status_label.configure(text="Translating... 100%"),
            ))
            time.sleep(0.6)
            self.master.after(0, lambda ri=result_info: self.translation_done(
                total_input_tokens, total_output_tokens, True, "Translation complete!",
                ri.get("output_dir"), ri.get("output_paths"), ri
            ))

    def translation_done(self, input_tokens, output_tokens, success, message,
                         output_dir=None, output_paths=None, result_info=None):
        self.translating = False
        self.translate_btn.configure(state="normal" if self.file_queue else "disabled")
        self.add_files_btn.configure(state="normal")

        if success and result_info:
            self.file_queue.clear()
            self.refresh_file_list()
            self.progress_frame.pack_forget()
            self.status_label.configure(text="Complete!")
            self.show_completion_window(result_info)
        elif not success:
            messagebox.showwarning("Cancelled", message)
            self.progress_frame.pack_forget()
            self.status_label.configure(text="Cancelled")
        else:
            messagebox.showinfo("Info", message)
            self.progress_frame.pack_forget()
            self.status_label.configure(text=message)

    def show_completion_window(self, info):
        win = ctk.CTkToplevel(self.master)
        win.title("Translation Complete")
        win.geometry("500x520")
        win.resizable(False, False)
        win.transient(self.master)
        win.grab_set()
        self._completion_win = win

        # Center on parent
        win.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - 500) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 520) // 2
        win.geometry(f"+{x}+{y}")

        pad = ctk.CTkFrame(win, fg_color="transparent")
        pad.pack(fill="both", expand=True, padx=28, pady=24)

        # Checkmark and title
        ctk.CTkLabel(
            pad, text="\u2713", font=ctk.CTkFont(size=48, weight="bold"),
            text_color=("#2ea043", "#3fb950"),
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            pad, text="Translation Complete",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(0, 16))

        # Stats card
        stats_frame = ctk.CTkFrame(pad, corner_radius=10)
        stats_frame.pack(fill="x", pady=(0, 20))

        rows = [
            ("Files", f"{info['completed_files']}/{info['total_files']}"),
            ("Paragraphs", f"{info['paragraphs']:,}"),
            ("Tokens", f"{info['tokens']:,}"),
            ("Cost", f"${info['cost']:.4f}"),
            ("Time", f"{info['elapsed']}s"),
            ("Output", info['output_dir']),
        ]
        for i, (label, value) in enumerate(rows):
            row = ctk.CTkFrame(stats_frame, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=(10 if i == 0 else 2, 10 if i == len(rows) - 1 else 2))
            ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=12),
                         text_color=("gray50", "gray60"), width=90, anchor="w").pack(side="left")
            val_label = ctk.CTkLabel(row, text=value, font=ctk.CTkFont(size=12, weight="bold"), anchor="w")
            val_label.pack(side="left", fill="x", expand=True)

        # Buttons
        btn_frame = ctk.CTkFrame(pad, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 12))

        output_dir = info["output_dir"]
        output_paths = info.get("output_paths", [])

        ctk.CTkButton(
            btn_frame, text="Open in Folder", height=44,
            fg_color=("gray75", "gray30"), hover_color=("gray65", "gray40"),
            text_color=("gray10", "gray90"),
            font=ctk.CTkFont(size=13),
            command=lambda: self.open_in_folder(output_dir),
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            btn_frame, text="Open File", height=44,
            font=ctk.CTkFont(size=13),
            command=lambda: self.open_file(output_paths[0] if output_paths else output_dir),
        ).pack(side="left", fill="x", expand=True, padx=(6, 0))

        # Rainbow AI button
        self._ai_btn = ctk.CTkButton(
            pad,
            text="\u2728  Ask AI about this file",
            height=50,
            border_width=3,
            border_color="#a855f7",
            fg_color="transparent",
            hover_color=("gray85", "gray25"),
            text_color=("gray10", "gray95"),
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda: self._open_ai_chat(info, win),
        )
        self._ai_btn.pack(fill="x", pady=(0, 0))

        # Start rainbow animation
        self._rainbow_hue = 0.0
        self._animate_rainbow()

    def _animate_rainbow(self):
        if not hasattr(self, '_ai_btn') or self._ai_btn is None:
            return
        try:
            if not self._ai_btn.winfo_exists():
                return
        except Exception:
            return

        self._rainbow_hue = (self._rainbow_hue + 0.012) % 1.0

        # Generate two colors offset by 0.33 for a gradient-like look
        h1 = self._rainbow_hue
        h2 = (self._rainbow_hue + 0.33) % 1.0

        r1, g1, b1 = colorsys.hsv_to_rgb(h1, 0.75, 1.0)
        r2, g2, b2 = colorsys.hsv_to_rgb(h2, 0.75, 1.0)

        color1 = f"#{int(r1*255):02x}{int(g1*255):02x}{int(b1*255):02x}"
        color2 = f"#{int(r2*255):02x}{int(g2*255):02x}{int(b2*255):02x}"

        self._ai_btn.configure(border_color=color1, text_color=(color2, color1))

        self._rainbow_job = self.master.after(40, self._animate_rainbow)

    def _open_ai_chat(self, info, completion_win):
        # Stop rainbow animation
        if self._rainbow_job:
            self.master.after_cancel(self._rainbow_job)
            self._rainbow_job = None
        self._ai_btn = None

        output_paths = info.get("output_paths", [])
        if not output_paths:
            messagebox.showwarning("No File", "No translated file found.")
            return

        pptx_path = output_paths[0]
        model_id = info.get("model_id", self.get_selected_model_id())
        api_key = info.get("api_key", self.api_key_var.get().strip())

        completion_win.destroy()
        ChatWindow(self.master, pptx_path, model_id, api_key)

    def open_in_folder(self, folder_path):
        os.startfile(folder_path)

    def open_file(self, file_path):
        os.startfile(file_path)

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
