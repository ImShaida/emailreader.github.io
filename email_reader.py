import os
import sys
import tempfile
import webbrowser
import traceback
from datetime import datetime
from email import policy
from email.parser import BytesParser

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except Exception as e:
    print("Tkinter is required. On Windows, install the official Python from python.org which includes tkinter.")
    raise

# Optional dependencies
try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None

try:
    import extract_msg
except Exception:
    extract_msg = None

APP_TITLE = "Email Reader & Attachment Extractor"


def parse_eml(path):
    """
    Parse an .eml file using the standard library and return
    (headers_dict, body_text, body_html, attachments_list)
    attachments_list := list of dict {'filename','content_type','payload' (bytes)}
    """
    with open(path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    headers = dict(msg.items())

    body_text = ""
    body_html = ""
    attachments = []

    if msg.is_multipart():
        for part in msg.walk():
            cdis = part.get_content_disposition()  # 'inline', 'attachment', or None
            ctype = part.get_content_type()
            # Attachments (or inline files with filename)
            filename = part.get_filename()
            if cdis == "attachment" or filename:
                payload = part.get_payload(decode=True)
                attachments.append({
                    "filename": filename or "attachment",
                    "content_type": ctype,
                    "payload": payload
                })
            else:
                # Capture first plain text and first html body
                if ctype == "text/plain" and not body_text:
                    try:
                        body_text = part.get_content()
                    except Exception:
                        body_text = part.get_payload(decode=True).decode(errors="replace")
                elif ctype == "text/html" and not body_html:
                    try:
                        body_html = part.get_content()
                    except Exception:
                        body_html = part.get_payload(decode=True).decode(errors="replace")
    else:
        ctype = msg.get_content_type()
        if ctype == "text/plain":
            body_text = msg.get_content()
        elif ctype == "text/html":
            body_html = msg.get_content()

    return headers, body_text, body_html, attachments


def parse_msg(path):
    """
    Parse a .msg file using extract_msg (optional dependency).
    Returns (headers_dict, body_text, body_html, attachments_list)
    attachments_list := list of dict {'filename','obj'}
    where 'obj' is the extract_msg attachment object (saveable)
    """
    if extract_msg is None:
        raise RuntimeError("extract_msg is not installed. Install with: pip install extract_msg")

    m = extract_msg.Message(path)
    headers = {
        "From": getattr(m, "sender", ""),
        "To": getattr(m, "to", ""),
        "Subject": getattr(m, "subject", ""),
        "Date": getattr(m, "date", ""),
    }
    body_text = getattr(m, "body", "") or ""
    body_html = getattr(m, "htmlBody", "") or ""

    attachments = []
    for att in getattr(m, "attachments", []) or []:
        fname = getattr(att, "longFilename", None) or getattr(att, "filename", None) or "attachment"
        attachments.append({"filename": fname, "obj": att})

    return headers, body_text, body_html, attachments


class EmailReaderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x600")

        self.current_path = None
        self.current_meta = {}  # {'headers','body_text','body_html','attachments','type'}

        self._build_ui()

    def _build_ui(self):
        # Top frame: buttons
        top = ttk.Frame(self)
        top.pack(side="top", fill="x", padx=8, pady=6)

        btn_open = ttk.Button(top, text="Open Email", command=self.open_email)
        btn_open.pack(side="left")

        btn_extract = ttk.Button(top, text="Extract Attachments", command=self.extract_attachments)
        btn_extract.pack(side="left", padx=(6,0))

        btn_savebody = ttk.Button(top, text="Save Body as .txt", command=self.save_body_as_txt)
        btn_savebody.pack(side="left", padx=(6,0))

        btn_openhtml = ttk.Button(top, text="Open HTML in browser", command=self.open_html_in_browser)
        btn_openhtml.pack(side="left", padx=(6,0))

        # Tabs: Headers | Body | Attachments | Raw
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=6)

        # Headers tab
        self.header_text = scrolledtext.ScrolledText(notebook, wrap="none")
        self.header_text.configure(state="disabled")
        notebook.add(self.header_text, text="Headers")

        # Body tab (plain text)
        self.body_text_widget = scrolledtext.ScrolledText(notebook)
        self.body_text_widget.configure(state="disabled")
        notebook.add(self.body_text_widget, text="Body (text)")

        # Attachments tab
        attach_frame = ttk.Frame(notebook)
        attach_frame.pack(fill="both", expand=True)
        self.attach_listbox = tk.Listbox(attach_frame, height=8)
        self.attach_listbox.pack(side="left", fill="both", expand=True, padx=(4,0), pady=4)
        self.attach_listbox.bind("<Double-Button-1>", self.on_attachment_doubleclick)
        scrollbar = ttk.Scrollbar(attach_frame, orient="vertical", command=self.attach_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.attach_listbox.config(yscrollcommand=scrollbar.set)

        right_panel = ttk.Frame(attach_frame)
        right_panel.pack(side="left", fill="y", padx=6)
        ttk.Label(right_panel, text="Attachment actions:").pack(anchor="nw", pady=(4,2))
        ttk.Button(right_panel, text="Save Selected", command=self.save_selected_attachment).pack(fill="x", pady=2)
        ttk.Button(right_panel, text="Save All", command=self.extract_attachments).pack(fill="x", pady=2)

        notebook.add(attach_frame, text="Attachments")

        # Raw tab (raw RFC content)
        self.raw_text = scrolledtext.ScrolledText(notebook, wrap="none")
        self.raw_text.configure(state="disabled")
        notebook.add(self.raw_text, text="Raw/Debug")

        # Status bar
        self.status = ttk.Label(self, text="No file loaded.", relief="sunken", anchor="w")
        self.status.pack(side="bottom", fill="x")

    def set_status(self, text):
        self.status.configure(text=text)
        self.update_idletasks()

    def open_email(self):
        path = filedialog.askopenfilename(title="Open email file", filetypes=[("Email files", "*.eml *.msg"), ("EML files", "*.eml"), ("MSG files", "*.msg"), ("All files", "*.*")])
        if not path:
            return
        self.current_path = path
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".eml":
                headers, body_text, body_html, attachments = parse_eml(path)
                filetype = "eml"
            elif ext == ".msg":
                if extract_msg is None:
                    messagebox.showwarning("Missing dependency", "MSG support requires 'extract_msg' package. Install with:\n\npip install extract_msg\n\nYou can still open EML files.")
                    return
                headers, body_text, body_html, attachments = parse_msg(path)
                filetype = "msg"
            else:
                # try eml parse as fallback
                headers, body_text, body_html, attachments = parse_eml(path)
                filetype = "eml"

            self.current_meta = {
                "headers": headers,
                "body_text": body_text,
                "body_html": body_html,
                "attachments": attachments,
                "type": filetype
            }
            self._render_loaded_email()
            self.set_status(f"Loaded: {os.path.basename(path)} ({filetype.upper()}) -- {len(attachments)} attachments")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open file:\n{e}")

    def _render_loaded_email(self):
        headers = self.current_meta.get("headers", {})
        body_text = self.current_meta.get("body_text", "")
        body_html = self.current_meta.get("body_html", "")
        attachments = self.current_meta.get("attachments", [])

        # Headers
        self.header_text.configure(state="normal")
        self.header_text.delete("1.0", "end")
        header_lines = [f"{k}: {v}" for k, v in headers.items()]
        self.header_text.insert("1.0", "\n".join(header_lines))
        self.header_text.configure(state="disabled")

        # Body (prefer plain text)
        display_text = body_text
        if not display_text and body_html:
            # Try to convert HTML to text if BeautifulSoup is available
            if BeautifulSoup:
                display_text = BeautifulSoup(body_html, "html.parser").get_text()
            else:
                display_text = "(No plain-text body; HTML present. Use 'Open HTML in browser' to view.)\n\n" + body_html

        self.body_text_widget.configure(state="normal")
        self.body_text_widget.delete("1.0", "end")
        self.body_text_widget.insert("1.0", display_text)
        self.body_text_widget.configure(state="disabled")

        # Raw debug (for advanced users)
        try:
            with open(self.current_path, "rb") as f:
                raw = f.read().decode(errors="replace")
        except Exception:
            raw = "(unable to load raw)"
        self.raw_text.configure(state="normal")
        self.raw_text.delete("1.0", "end")
        self.raw_text.insert("1.0", raw)
        self.raw_text.configure(state="disabled")

        # Attachments listbox
        self.attach_listbox.delete(0, "end")
        for a in attachments:
            if isinstance(a, dict):
                fname = a.get("filename", "attachment")
                size = len(a.get("payload") or b"")
                self.attach_listbox.insert("end", f"{fname} ({size} bytes)")
            else:
                self.attach_listbox.insert("end", str(a))

    def open_html_in_browser(self):
        body_html = self.current_meta.get("body_html") or ""
        if not body_html:
            messagebox.showinfo("No HTML body", "No HTML body was detected for this message.")
            return
        fd, tmp = tempfile.mkstemp(suffix=".html", text=True)
        with os.fdopen(fd, "w", encoding="utf-8", errors="replace") as f:
            f.write(body_html)
        webbrowser.open("file://" + tmp)

    def save_body_as_txt(self):
        body = self.current_meta.get("body_text") or ""
        if not body and self.current_meta.get("body_html"):
            if BeautifulSoup:
                body = BeautifulSoup(self.current_meta.get("body_html"), "html.parser").get_text()
            else:
                body = "(HTML body present; install beautifulsoup4 to convert to text.)"
        if not body:
            messagebox.showinfo("No body", "No body content available to save.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files","*.txt"),("All files","*.*")])
        if not path:
            return
        with open(path, "w", encoding="utf-8", errors="replace") as f:
            f.write(body)
        messagebox.showinfo("Saved", f"Body saved to: {path}")

    def extract_attachments(self):
        attachments = self.current_meta.get("attachments", [])
        if not attachments:
            messagebox.showinfo("No attachments", "No attachments were found in the loaded message.")
            return
        dest = filedialog.askdirectory(title="Select folder to save attachments")
        if not dest:
            return
        saved = 0
        for a in attachments:
            try:
                if self.current_meta.get("type") == "eml":
                    fname = a.get("filename") or f"attachment_{saved}"
                    payload = a.get("payload") or b""
                    safe_name = self._sanitize_filename(fname)
                    full = os.path.join(dest, safe_name)
                    with open(full, "wb") as fh:
                        fh.write(payload)
                    saved += 1
                elif self.current_meta.get("type") == "msg":
                    obj = a.get("obj")
                    fname = a.get("filename") or getattr(obj, "filename", "attachment")
                    safe_name = self._sanitize_filename(fname)
                    full = os.path.join(dest, safe_name)
                    if hasattr(obj, "save"):
                        try:
                            obj.save(customPath=full)
                        except TypeError:
                            try:
                                obj.save(full)
                            except Exception:
                                data = getattr(obj, "data", None) or getattr(obj, "payload", None)
                                if data:
                                    with open(full, "wb") as fh:
                                        fh.write(data)
                                else:
                                    raise
                        saved += 1
                    else:
                        data = getattr(obj, "data", None) or getattr(obj, "payload", None)
                        if data:
                            with open(full, "wb") as fh:
                                fh.write(data)
                            saved += 1
                        else:
                            messagebox.showwarning("Warning", f"Could not save attachment: {fname}")
                else:
                    if isinstance(a, dict):
                        fname = a.get("filename","attachment")
                        payload = a.get("payload") or b""
                        with open(os.path.join(dest, fname), "wb") as fh:
                            fh.write(payload)
                        saved += 1
            except Exception as e:
                traceback.print_exc()
                messagebox.showwarning("Save error", f"Failed to save one attachment: {e}")
        messagebox.showinfo("Done", f"Saved {saved} attachment(s) to:\n{dest}")
        self.set_status(f"Saved {saved} attachment(s) to {dest}")

    def save_selected_attachment(self):
        sel = self.attach_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select", "Select an attachment from the list first (double-click to save).")
            return
        idx = sel[0]
        attachments = self.current_meta.get("attachments", [])
        if idx >= len(attachments):
            messagebox.showerror("Index error", "Selected index out of range.")
            return
        dest_file = filedialog.asksaveasfilename(defaultextension="", initialfile=attachments[idx].get("filename") if isinstance(attachments[idx], dict) else str(attachments[idx]))
        if not dest_file:
            return
        a = attachments[idx]
        try:
            if self.current_meta.get("type") == "eml" and isinstance(a, dict):
                payload = a.get("payload") or b""
                with open(dest_file, "wb") as fh:
                    fh.write(payload)
            elif self.current_meta.get("type") == "msg":
                obj = a.get("obj")
                if hasattr(obj, "save"):
                    try:
                        obj.save(customPath=dest_file)
                    except TypeError:
                        try:
                            obj.save(dest_file)
                        except Exception:
                            data = getattr(obj, "data", None) or getattr(obj, "payload", None)
                            if data:
                                with open(dest_file, "wb") as fh:
                                    fh.write(data)
                            else:
                                raise
                else:
                    data = getattr(obj, "data", None) or getattr(obj, "payload", None)
                    if data:
                        with open(dest_file, "wb") as fh:
                            fh.write(data)
                    else:
                        raise RuntimeError("Attachment object not writable")
            else:
                if isinstance(a, dict):
                    payload = a.get("payload") or b""
                    with open(dest_file, "wb") as fh:
                        fh.write(payload)
                else:
                    raise RuntimeError("Unknown attachment type")
            messagebox.showinfo("Saved", f"Attachment saved to: {dest_file}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to save selected attachment:\n{e}")

    def on_attachment_doubleclick(self, event):
        self.save_selected_attachment()

    @staticmethod
    def _sanitize_filename(name):
        return "".join(c for c in name if c not in "\\/:*?\"<>|").strip() or "attachment"

def main():
    app = EmailReaderApp()
    app.mainloop()

if __name__ == "__main__":
    main()
