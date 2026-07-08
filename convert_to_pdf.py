"""Converter em PDF - context menu batch merger.

Receives selected files as argv, converts each one to PDF internally, then
merges everything (in the order the shell provided) into a single PDF that
the user names via a Save As dialog. Built to run standalone via PyInstaller
(--onedir --noconsole -- onedir instead of onefile because onefile re-extracts
~46MB to a temp folder on every single launch, which is too slow on HDD to
survive 10-20 near-simultaneous cold starts), so it must not depend on a
console or on any Python install being present on the target machine.

Windows Explorer, on this shell version, does NOT honor MultiSelectModel=Player:
selecting N files still launches this exe N times (once per file) instead of
once with all of them. get_batch_files() works around that: the first
instance to grab a named mutex becomes the "leader" and briefly waits for its
siblings to report in via a shared queue file; every other instance is a
"follower" that just appends its file(s) to that queue and exits immediately.
"""
import ctypes
import datetime
import os
import shutil
import sys
import tempfile
import time

BATCH_MUTEX_NAME = "ConversorPDF_LeaderMutex"
BATCH_QUEUE_PATH = os.path.join(tempfile.gettempdir(), "conversorpdf_batch_queue.txt")
BATCH_MIN_WAIT_SECONDS = 1.5   # always wait at least this long, covers slow sibling cold-starts
BATCH_QUIET_SECONDS = 0.6      # stop once the queue file hasn't grown for this long
BATCH_MAX_WAIT_SECONDS = 6.0   # hard cap so the leader never hangs indefinitely
BATCH_POLL_SECONDS = 0.15

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp"}
WORD_EXT = {".doc", ".docx", ".rtf"}
EXCEL_EXT = {".xls", ".xlsx", ".csv"}
POWERPOINT_EXT = {".ppt", ".pptx"}
TEXT_EXT = {".txt"}

LOG_PATH = os.path.join(tempfile.gettempdir(), "pdf-context-menu.log")

MB_ICONINFORMATION = 0x40
MB_ICONERROR = 0x10


def log(message):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] {message}\n")


def show_message(title, text, icon=MB_ICONINFORMATION):
    ctypes.windll.user32.MessageBoxW(0, text, title, icon)


def _write_lines_retry(path, mode, lines, attempts=20, delay=0.05):
    for _ in range(attempts):
        try:
            with open(path, mode, encoding="utf-8") as f:
                for line in lines:
                    f.write(line + "\n")
            return True
        except OSError:
            time.sleep(delay)
    return False


def get_batch_files():
    """Returns (files, mutex) for the leader, or (None, mutex) for a follower
    (caller should just exit). `mutex` must stay referenced until the leader
    is done, otherwise its handle closes and a later invocation could wrongly
    become leader again mid-batch."""
    import win32api
    import win32event
    import winerror

    my_files = [a for a in sys.argv[1:] if a]

    mutex = win32event.CreateMutex(None, False, BATCH_MUTEX_NAME)
    is_leader = win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS

    if is_leader:
        _write_lines_retry(BATCH_QUEUE_PATH, "w", my_files)

        start = time.monotonic()
        last_size = os.path.getsize(BATCH_QUEUE_PATH) if os.path.exists(BATCH_QUEUE_PATH) else 0
        last_growth = start
        while True:
            time.sleep(BATCH_POLL_SECONDS)
            now = time.monotonic()
            try:
                size = os.path.getsize(BATCH_QUEUE_PATH)
            except OSError:
                size = last_size
            if size != last_size:
                last_size = size
                last_growth = now
            elapsed = now - start
            quiet_for = now - last_growth
            if elapsed >= BATCH_MAX_WAIT_SECONDS:
                break
            if elapsed >= BATCH_MIN_WAIT_SECONDS and quiet_for >= BATCH_QUIET_SECONDS:
                break

        try:
            with open(BATCH_QUEUE_PATH, "r", encoding="utf-8") as f:
                all_files = [line.strip() for line in f if line.strip()]
        except OSError:
            all_files = my_files
        try:
            os.remove(BATCH_QUEUE_PATH)
        except OSError:
            pass
        log(f"Lider: {len(all_files)} arquivo(s) recebidos no lote (esperou {time.monotonic() - start:.1f}s)")
        return all_files, mutex

    for _ in range(40):
        if os.path.exists(BATCH_QUEUE_PATH):
            break
        time.sleep(0.05)
    _write_lines_retry(BATCH_QUEUE_PATH, "a", my_files)
    return None, mutex


def ask_file_order(files):
    """Lets the user reorder the collected files before merging. Returns the
    reordered list, or None if the user cancelled."""
    import tkinter as tk

    result = {"files": None}

    root = tk.Tk()
    root.title("Converter em PDF - ordem dos arquivos")
    root.attributes("-topmost", True)
    root.geometry("480x420")

    tk.Label(
        root,
        text="Ajuste a ordem dos arquivos no PDF final (selecione um e use os botoes):",
        wraplength=440,
        justify="left",
    ).pack(padx=10, pady=(10, 5), anchor="w")

    list_frame = tk.Frame(root)
    list_frame.pack(fill="both", expand=True, padx=10)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(list_frame, selectmode="browse", yscrollcommand=scrollbar.set)
    for f in files:
        listbox.insert("end", os.path.basename(f))
    listbox.selection_set(0)
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    order = list(files)

    def move(delta):
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= len(order):
            return
        order[i], order[j] = order[j], order[i]
        text = listbox.get(i)
        listbox.delete(i)
        listbox.insert(j, text)
        listbox.selection_set(j)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=8)
    tk.Button(btn_frame, text="Mover para cima", width=16, command=lambda: move(-1)).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Mover para baixo", width=16, command=lambda: move(1)).pack(side="left", padx=5)

    def confirm():
        result["files"] = order
        root.destroy()

    def cancel():
        result["files"] = None
        root.destroy()

    action_frame = tk.Frame(root)
    action_frame.pack(pady=(0, 10))
    tk.Button(action_frame, text="Continuar", width=14, command=confirm).pack(side="left", padx=5)
    tk.Button(action_frame, text="Cancelar", width=14, command=cancel).pack(side="left", padx=5)

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.mainloop()

    return result["files"]


def ask_output_path(first_file):
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    initial_dir = os.path.dirname(first_file) or "."
    base_name = os.path.splitext(os.path.basename(first_file))[0]
    initial_name = f"{base_name}_unificado.pdf"

    path = filedialog.asksaveasfilename(
        title="Salvar PDF unificado como",
        initialdir=initial_dir,
        initialfile=initial_name,
        defaultextension=".pdf",
        filetypes=[("Arquivo PDF", "*.pdf")],
    )
    root.destroy()
    return path or None


def convert_image(src_path, dst_path):
    from PIL import Image

    with Image.open(src_path) as img:
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        img.save(dst_path, "PDF")


def convert_text(src_path, dst_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    with open(src_path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.read().splitlines()

    c = canvas.Canvas(dst_path, pagesize=A4)
    width, height = A4
    left_margin = 40
    top_margin = height - 40
    line_height = 14
    c.setFont("Courier", 10)

    y = top_margin
    for line in lines:
        if y < 40:
            c.showPage()
            c.setFont("Courier", 10)
            y = top_margin
        c.drawString(left_margin, y, line[:110])
        y -= line_height
    c.save()


class OfficeApps:
    """Lazily opens Word/Excel/PowerPoint once and reuses them for the whole batch."""

    def __init__(self):
        self._word = None
        self._excel = None
        self._powerpoint = None

    def word(self):
        if self._word is None:
            import win32com.client

            self._word = win32com.client.Dispatch("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = False
            try:
                self._word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
        return self._word

    def excel(self):
        if self._excel is None:
            import win32com.client

            self._excel = win32com.client.Dispatch("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            self._excel.AskToUpdateLinks = False
            try:
                self._excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
        return self._excel

    def powerpoint(self):
        if self._powerpoint is None:
            import win32com.client

            self._powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            self._powerpoint.DisplayAlerts = False
            try:
                self._powerpoint.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
        return self._powerpoint

    def quit_all(self):
        for app in (self._word, self._excel, self._powerpoint):
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass


def convert_word(office, src_path, dst_path):
    word = office.word()
    doc = word.Documents.Open(src_path, ReadOnly=True)
    try:
        doc.ExportAsFixedFormat(dst_path, 17)  # wdExportFormatPDF
    finally:
        doc.Close(False)


def convert_excel(office, src_path, dst_path):
    excel = office.excel()
    wb = excel.Workbooks.Open(src_path, ReadOnly=True, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
    try:
        wb.ExportAsFixedFormat(0, dst_path)  # xlTypePDF
    finally:
        wb.Close(False)


def convert_powerpoint(office, src_path, dst_path):
    powerpoint = office.powerpoint()
    pres = powerpoint.Presentations.Open(src_path, WithWindow=False)
    try:
        pres.SaveAs(dst_path, 32)  # ppSaveAsPDF
    finally:
        pres.Close()


def describe_error(exc):
    """pywintypes.com_error hides the real Office message behind a generic
    HRESULT; excepinfo[2] usually has the human-readable text (e.g. 'this
    file is protected' or 'file already open')."""
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return excepinfo[2].strip()
    args = getattr(exc, "args", None)
    if args and len(args) > 2 and isinstance(args[2], (tuple, list)) and len(args[2]) > 2 and args[2][2]:
        return str(args[2][2]).strip()
    return str(exc)


def to_pdf_piece(office, src_path, dst_path):
    """Converts one selected file into a standalone PDF at dst_path."""
    ext = os.path.splitext(src_path)[1].lower()

    if not os.path.isfile(src_path):
        return "fail", "arquivo nao encontrado"

    if ext == ".pdf":
        shutil.copyfile(src_path, dst_path)
        return "ok", dst_path

    try:
        if ext in IMAGE_EXT:
            convert_image(src_path, dst_path)
        elif ext in TEXT_EXT:
            convert_text(src_path, dst_path)
        elif ext in WORD_EXT:
            convert_word(office, src_path, dst_path)
        elif ext in EXCEL_EXT:
            convert_excel(office, src_path, dst_path)
        elif ext in POWERPOINT_EXT:
            convert_powerpoint(office, src_path, dst_path)
        else:
            return "skip", f"tipo nao suportado ({ext})"
    except Exception as exc:
        reason = describe_error(exc)
        log(f"FALHOU {src_path}: {exc!r} | motivo: {reason}")
        return "fail", reason

    return "ok", dst_path


def merge_pdfs(pdf_paths, output_path):
    from pypdf import PdfWriter

    writer = PdfWriter()
    try:
        for p in pdf_paths:
            writer.append(p)
        writer.write(output_path)
    finally:
        writer.close()


def main():
    files, _mutex = get_batch_files()
    if files is None:
        return  # follower: just reported its files to the leader, nothing else to do

    if not files:
        show_message("Converter em PDF", "Nenhum arquivo selecionado.", MB_ICONERROR)
        return

    if len(files) > 1:
        ordered = ask_file_order(files)
        if ordered is None:
            log("Usuario cancelou a reordenacao dos arquivos")
            return
        files = ordered

    output_path = ask_output_path(files[0])
    if not output_path:
        log("Usuario cancelou a escolha do arquivo de saida")
        return

    log(f"Iniciando juncao de {len(files)} arquivo(s) -> {output_path}")

    work_dir = tempfile.mkdtemp(prefix="pdf-context-menu-")
    office = OfficeApps()
    ok_pieces = []
    skipped, failed = [], []
    try:
        for i, f in enumerate(files):
            piece_path = os.path.join(work_dir, f"{i:03d}.pdf")
            status, detail = to_pdf_piece(office, f, piece_path)
            name = os.path.basename(f)
            if status == "ok":
                ok_pieces.append(detail)
                log(f"OK {f}")
            elif status == "skip":
                skipped.append(f"{name} ({detail})")
            else:
                failed.append(f"{name}: {detail}")
    finally:
        office.quit_all()

    if not ok_pieces:
        log("Nenhum arquivo pode ser convertido, PDF nao foi gerado")
        show_message("Converter em PDF", "Nenhum arquivo pode ser convertido. Nada foi gerado.", MB_ICONERROR)
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    try:
        merge_pdfs(ok_pieces, output_path)
        log(f"PDF final gerado: {output_path}")
    except Exception as exc:
        log(f"ERRO ao juntar PDFs: {exc}")
        show_message("Converter em PDF", f"Erro ao juntar os PDFs: {exc}", MB_ICONERROR)
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    shutil.rmtree(work_dir, ignore_errors=True)

    lines = [f"PDF criado: {os.path.basename(output_path)}", f"Arquivos incluidos: {len(ok_pieces)}"]
    if skipped:
        lines.append(f"Pulados: {len(skipped)}")
    if failed:
        lines.append(f"Falharam: {len(failed)}")
    details = skipped + failed
    if details:
        lines.append("")
        lines.extend(details[:15])
        if len(details) > 15:
            lines.append(f"... e mais {len(details) - 15} (ver log em {LOG_PATH})")

    icon = MB_ICONERROR if failed else MB_ICONINFORMATION
    show_message("Converter em PDF", "\n".join(lines), icon)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"ERRO FATAL: {exc}")
        show_message("Converter em PDF", f"Erro inesperado: {exc}", MB_ICONERROR)
