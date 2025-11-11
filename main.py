import os
import platform
import subprocess
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List

"""
Universal Label Printer
=======================

This script provides a cross‑platform graphical interface for sending ZPL label
files (commonly delivered as `.txt` or `.zpl`) to any available printer. The
application detects installed printers, allows the user to select one from a
list, and offers a "Test mode" which prints only the first label for
verification before committing to print the entire set.

It is aimed at Mercado Livre "Full" product labels printed on thermal
printers, but because the printing routine sends raw data, it should work with
any printer that can interpret ZPL commands. If the printer does not support
ZPL, the output may be gibberish. In such cases, use software that converts
ZPL to PDF or the vendor's recommended solution.

Features:

* **Printer selection**: The app lists all installed printers on the host
  operating system and lets the user choose one.
* **Test mode**: Enabled by default; prints only the first label (`^XA…^XZ` block).
  Disable test mode via the checkbox to print all labels in the file.
* **Cross‑platform**: Uses `pywin32` on Windows and the `lp` command on
  macOS/Linux. There is no need to hardcode a printer name.

Note: On Windows, this script depends on the `pywin32` package to access
`win32print`. Ensure it is installed when running the script or bundling it
into an executable.
"""

# Global: hold selected printer name and test mode state
selected_printer = None  # type: str


def list_printers() -> List[str]:
    """Return a list of installed printer names on the host system."""
    system_name = platform.system().lower()
    printers = []
    if system_name.startswith("win"):
        try:
            import win32print  # type: ignore
        except ImportError:
            return []
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            printers.append(p[2])
    else:
        # Use lpstat to list printers on Unix systems
        try:
            result = subprocess.run(
                ["lpstat", "-p"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=False,
            )
            for line in result.stdout.splitlines():
                parts = line.split()
                if parts and parts[0] == "printer":
                    printers.append(parts[1])
        except Exception:
            pass
    return printers


def split_zpl_labels(content: str) -> List[str]:
    """Split raw ZPL content into individual label blocks."""
    parts = content.split("^XZ")
    labels = []
    for part in parts:
        if part.strip():
            labels.append(part + "^XZ\n")
    return labels


def send_raw_to_printer(printer: str, data: str) -> None:
    """Send raw ZPL data to the specified printer name."""
    system_name = platform.system().lower()
    if system_name.startswith("win"):
        try:
            import win32print  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 is required on Windows to send raw data to printers"
            ) from exc
        # Open printer
        try:
            handle = win32print.OpenPrinter(printer)
        except win32print.error as exc:
            raise RuntimeError(
                f"Could not open printer '{printer}'. Check the name and driver."
            ) from exc
        try:
            job = win32print.StartDocPrinter(handle, 1, ("Label Print", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, data.encode("latin-1"))
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
        finally:
            win32print.ClosePrinter(handle)
    else:
        # macOS/Linux: write to temp file and invoke lp
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zpl") as tmp:
            tmp.write(data.encode("latin-1"))
            tmp_path = tmp.name
        try:
            result = subprocess.run(
                ["lp", "-d", printer, "-o", "raw", tmp_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=False,
            )
            if result.returncode != 0:
                raise RuntimeError(f"Print error: {result.stderr.strip()}")
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


def print_labels(file_path: str, printer: str, test_only: bool) -> None:
    """Read a ZPL file, split into labels, and send to the printer."""
    try:
        with open(file_path, "r", encoding="latin-1") as f:
            content = f.read()
    except OSError as exc:
        raise RuntimeError(f"Could not read file: {file_path}") from exc

    if "^XA" not in content:
        raise RuntimeError("The file does not contain ZPL commands (missing ^XA).")
    labels = split_zpl_labels(content)
    if not labels:
        raise RuntimeError("No labels found in the file.")
    to_print = labels[:1] if test_only else labels
    for label in to_print:
        send_raw_to_printer(printer, label)


def browse_and_print(printers: List[str], test_var: tk.IntVar, file_label: tk.Label):
    """Handle the file selection and printing logic from the GUI."""
    global selected_printer
    # Choose file
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo de etiquetas (.txt ou .zpl)",
        filetypes=(("Arquivos de etiqueta", "*.txt *.zpl"), ("Todos", "*.*")),
    )
    if not file_path:
        return
    # Determine selected printer
    if not selected_printer:
        messagebox.showerror(
            "Erro", "Por favor, selecione uma impressora antes de imprimir.",
        )
        return
    try:
        print_labels(file_path, selected_printer, bool(test_var.get()))
    except Exception as exc:
        messagebox.showerror("Erro", str(exc))
        return
    if test_var.get():
        messagebox.showinfo("Sucesso", "A primeira fileira foi enviada para a impressora.")
    else:
        messagebox.showinfo("Sucesso", "Todas as etiquetas foram enviadas para a impressora.")


def on_printer_select(event, printer_var: tk.StringVar):
    """Update the selected_printer global when the user selects a printer from the dropdown."""
    global selected_printer
    selected_printer = printer_var.get()


def create_gui():
    """Set up and run the Tkinter GUI."""
    printers = list_printers()
    root = tk.Tk()
    root.title("Impressora Universal de Etiquetas")
    root.geometry("560x240")

    description = (
        "Selecione uma impressora, escolha o arquivo (.txt ou .zpl) e imprima etiquetas.\n"
        "Modo teste (marcado) imprimirá apenas a primeira fileira para verificação."
    )
    tk.Label(root, text=description, wraplength=540, justify="left", pady=10).pack()

    # Printer selection
    frame = tk.Frame(root)
    frame.pack(pady=5)
    tk.Label(frame, text="Impressora:").pack(side="left")
    printer_var = tk.StringVar(value=printers[0] if printers else "")
    if printers:
        global selected_printer
        selected_printer = printers[0]
    printer_menu = tk.OptionMenu(frame, printer_var, *printers, command=lambda _: on_printer_select(_, printer_var))
    printer_menu.config(width=40)
    printer_menu.pack(side="left", padx=10)

    # Test mode checkbox
    test_var = tk.IntVar(value=1)
    tk.Checkbutton(root, text="Modo Teste (imprimir apenas a primeira fileira)", variable=test_var).pack()

    # Print button
    tk.Button(
        root,
        text="Selecionar arquivo e Imprimir",
        command=lambda: browse_and_print(printers, test_var, None),
        height=2,
        width=35,
    ).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    create_gui()
