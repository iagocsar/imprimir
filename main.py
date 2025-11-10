import os, sys, platform, tkinter as tk
from tkinter import filedialog, messagebox

PRINTER_NAME = "MDK-006"
BATCH_SIZE = 50

def split_labels(content):
    parts = content.split("^XZ")
    parts = [p + "^XZ\n" for p in parts if p.strip()]
    return [parts[i:i + BATCH_SIZE] for i in range(0, len(parts), BATCH_SIZE)]

def send_to_printer(zpl_batches):
    system_name = platform.system().lower()
    if "windows" in system_name:
        import win32print
        hPrinter = win32print.OpenPrinter(PRINTER_NAME)
        for idx, batch in enumerate(zpl_batches[:1]):  # imprime o primeiro lote
            job_name = f"Lote {idx+1}"
            hJob = win32print.StartDocPrinter(hPrinter, 1, (job_name, None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            data = "".join(batch)
            win32print.WritePrinter(hPrinter, data.encode("utf-8"))
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
        win32print.ClosePrinter(hPrinter)
    else:
        messagebox.showerror("Erro", "Este aplicativo é para Windows.")

def select_and_print():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo de etiquetas (.txt ou .zpl)",
        filetypes=(("Arquivos de texto", "*.txt *.zpl"),)
    )
    if not file_path: return
    try:
        with open(file_path, "r", encoding="latin-1") as f: content = f.read()
        if "^XA" not in content:
            messagebox.showerror("Erro", "Arquivo inválido (sem comandos ZPL).")
            return
        batches = split_labels(content)
        send_to_printer(batches)
        messagebox.showinfo("Sucesso", "Primeira página enviada à impressora!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao imprimir: {e}")

root = tk.Tk()
root.title("Impressora Full MDK006")
root.geometry("420x200")

label = tk.Label(root, text="Selecione o arquivo de etiquetas (.txt ou .zpl)",
                 wraplength=350, justify="center", pady=20)
label.pack()
btn = tk.Button(root, text="Selecionar e Imprimir",
                command=select_and_print, height=2, width=25)
btn.pack()
root.mainloop()
