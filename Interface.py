import os
import subprocess
import tkinter as tk
from pathlib import Path
from ordenador import pdfs_to_single_docx
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import ttk, filedialog, messagebox

# Lista para almacenar los archivos PDF arrastrados
pdf_files = []

def on_drop(event):
    files = root.tk.splitlist(event.data)
    for file in files:
        if file.endswith('.pdf'):
            pdf_files.append(file)
            listbox.insert(tk.END, file)  # Añade el archivo a la lista mostrada en la interfaz

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    entry_output_file.delete(0, tk.END)
    entry_output_file.insert(0, file_path)
    output_option.set("specify")

def remove_selected_file():
    selected_indices = listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("Advertencia", "Seleccione un archivo para eliminar.")
        return
    for index in selected_indices:
        pdf_files.pop(index)
        listbox.delete(index)

def remove_all_files():
    pdf_files.clear()
    listbox.delete(0, tk.END)

def exit_program():
    root.quit()

def start_conversion():
    if not pdf_files:
        messagebox.showerror("Error", "Debe arrastrar archivos PDF.")
        return

    # Determina el camino del archivo de salida
    if output_option.get() == "specify":
        docx_path = entry_output_file.get()
    else:
        downloads_folder = str(Path.home() / "Downloads")
        docx_path = os.path.join(downloads_folder, "Recorador.docx")

    crop_dimensions = (
        int(entry_crop_x1.get()),
        int(entry_crop_y1.get()),
        int(entry_crop_x2.get()),
        int(entry_crop_y2.get())
    )
    images_per_row = int(entry_images_per_row.get())

    if not pdf_files:
        messagebox.showerror("Error", "Debe arrastrar archivos PDF")
    elif not docx_path:
        messagebox.showerror("Error", "Debe seleccionar un archivo de salida.")
        return
    
    pdfs_to_single_docx(pdf_files, docx_path, crop_dimensions, images_per_row)
    messagebox.showinfo("Éxito", "El archivo se ha creado correctamente.")
    print("El documento se guardó correctamente en la ruta: ", docx_path)

    # Pregunta si quiere abrir el archivo
    if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo recién creado?"):
        try:
            subprocess.Popen([docx_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

# Inicializa la ventana principal de la interfaz gráfica
root = TkinterDnD.Tk()
# Recolector y ordenador de comprobantes
root.title("Recorador")

frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

# Listbox para mostrar los archivos PDF arrastrados
listbox = tk.Listbox(frame, width=50, height=10)
listbox.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
listbox.drop_target_register(DND_FILES)
listbox.dnd_bind('<<Drop>>', on_drop)

ttk.Button(frame, text="Eliminar archivos seleccionados", command=remove_selected_file).grid(row=0, column=2, padx=5, pady=5)
ttk.Button(frame, text="Eliminar todos los archivos", command=remove_all_files).grid(row=0, column=3, padx=5, pady=5)

# Opción de salida
output_option = tk.StringVar(value="specify")

ttk.Radiobutton(frame, text="Especificar archivo de salida", variable=output_option, value="specify").grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=5)
ttk.Radiobutton(frame, text="Guardar en Descargas como 'Recorador.docx'", variable=output_option, value="default").grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=5)


tk.Label(frame, text="Archivo DOCX de salida:").grid(row=5, column=0, sticky="e")
entry_output_file = tk.Entry(frame, width=50)
entry_output_file.grid(row=5, column=1, padx=5, pady=5)
ttk.Button(frame, text="Seleccionar", command=select_output_file).grid(row=5, column=2, padx=5, pady=5)

ttk.Label(frame, text="Dimensiones de recorte (x1, y1, x2, y2):").grid(row=2, column=0, sticky="e")
entry_crop_x1 = ttk.Entry(frame, width=5)
entry_crop_x1.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=5)
entry_crop_x1.insert(0, "150")
entry_crop_y1 = ttk.Entry(frame, width=5)
entry_crop_y1.grid(row=2, column=1, padx=(55, 0), pady=5)
entry_crop_y1.insert(0, "150")
entry_crop_x2 = ttk.Entry(frame, width=5)
entry_crop_x2.grid(row=2, column=1, padx=(105, 0), pady=5)
entry_crop_x2.insert(0, "850")
entry_crop_y2 = ttk.Entry(frame, width=5)
entry_crop_y2.grid(row=2, column=1, padx=(155, 0), pady=5)
entry_crop_y2.insert(0, "1200")

ttk.Label(frame, text="Imágenes por fila:").grid(row=3, column=0, sticky="e")
entry_images_per_row = ttk.Entry(frame, width=5)
entry_images_per_row.grid(row=3, column=1, sticky="w", padx=(5, 0), pady=5)
entry_images_per_row.insert(0, "4")

ttk.Button(frame, text="Convertir", command=start_conversion).grid(row=7, column=0, columnspan=3, pady=10)
ttk.Button(frame, text="Cerrar programa", command=exit_program).grid(row=7, column=3, pady=10)

root.mainloop()