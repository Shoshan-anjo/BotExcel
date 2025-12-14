# gui/excel_manager_gui.py

import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from application.scheduler_uc import SchedulerUseCase
from infrastructure.config_loader import ConfigLoader

# Requiere instalar: pip install tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    messagebox.showerror("tkinterdnd2 no instalado", "Instala tkinterdnd2: pip install tkinterdnd2")
    raise

CONFIG_PATH = "config/excels.json"

class ExcelManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BotExcel - Gestión de Excels y Horarios")
        self.root.geometry("800x550")

        self.config_loader = ConfigLoader()
        self.scheduler_uc = SchedulerUseCase()
        self.excels = self.load_excels()

        # Sincronizar con horarios guardados
        self.sync_with_schedule()

        # Lista de Excels
        self.listbox = tk.Listbox(root, width=120)
        self.listbox.pack(pady=10)
        # Habilitar drag & drop
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.drop_files)

        # Botones de acciones
        btn_frame = tk.Frame(root)
        btn_frame.pack()
        tk.Button(btn_frame, text="Agregar Excel", command=self.add_excel).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Eliminar", command=self.remove_excel).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Asignar Backup", command=self.assign_backup).grid(row=0, column=2, padx=5)
        tk.Button(btn_frame, text="Guardar Config", command=self.save_excels).grid(row=0, column=3, padx=5)
        tk.Button(btn_frame, text="Activar/Desactivar Horario", command=self.toggle_horario).grid(row=0, column=4, padx=5)

        # Entrada de horario
        horario_frame = tk.Frame(root)
        horario_frame.pack(pady=10)
        tk.Label(horario_frame, text="Horario (HH:MM)").grid(row=0, column=0)
        self.horario_entry = tk.Entry(horario_frame, width=10)
        self.horario_entry.grid(row=0, column=1, padx=5)
        tk.Button(horario_frame, text="Asignar Horario", command=self.set_horario).grid(row=0, column=2, padx=5)

        self.refresh_list()

    # --------------------------
    # Manejar archivos arrastrados
    # --------------------------
    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        added = 0
        for f in files:
            if f.lower().endswith((".xlsx", ".xlsm")) and not any(e['path'] == f for e in self.excels):
                self.excels.append({"path": f, "backup": "", "horario": "", "activo": True})
                added += 1
        if added > 0:
            self.refresh_list()
            messagebox.showinfo("Archivos agregados", f"{added} archivos Excel agregados mediante drag & drop.")

    # --------------------------
    # Cargar Excels desde JSON
    # --------------------------
    def load_excels(self):
        if not os.path.exists(CONFIG_PATH):
            return []
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("excels", [])

    # --------------------------
    # Guardar Excels en JSON
    # --------------------------
    def save_excels(self):
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"excels": self.excels}, f, indent=4, ensure_ascii=False)
        messagebox.showinfo("Guardado", "Configuración guardada correctamente.")

    # --------------------------
    # Sincronizar Excel con horarios
    # --------------------------
    def sync_with_schedule(self):
        for e in self.excels:
            job = self.scheduler_uc.get_job(e["path"])
            if job:
                e["horario"] = job.get("horario", "")
                e["backup"] = job.get("backup_path", "")
                e["activo"] = job.get("activo", True)
            else:
                e.setdefault("horario", "")
                e.setdefault("backup", "")
                e.setdefault("activo", True)

    # --------------------------
    # Agregar Excel (por botón)
    # --------------------------
    def add_excel(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel Files", "*.xlsx *.xlsm")])
        if path and not any(e['path'] == path for e in self.excels):
            self.excels.append({"path": path, "backup": "", "horario": "", "activo": True})
            self.refresh_list()

    # --------------------------
    # Eliminar Excel
    # --------------------------
    def remove_excel(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        excel_path = self.excels[idx]["path"]
        # Eliminar también de scheduler
        try:
            self.scheduler_uc.remove_job(excel_path)
        except Exception:
            pass
        del self.excels[idx]
        self.refresh_list()

    # --------------------------
    # Asignar Backup
    # --------------------------
    def assign_backup(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("Seleccionar", "Selecciona primero un Excel de la lista.")
            return
        idx = selected[0]
        path = filedialog.askopenfilename(title="Seleccionar archivo Backup", filetypes=[("Excel Files", "*.xlsx *.xlsm")])
        if path:
            self.excels[idx]["backup"] = path
            excel_path = self.excels[idx]["path"]
            job = self.scheduler_uc.get_job(excel_path)
            if job:
                job["backup_path"] = path
                self.scheduler_uc._save_jobs()
            self.refresh_list()
            messagebox.showinfo("Backup asignado", f"Backup {path} asignado al Excel seleccionado.")

    # --------------------------
    # Asignar Horario
    # --------------------------
    def set_horario(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("Seleccionar", "Primero selecciona un Excel de la lista.")
            return
        idx = selected[0]
        horario = self.horario_entry.get().strip()
        if not horario:
            messagebox.showwarning("Horario vacío", "Ingresa un horario en formato HH:MM.")
            return
        try:
            h, m = map(int, horario.split(":"))
            if not (0 <= h <= 23 and 0 <= m <= 59):
                raise ValueError
        except Exception:
            messagebox.showerror("Error", f"Horario inválido: {horario}")
            return

        self.excels[idx]["horario"] = horario
        excel_path = self.excels[idx]["path"]
        backup = self.excels[idx].get("backup")
        # Guardar en SchedulerUseCase
        try:
            job = self.scheduler_uc.get_job(excel_path)
            if job:
                job["horario"] = horario
                job["backup_path"] = backup
                job["activo"] = True
                self.scheduler_uc._save_jobs()
            else:
                self.scheduler_uc.add_job(excel_path, horario, backup)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el horario: {str(e)}")
            return

        self.refresh_list()
        messagebox.showinfo("Horario asignado", f"Horario {horario} asignado a {excel_path}")

    # --------------------------
    # Activar / Desactivar Horario
    # --------------------------
    def toggle_horario(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("Seleccionar", "Primero selecciona un Excel de la lista.")
            return
        idx = selected[0]
        excel_path = self.excels[idx]["path"]
        job = self.scheduler_uc.get_job(excel_path)
        if not job:
            messagebox.showwarning("No programado", "Este Excel aún no tiene horario asignado.")
            return
        job["activo"] = not job.get("activo", True)
        self.scheduler_uc._save_jobs()
        self.excels[idx]["activo"] = job["activo"]
        status = "ACTIVO" if job["activo"] else "INACTIVO"
        self.refresh_list()
        messagebox.showinfo("Estado cambiado", f"El Excel {excel_path} ahora está {status}.")

    # --------------------------
    # Refrescar Listbox
    # --------------------------
    def refresh_list(self):
        self.listbox.delete(0, tk.END)
        for e in self.excels:
            excel_path = e["path"]
            backup = e.get("backup", "")
            horario = e.get("horario", "")
            activo = "ACTIVO" if e.get("activo", True) else "INACTIVO"
            display = f"{excel_path} | Backup: {backup} | Horario: {horario} | {activo}"
            self.listbox.insert(tk.END, display)


if __name__ == "__main__":
    root = TkinterDnD.Tk()  # <--- reemplazamos Tk() por TkinterDnD.Tk()
    ExcelManagerGUI(root)
    root.mainloop()
