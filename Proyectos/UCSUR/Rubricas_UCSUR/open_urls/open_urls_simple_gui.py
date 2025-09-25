# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import webbrowser
import time
import os

APP_TITLE = "Abrir URLs por Carrera (3 preguntas)"
URL_COLS = ["Home", "URL /outcomes", "URL /assignments", "URL /rubrics", "URL /content_migrations"]
BROWSERS = ["Chrome", "Firefox"]

def get_browser(name: str):
    name = (name or "").lower()
    try:
        if name == "chrome":
            return webbrowser.get("chrome")
        if name == "firefox":
            return webbrowser.get("firefox")
    except Exception:
        pass
    # Intentos por ruta en Windows
    try:
        if name == "chrome":
            for p in [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]:
                if os.path.exists(p):
                    webbrowser.register("chrome_win", None, webbrowser.BackgroundBrowser(p))
                    return webbrowser.get("chrome_win")
        if name == "firefox":
            for p in [
                r"C:\Program Files\Mozilla Firefox\firefox.exe",
                r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
            ]:
                if os.path.exists(p):
                    webbrowser.register("firefox_win", None, webbrowser.BackgroundBrowser(p))
                    return webbrowser.get("firefox_win")
    except Exception:
        pass
    # Fallback: navegador por defecto
    try:
        return webbrowser.get()
    except Exception:
        return None

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("700x360")
        self.df = None
        self.file = None

        self.carrera = tk.StringVar()
        self.tipo_url = tk.StringVar(value=URL_COLS[0])
        self.navegador = tk.StringVar(value="Chrome")

        self.build_ui()

    def build_ui(self):
        pad = {"padx":10, "pady":8}

        frm_file = ttk.LabelFrame(self, text="Excel de origen")
        frm_file.pack(fill="x", **pad)

        self.txt_file = ttk.Entry(frm_file, width=70)
        self.txt_file.grid(row=0, column=0, padx=8, pady=6, sticky="we")
        ttk.Button(frm_file, text="Seleccionar Excel...", command=self.pick_excel).grid(row=0, column=1, padx=8, pady=6)

        frm = ttk.LabelFrame(self, text="Preguntas")
        frm.pack(fill="x", **pad)

        ttk.Label(frm, text="(1) ¿Qué carrera?").grid(row=0, column=0, sticky="w")
        self.cbo_carrera = ttk.Combobox(frm, textvariable=self.carrera, state="readonly", width=50)
        self.cbo_carrera.grid(row=0, column=1, sticky="w")

        ttk.Label(frm, text="(2) ¿Qué URL?").grid(row=1, column=0, sticky="w")
        self.cbo_url = ttk.Combobox(frm, textvariable=self.tipo_url, state="readonly", values=URL_COLS, width=50)
        self.cbo_url.grid(row=1, column=1, sticky="w")

        ttk.Label(frm, text="(3) ¿Qué navegador?").grid(row=2, column=0, sticky="w")
        self.cbo_nav = ttk.Combobox(frm, textvariable=self.navegador, state="readonly", values=["Chrome","Firefox"], width=50)
        self.cbo_nav.grid(row=2, column=1, sticky="w")

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", **pad)

        self.btn_open = ttk.Button(frm_btn, text="Abrir URLs", command=self.open_urls, state="disabled")
        self.btn_open.pack(side="right")

        self.status = tk.StringVar(value="Seleccione el Excel.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x", padx=10, pady=(0,10))

    def pick_excel(self):
        path = filedialog.askopenfilename(title="Seleccione Excel", filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        self.txt_file.delete(0, tk.END)
        self.txt_file.insert(0, path)
        self.load_df(path)

    def load_df(self, path):
        try:
            xls = pd.ExcelFile(path)
            sheet = "Sheet1" if "Sheet1" in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(path, sheet_name=sheet)
            # Validar columnas
            needed = ["Carrera del curso"] + URL_COLS
            missing = [c for c in needed if c not in df.columns]
            if missing:
                raise ValueError(f"Faltan columnas: {missing}")
            self.df = df.copy()
            carreras = sorted(self.df["Carrera del curso"].dropna().astype(str).unique())
            self.cbo_carrera["values"] = carreras
            if carreras:
                self.cbo_carrera.set(carreras[0])
                self.carrera.set(carreras[0])
            self.btn_open["state"] = "normal"
            self.status.set("Excel cargado. Responda las 3 preguntas y presione 'Abrir URLs'.")
        except Exception as e:
            self.df = None
            self.btn_open["state"] = "disabled"
            self.status.set("Error al cargar Excel.")
            messagebox.showerror("Error", str(e))

    def open_urls(self):
        if self.df is None:
            return
        car = self.carrera.get()
        col = self.tipo_url.get()
        nav = self.navegador.get()

        dff = self.df[self.df["Carrera del curso"].astype(str) == str(car)]
        urls = dff[col].dropna().astype(str)
        urls = [u.strip() for u in urls if u.startswith(("http://","https://"))]

        if not urls:
            messagebox.showinfo("Sin URLs", "No hay URLs para los filtros elegidos.")
            return

        prev = "\n".join(urls[:5])
        extra = "" if len(urls) <= 5 else f"\n... y {len(urls)-5} más."
        if not messagebox.askyesno("Confirmar", f"Se abrirán {len(urls)} URL(s) en {nav} ({col}).\n\nVista previa:\n{prev}{extra}\n\n¿Continuar?"):
            return

        ctrl = get_browser(nav) or webbrowser
        opened = 0
        for u in urls:
            try:
                ctrl.open_new_tab(u)
                opened += 1
                time.sleep(0.25)
            except Exception:
                try:
                    webbrowser.open_new_tab(u)
                    opened += 1
                    time.sleep(0.25)
                except Exception:
                    pass
        messagebox.showinfo("Listo", f"Se intentaron abrir {opened}/{len(urls)} URL(s).")

if __name__ == "__main__":
    app = App()
    try:
        from tkinter import ttk
        style = ttk.Style()
        for theme in ["vista","clam","default"]:
            try:
                style.theme_use(theme)
                break
            except Exception:
                continue
    except Exception:
        pass
    app.mainloop()
