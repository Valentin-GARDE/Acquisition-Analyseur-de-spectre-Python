import tkinter as tk
from tkinter import ttk, messagebox
import wmi # Pour la détection des périphériques USB
import sys # Pour la gestion de la fermeture propre du programme
import numpy as np
import time
from math import degrees
from Acq_Fonctions import (
    connecter_appareil_usb,
    send_usb_command,
    connect_to_device,
    send_scpi_command,
    receive_data,
    save_data_to_file,
    configure_spectrum
)
import matplotlib.pyplot as plt
from matplotlib.ticker import EngFormatter  # Pour formater automatiquement les unités
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class AppMesure:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface de Mesure")
        try:
            icon = tk.PhotoImage(file="Logo_UGE.png")
            self.root.iconphoto(True, icon)
        except Exception:
            pass
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        # Taille initiale de la fenêtre
        self.root.geometry("1200x668")

        # Connection state
        self.connexion_etablie = False
        self.device_connection = None

        # Variables
        self.conn_type = tk.StringVar(value="USB")
        self.ip_address = tk.StringVar()
        self.usb_device = tk.StringVar()
        self.param1 = tk.DoubleVar(value=1.0)
        self.param2 = tk.DoubleVar(value=5.0)
        self.start_freq_var = tk.DoubleVar(value=1e6)
        self.span_var = tk.DoubleVar(value=1e6)
        self.ref_level_var = tk.DoubleVar(value=0.0)
        self.rbw_var = tk.DoubleVar(value=3e+6)
        self.vbw_var = tk.DoubleVar(value=1e+6)
        self.status_var = tk.StringVar(value="Prêt.")
        self.running = False

        self.build_interface()
        self.setup_plot()
        self.setup_status_bar()

    def build_interface(self):
        # GPS display
        self.gps_label = tk.Label(self.root, text="GPS: --", bg="lightgray", anchor="w")
        self.gps_label.pack(fill=tk.X)

        # Control panel
        panel = tk.Frame(self.root, bg="#2f2a85", width=300)
        panel.pack_propagate(False)
        panel.pack(side=tk.LEFT, fill=tk.Y)

        # Connection
        ttk.Label(panel, text="Connexion:", background="#2f2a85", foreground="white").pack(pady=5)
        ttk.Combobox(panel, textvariable=self.conn_type, values=["USB","LAN"], state="readonly").pack(pady=2)
        self.ip_entry = ttk.Entry(panel, textvariable=self.ip_address)
        self.ip_entry.pack(pady=2)
        self.usb_combo = ttk.Combobox(panel, textvariable=self.usb_device, state="readonly")
        # Initialiser la liste et sélectionner l'appareil par défaut
        devices = self.get_usb_devices()
        self.usb_combo["values"] = devices
        # Sélectionner le périphérique contenant 'USB Test and Measurement Device' si présent
        default = next((d for d in devices if "USB Test and Measurement Device" in d), None)
        if default:
            self.usb_combo.set(default)
        else:
            if devices:
                self.usb_combo.current(0)
        self.usb_combo.pack(pady=5)
        tk.Button(panel, text="Rafraîchir USB", command=self.refresh_usb_list, bg="#2f2a85", fg="white").pack(pady=5)
        tk.Button(panel, text="Connecter", command=self.connexion, bg="#2f2a85", fg="white").pack(pady=5)

        # Device config
        ttk.Label(panel, text="Configuration Appareil:", background="#2f2a85", foreground="white").pack(pady=5)
        self._add_entry(panel, "Start Freq (Hz):", self.start_freq_var)
        self._add_entry(panel, "Span (Hz):", self.span_var)
        self._add_entry(panel, "Ref Level (dBm):", self.ref_level_var)
        self._add_entry(panel, "RBW (Hz):", self.rbw_var)
        self._add_entry(panel, "VBW (Hz):", self.vbw_var)
        tk.Button(panel, text="Config Device", command=self.config_device, bg="#2f2a85", fg="white").pack(pady=5)

        # Acquisition
        ttk.Label(panel, text="Acquisition:", background="#2f2a85", foreground="white").pack(pady=5)
        tk.Button(panel, text="Acquérir", command=self.acquerir_donnees, bg="#2f2a85", fg="white").pack(pady=2)
        self.start_button = tk.Button(panel, text="Start", command=self.toggle_acquisition_loop, bg="#2f2a85", fg="white")
        self.start_button.pack(pady=2)

        self.conn_type.trace("w", self.toggle_connexion_fields)
        self.toggle_connexion_fields()

    def _add_entry(self, parent, label, var):
        ttk.Label(parent, text=label, background="#2f2a85", foreground="white").pack(pady=(5,0))
        ttk.Entry(parent, textvariable=var).pack(pady=2)

    def setup_plot(self):
        frame = tk.Frame(self.root, bg="lightgray")
        frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.fig, self.ax = plt.subplots()
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def setup_status_bar(self):
        bar = tk.Frame(self.root, relief=tk.SUNKEN, bd=1, bg="#2f2a85")
        bar.place(relx=0, rely=1, relwidth=1, anchor="sw")
        tk.Label(bar, textvariable=self.status_var, bg="#2f2a85", fg="white", anchor="w").pack(fill=tk.X)

    def get_usb_devices(self):
        c = wmi.WMI()
        return [d.Name for d in c.Win32_PnPEntity() if d.PNPDeviceID and "USB" in d.PNPDeviceID] or ["Aucun détecté"]

    def refresh_usb_list(self):
        vals = self.get_usb_devices()
        self.usb_combo["values"] = vals
        # Sélectionner le périphérique par défaut si présent
        default = next((d for d in vals if "USB Test and Measurement Device" in d), None)
        if default:
            self.usb_combo.set(default)
        elif vals:
            self.usb_combo.current(0)

    def toggle_connexion_fields(self, *args):
        if self.conn_type.get() == "LAN":
            self.ip_entry.configure(state="normal")
            self.usb_combo.configure(state="disabled")
        else:
            self.ip_entry.configure(state="disabled")
            self.usb_combo.configure(state="readonly")
            self.refresh_usb_list()

    def connexion(self):
        self.status_var.set("Connexion en cours...")
        self.root.update_idletasks()
        if self.conn_type.get() == "LAN":
            ip = self.ip_address.get()
            if not ip:
                messagebox.showerror("Erreur","Adresse IP manquante.")
                return
            self.device_connection = connect_to_device(ip, 5025)
        else:
            self.device_connection = connecter_appareil_usb()
            if self.device_connection:
                self.device_connection.timeout = 5000
        if not self.device_connection:
            self.status_var.set("Connexion échouée.")
            messagebox.showerror("Erreur","Connexion échouée.")
            return
        self.connexion_etablie = True
        try:
            idn = self.device_connection.query("*IDN?")
            self.status_var.set(f"Connecté : {idn.strip()}")
        except:
            self.status_var.set("Connecté (sans IDN)")

    def initialiser(self):
        if not self.connexion_etablie:
            return
        cmd = f"INIT:P1 {self.param1.get()};P2 {self.param2.get()}"
        if self.conn_type.get() == "LAN":
            send_scpi_command(self.device_connection, cmd)
        else:
            send_usb_command(self.device_connection, cmd)
        self.status_var.set("Paramètres initialisés.")

    def config_device(self):
        if not self.connexion_etablie:
            return
        try:
            configure_spectrum(
                self.device_connection,
                start_freq_hz=self.start_freq_var.get(),
                span_hz=self.span_var.get(),
                rbw_hz=self.rbw_var.get(),
                vbw_hz=self.vbw_var.get(),
                input_atten_db=self.ref_level_var.get(),
                continuous=False
            )
            self.status_var.set("Configuration appliquée.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de configurer : {e}")
            self.status_var.set("Erreur config")

    def acquerir_donnees(self):
        if not self.connexion_etablie:
            return
        self.status_var.set("Acquisition en cours...")
        self.root.update()

        # Helper pour parser les nombres avec suffixe
        def parse_num(s):
            units = {'':1, 'K':1e3, 'M':1e6, 'G':1e9}
            import re
            m = re.match(r"\s*([0-9.]+)\s*([kKmMgG]?)", s)
            if not m:
                raise ValueError(f"Format invalide: {s!r}")
            val, suf = m.groups()
            return float(val) * units[suf.upper()]

        try:
            # Initialisation du sweep
            self.device_connection.write(":FORMat:DATA ASCii")
            self.device_connection.write(":INIT:IMM")
            time.sleep(0.7)
            raw = self.device_connection.query(":TRACe:DATA? 1", delay=1)
            if raw.startswith("#"):
                hd = int(raw[1]); cnt = int(raw[2:2+hd])
                vals = [float(v) for v in raw[2+hd:2+hd+cnt].split(",") if v]
            else:
                raise ValueError("Header SCPI invalide")

            # Lecture du préambule pour les fréquences
            pre = self.device_connection.query(":TRACe:PREamble? 1")
            sf_str = pre.split("START_FREQ=")[1].split(",")[0]   # e.g. "0.000000 M"
            ef_str = pre.split("STOP_FREQ=")[1].split(",")[0]
            sf = parse_num(sf_str)
            ef = parse_num(ef_str)

            np_str = pre.split("UI_DATA_POINTS=")[1].split(",")[0]
            num = int(float(np_str)) if np_str.strip() else 0
            freqs = np.linspace(sf, ef, num)

            # Affichage
            self.ax.clear()
            self.ax.plot(freqs, vals)
            # Formatter automatique de l'axe des X (Hz, kHz, MHz, ...)
            self.ax.xaxis.set_major_formatter(EngFormatter(unit='Hz'))
            # Titres des axes
            self.ax.set_xlabel('Fréquence')
            self.ax.set_ylabel('Amplitude (dBm)')
            self.canvas.draw()

            try:
                resp = self.device_connection.query(":FETCh:GPS?")
                parts = resp.strip().split(",")
                if parts[0] == "GOOD FIX":
                    dt, lat, lon = parts[1], float(parts[2]), float(parts[3])
                    self.gps_label.config(text=f"GPS: {dt} | Lat:{degrees(lat):.6f}° Lon:{degrees(lon):.6f}°")
                else:
                    self.gps_label.config(text="GPS: NO FIX")
            except:
                self.gps_label.config(text="GPS: Erreur")

        except Exception as e:
            self.status_var.set("Erreur acquisition")
            messagebox.showerror("Erreur", f"Erreur: {e}")


    def toggle_acquisition_loop(self):
        """
        Lance ou arrête la boucle d'acquisition continue.
        """
        if not self.connexion_etablie:
            return
        # Démarrage de la boucle si arrêtée
        if not self.running:
            self.running = True
            self.start_button.config(text="Stop")
            # Appel initial de la boucle d'acquisition
            self.loop_acquisition()
        # Arrêt de la boucle
        else:
            self.running = False
            self.start_button.config(text="Start")

    def loop_acquisition(self):
        """
        Effectue des acquisitions répétées tant que self.running est True.
        """
        if not self.running:
            return
        self.acquerir_donnees()
        # Planifie l'appel suivant après 1000 ms
        self.root.after(1000, self.loop_acquisition)

    def generate_unique_filename(self):
        import os
        os.makedirs("results", exist_ok=True)
        i = 1
        while True:
            f = os.path.join("results", f"measure_{i}.txt")
            if not os.path.exists(f):
                return f
            i += 1

    def on_close(self):
        self.running = False
        self.root.destroy()
        sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = AppMesure(root)
    root.mainloop()
