import tkinter as tk
from tkinter import ttk, messagebox
import socket
import platform
import json
import wmi
import requests

try:
    import psutil
except ImportError:
    psutil = None

# --- CONFIGURATION ---
USER_ENDPOINT = "http://192.168.201.39:3001/addUsuario"
PC_ENDPOINT = "http://192.168.201.39:3001/addPC"

# --- DATA GATHERING ---
def gather_pc_info():
    """
    Gathers detailed hardware and software information from the computer.
    """
    try:
        c = wmi.WMI()
    except wmi.x_wmi:
        return {"error": "Could not connect to WMI. Please run as administrator."}

    info = {}

    # Basic Info
    info['nome'] = socket.gethostname()
    info['tipo_OS'] = platform.platform()

    # CPU (Processor)
    try:
        cpu_info = c.Win32_Processor()[0]
        info['processador'] = cpu_info.Name.strip()
    except IndexError:
        info['processador'] = platform.processor() # Fallback

    # Motherboard
    try:
        board_info = c.Win32_BaseBoard()[0]
        info['placa_mae'] = f"{board_info.Manufacturer} {board_info.Product}"
    except IndexError:
        info['placa_mae'] = "Não disponível"

    # RAM (Memory)
    info['memorias'] = []
    try:
        for stick in c.Win32_PhysicalMemory():
            info['memorias'].append(
                f"{stick.Manufacturer} {int(stick.Capacity) // (1024**3)}GB {stick.Speed}MHz"
            )
    except Exception:
        # Fallback to psutil for total RAM if WMI fails
        if psutil:
             mem = psutil.virtual_memory()
             info['memorias'].append(f"Total: {mem.total // (1024**3)} GB")

    # Storage (Disks)
    info['discos'] = []
    try:
        for disk in c.Win32_DiskDrive():
            disk_type = "SSD" if "SSD" in disk.Model.upper() else "HDD"
            info['discos'].append({
                'tipo': disk_type,
                'info': f"{disk.Model.strip()} ({int(disk.Size) // (1024**3)} GB)"
            })
    except Exception:
        info['discos'].append({'tipo': 'N/A', 'info': 'Não disponível'})


    # OS License Key (Note: This is not always reliable)
    try:
        os_key = c.Win32_SoftwareLicensingService.GetOA3xOriginalProductKey()
        info['licenca_OS'] = os_key[0] if os_key[0] else "Não encontrada"
    except Exception:
        info['licenca_OS'] = "Não disponível (requer admin)"

    return info


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PC Info Uploader")
        self.resizable(False, False)

        notebook = ttk.Notebook(self)
        self.user_frame = ttk.Frame(notebook)
        self.pc_frame = ttk.Frame(notebook)
        notebook.add(self.user_frame, text="Usuário")
        notebook.add(self.pc_frame, text="PC")
        notebook.pack(padx=10, pady=10)

        self.create_user_form()
        self.create_pc_form()

    def create_user_form(self):
        labels = ["Nome", "Nome do PC", "Monitor 1", "Monitor 2", "Mesa", "Licença 1", "Licença 2", "Licença 3"]
        self.user_entries = {}
        for i, label in enumerate(labels):
            ttk.Label(self.user_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2, padx=5)
            entry = ttk.Entry(self.user_frame, width=40)
            entry.grid(row=i, column=1, pady=2, padx=5)
            self.user_entries[label] = entry

        self.user_entries["Nome do PC"].insert(0, socket.gethostname())
        submit = ttk.Button(self.user_frame, text="Enviar", command=self.submit_user)
        submit.grid(row=len(labels), column=0, columnspan=2, pady=10)

    def create_pc_form(self):
        labels = ["Nome", "Número", "Placa mãe", "Processador", "Memória 1", "Memória 2", "HDD 1", "HDD 2", "SSD 1", "SSD 2", "Tipo OS", "Licença OS", "Ramal"]
        self.pc_entries = {}
        for i, label in enumerate(labels):
            ttk.Label(self.pc_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2, padx=5)
            entry = ttk.Entry(self.pc_frame, width=40)
            entry.grid(row=i, column=1, pady=2, padx=5)
            self.pc_entries[label] = entry

        # --- POPULATE FORM WITH GATHERED INFO ---
        info = gather_pc_info()

        # Check for WMI errors
        if "error" in info:
            messagebox.showwarning("WMI Error", info["error"])
            return

        self.pc_entries["Nome"].insert(0, info.get("nome", ""))
        self.pc_entries["Processador"].insert(0, info.get("processador", ""))
        self.pc_entries["Placa mãe"].insert(0, info.get("placa_mae", ""))
        self.pc_entries["Tipo OS"].insert(0, info.get("tipo_OS", ""))
        self.pc_entries["Licença OS"].insert(0, info.get("licenca_OS", ""))

        # Populate memory slots
        memorias = info.get('memorias', [])
        if len(memorias) > 0:
            self.pc_entries["Memória 1"].insert(0, memorias[0])
        if len(memorias) > 1:
            self.pc_entries["Memória 2"].insert(0, memorias[1])

        # Populate storage slots
        discos = info.get('discos', [])
        ssds = [d['info'] for d in discos if d['tipo'] == 'SSD']
        hdds = [d['info'] for d in discos if d['tipo'] == 'HDD']

        if len(ssds) > 0: self.pc_entries["SSD 1"].insert(0, ssds[0])
        if len(ssds) > 1: self.pc_entries["SSD 2"].insert(0, ssds[1])
        if len(hdds) > 0: self.pc_entries["HDD 1"].insert(0, hdds[0])
        if len(hdds) > 1: self.pc_entries["HDD 2"].insert(0, hdds[1])

        submit = ttk.Button(self.pc_frame, text="Enviar", command=self.submit_pc)
        submit.grid(row=len(labels), column=0, columnspan=2, pady=10)

    def submit_user(self):
        data = {key.replace(" ", "_").lower(): entry.get() for key, entry in self.user_entries.items()}
        data['nome_do_pc'] = data.pop('nome_do_pc') # Adjust key if needed
        self.send_request(USER_ENDPOINT, data)

    def submit_pc(self):
        data = {key.replace(" ", "_").lower(): entry.get() for key, entry in self.pc_entries.items()}
        self.send_request(PC_ENDPOINT, data)

    def send_request(self, url, data):
        headers = {'Content-Type': 'application/json'}
        try:
            # Filter out empty values before sending
            payload = {k: v for k, v in data.items() if v}
            response = requests.post(url, data=json.dumps(payload), headers=headers)
            response.raise_for_status()
            messagebox.showinfo("Sucesso", f"Dados enviados com sucesso!\nResposta: {response.text}")
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Erro de Rede", f"Falha ao enviar dados: {e}")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()