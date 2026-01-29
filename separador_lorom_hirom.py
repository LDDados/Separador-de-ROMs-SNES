import os
import zipfile
import shutil
import tempfile
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Label
from collections import defaultdict

# ------------------------------
# Localizar 7-Zip
# ------------------------------
def localizar_7z():
    possiveis = [
        "7z",
        "7z.exe",
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe"
    ]
    for p in possiveis:
        if shutil.which(p) or os.path.exists(p):
            return p
    return None

CAMINHO_7Z = localizar_7z()

# ------------------------------
# Detectar LoROM / HiROM / ExHiROM (MELHORADO)
# ------------------------------
def detectar_mapa(rom_path):
    try:
        with open(rom_path, "rb") as f:
            data = f.read()

        tamanho = len(data)

        # Detectar header de 512 bytes
        tem_header = (tamanho % 0x8000) == 512
        base = 512 if tem_header else 0

        def score_lorom():
            try:
                off = 0x7FC0 + base
                checksum = data[off + 0x1C] | (data[off + 0x1D] << 8)
                complement = data[off + 0x1E] | (data[off + 0x1F] << 8)
                score = 0
                if checksum ^ complement == 0xFFFF:
                    score += 2
                if data[off + 0x15] in (0x20, 0x21, 0x30):
                    score += 1
                if data[off + 0x18] < 0x10:
                    score += 1
                return score
            except:
                return 0

        def score_hirom():
            try:
                off = 0xFFC0 + base
                checksum = data[off + 0x1C] | (data[off + 0x1D] << 8)
                complement = data[off + 0x1E] | (data[off + 0x1F] << 8)
                score = 0
                if checksum ^ complement == 0xFFFF:
                    score += 2
                if data[off + 0x15] in (0x21, 0x31):
                    score += 1
                if data[off + 0x18] < 0x10:
                    score += 1
                return score
            except:
                return 0

        def score_exhirom():
            try:
                if tamanho < 0x410000:
                    return 0
                off = 0x40FFC0 + base
                checksum = data[off + 0x1C] | (data[off + 0x1D] << 8)
                complement = data[off + 0x1E] | (data[off + 0x1F] << 8)
                score = 0
                if checksum ^ complement == 0xFFFF:
                    score += 2
                if data[off + 0x15] in (0x25, 0x35):
                    score += 2
                return score
            except:
                return 0

        s_lo = score_lorom()
        s_hi = score_hirom()
        s_ex = score_exhirom()

        if s_ex >= max(s_lo, s_hi) and s_ex >= 3:
            return "ExHiRom"
        elif s_lo > s_hi and s_lo >= 2:
            return "LoRom"
        elif s_hi >= 2:
            return "HiRom"

    except:
        pass

    return "Desconhecido"

# ------------------------------
# Extrair arquivos
# ------------------------------
def extrair_arquivo(arquivo, temp_dir):
    extraidos = []
    nome = arquivo.lower()

    try:
        pasta_temp = tempfile.mkdtemp(dir=temp_dir)

        if nome.endswith(".zip"):
            with zipfile.ZipFile(arquivo, "r") as z:
                z.extractall(pasta_temp)

        elif nome.endswith((".rar", ".7z")):
            if not CAMINHO_7Z:
                messagebox.showerror(
                    "7-Zip não encontrado",
                    "7-Zip não encontrado. Instale ou adicione ao PATH, "
                    "ou extraia a ROM .7z/.rar manualmente!"
                )
                return []

            subprocess.run(
                [CAMINHO_7Z, "x", arquivo, f"-o{pasta_temp}", "-y"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )

        else:
            destino = os.path.join(pasta_temp, os.path.basename(arquivo))
            shutil.copy(arquivo, destino)

        for root, _, files in os.walk(pasta_temp):
            for f in files:
                path = os.path.join(root, f)
                if os.path.getsize(path) > 0:
                    extraidos.append(path)

    except Exception as e:
        print("Erro ao extrair:", arquivo, e)
        return []

    return extraidos

# ------------------------------
# Formatar tamanho
# ------------------------------
def nome_tamanho(bytes_):
    kb = bytes_ / 1024
    mb = kb / 1024
    if mb >= 1:
        return f"{mb:.1f}MB" if not mb.is_integer() else f"{int(mb)}MB"
    else:
        return f"{int(kb)}KB" if kb >= 1 else "1KB"

# ------------------------------
# Interface
# ------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Separador SNES - LoROM / HiROM / ExHiROM")
        self.root.geometry("520x260")
        self.root.resizable(False, False)

        self.arquivos = []
        self.destino = ""

        tk.Button(root, text="Selecionar arquivos", command=self.selecionar_roms).pack(pady=5)
        tk.Button(root, text="Selecionar destino", command=self.selecionar_destino).pack(pady=5)
        tk.Button(root, text="Separar ROMs", command=self.iniciar_thread).pack(pady=8)

        self.status = Label(root, text="Aguardando...")
        self.status.pack()

        self.contador = Label(root, text="0 / 0")
        self.contador.pack()

        self.progress = Progressbar(root, length=480, mode="determinate")
        self.progress.pack(pady=10)

    def selecionar_roms(self):
        self.arquivos = filedialog.askopenfilenames(
            title="Selecionar arquivos",
            filetypes=[("Todos os arquivos", "*.*")]
        )
        messagebox.showinfo("Arquivos", f"{len(self.arquivos)} arquivo(s) selecionado(s)")

    def selecionar_destino(self):
        self.destino = filedialog.askdirectory(title="Selecionar pasta destino")

    def iniciar_thread(self):
        if not self.arquivos or not self.destino:
            messagebox.showerror("Erro", "Selecione arquivos e destino")
            return
        threading.Thread(target=self.processar, daemon=True).start()

    def processar(self):
        base_dirs = {
            "LoRom": os.path.join(self.destino, "LoRom"),
            "HiRom": os.path.join(self.destino, "HiRom"),
            "ExHiRom": os.path.join(self.destino, "ExHiRom"),
            "Desconhecido": os.path.join(self.destino, "Desconhecido")
        }

        for d in base_dirs.values():
            os.makedirs(d, exist_ok=True)

        arquivos_por_tipo = defaultdict(list)

        with tempfile.TemporaryDirectory() as temp:
            for arq in self.arquivos:
                extraidos = extrair_arquivo(arq, temp)
                for r in extraidos:
                    mapa = detectar_mapa(r)
                    arquivos_por_tipo[mapa].append(r)

            total = sum(len(v) for v in arquivos_por_tipo.values())
            self.root.after(0, self.progress.config, {"maximum": total})

            i = 0
            for mapa, arquivos in arquivos_por_tipo.items():
                base = base_dirs.get(mapa, base_dirs["Desconhecido"])

                tamanho_grupos = defaultdict(list)
                for arq in arquivos:
                    tamanho_grupos[os.path.getsize(arq)].append(arq)

                for tamanho, grup in tamanho_grupos.items():
                    if len(grup) > 1:
                        pasta_tam = os.path.join(base, nome_tamanho(tamanho))
                        os.makedirs(pasta_tam, exist_ok=True)
                    else:
                        pasta_tam = base

                    for arq in grup:
                        i += 1
                        nome = os.path.basename(arq)
                        self.root.after(0, self.status.config, {"text": f"Analisando: {nome} → {mapa}"})
                        self.root.after(0, self.contador.config, {"text": f"{i:02d} / {total}"})
                        self.root.after(0, self.progress.config, {"value": i})

                        shutil.copy(arq, pasta_tam)

        self.root.after(0, self.status.config, {"text": "Concluído!"})
        self.root.after(0, messagebox.showinfo, "Finalizado", "Separação concluída!")

# ------------------------------
# Main
# ------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()