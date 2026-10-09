import tkinter as tk
from tkinter import ttk, messagebox
from utils import APP_NAME, obter_token, obter_certificado_padrao, HAS_KEYRING

try:
    import keyring
except ImportError:
    pass

class AbaConfig:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Configurações e Integrações", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        tk.Label(frame_top, text="Segurança e Autenticação (Cofre Local do Windows)", font=("Segoe UI", 12, "bold"), bg="white", fg="#333").pack(anchor="w", pady=(0, 10))
        
        frame_token = tk.Frame(frame_top, bg="white")
        frame_token.pack(fill=tk.X, pady=2)
        tk.Label(frame_token, text="Chave Token API (Airtable):", bg="white", width=25, anchor="w", font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.ent_token = ttk.Entry(frame_token, width=55, show="*")
        self.ent_token.pack(side=tk.LEFT, padx=5)
        if obter_token(): self.ent_token.insert(0, "********************************")

        frame_cert = tk.Frame(frame_top, bg="white")
        frame_cert.pack(fill=tk.X, pady=2)
        tk.Label(frame_cert, text="Certificado Digital Padrão:", bg="white", width=25, anchor="w", font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.ent_cert = ttk.Entry(frame_cert, width=55)
        self.ent_cert.pack(side=tk.LEFT, padx=5)
        cert_salvo = obter_certificado_padrao()
        if cert_salvo: self.ent_cert.insert(0, cert_salvo)

        frame_btns_conf = tk.Frame(frame_top, bg="white")
        frame_btns_conf.pack(fill=tk.X, pady=(10, 0))
        tk.Button(frame_btns_conf, text="Salvar Cofre", command=self.salvar_configuracoes, bg="#2c3e50", fg="white", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=15).pack(side=tk.LEFT)
        tk.Button(frame_btns_conf, text="Atualizar Sistema Github", command=lambda: self.app.verificar_atualizacao(checagem_manual=True), bg="#FF9800", fg="white", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=15).pack(side=tk.RIGHT)

        frame_bot = tk.Frame(self.parent, bg="#f4f6f9")
        frame_bot.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        frame_header_clientes = tk.Frame(frame_bot, bg="#f4f6f9")
        frame_header_clientes.pack(fill=tk.X, pady=(10, 5))
        
        tk.Label(frame_header_clientes, text="Base de Clientes Mapeada", font=("Segoe UI", 12, "bold"), bg="#f4f6f9", fg="#333").pack(side=tk.LEFT)
        self.app.lbl_status_clientes = tk.Label(frame_header_clientes, text="Aguardando conexão...", font=("Segoe UI", 10), bg="#f4f6f9", fg="gray")
        self.app.lbl_status_clientes.pack(side=tk.RIGHT)

        colunas_cli = ("Razão Social", "CNPJ", "Status", "Regime", "Diretório (Cód-Apelido)")
        self.app.tree_clientes = ttk.Treeview(frame_bot, columns=colunas_cli, show="headings", height=12)
        for col in colunas_cli:
            self.app.tree_clientes.heading(col, text=col)
            self.app.tree_clientes.column(col, minwidth=100, width=150)
        
        self.app.tree_clientes.column("Razão Social", width=250)
        self.app.tree_clientes.pack(fill=tk.BOTH, expand=True)
        tk.Button(frame_bot, text="↻ Forçar Sincronização Airtable", command=self.app.auto_carregar_clientes, bg="#008CBA", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=5).pack(fill=tk.X, pady=10)

    def salvar_configuracoes(self):
        novo_token = self.ent_token.get().strip()
        novo_cert = self.ent_cert.get().strip()
        if HAS_KEYRING:
            try:
                if novo_token and novo_token != "********************************":
                    keyring.set_password(APP_NAME, "AirtableToken", novo_token)
                keyring.set_password(APP_NAME, "CertificadoPadrão", novo_cert)
                messagebox.showinfo("Segurança", "Configurações armazenadas no cofre criptografado do Windows com sucesso!")
                if novo_token and novo_token != "********************************":
                    self.app.auto_carregar_clientes()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar no Windows:\n{e}")
        else:
            messagebox.showerror("Erro", "Biblioteca 'keyring' não instalada.")