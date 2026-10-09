import os
import sys
import re
import threading
import logging
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import requests

from utils import (
    APP_NAME, VERSAO_ATUAL, GITHUB_REPO, obter_token, COLUNAS_AIRTABLE, 
    limpar_cnpj, HAS_KEYRING, AIRTABLE_BASE_ID, AIRTABLE_TABLE_ID, 
    AIRTABLE_CIDADES_TABLE_ID, AIRTABLE_VIEW_ID
)

try:
    import fitz
    import pytesseract
    from PIL import Image, ImageTk 
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# =====================================================================
# IMPORTAÇÃO DOS MÓDULOS
# =====================================================================
from mod_config import AbaConfig
from mod_nfse_xml import AbaXML
from mod_emissor_nfts import AbaEmissorNFTS
from mod_auditoria_nfts import AbaAuditoriaNFTS
from mod_pgdas import AbaPGDAS

# Módulos de Guias Desmembrados
from mod_das import AbaDAS
from mod_darf import AbaDARF
from mod_iss import AbaISS
from mod_prolabore import AbaProLabore

# O NOVO MÓDULO DE FECHAMENTO!
from mod_fechamento import AbaFechamento
# =====================================================================

if getattr(sys, 'frozen', False): diretorio_base = os.path.dirname(sys.executable)
else: diretorio_base = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(
    filename='robo_execucao.log', level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
)

class HubAutomacaoApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} (v{VERSAO_ATUAL})")
        self.root.geometry("1250x750")
        self.root.configure(bg="#f4f6f9")
        
        try:
            def resource_path(relative_path):
                try: base_path = sys._MEIPASS
                except Exception: base_path = os.path.dirname(os.path.abspath(__file__))
                return os.path.join(base_path, relative_path)
            
            caminho_icone = resource_path('icone.ico')
            self.root.iconbitmap(default=caminho_icone)
            if HAS_OCR: 
                img_icone = ImageTk.PhotoImage(Image.open(caminho_icone))
                self.root.iconphoto(True, img_icone) 
        except Exception as e:
            self.log_backend(f"Ícone visual não carregado: {e}")
        
        self.style = ttk.Style()
        if "clam" in self.style.theme_names(): self.style.theme_use("clam")
        self.style.configure('TFrame', background="#f4f6f9")
        self.style.configure('Treeview', font=('Segoe UI', 10), rowheight=28)
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background="#e0e0e0")
        self.style.configure('Menu.Treeview', font=('Segoe UI', 11), rowheight=35, background="#2c3e50", foreground="white", fieldbackground="#2c3e50")
        self.style.map('Menu.Treeview', background=[('selected', '#34495e')])
        
        self.mapping_cnpj = {}
        self.mapping_ccm = {}

        self.construir_tela_login()

    def log_backend(self, mensagem):
        logging.info(mensagem)
        print(mensagem) 

    def construir_tela_login(self):
        self.frame_login = tk.Frame(self.root, bg="#2c3e50")
        self.frame_login.pack(fill=tk.BOTH, expand=True)
        caixa_centro = tk.Frame(self.frame_login, bg="white", padx=50, pady=50, relief=tk.RAISED, borderwidth=1)
        caixa_centro.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        tk.Label(caixa_centro, text="HUB DE AUTOMAÇÃO", font=("Segoe UI", 20, "bold"), bg="white", fg="#2c3e50").pack(pady=(0, 5))
        tk.Label(caixa_centro, text="Acesso Restrito", font=("Segoe UI", 12), bg="white", fg="#7f8c8d").pack(pady=(0, 20))
        btn_bypass = tk.Button(caixa_centro, text="ACESSAR SISTEMA", command=self.iniciar_sistema, bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=5)
        btn_bypass.pack(fill=tk.X, pady=(20, 0))

    def iniciar_sistema(self):
        self.frame_login.destroy()
        self.construir_layout_principal()
        self.auto_carregar_clientes()
        threading.Thread(target=self.verificar_atualizacao, daemon=True).start()

    def construir_layout_principal(self):
        self.sidebar = tk.Frame(self.root, bg="#2c3e50", width=340)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False) 
        tk.Label(self.sidebar, text=APP_NAME, font=("Segoe UI", 14, "bold"), bg="#2c3e50", fg="white", pady=15).pack(fill=tk.X)
        self.nav_tree = ttk.Treeview(self.sidebar, show="tree", style="Menu.Treeview")
        self.nav_tree.column("#0", width=310, minwidth=310, stretch=True)
        self.nav_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.nav_tree.bind("<<TreeviewSelect>>", self.ao_clicar_no_menu)
        self.montar_itens_do_menu()

        self.content_area = tk.Frame(self.root, bg="#f4f6f9")
        self.content_area.pack(side="left", fill="both", expand=True)

        # ----------------------------------------------------
        # DEFINIÇÃO DOS CONTAINERS (TELAS)
        # ----------------------------------------------------
        self.tela_pdf = tk.Frame(self.content_area, bg="#f4f6f9") 
        self.tela_xml = tk.Frame(self.content_area, bg="#f4f6f9")
        self.tela_nfts = tk.Frame(self.content_area, bg="#f4f6f9") 
        self.tela_auditoria = tk.Frame(self.content_area, bg="#f4f6f9") 
        self.tela_pgdas = tk.Frame(self.content_area, bg="#f4f6f9")
        self.tela_fechamento = tk.Frame(self.content_area, bg="#f4f6f9") # Nova tela do Fechamento
        self.tela_config = tk.Frame(self.content_area, bg="#f4f6f9")
        self.tela_dev = tk.Frame(self.content_area, bg="#f4f6f9") 

        # Sub-abas dentro de Guias
        self.notebook_guias = ttk.Notebook(self.tela_pdf)
        self.notebook_guias.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        frame_das = ttk.Frame(self.notebook_guias)
        frame_darf = ttk.Frame(self.notebook_guias)
        frame_iss = ttk.Frame(self.notebook_guias)
        frame_prolabore = ttk.Frame(self.notebook_guias)

        self.notebook_guias.add(frame_das, text="📄 Guias DAS")
        self.notebook_guias.add(frame_darf, text="📄 Guias DARF")
        self.notebook_guias.add(frame_iss, text="🏛️ Guias ISS")
        self.notebook_guias.add(frame_prolabore, text="💼 Pró-Labore")

        self.aba_das = AbaDAS(frame_das, self)
        self.aba_darf = AbaDARF(frame_darf, self)
        self.aba_iss = AbaISS(frame_iss, self)
        self.aba_prolabore = AbaProLabore(frame_prolabore, self)

        # Instanciando as outras classes isoladas normais
        self.mod_config = AbaConfig(self.tela_config, self)
        self.mod_xml = AbaXML(self.tela_xml, self)
        self.mod_nfts = AbaEmissorNFTS(self.tela_nfts, self)
        self.mod_auditoria = AbaAuditoriaNFTS(self.tela_auditoria, self)
        self.mod_pgdas = AbaPGDAS(self.tela_pgdas, self)
        
        # Conectando a Classe de Fechamento ao frame do Content Area
        self.mod_fechamento = AbaFechamento(self.tela_fechamento, self)

        self.tela_atual = None
        self.mostrar_tela(self.tela_pdf)

    def montar_itens_do_menu(self):
        n_fiscal = self.nav_tree.insert("", "end", text="📂 FISCAL", open=True)
        self.nav_tree.insert(n_fiscal, "end", text=" 📄 Processador de Guias e Recibos", tags=('tela_pdf',))
        self.nav_tree.insert(n_fiscal, "end", text=" 🧾 Processador de NFSe (XML)", tags=('tela_xml',))
        self.nav_tree.insert(n_fiscal, "end", text=" 🤖 Emissor de NFTS (São Paulo)", tags=('tela_nfts',))
        self.nav_tree.insert(n_fiscal, "end", text=" 🔍 Conferência de NFTS (São Paulo)", tags=('tela_auditoria',))
        self.nav_tree.insert(n_fiscal, "end", text=" 🚧 Alteração Pró-labore Domínio", tags=('dev', 'Alteração Pró-labore na Domínio'))
        self.nav_tree.insert(n_fiscal, "end", text=" 🚧 Emissão de CNDs", tags=('dev', 'Emissão de CNDs e Simples Nacional'))

        n_declaracoes = self.nav_tree.insert("", "end", text="📂 DECLARAÇÕES E RELATÓRIOS", open=True)
        self.nav_tree.insert(n_declaracoes, "end", text=" 📄 Declaração PGDAS-D", tags=('tela_pgdas',))
        self.nav_tree.insert(n_declaracoes, "end", text=" 📊 Relatório de Fechamento", tags=('tela_fechamento',)) # <--- Link Módulo Fechamento

        n_contabil = self.nav_tree.insert("", "end", text="📂 CONTÁBIL / FINANCEIRO", open=False)
        self.nav_tree.insert(n_contabil, "end", text=" 🚧 Conciliação Contábil", tags=('dev', 'Robô de Conciliação Contábil'))
        self.nav_tree.insert(n_contabil, "end", text=" 🚧 Importação de OFX Domínio", tags=('dev', 'Importação Automática de OFX'))
        self.nav_tree.insert(n_contabil, "end", text=" 🚧 Conversor PDF > Excel", tags=('dev', 'Conversor de Extratos Bancários'))
        self.nav_tree.insert(n_contabil, "end", text=" 🚧 Conciliador de Cartões", tags=('dev', 'Conciliador de Cartões'))

        n_societario = self.nav_tree.insert("", "end", text="📂 SOCIETÁRIO", open=False)
        self.nav_tree.insert(n_societario, "end", text=" 🚧 Geração de Documentos", tags=('dev', 'Robô de Geração de Documentos (Modelos)'))

        n_atendimento = self.nav_tree.insert("", "end", text="📂 ATENDIMENTO", open=False)
        self.nav_tree.insert(n_atendimento, "end", text=" 🚧 Robô Carteiro (E-mail)", tags=('dev', 'Disparo de Guias por E-mail'))
        self.nav_tree.insert(n_atendimento, "end", text=" 🚧 Alertas de Vencimento", tags=('dev', 'Alertas Automáticos de Vencimentos'))

        n_outros = self.nav_tree.insert("", "end", text="📂 OUTRAS AUTOMAÇÕES", open=False)
        self.nav_tree.insert(n_outros, "end", text=" 🚧 Novos Robôs", tags=('dev', 'Repositório de Outros Robôs'))

        n_sistema = self.nav_tree.insert("", "end", text="📂 SISTEMA", open=True)
        self.nav_tree.insert(n_sistema, "end", text=" ⚙️ Banco de Clientes (Airtable)", tags=('tela_config',))
        self.nav_tree.insert(n_sistema, "end", text=" 🚧 Dashboard & Métricas", tags=('dev', 'Dashboard de Produtividade'))
        self.nav_tree.insert(n_sistema, "end", text=" 🚧 Agendador Automático", tags=('dev', 'Agendador de Tarefas Automáticas'))

    def ao_clicar_no_menu(self, event):
        selecionado = self.nav_tree.selection()
        if not selecionado: return
        tags = self.nav_tree.item(selecionado[0], "tags")
        if not tags: return
        acao = tags[0]
        if acao == 'tela_pdf': self.mostrar_tela(self.tela_pdf)
        elif acao == 'tela_xml': self.mostrar_tela(self.tela_xml)
        elif acao == 'tela_nfts': self.mostrar_tela(self.tela_nfts)
        elif acao == 'tela_auditoria': self.mostrar_tela(self.tela_auditoria)
        elif acao == 'tela_pgdas': self.mostrar_tela(self.tela_pgdas)
        elif acao == 'tela_fechamento': self.mostrar_tela(self.tela_fechamento) # <--- Navegação Fechamento
        elif acao == 'tela_config': self.mostrar_tela(self.tela_config)
        elif acao == 'dev':
            for widget in self.tela_dev.winfo_children(): widget.destroy()
            tk.Label(self.tela_dev, text="🚧", font=("Segoe UI", 55), bg="#f4f6f9").pack(pady=(150, 20))
            tk.Label(self.tela_dev, text=f"Módulo: {tags[1] if len(tags)>1 else 'Em Breve'}", font=("Segoe UI", 22, "bold"), bg="#f4f6f9").pack()
            self.mostrar_tela(self.tela_dev)

    def mostrar_tela(self, nova_tela):
        if self.tela_atual == nova_tela: return
        if self.tela_atual: self.tela_atual.pack_forget()
        nova_tela.pack(fill="both", expand=True)
        self.tela_atual = nova_tela

    # ================== MOTOR AIRTABLE ==================
    def auto_carregar_clientes(self):
        token = obter_token()
        if token:
            if hasattr(self, 'lbl_status_clientes'): self.lbl_status_clientes.config(text="Sincronizando clientes com o Airtable...", fg="#FF9800")
            threading.Thread(target=self.carregar_dados_airtable_thread, args=(token,), daemon=True).start()
        else:
            if hasattr(self, 'lbl_status_clientes'): self.lbl_status_clientes.config(text="Token não configurado. Vá na aba Configurações.", fg="red")

    def carregar_dados_airtable_thread(self, token):
        self.log_backend("Iniciando sincronização com Airtable...")
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        
        url_cidades = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_CIDADES_TABLE_ID}"
        mapping_cidades = {}
        offset_cid = None
        while True:
            params_cid = {}
            if offset_cid: params_cid["offset"] = offset_cid
            try:
                resp_cid = requests.get(url_cidades, headers=headers, params=params_cid)
                if resp_cid.status_code == 200:
                    data_cid = resp_cid.json()
                    for rec in data_cid.get("records", []):
                        flds = rec.get("fields", {})
                        if flds:
                            cidade_nome = flds.get("Name") or flds.get("Nome") or flds.get("Cidade") or list(flds.values())[0]
                            mapping_cidades[rec["id"]] = str(cidade_nome).strip()
                    offset_cid = data_cid.get("offset")
                    if not offset_cid: break
                else: break 
            except: break

        url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_ID}"
        todos_registros = []
        offset = None
        while True:
            params = {}
            if offset: params["offset"] = offset
            if AIRTABLE_VIEW_ID: params["view"] = AIRTABLE_VIEW_ID
            try:
                response = requests.get(url, headers=headers, params=params)
                if response.status_code == 404:
                    if hasattr(self, 'lbl_status_clientes'): self.root.after(0, lambda: self.lbl_status_clientes.config(text="Erro: Tabela não encontrada.", fg="red"))
                    return
                response.raise_for_status()
                data = response.json()
                todos_registros.extend(data.get("records", []))
                offset = data.get("offset")
                if not offset: break
            except:
                if hasattr(self, 'lbl_status_clientes'): self.root.after(0, lambda: self.lbl_status_clientes.config(text="Erro de conexão com Airtable.", fg="red"))
                return
        
        temp_mapping_cnpj = {}
        temp_mapping_ccm = {}
        if hasattr(self, 'tree_clientes'): self.root.after(0, lambda: self.tree_clientes.delete(*self.tree_clientes.get_children()))
        
        for record in todos_registros:
            fields = record.get("fields", {})
            status_raw = fields.get("Status da Empresa", "")
            if isinstance(status_raw, list): status_empresa = ", ".join([str(v) for v in status_raw]).lower()
            else: status_empresa = str(status_raw).lower()
                
            if not re.search(r'\b(ativo|ultimo mês|ultimo mes|último mês|último mes)\b', status_empresa): continue
            
            val_cnpj = fields.get("CNPJ", "")
            if isinstance(val_cnpj, list) and len(val_cnpj) > 0: val_cnpj = val_cnpj[0]
            cnpj_limpo = limpar_cnpj(val_cnpj)
            ccm = str(fields.get("CCM", "")).strip()
            
            dados_empresa = {}
            for col in COLUNAS_AIRTABLE:
                val = fields.get(col, "")
                if col == "Cidade da Empresa" and isinstance(val, list): val = ", ".join([mapping_cidades.get(v, str(v)) for v in val])
                elif isinstance(val, list): val = ", ".join([str(v) for v in val])
                dados_empresa[col] = str(val).strip()
                
            if len(cnpj_limpo) >= 11: temp_mapping_cnpj[cnpj_limpo] = dados_empresa
            if ccm:
                temp_mapping_ccm[ccm] = dados_empresa
                temp_mapping_ccm[ccm.lstrip('0')] = dados_empresa
                
            if hasattr(self, 'tree_clientes'):
                self.root.after(0, lambda d=dados_empresa, c=val_cnpj: self.tree_clientes.insert("", "end", values=(
                    d.get("Razão Social", ""), c, d.get("Status da Empresa", ""), d.get("Tipo de Tributação", ""), f'{d.get("Código", "")} - {d.get("Apelido", "")}')))

        self.mapping_cnpj = temp_mapping_cnpj
        self.mapping_ccm = temp_mapping_ccm
        self.log_backend(f"Sucesso! {len(self.mapping_cnpj)} clientes mapeados.")
        if hasattr(self, 'lbl_status_clientes'): self.root.after(0, lambda: self.lbl_status_clientes.config(text=f"Conectado: {len(self.mapping_cnpj)} clientes mapeados.", fg="#4CAF50"))

    # ================== ATUALIZADOR ==================
    def verificar_atualizacao(self, checagem_manual=False):
        if not getattr(sys, 'frozen', False): return
        url_api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        try:
            response = requests.get(url_api, timeout=5)
            response.raise_for_status()
            dados = response.json()
            versao_nuvem = dados.get("tag_name", "").replace("v", "")
            if versao_nuvem > VERSAO_ATUAL:
                for asset in dados.get("assets", []):
                    if asset["name"].endswith(".exe"): return self.perguntar_se_atualiza(versao_nuvem, asset["browser_download_url"])
            if checagem_manual: messagebox.showinfo("Atualização", "Você já está na versão mais recente!")
        except Exception as e:
            if checagem_manual: messagebox.showwarning("Erro", "Não foi possível verificar atualizações.")

    def perguntar_se_atualiza(self, nova_versao, url_download):
        if messagebox.askyesno("Nova Versão Disponível!", f"A versão {nova_versao} do Robô está disponível.\n\nDeseja atualizar agora?"):
            threading.Thread(target=self.baixar_e_atualizar, args=(url_download,), daemon=True).start()

    def baixar_e_atualizar(self, url_download):
        try:
            exe_atual = sys.executable
            exe_novo = os.path.join(os.path.dirname(exe_atual), "update_temporario.exe")
            script_bat = os.path.join(os.path.dirname(exe_atual), "atualizador.bat")

            resposta = requests.get(url_download, stream=True, allow_redirects=True)
            resposta.raise_for_status()
            with open(exe_novo, 'wb') as f:
                for chunk in resposta.iter_content(chunk_size=8192):
                    if chunk: f.write(chunk)

            bat_codigo = f'@echo off\ntimeout /t 3 /nobreak > NUL\ndel "{exe_atual}"\nren "{exe_novo}" "{os.path.basename(exe_atual)}"\nstart "" "{exe_atual}"\ndel "%~f0"\n'
            with open(script_bat, 'w') as f: f.write(bat_codigo)
                
            subprocess.Popen(script_bat, shell=True)
            os._exit(0) 
        except: pass

if __name__ == '__main__':
    if os.name == 'nt':
        import ctypes
        try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('escritorio.hubautomacao.dominio.100')
        except: pass
    root = tk.Tk()
    app = HubAutomacaoApp(root)
    root.mainloop()