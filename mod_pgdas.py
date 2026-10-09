import os, re, glob, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from utils import limpar_cnpj, limpar_nome_arquivo, COLUNAS_AIRTABLE, validar_tributacao

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

class AbaPGDAS:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.pasta_pdfs = ""
        self.dados_pre_pdf = []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Processador de Declarações (PGDAS-D)", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        tk.Label(frame_top, text="Selecione o diretório com as Declarações PGDAS (.pdf)", font=("Segoe UI", 9), bg="white").pack(anchor="w", pady=(5, 5))
        
        frame_input = tk.Frame(frame_top, bg="white")
        frame_input.pack(fill=tk.X)
        self.btn_pasta_pdf = tk.Button(frame_input, text="📁 Procurar Pasta", command=self.selecionar_pasta_pdf, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10)
        self.btn_pasta_pdf.pack(side=tk.LEFT)
        self.lbl_pasta_pdf = tk.Label(frame_input, text="Nenhuma pasta selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_pasta_pdf.pack(side=tk.LEFT, padx=15)
        
        self.var_apenas_leitura = tk.BooleanVar(value=False)
        tk.Checkbutton(frame_top, text="Fazer apenas Leitura (Não Renomear)", variable=self.var_apenas_leitura, bg="white", font=("Segoe UI", 9, "italic"), fg="#d35400").pack(side=tk.RIGHT, padx=5)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        
        self.tree_pdf = ttk.Treeview(frame_table, show="headings", height=10)
        
        colunas = ("Arquivo Original", "Empresa", "Competência", "Débito Exigível", "Status de Ação", "Status")
        self.tree_pdf["columns"] = colunas
        for col in colunas:
            self.tree_pdf.heading(col, text=col)
            w = 200 if col == "Empresa" else 120
            self.tree_pdf.column(col, minwidth=80, width=w)
            
        self.tree_pdf.pack(fill=tk.BOTH, expand=True, pady=5)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_ler_pdf = tk.Button(frame_btns, text="🔍 INICIAR LEITURA DE PGDAS", command=self.iniciar_leitura_pdf, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
        self.btn_ler_pdf.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)
        self.btn_exec_pdf = tk.Button(frame_btns, text="✅ APROVAR E EXECUTAR ROTINA", command=self.executar_renomeacao_pdf, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_exec_pdf.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def selecionar_pasta_pdf(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta de PGDAS")
        if pasta:
            self.pasta_pdfs = pasta
            self.lbl_pasta_pdf.config(text=pasta, fg="#2c3e50")

    def formatar_valor(self, val):
        if not val: return ""
        val_clean = re.sub(r'[^\d.,]', '', str(val).strip())
        if len(val_clean) >= 3 and val_clean[-3] in ['.', ',']:
            inteiro = val_clean[:-3].replace('.', '').replace(',', '')
            decimal = val_clean[-2:]
            try: return f"{float(f'{inteiro}.{decimal}'):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            except: pass
        return val

    def extrair_dados_pgdas(self, pdf_path):
        info = {
            "CNPJ": None, "Valor": None, "Competencia": None, 
            "Tipo_Guia": "Declaração PGDAS-D", "Sufixo_Arquivo": "PGDAS",
            "Regra_Usada": "Declaração PGDAS-D"
        }
        
        if not HAS_PDFPLUMBER:
            return info
            
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text(layout=False) or ""
        except Exception as e:
            self.app.log_backend(f"[ERRO] Leitura PGDAS {os.path.basename(pdf_path)}: {e}")
            return info
            
        if not text.strip():
            return info

        # CNPJ
        m_cnpj = re.search(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', text)
        if m_cnpj: 
            info["CNPJ"] = m_cnpj.group(0)
            
        # Competência
        m_comp = re.search(r'Per[íi]odo de Apura[çc][ãa]o.*?(?:a\s+\d{2}/| )(\d{2}/\d{4})', text, re.IGNORECASE | re.DOTALL)
        if not m_comp:
            m_comp = re.search(r'Per[íi]odo de Apura[çc][ãa]o[^\d]*(\d{2}/\d{4})', text, re.IGNORECASE)
        if m_comp: 
            info["Competencia"] = m_comp.group(1)

        # Valor (Débito Exigível - Leitura aprimorada para evitar erro com valor do IRPJ)
        for match in re.finditer(r'Total do D[é]bito Exig[íi]vel(.*)', text, re.IGNORECASE | re.DOTALL):
            sub_text = match.group(1)[:500] 
            for line in sub_text.split('\n'):
                nums = re.findall(r'(?<!\d)\d+(?:[.,]\d{3})*[.,]\d{2}(?!\d)', line)
                if nums:
                    info["Valor"] = self.formatar_valor(nums[-1])
                    break
            if info["Valor"]: break
            
        # Fallback para o valor Declarado
        if not info["Valor"]:
            for match in re.finditer(r'Total do D[é]bito Declarado(.*)', text, re.IGNORECASE | re.DOTALL):
                sub_text = match.group(1)[:500]
                for line in sub_text.split('\n'):
                    nums = re.findall(r'(?<!\d)\d+(?:[.,]\d{3})*[.,]\d{2}(?!\d)', line)
                    if nums:
                        info["Valor"] = self.formatar_valor(nums[-1])
                        break
                if info["Valor"]: break
            
        return info

    def iniciar_leitura_pdf(self):
        if not self.pasta_pdfs: return messagebox.showwarning("Atenção", "Selecione a pasta dos PGDAS.")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "Aguarde carregar os clientes do Airtable.")
        self.btn_ler_pdf.config(state=tk.DISABLED, text="LENDO...")
        threading.Thread(target=self.ler_pdfs_thread, daemon=True).start()

    def ler_pdfs_thread(self):
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children()))
        self.dados_pre_pdf = []
        pdfs = glob.glob(os.path.join(self.pasta_pdfs, '*.pdf'))
        
        if not pdfs: 
            return self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 INICIAR LEITURA DE PGDAS"))
        
        for filepath in pdfs:
            filename = os.path.basename(filepath)
            info = self.extrair_dados_pgdas(filepath)
            
            cnpj_valido = None
            dados_empresa = {}
            tipo_guia = info["Tipo_Guia"]
            sufixo = info["Sufixo_Arquivo"]
            status, apelido, novo_nome = "OK", "", ""

            if info.get("CNPJ"):
                cl = limpar_cnpj(info["CNPJ"])
                if cl in self.app.mapping_cnpj:
                    d_temp = self.app.mapping_cnpj[cl]
                    if validar_tributacao(tipo_guia, d_temp.get("Tipo de Tributação", "")):
                        cnpj_valido = info["CNPJ"]
                        dados_empresa = d_temp

            if cnpj_valido:
                apelido = limpar_nome_arquivo(dados_empresa.get("Apelido", ""))
                novo_nome = f"{apelido}_{sufixo}.pdf"
            else: 
                status = "Não mapeado/Trib. errada"

            msg_acao = "Pronto p/ Leitura" if self.var_apenas_leitura.get() else "Pronto p/ Renomear"
            if status != "OK": msg_acao = "Ação Cancelada"

            base_excel = {col: dados_empresa.get(col, "") for col in COLUNAS_AIRTABLE}
            base_excel.update({
                "Arquivo Original": filename, "Arquivo Renomeado": novo_nome if (novo_nome and not self.var_apenas_leitura.get()) else "APENAS LEITURA",
                "Identificação do Robô": info.get("Regra_Usada", ""), "Tipo de Guia": tipo_guia,
                "Mês de Competência": info.get("Competencia", ""), "Data de Vencimento": "N/A", "Valor (R$)": info.get("Valor", "")
            })

            self.dados_pre_pdf.append({"filepath": filepath, "filename": filename, "novo_nome": novo_nome, "status": status, "linhas_excel": [base_excel]})
            self.app.root.after(0, lambda f=filename, a=apelido, c=info.get("Competencia",""), v=info.get("Valor",""), act=msg_acao, s=status: 
                            self.tree_pdf.insert("", "end", values=(f, a, c, v, act, s)))

        self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 INICIAR LEITURA DE PGDAS"))
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(state=tk.NORMAL))

    def executar_renomeacao_pdf(self):
        msg = "Deseja renomear os arquivos e gerar o Excel?" if not self.var_apenas_leitura.get() else "Modo Leitura: Deseja gerar o Excel com os dados?"
        if not messagebox.askyesno("Confirmar", msg): return
        self.btn_exec_pdf.config(state=tk.DISABLED, text="PROCESSANDO...")
        threading.Thread(target=self.renomear_pdfs_thread, daemon=True).start()

    def renomear_pdfs_thread(self):
        sucesso = 0
        dados_excel = []
        apenas_leitura = self.var_apenas_leitura.get()
        for item in self.dados_pre_pdf:
            if item["status"] == "OK" and item["novo_nome"]:
                if not apenas_leitura:
                    try:
                        novo_caminho = os.path.join(self.pasta_pdfs, item["novo_nome"])
                        if item["filepath"] != novo_caminho and not os.path.exists(novo_caminho):
                            os.rename(item["filepath"], novo_caminho)
                        sucesso += 1
                    except Exception: pass
                else: sucesso += 1 
            dados_excel.extend(item["linhas_excel"])
        if dados_excel:
            df = pd.DataFrame(dados_excel)
            df.to_excel(os.path.join(self.pasta_pdfs, "Relatorio_Declaracoes_PGDAS.xlsx"), index=False)
        
        acao_feita = "Arquivos processados" if apenas_leitura else "PGDAS renomeados"
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", f"{sucesso} {acao_feita} com sucesso! Relatório gerado."))
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children())) 
        self.dados_pre_pdf = []
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(text="✅ APROVAR E EXECUTAR ROTINA", state=tk.DISABLED))