import os, re, glob, shutil, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from utils import limpar_cnpj, limpar_nome_arquivo, COLUNAS_AIRTABLE

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

class AbaProLabore:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.pasta_pdfs = ""
        self.dados_pre_pdf = []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Processador de Recibos Pró-Labore", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        self.var_apenas_leitura = tk.BooleanVar(value=False)
        tk.Checkbutton(frame_top, text="Fazer apenas Leitura (Não Renomear)", variable=self.var_apenas_leitura, bg="white", font=("Segoe UI", 9, "italic"), fg="#d35400").pack(side=tk.RIGHT, padx=5)

        tk.Label(frame_top, text="Selecione o diretório com os Recibos (.pdf)", font=("Segoe UI", 9), bg="white").pack(anchor="w", pady=(5, 5))
        frame_input = tk.Frame(frame_top, bg="white")
        frame_input.pack(fill=tk.X)
        self.btn_pasta_pdf = tk.Button(frame_input, text="📁 Procurar Pasta", command=self.selecionar_pasta_pdf, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10)
        self.btn_pasta_pdf.pack(side=tk.LEFT)
        self.lbl_pasta_pdf = tk.Label(frame_input, text="Nenhuma pasta selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_pasta_pdf.pack(side=tk.LEFT, padx=15)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        self.tree_pdf = ttk.Treeview(frame_table, show="headings", height=10)
        colunas = ("Arquivo Original", "Sócio", "Competência", "Pró-Labore", "Status de Ação", "Status")
        self.tree_pdf["columns"] = colunas
        for col in colunas:
            self.tree_pdf.heading(col, text=col)
            self.tree_pdf.column(col, minwidth=80, width=200 if col == "Sócio" else 120)
        self.tree_pdf.pack(fill=tk.BOTH, expand=True, pady=5)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_ler_pdf = tk.Button(frame_btns, text="🔍 INICIAR LEITURA DE PRÓ-LABORE", command=self.iniciar_leitura_pdf, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
        self.btn_ler_pdf.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)
        self.btn_exec_pdf = tk.Button(frame_btns, text="✅ APROVAR E EXECUTAR ROTINA", command=self.executar_renomeacao_pdf, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_exec_pdf.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def selecionar_pasta_pdf(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta de PDFs")
        if pasta:
            self.pasta_pdfs = pasta
            self.lbl_pasta_pdf.config(text=pasta, fg="#2c3e50")

    def formatar_competencia(self, comp):
        if not comp: return ""
        comp = comp.upper().replace(" ", "")
        meses_texto = {
            'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'MARCO': '03', 'ABRIL': '04', 'MAIO': '05', 
            'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08', 'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12',
            'JAN': '01', 'FEV': '02', 'MAR': '03', 'ABR': '04', 'MAI': '05', 'JUN': '06', 'JUL': '07', 'AGO': '08', 'SET': '09', 'OUT': '10', 'NOV': '11', 'DEZ': '12'
        }
        for m, num in meses_texto.items():
            if m in comp:
                comp = comp.replace(m, "")
                comp = comp.replace("DE", "")
                return f"{num}/{comp[-4:]}"
                
        parts = comp.split('/')
        if len(parts) == 2:
            return f"{parts[0].zfill(2)}/{parts[1]}"
            
        m_dig = re.search(r'(\d{1,2})[^\d]*(\d{4})', comp)
        if m_dig: return f"{m_dig.group(1).zfill(2)}/{m_dig.group(2)}"
        return comp

    def extrair_dados_prolabore(self, pdf_path):
        resultados = []
        if not HAS_PDFPLUMBER: return resultados
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text(layout=False) or ""
                    if not texto.strip(): continue 
                    info = {"CNPJ": None, "Nome_Socio": None, "Competencia": None, "Valor_Prolabore": "0,00", "Valor_INSS": "0,00", "Valor_IRRF": "0,00", "Texto_Bruto": texto}
                    
                    match_cnpj = re.search(r'CNPJ:\s*([\d\.\-\/]+)', texto, re.IGNORECASE)
                    if match_cnpj: info["CNPJ"] = match_cnpj.group(1)
                    
                    match_comp = re.search(r'\b(JANEIRO|FEVEREIRO|MARÇO|MARCO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s+(?:de\s+)?(20\d{2})\b', texto, re.IGNORECASE)
                    if match_comp: info["Competencia"] = self.formatar_competencia(match_comp.group(0))
                    
                    linhas = [l.strip() for l in texto.split('\n') if l.strip()]
                    for i, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        if "NOME DO FUNCION" in linha_up:
                            nome_part = re.sub(r'(?i)\b(C[OÓ]DIGO|NOME DO FUNCION[AÁ]RIO|CBO|DEPARTAMENTO|FILIAL)\b', '', linha)
                            nome_part = re.sub(r'\d+', '', nome_part).strip() 
                            if len(nome_part) > 3: info["Nome_Socio"] = re.sub(r'\s+', ' ', nome_part).strip()
                            elif i + 1 < len(linhas): info["Nome_Socio"] = re.sub(r'\s+', ' ', re.sub(r'\d+', '', linhas[i+1]).strip()).strip()
                        if "PRO-LABORE" in linha_up or "PRÓ-LABORE" in linha_up or "PRO LABORE" in linha_up:
                            nums = re.findall(r'\b\d+(?:\.\d{3})*,\d{2}\b', linha)
                            if nums: info["Valor_Prolabore"] = nums[-1] 
                        if "INSS" in linha_up and not re.search(r'MATR|SAL|BASE', linha_up):
                            nums = re.findall(r'\b\d+(?:\.\d{3})*,\d{2}\b', linha)
                            if nums: info["Valor_INSS"] = nums[-1]
                        if "IRRF" in linha_up and not re.search(r'BASE|FAIXA', linha_up):
                            nums = re.findall(r'\b\d+(?:\.\d{3})*,\d{2}\b', linha)
                            if nums: info["Valor_IRRF"] = nums[-1]
                    if info["CNPJ"] and info["Nome_Socio"]: resultados.append(info)
        except Exception as e:
            self.app.log_backend(f"[ERRO] Pro-labore: {e}")
        return resultados

    def iniciar_leitura_pdf(self):
        if not self.pasta_pdfs: return messagebox.showwarning("Atenção", "Selecione a pasta de PDFs.")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "Aguarde carregar os clientes do Airtable.")
        self.btn_ler_pdf.config(state=tk.DISABLED, text="LENDO...")
        threading.Thread(target=self.ler_pdfs_thread, daemon=True).start()

    def ler_pdfs_thread(self):
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children()))
        self.dados_pre_pdf = []
        pdfs = glob.glob(os.path.join(self.pasta_pdfs, '*.pdf'))
        
        for filepath in pdfs:
            filename = os.path.basename(filepath)
            res_prolabore = self.extrair_dados_prolabore(filepath)
            
            if not res_prolabore:
                self.dados_pre_pdf.append({"filepath": filepath, "filename": filename, "novo_nome": "", "status": "Falha na Leitura", "linhas_excel": []})
                self.app.root.after(0, lambda f=filename: self.tree_pdf.insert("", "end", values=(f, "-", "-", "-", "-", "Falha")))
                continue
            
            cnpj_limpo = limpar_cnpj(res_prolabore[0]["CNPJ"])
            dados_empresa = self.app.mapping_cnpj.get(cnpj_limpo, {})
            if dados_empresa:
                apelido = limpar_nome_arquivo(dados_empresa.get("Apelido", ""))
                novo_nome = f"{apelido}_Recibo_pro_labore.pdf"
                status = "OK"
            else: apelido, novo_nome, status = "", "", "CNPJ não mapeado"
            
            msg_acao = "Pronto p/ Leitura" if self.var_apenas_leitura.get() else "Pronto p/ Renomear"
            if status != "OK": msg_acao = "Ação Cancelada"

            lista_excel = []
            for socio in res_prolabore:
                base_excel = {col: dados_empresa.get(col, "") for col in COLUNAS_AIRTABLE}
                base_excel.update({"Nome do Sócio": socio["Nome_Socio"], "Competência": socio["Competencia"], "Pró-Labore (R$)": socio["Valor_Prolabore"], "INSS (R$)": socio["Valor_INSS"], "IRRF (R$)": socio["Valor_IRRF"], "Arquivo Original": filename, "Arquivo Renomeado": novo_nome if (novo_nome and not self.var_apenas_leitura.get()) else "APENAS LEITURA"})
                lista_excel.append(base_excel)
                self.app.root.after(0, lambda f=filename, s=socio["Nome_Socio"], c=socio["Competencia"], p=socio["Valor_Prolabore"], act=msg_acao, st=status: self.tree_pdf.insert("", "end", values=(f, s, c, p, act, st)))
                        
            self.dados_pre_pdf.append({"filepath": filepath, "filename": filename, "novo_nome": novo_nome, "status": status, "linhas_excel": lista_excel})

        self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 INICIAR LEITURA DE PRÓ-LABORE"))
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(state=tk.NORMAL))

    def executar_renomeacao_pdf(self):
        msg = "Deseja renomear os arquivos e gerar o Excel?" if not self.var_apenas_leitura.get() else "Modo Leitura: Deseja gerar o Excel com os dados?"
        if not messagebox.askyesno("Confirmar", msg): return
        self.btn_exec_pdf.config(state=tk.DISABLED, text="PROCESSANDO...")
        threading.Thread(target=self.renomear_pdfs_thread, daemon=True).start()

    def renomear_pdfs_thread(self):
        sucesso, dados_excel = 0, []
        for item in self.dados_pre_pdf:
            if item["status"] == "OK" and item["novo_nome"]:
                if not self.var_apenas_leitura.get():
                    try:
                        novo_caminho = os.path.join(self.pasta_pdfs, item["novo_nome"])
                        if item["filepath"] != novo_caminho and not os.path.exists(novo_caminho): os.rename(item["filepath"], novo_caminho)
                        sucesso += 1
                    except Exception: pass
                else: sucesso += 1 
            dados_excel.extend(item["linhas_excel"])
        if dados_excel:
            pd.DataFrame(dados_excel).to_excel(os.path.join(self.pasta_pdfs, "Relatorio_Prolabore.xlsx"), index=False)
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", f"{sucesso} PDFs processados!"))
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children())) 
        self.dados_pre_pdf = []
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(text="✅ APROVAR E EXECUTAR ROTINA", state=tk.DISABLED))