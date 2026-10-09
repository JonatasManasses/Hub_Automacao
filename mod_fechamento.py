import os, re, glob, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from utils import limpar_cnpj, limpar_nome_arquivo

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import fitz  # PyMuPDF, usado para fatiar o PDF página a página
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


class AbaFechamento:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.pasta_pdfs = ""
        self.dados_pre_pdf = []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Relatório de Fechamento (Separador de Lotes)", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        tk.Label(frame_top, text="Selecione o diretório com os Relatórios em Lote (.pdf)", font=("Segoe UI", 9), bg="white").pack(anchor="w", pady=(5, 5))
        
        frame_input = tk.Frame(frame_top, bg="white")
        frame_input.pack(fill=tk.X)
        self.btn_pasta_pdf = tk.Button(frame_input, text="📁 Procurar Pasta", command=self.selecionar_pasta_pdf, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10)
        self.btn_pasta_pdf.pack(side=tk.LEFT)
        self.lbl_pasta_pdf = tk.Label(frame_input, text="Nenhuma pasta selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_pasta_pdf.pack(side=tk.LEFT, padx=15)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        
        self.tree_pdf = ttk.Treeview(frame_table, show="headings", height=10)
        
        colunas = ("Arquivo / Página", "Empresa", "Competência", "Regime Identificado", "Status de Ação")
        self.tree_pdf["columns"] = colunas
        for col in colunas:
            self.tree_pdf.heading(col, text=col)
            w = 250 if col == "Empresa" else 150
            self.tree_pdf.column(col, minwidth=80, width=w)
            
        self.tree_pdf.pack(fill=tk.BOTH, expand=True, pady=5)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_ler_pdf = tk.Button(frame_btns, text="🔍 LER E CLASSIFICAR PÁGINAS", command=self.iniciar_leitura_pdf, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
        self.btn_ler_pdf.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)
        self.btn_exec_pdf = tk.Button(frame_btns, text="✅ FATIAR, SALVAR E ORGANIZAR EM PASTAS", command=self.executar_renomeacao_pdf, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_exec_pdf.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def selecionar_pasta_pdf(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta de Relatórios")
        if pasta:
            self.pasta_pdfs = pasta
            self.lbl_pasta_pdf.config(text=pasta, fg="#2c3e50")

    def extrair_dados_por_pagina(self, pdf_path):
        resultados_paginas = []
        if not HAS_PDFPLUMBER:
            messagebox.showerror("Erro", "A biblioteca pdfplumber não está instalada.")
            return resultados_paginas
            
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for num_pagina, page in enumerate(pdf.pages):
                    texto = page.extract_text(layout=False) or ""
                    if not texto.strip():
                        continue

                    info = {
                        "page_num": num_pagina,
                        "CNPJ": None,
                        "Competencia": None,
                        "Regime": "Desconhecido"
                    }

                    # Extrai o CNPJ da página atual
                    m_cnpj = re.search(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', texto)
                    if m_cnpj: 
                        info["CNPJ"] = m_cnpj.group(0)
                        
                    # Extrai a competência
                    m_comp = re.search(r'M[êe]s de Refer[êe]ncia:\s*(\d{2}/\d{4})', texto, re.IGNORECASE)
                    if m_comp: 
                        info["Competencia"] = m_comp.group(1)

                    # Classifica o Regime da página atual
                    texto_minusculo = texto.lower()
                    if "%total impostos sobre faturamento (das)" in texto_minusculo:
                        info["Regime"] = "Simples Nacional"
                    elif "%total impostos sobre faturamento (brasil)" in texto_minusculo:
                        info["Regime"] = "Lucro Presumido"
                        
                    # Só adiciona à lista se encontrou uma empresa (CNPJ) nessa página
                    if info["CNPJ"]:
                        resultados_paginas.append(info)
                        
        except Exception as e:
            self.app.log_backend(f"[ERRO] Leitura Fechamento {os.path.basename(pdf_path)}: {e}")
            
        return resultados_paginas

    def iniciar_leitura_pdf(self):
        if not self.pasta_pdfs: return messagebox.showwarning("Atenção", "Selecione a pasta dos Relatórios.")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "Aguarde carregar os clientes do Airtable.")
        if not HAS_FITZ: return messagebox.showerror("Erro", "A biblioteca PyMuPDF (fitz) é necessária para cortar os PDFs.")
        
        self.btn_ler_pdf.config(state=tk.DISABLED, text="LENDO PÁGINAS...")
        threading.Thread(target=self.ler_pdfs_thread, daemon=True).start()

    def ler_pdfs_thread(self):
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children()))
        self.dados_pre_pdf = []
        pdfs = glob.glob(os.path.join(self.pasta_pdfs, '*.pdf'))
        
        if not pdfs: 
            return self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 LER E CLASSIFICAR PÁGINAS"))
        
        for filepath in pdfs:
            filename = os.path.basename(filepath)
            
            # Vai retornar uma lista de dicionários (um por cada página que tem CNPJ)
            paginas_extraidas = self.extrair_dados_por_pagina(filepath)
            
            for page_info in paginas_extraidas:
                cnpj_valido = None
                dados_empresa = {}
                regime = page_info["Regime"]
                status, apelido = "OK", ""

                cl = limpar_cnpj(page_info["CNPJ"])
                if cl in self.app.mapping_cnpj:
                    dados_empresa = self.app.mapping_cnpj[cl]
                    cnpj_valido = page_info["CNPJ"]

                if cnpj_valido:
                    apelido = limpar_nome_arquivo(dados_empresa.get("Apelido", ""))
                else: 
                    status = "Erro: CNPJ Não Mapeado"
                    
                # Trava Airtable x PDF
                if cnpj_valido and status == "OK":
                    trib_airtable = str(dados_empresa.get("Tipo de Tributação", "")).upper()
                    
                    if regime == "Simples Nacional" and "SIMPLES NACIONAL" not in trib_airtable and "SIMPLES NACIONAL ANEXO V" not in trib_airtable:
                        status = f"Bloqueado: PDF é SN, Airtable é {trib_airtable}"
                    elif regime == "Lucro Presumido" and "LUCRO PRESUMIDO" not in trib_airtable and "LUCRO REAL" not in trib_airtable:
                        status = f"Bloqueado: PDF é LP, Airtable é {trib_airtable}"
                    elif regime == "Desconhecido":
                        status = "Erro: Frase identificadora ausente"

                # Guarda o filepath original e o número da página que deverá ser cortada
                self.dados_pre_pdf.append({
                    "filepath": filepath,
                    "filename": filename,
                    "page_num": page_info["page_num"],
                    "apelido": apelido,
                    "regime": regime,
                    "status": status,
                    "competencia": page_info.get("Competencia", "")
                })
                
                num_exibicao = page_info["page_num"] + 1
                self.app.root.after(0, lambda f=filename, p=num_exibicao, a=apelido, c=page_info.get("Competencia",""), r=regime, s=status: 
                                self.tree_pdf.insert("", "end", values=(f"Pág {p} | {f}", a, c, r, s)))

        self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 LER E CLASSIFICAR PÁGINAS"))
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(state=tk.NORMAL))

    def executar_renomeacao_pdf(self):
        if not messagebox.askyesno("Confirmar", "Deseja separar as páginas e salvar os relatórios nas pastas LP e SN?"): return
        self.btn_exec_pdf.config(state=tk.DISABLED, text="FATIANDO PDFs...")
        threading.Thread(target=self.renomear_pdfs_thread, daemon=True).start()

    def renomear_pdfs_thread(self):
        sucesso = 0
        pasta_lp = os.path.join(self.pasta_pdfs, "LP")
        pasta_sn = os.path.join(self.pasta_pdfs, "SN")
        
        os.makedirs(pasta_lp, exist_ok=True)
        os.makedirs(pasta_sn, exist_ok=True)
        
        # Agrupa as páginas a serem cortadas pelo caminho do arquivo original.
        # Assim abrimos o PDF gigante (fitz.open) apenas uma vez por arquivo!
        arquivos_para_processar = {}
        for item in self.dados_pre_pdf:
            if item["status"] == "OK" and item["apelido"] and item["regime"] != "Desconhecido":
                fp = item["filepath"]
                if fp not in arquivos_para_processar:
                    arquivos_para_processar[fp] = []
                arquivos_para_processar[fp].append(item)
                
        for fp, lista_paginas in arquivos_para_processar.items():
            try:
                # Abre o bloco de PDF original
                doc_original = fitz.open(fp)
                
                for item in lista_paginas:
                    try:
                        novo_nome = f'{item["apelido"]}_fechamento.pdf'
                        
                        if item["regime"] == "Simples Nacional":
                            novo_caminho = os.path.join(pasta_sn, novo_nome)
                        else:
                            novo_caminho = os.path.join(pasta_lp, novo_nome)
                            
                        # Cria um PDF novo e vazio
                        doc_individual = fitz.open()
                        # Insere apenas a página específica do cliente neste novo PDF
                        doc_individual.insert_pdf(doc_original, from_page=item["page_num"], to_page=item["page_num"])
                        
                        # Salva o arquivo fatiado na pasta final e fecha-o
                        doc_individual.save(novo_caminho)
                        doc_individual.close()
                        
                        sucesso += 1
                    except Exception as e:
                        self.app.log_backend(f"Erro ao fatiar página {item['page_num']} de {item['filename']}: {e}")
                
                # Fecha o documento original após processar todas as páginas
                doc_original.close()
            except Exception as e:
                self.app.log_backend(f"Erro ao abrir arquivo base {fp}: {e}")
                    
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", f"{sucesso} Relatórios individualizados com sucesso nas pastas LP e SN!"))
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children())) 
        self.dados_pre_pdf = []
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(text="✅ FATIAR, SALVAR E ORGANIZAR EM PASTAS", state=tk.DISABLED))