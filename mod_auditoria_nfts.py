import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from utils import limpar_cnpj, formatar_data_br, seguro_float

class AbaAuditoriaNFTS:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.caminho_relatorio_aud = ""
        self.caminho_notas_aud = ""
        self.dados_auditoria_f1, self.dados_auditoria_f2, self.dados_auditoria_f3 = [], [], []
        self.dados_auditoria_f4, self.dados_auditoria_f5 = [], []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Auditoria Avançada: Validação de Emissões NFTS", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        tk.Label(frame_top, text="Selecione as bases de dados para o cruzamento de informações:", font=("Segoe UI", 10, "bold"), bg="white", fg="#333").pack(anchor="w", pady=(0, 10))

        frame_f1 = tk.Frame(frame_top, bg="white")
        frame_f1.pack(fill=tk.X, pady=5)
        tk.Button(frame_f1, text="📁 Relatório Consolidado (NFTS)", command=self.selecionar_aud_relatorio, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, width=30).pack(side=tk.LEFT)
        self.lbl_aud_relatorio = tk.Label(frame_f1, text="Nenhum arquivo selecionado", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_aud_relatorio.pack(side=tk.LEFT, padx=15)

        frame_f2 = tk.Frame(frame_top, bg="white")
        frame_f2.pack(fill=tk.X, pady=5)
        tk.Button(frame_f2, text="📁 Notas Fiscais Manassés", command=self.selecionar_aud_notas, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, width=30).pack(side=tk.LEFT)
        self.lbl_aud_notas = tk.Label(frame_f2, text="Nenhum arquivo selecionado", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_aud_notas.pack(side=tk.LEFT, padx=15)

        self.notebook_auditoria = ttk.Notebook(self.parent)
        self.notebook_auditoria.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        self.tab_aud1 = ttk.Frame(self.notebook_auditoria); self.notebook_auditoria.add(self.tab_aud1, text="1. Não Emitida")
        self.tab_aud2 = ttk.Frame(self.notebook_auditoria); self.notebook_auditoria.add(self.tab_aud2, text="2. Outro Município")
        self.tab_aud3 = ttk.Frame(self.notebook_auditoria); self.notebook_auditoria.add(self.tab_aud3, text="3. Valor Errado")
        self.tab_aud4 = ttk.Frame(self.notebook_auditoria); self.notebook_auditoria.add(self.tab_aud4, text="4. Sem Nota Correspond.")
        self.tab_aud5 = ttk.Frame(self.notebook_auditoria); self.notebook_auditoria.add(self.tab_aud5, text="5. Dado Incorreto")

        self.col_aud1 = ("Razão Social", "CNPJ", "Cidade", "Mensal. (Airtable)", "Abertura", "Inativação", "Cliente (Manassés)", "NF (Manassés)", "Data NF", "Valor NF")
        self.tree_aud1 = ttk.Treeview(self.tab_aud1, columns=self.col_aud1, show="headings"); self.tree_aud1.pack(fill=tk.BOTH, expand=True)
        for c in self.col_aud1: self.tree_aud1.heading(c, text=c); self.tree_aud1.column(c, width=90)

        self.col_aud2 = ("Razão Social", "CNPJ", "Cidade (Airtable)", "Valor na NFTS")
        self.tree_aud2 = ttk.Treeview(self.tab_aud2, columns=self.col_aud2, show="headings"); self.tree_aud2.pack(fill=tk.BOTH, expand=True)
        for c in self.col_aud2: self.tree_aud2.heading(c, text=c); self.tree_aud2.column(c, width=150)

        self.col_aud3 = ("Razão Social", "CNPJ", "Prestador (Relatório)", "Airtable (R$)", "NFTS (R$)", "Manassés (R$)", "Divergência", "Abertura", "Inativação")
        self.tree_aud3 = ttk.Treeview(self.tab_aud3, columns=self.col_aud3, show="headings"); self.tree_aud3.pack(fill=tk.BOTH, expand=True)
        for c in self.col_aud3: self.tree_aud3.heading(c, text=c); self.tree_aud3.column(c, width=100)

        self.col_aud4 = ("Nº NFTS", "Razão Social Tomador", "CNPJ", "Prestador (Relatório)", "Valor dos Serviços")
        self.tree_aud4 = ttk.Treeview(self.tab_aud4, columns=self.col_aud4, show="headings"); self.tree_aud4.pack(fill=tk.BOTH, expand=True)
        for c in self.col_aud4: self.tree_aud4.heading(c, text=c); self.tree_aud4.column(c, width=130)

        self.col_aud5 = ("Razão Social Tomador", "CNPJ", "Prestador (Relatório)", "Dado Divergente", "Info NFTS (Prefeitura)", "Info Sist. Manassés")
        self.tree_aud5 = ttk.Treeview(self.tab_aud5, columns=self.col_aud5, show="headings"); self.tree_aud5.pack(fill=tk.BOTH, expand=True)
        for c in self.col_aud5: self.tree_aud5.heading(c, text=c); self.tree_aud5.column(c, width=130)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_exec_aud = tk.Button(frame_btns, text="🔍 EXECUTAR AUDITORIA", command=self.executar_auditoria, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
        self.btn_exec_aud.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,10))
        self.btn_export_aud = tk.Button(frame_btns, text="✅ EXPORTAR RELATÓRIO", command=self.exportar_auditoria, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_export_aud.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def selecionar_aud_relatorio(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.caminho_relatorio_aud = path
            self.lbl_aud_relatorio.config(text=os.path.basename(path), fg="#2c3e50")

    def selecionar_aud_notas(self):
        path = filedialog.askopenfilename(filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls")])
        if path:
            self.caminho_notas_aud = path
            self.lbl_aud_notas.config(text=os.path.basename(path), fg="#2c3e50")

    def executar_auditoria(self):
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "Carregue os clientes do Airtable.")
        if not self.caminho_relatorio_aud or not self.caminho_notas_aud: return messagebox.showwarning("Atenção", "Selecione as DUAS planilhas.")
        self.btn_exec_aud.config(state=tk.DISABLED, text="CRUZANDO DADOS...")
        threading.Thread(target=self.auditoria_thread, daemon=True).start()

    def auditoria_thread(self):
        for tree in [self.tree_aud1, self.tree_aud2, self.tree_aud3, self.tree_aud4, self.tree_aud5]:
            self.app.root.after(0, lambda t=tree: t.delete(*t.get_children()))
        self.dados_auditoria_f1.clear(); self.dados_auditoria_f2.clear(); self.dados_auditoria_f3.clear()
        self.dados_auditoria_f4.clear(); self.dados_auditoria_f5.clear()

        try:
            df_rel = pd.read_excel(self.caminho_relatorio_aud)
            if 'Tipo de Registro' in df_rel.columns: df_rel = df_rel[df_rel['Tipo de Registro'] != 'Total']
            
            if self.caminho_notas_aud.endswith('.csv'): df_notas = pd.read_csv(self.caminho_notas_aud, sep=';', encoding='latin-1', on_bad_lines='skip')
            else: df_notas = pd.read_excel(self.caminho_notas_aud)

            df_rel['CNPJ_Limpo'] = df_rel['CPF/CNPJ do Tomador'].apply(limpar_cnpj)
            df_notas['CNPJ_Limpo'] = df_notas['Documento do cliente'].apply(limpar_cnpj)

            dict_rel = df_rel.drop_duplicates('CNPJ_Limpo').set_index('CNPJ_Limpo').to_dict('index')
            dict_notas = df_notas.drop_duplicates('CNPJ_Limpo').set_index('CNPJ_Limpo').to_dict('index')

            clientes_sp, clientes_outros, dict_airtable = [], [], {}

            for cnpj, dados in self.app.mapping_cnpj.items():
                cidade = str(dados.get("Cidade da Empresa", "")).lower()
                mens_val = seguro_float(dados.get("Mensalidade", "0"))
                dict_airtable[cnpj] = {
                    "Razão Social": dados.get("Razão Social", ""), "Cidade": dados.get("Cidade da Empresa", ""),
                    "Mensalidade": mens_val, "Data Abertura": formatar_data_br(dados.get("Data de Abertura", "")),
                    "Data Inativacao": formatar_data_br(dados.get("Data de Inativação", ""))
                }
                if any(x in cidade for x in ['são paulo', 'sao paulo']): clientes_sp.append(cnpj)
                elif cidade.strip(): clientes_outros.append(cnpj)

            for cnpj in clientes_sp:
                dados = dict_airtable[cnpj]
                if dados["Mensalidade"] > 0 and cnpj not in dict_rel:
                    nome_man = str(dict_notas[cnpj].get('Nome do cliente', '-')) if cnpj in dict_notas else "-"
                    nf_man = str(dict_notas[cnpj].get('Número', '-')).split('.')[0].strip() if cnpj in dict_notas else "-"
                    dt_man = formatar_data_br(dict_notas[cnpj].get('Data de emissão', '')) if cnpj in dict_notas else "-"
                    val_man_raw = seguro_float(dict_notas[cnpj].get('Valor total')) if cnpj in dict_notas else 0
                    val_man = f"R$ {val_man_raw:.2f}".replace('.', ',') if cnpj in dict_notas else "-"
                    
                    linha = [dados["Razão Social"], cnpj, dados["Cidade"], f"R$ {dados['Mensalidade']:.2f}".replace('.', ','), 
                             dados["Data Abertura"], dados["Data Inativacao"], nome_man, nf_man, dt_man, val_man]
                    self.dados_auditoria_f1.append(linha)
                    self.app.root.after(0, lambda l=linha: self.tree_aud1.insert("", "end", values=l))

            for cnpj in clientes_outros:
                if cnpj in dict_rel:
                    val_rel = seguro_float(dict_rel[cnpj].get('Valor dos Serviços'))
                    linha = [dict_airtable[cnpj]["Razão Social"], cnpj, dict_airtable[cnpj]["Cidade"], f"R$ {val_rel:.2f}".replace('.', ',')]
                    self.dados_auditoria_f2.append(linha)
                    self.app.root.after(0, lambda l=linha: self.tree_aud2.insert("", "end", values=l))

            for cnpj in clientes_sp:
                dados = dict_airtable[cnpj]
                v_air = dados["Mensalidade"]
                if v_air <= 0: continue
                
                v_rel = seguro_float(dict_rel[cnpj].get('Valor dos Serviços')) if cnpj in dict_rel else None
                v_nf = seguro_float(dict_notas[cnpj].get('Valor total')) if cnpj in dict_notas else None
                prestador = str(dict_rel[cnpj].get('Razão Social do Prestador', '-')) if cnpj in dict_rel else "-"
                
                msg_erro = []
                if v_rel is None: msg_erro.append("Ausente na Prefeitura")
                if v_nf is None: msg_erro.append("Ausente no Sist. Manassés")
                if v_rel is not None and v_air != v_rel: msg_erro.append("Airtable ≠ Prefeitura")
                if v_nf is not None and v_air != v_nf: msg_erro.append("Airtable ≠ Manassés")
                if v_rel is not None and v_nf is not None and v_rel != v_nf: 
                    if "Prefeitura ≠ Manassés" not in msg_erro: msg_erro.append("Prefeitura ≠ Manassés")
                
                if msg_erro:
                    s_rel = f"R$ {v_rel:.2f}".replace('.', ',') if v_rel is not None else "-"
                    s_nf = f"R$ {v_nf:.2f}".replace('.', ',') if v_nf is not None else "-"
                    linha = [dados["Razão Social"], cnpj, prestador, f"R$ {v_air:.2f}".replace('.', ','), s_rel, s_nf, " | ".join(msg_erro), dados["Data Abertura"], dados["Data Inativacao"]]
                    self.dados_auditoria_f3.append(linha)
                    self.app.root.after(0, lambda l=linha: self.tree_aud3.insert("", "end", values=l))

            for cnpj, row_rel in dict_rel.items():
                if cnpj not in dict_notas:
                    val_rel = seguro_float(row_rel.get('Valor dos Serviços'))
                    prestador = str(row_rel.get('Razão Social do Prestador', '-'))
                    linha = [str(row_rel.get('Nº NFTS', '')), str(row_rel.get('Razão Social do Tomador', '')), cnpj, prestador, f"R$ {val_rel:.2f}".replace('.', ',')]
                    self.dados_auditoria_f4.append(linha)
                    self.app.root.after(0, lambda l=linha: self.tree_aud4.insert("", "end", values=l))

            for cnpj, row_rel in dict_rel.items():
                if cnpj in dict_notas:
                    row_nf = dict_notas[cnpj]
                    raz_soc = str(row_rel.get('Razão Social do Tomador', ''))
                    prestador = str(row_rel.get('Razão Social do Prestador', '-'))
                    
                    d_rel = formatar_data_br(row_rel.get('Data da Prestação de Serviços'))
                    d_nf = formatar_data_br(row_nf.get('Data de emissão'))
                    if d_rel != d_nf:
                        linha = [raz_soc, cnpj, prestador, "Divergência de Data", d_rel, d_nf]
                        self.dados_auditoria_f5.append(linha)
                        self.app.root.after(0, lambda l=linha: self.tree_aud5.insert("", "end", values=l))
                    
                    n_rel = str(row_rel.get('Número do Documento', '')).split('.')[0].strip()
                    n_nf = str(row_nf.get('Número', '')).split('.')[0].strip()
                    if n_rel != n_nf:
                        linha = [raz_soc, cnpj, prestador, "Número da Nota", f"NFTS {n_rel}", f"NF {n_nf}"]
                        self.dados_auditoria_f5.append(linha)
                        self.app.root.after(0, lambda l=linha: self.tree_aud5.insert("", "end", values=l))
                    
                    v_rel = seguro_float(row_rel.get('Valor dos Serviços'))
                    v_nf = seguro_float(row_nf.get('Valor total'))
                    if v_rel != v_nf:
                        linha = [raz_soc, cnpj, prestador, "Valor Total", f"R$ {v_rel:.2f}".replace('.', ','), f"R$ {v_nf:.2f}".replace('.', ',')]
                        self.dados_auditoria_f5.append(linha)
                        self.app.root.after(0, lambda l=linha: self.tree_aud5.insert("", "end", values=l))

        except Exception as e:
            self.app.log_backend(f"[ERRO] Auditoria: {e}")
            self.app.root.after(0, lambda: messagebox.showerror("Erro de Processamento", f"Ocorreu um erro ao processar os arquivos.\n{e}"))

        self.app.root.after(0, lambda: self.btn_exec_aud.config(state=tk.NORMAL, text="🔍 EXECUTAR AUDITORIA"))
        self.app.root.after(0, lambda: self.btn_export_aud.config(state=tk.NORMAL))
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", "Auditoria finalizada! Verifique as abas."))

    def exportar_auditoria(self):
        pasta = filedialog.askdirectory(title="Salvar Relatório de Auditoria")
        if not pasta: return
        
        caminho = os.path.join(pasta, "Relatorio_Auditoria_NFTS.xlsx")
        try:
            with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
                pd.DataFrame(self.dados_auditoria_f1, columns=self.col_aud1).to_excel(writer, sheet_name="Não Emitida", index=False)
                pd.DataFrame(self.dados_auditoria_f2, columns=self.col_aud2).to_excel(writer, sheet_name="Outro Município", index=False)
                pd.DataFrame(self.dados_auditoria_f3, columns=self.col_aud3).to_excel(writer, sheet_name="Valor Errado", index=False)
                pd.DataFrame(self.dados_auditoria_f4, columns=self.col_aud4).to_excel(writer, sheet_name="Sem Nota Correspondente", index=False)
                pd.DataFrame(self.dados_auditoria_f5, columns=self.col_aud5).to_excel(writer, sheet_name="Dado Incorreto", index=False)
            messagebox.showinfo("Sucesso", f"Relatório exportado em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar Excel:\n{e}")