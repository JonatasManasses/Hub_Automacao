import os, glob, shutil, threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
from utils import limpar_nome_arquivo

class AbaXML:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.pasta_xml_origem = ""
        self.pasta_xml_destino_base = r"C:\DominioWeb\TOMADOS"
        self.dados_pre_xml = []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Organizador de XML (NFSe Tomados)", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        frame_o = tk.Frame(frame_top, bg="white")
        frame_o.pack(fill=tk.X, pady=5)
        self.btn_xml_origem = tk.Button(frame_o, text="📁 Pasta com XMLs Soltos", command=self.selecionar_pasta_xml_origem, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, width=25)
        self.btn_xml_origem.pack(side=tk.LEFT)
        self.lbl_xml_origem = tk.Label(frame_o, text="Nenhuma pasta selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_xml_origem.pack(side=tk.LEFT, padx=15)

        frame_d = tk.Frame(frame_top, bg="white")
        frame_d.pack(fill=tk.X, pady=5)
        self.btn_xml_destino = tk.Button(frame_d, text="🎯 Selecionar Destino (Domínio)", command=self.selecionar_pasta_xml_destino, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, width=25)
        self.btn_xml_destino.pack(side=tk.LEFT)
        self.lbl_xml_destino = tk.Label(frame_d, text=self.pasta_xml_destino_base, fg="#2980b9", bg="white", font=("Segoe UI", 10, "bold"))
        self.lbl_xml_destino.pack(side=tk.LEFT, padx=15)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        colunas_xml = ("Arquivo XML", "CNPJ Tomador", "Competência", "Empresa Mapeada", "Pasta Destino", "Status")
        self.tree_xml = ttk.Treeview(frame_table, columns=colunas_xml, show="headings", height=10)
        for col in colunas_xml:
            self.tree_xml.heading(col, text=col)
            self.tree_xml.column(col, minwidth=100, width=150)
        self.tree_xml.column("Pasta Destino", width=250)
        self.tree_xml.pack(fill=tk.BOTH, expand=True, pady=5)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_ler_xml = tk.Button(frame_btns, text="🔍 ANALISAR ARQUIVOS XML", command=self.iniciar_leitura_xml, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
        self.btn_ler_xml.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)
        self.btn_exec_xml = tk.Button(frame_btns, text="✅ APROVAR E MOVER", command=self.executar_movimentacao_xml, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_exec_xml.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def selecionar_pasta_xml_origem(self):
        pasta = filedialog.askdirectory(title="Onde estão os XMLs soltos?")
        if pasta:
            self.pasta_xml_origem = pasta
            self.lbl_xml_origem.config(text=pasta, fg="#2c3e50")

    def selecionar_pasta_xml_destino(self):
        pasta = filedialog.askdirectory(title="Pasta Base do Domínio")
        if pasta:
            self.pasta_xml_destino_base = pasta
            self.lbl_xml_destino.config(text=pasta)

    def extrair_dados_nfse(self, filepath):
        info = {"CNPJ_Tomador": None, "DataCompetencia": None, "Erro": ""}
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f: xml_string = f.read()
            bloco_tomador = re.search(r'<toma(?:>|\s[^>]*>).*?</toma>', xml_string, re.IGNORECASE | re.DOTALL)
            if bloco_tomador:
                match_cnpj = re.search(r'<CNPJ[^>]*>(\d{14})</CNPJ>', bloco_tomador.group(0), re.IGNORECASE)
                if match_cnpj: info["CNPJ_Tomador"] = match_cnpj.group(1)
            match_data = re.search(r'<dCompet[^>]*>(\d{4}-\d{2}-\d{2})</dCompet>', xml_string, re.IGNORECASE)
            if match_data: info["DataCompetencia"] = match_data.group(1)
        except Exception as e: info["Erro"] = str(e)
        return info

    def iniciar_leitura_xml(self):
        if not self.pasta_xml_origem: return messagebox.showwarning("Atenção", "Selecione a pasta de Origem dos XMLs.")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "Aguarde o carregamento do Airtable.")
        self.btn_ler_xml.config(state=tk.DISABLED, text="LENDO...")
        threading.Thread(target=self.ler_xmls_thread, daemon=True).start()

    def ler_xmls_thread(self):
        self.app.root.after(0, lambda: self.tree_xml.delete(*self.tree_xml.get_children()))
        self.dados_pre_xml = []
        xmls = glob.glob(os.path.join(self.pasta_xml_origem, '*.xml'))
        if not xmls: return self.app.root.after(0, lambda: self.btn_ler_xml.config(state=tk.NORMAL, text="🔍 ANALISAR ARQUIVOS XML"))
        
        for filepath in xmls:
            filename = os.path.basename(filepath)
            info = self.extrair_dados_nfse(filepath)
            cnpj = info["CNPJ_Tomador"]
            data_comp = info["DataCompetencia"]
            status, apelido, codigo, destino_final, mesano = "OK", "", "", "", ""
            
            if info["Erro"]: status = f"Erro Leitura: {info['Erro']}"
            elif not cnpj: status = "CNPJ Tomador não localizado"
            elif not data_comp: status = "Data Competência ausente"
            else:
                empresa = self.app.mapping_cnpj.get(cnpj)
                if empresa:
                    apelido = limpar_nome_arquivo(empresa.get("Apelido", ""))
                    codigo = limpar_nome_arquivo(empresa.get("Código", ""))
                    ano, mes, dia = data_comp.split('-')
                    mesano = f"{mes}{ano}"
                    nome_pasta_cliente = f"{codigo}-{apelido}" if codigo else apelido
                    destino_final = os.path.join(self.pasta_xml_destino_base, nome_pasta_cliente, mesano)
                else: status = "CNPJ não mapeado"

            self.dados_pre_xml.append({"filepath": filepath, "filename": filename, "destino_final": destino_final, "status": status})
            self.app.root.after(0, lambda f=filename, c=cnpj, ma=(f"{mes}/{ano}" if mesano else ""), a=apelido, d=destino_final, s=status: 
                            self.tree_xml.insert("", "end", values=(f, c, ma, a, d, s)))

        self.app.root.after(0, lambda: self.btn_ler_xml.config(state=tk.NORMAL, text="🔍 ANALISAR ARQUIVOS XML"))
        self.app.root.after(0, lambda: self.btn_exec_xml.config(state=tk.NORMAL))

    def executar_movimentacao_xml(self):
        if not messagebox.askyesno("Confirmar", "Deseja organizar os XMLs?"): return
        self.btn_exec_xml.config(state=tk.DISABLED, text="MOVENDO...")
        threading.Thread(target=self.mover_xmls_thread, daemon=True).start()

    def mover_xmls_thread(self):
        sucesso = 0
        for item in self.dados_pre_xml:
            if item["status"] == "OK" and item["destino_final"]:
                try:
                    os.makedirs(item["destino_final"], exist_ok=True)
                    novo_caminho = os.path.join(item["destino_final"], item["filename"])
                    if not os.path.exists(novo_caminho):
                        shutil.move(item["filepath"], novo_caminho)
                        sucesso += 1
                except: pass
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", f"{sucesso} XMLs organizados com sucesso!"))
        self.app.root.after(0, lambda: self.tree_xml.delete(*self.tree_xml.get_children()))
        self.dados_pre_xml = []
        self.app.root.after(0, lambda: self.btn_exec_xml.config(text="✅ APROVAR E MOVER", state=tk.DISABLED))