import os, re, glob, shutil, threading, unicodedata, tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from utils import limpar_cnpj, limpar_nome_arquivo, validar_tributacao, COLUNAS_AIRTABLE

try:
    import fitz, pytesseract, ctypes
    from PIL import Image

    def obter_caminho_curto(caminho):
        try:
            buf_size = ctypes.windll.kernel32.GetShortPathNameW(caminho, None, 0)
            if buf_size > 0:
                buf = ctypes.create_unicode_buffer(buf_size)
                ctypes.windll.kernel32.GetShortPathNameW(caminho, buf, buf_size)
                return buf.value
        except: pass
        return caminho

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR_CURTO = obter_caminho_curto(BASE_DIR).replace("\\", "/")
    
    caminho_tess_pasta = f"{BASE_DIR_CURTO}/Tesseract-OCR/tesseract.exe"
    caminho_tess_raiz = f"{BASE_DIR_CURTO}/tesseract.exe"
    
    if os.path.exists(caminho_tess_pasta):
        pytesseract.pytesseract.tesseract_cmd = caminho_tess_pasta
        os.environ["TESSDATA_PREFIX"] = f"{BASE_DIR_CURTO}/Tesseract-OCR/tessdata"
    elif os.path.exists(caminho_tess_raiz):
        pytesseract.pytesseract.tesseract_cmd = caminho_tess_raiz
        os.environ["TESSDATA_PREFIX"] = f"{BASE_DIR_CURTO}/tessdata"

    pasta_segura_ocr = "C:\\Temp_OCR_Hub"
    os.makedirs(pasta_segura_ocr, exist_ok=True)
    tempfile.tempdir = pasta_segura_ocr
    os.environ['TMP'] = pasta_segura_ocr
    os.environ['TEMP'] = pasta_segura_ocr

    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

class AbaDARF:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.pasta_pdfs = ""
        self.dados_pre_pdf = []
        self.construir_tela()

    def construir_tela(self):
        tk.Label(self.parent, text="Processador de Guias - DARF", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)
        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        self.var_apenas_leitura = tk.BooleanVar(value=False)
        tk.Checkbutton(frame_top, text="Fazer apenas Leitura (Não Renomear)", variable=self.var_apenas_leitura, bg="white", font=("Segoe UI", 9, "italic"), fg="#d35400").pack(side=tk.RIGHT, padx=5)

        tk.Label(frame_top, text="Selecione o diretório com os arquivos (.pdf)", font=("Segoe UI", 9), bg="white").pack(anchor="w", pady=(5, 5))
        frame_input = tk.Frame(frame_top, bg="white")
        frame_input.pack(fill=tk.X)
        self.btn_pasta_pdf = tk.Button(frame_input, text="📁 Procurar Pasta", command=self.selecionar_pasta_pdf, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10)
        self.btn_pasta_pdf.pack(side=tk.LEFT)
        self.lbl_pasta_pdf = tk.Label(frame_input, text="Nenhuma pasta selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_pasta_pdf.pack(side=tk.LEFT, padx=15)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        self.tree_pdf = ttk.Treeview(frame_table, show="headings", height=10)
        colunas = ("Arquivo Original", "Empresa", "Competência", "Valor", "Status de Ação", "Status")
        self.tree_pdf["columns"] = colunas
        for col in colunas:
            self.tree_pdf.heading(col, text=col)
            self.tree_pdf.column(col, minwidth=80, width=200 if col == "Empresa" else 120)
        self.tree_pdf.pack(fill=tk.BOTH, expand=True, pady=5)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)
        self.btn_ler_pdf = tk.Button(frame_btns, text="🔍 INICIAR LEITURA DARF", command=self.iniciar_leitura_pdf, bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8)
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
        
    def formatar_valor(self, val):
        if not val: return ""
        val_clean = re.sub(r'[^\d.,]', '', str(val).strip())
        if len(val_clean) >= 3 and val_clean[-3] in ['.', ',']:
            inteiro = val_clean[:-3].replace('.', '').replace(',', '')
            decimal = val_clean[-2:]
            try: return f"{float(f'{inteiro}.{decimal}'):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            except: pass
        return val

    def extrair_dados_darf(self, pdf_path):
        info = {"CNPJ": None, "Todos_CNPJs": [], "Valor": None, "Competencia": None, "Vencimento": None, "Tipo_Guia": "DARF DCTFWeb", "Sufixo_Arquivo": "DARF_DCTF_WEB", "Regra_Usada": ""}
        text_para_valor = ""
        
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    text_para_valor = pdf.pages[0].extract_text(layout=False) or ""
            except: pass

        if not re.search(r'\d+(?:\.\d{3})*,\d{2}', text_para_valor) and HAS_OCR:
            try:
                temp_antigo = tempfile.tempdir
                tempfile.tempdir = "C:\\Temp_OCR_Hub"
                try:
                    doc = fitz.open(pdf_path)
                    text_ocr = ""
                    tess_config = f'--tessdata-dir "{os.environ.get("TESSDATA_PREFIX")}"' if os.environ.get("TESSDATA_PREFIX") else ""
                    for page in doc:
                        pix = page.get_pixmap(dpi=200)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        text_ocr += pytesseract.image_to_string(img, lang='por', config=tess_config)
                    text_para_valor = text_ocr
                finally:
                    tempfile.tempdir = temp_antigo
            except Exception as e:
                self.app.log_backend(f"[AVISO] Erro no OCR: {e}")

        t_cnpj_limpo = text_para_valor.replace(' ', '').replace('\n', '')
        matches_perfeitos = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', t_cnpj_limpo)
        for p in matches_perfeitos:
            num = re.sub(r'\D', '', p).zfill(14)
            fmt = f"{num[:2]}.{num[2:5]}.{num[5:8]}/{num[8:12]}-{num[12:]}"
            if fmt not in info["Todos_CNPJs"]: info["Todos_CNPJs"].append(fmt)

        if not info["Todos_CNPJs"]:
            text_num = t_cnpj_limpo.replace('.', '').replace('-', '').replace('/', '')
            near_matches = re.finditer(r'(?:CNPJ|CPF|INSCRI[ÇC][ÃA]O)[\s\S]{0,50}?(\d{13,14})(?!\d)', text_num, re.IGNORECASE)
            for m in near_matches:
                num = m.group(1).zfill(14)
                fmt = f"{num[:2]}.{num[2:5]}.{num[5:8]}/{num[8:12]}-{num[12:]}"
                if fmt not in info["Todos_CNPJs"]: info["Todos_CNPJs"].append(fmt)

        if info["Todos_CNPJs"]: info["CNPJ"] = info["Todos_CNPJs"][-1]

        def normalizar_texto(t):
            t = t.upper()
            t = ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
            return re.sub(r'[^A-Z0-9]', '', t)

        t_norm = normalizar_texto(text_para_valor)

        if "DOCUMENTODEARRECADACAODERECEITASFEDERAIS" in t_norm:
            info["Regra_Usada"] = "Receita Federal - DARF DCTF-WEB"
            def extracao_segura(padroes):
                t_no_space = text_para_valor.replace(' ', '')
                for p in padroes:
                    m = re.search(p, text_para_valor, re.IGNORECASE | re.DOTALL)
                    if m: return m.group(1)
                    p_no_space = p.replace(r'\s*', '').replace(' ', '')
                    m2 = re.search(p_no_space, t_no_space, re.IGNORECASE | re.DOTALL)
                    if m2: return m2.group(1)
                return None

            info["Valor"] = extracao_segura([r'(?:Valor Total do Documento|Valor)[^\d]*(\d+(?:[.,]\d{3})*[.,]\d{2})'])
            
            c_raw = extracao_segura([
                r'((?:janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)/\d{4})',
                r'(?:PER[ÍI]ODO DE APURA[ÇC][ÃA]O|PA)[^\w]*([A-Za-z]+/\d{4}|\d{2}/\d{4}|\d{2}/\d{2}/\d{4})'
            ])
            if c_raw:
                if len(c_raw) >= 10 and not re.search(r'[A-Za-z]', c_raw): c_raw = c_raw[-7:]
                info["Competencia"] = self.formatar_competencia(c_raw)
                
            info["Vencimento"] = extracao_segura([r'(?:Pagar|Data de Vencimento)[^\d]*(\d{2}/\d{2}/(?:19|20)\d{2})'])
        else:
            info["Regra_Usada"] = "Arquivo não reconhecido como DARF"

        info["Valor"] = self.formatar_valor(info["Valor"])
        return info

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
            info = self.extrair_dados_darf(filepath) 
            
            cnpj_valido, dados_empresa, status, apelido, novo_nome = None, {}, "OK", "", ""
            tipo_guia = info.get("Tipo_Guia", "")

            if info["Regra_Usada"] == "Arquivo não reconhecido como DARF": status = info["Regra_Usada"]
            else:
                for c in reversed(info.get("Todos_CNPJs", [])):
                    cl = limpar_cnpj(c)
                    if cl in self.app.mapping_cnpj:
                        d_temp = self.app.mapping_cnpj[cl]
                        if validar_tributacao(tipo_guia, d_temp.get("Tipo de Tributação", "")):
                            cnpj_valido = c; dados_empresa = d_temp; break
                
                if cnpj_valido:
                    apelido = limpar_nome_arquivo(dados_empresa.get("Apelido", ""))
                    novo_nome = f"{apelido}_{info.get('Sufixo_Arquivo')}.pdf"
                else: status = "Não mapeado/Trib. errada"

            msg_acao = "Pronto p/ Leitura" if self.var_apenas_leitura.get() else "Pronto p/ Renomear"
            if status != "OK": msg_acao = "Ação Cancelada"

            base_excel = {col: dados_empresa.get(col, "") for col in COLUNAS_AIRTABLE}
            base_excel.update({"Arquivo Original": filename, "Arquivo Renomeado": novo_nome if (novo_nome and not self.var_apenas_leitura.get()) else "APENAS LEITURA", "Tipo de Guia": tipo_guia, "Mês de Competência": info.get("Competencia", ""), "Data de Vencimento": info.get("Vencimento", ""), "Valor (R$)": info.get("Valor", "")})

            self.dados_pre_pdf.append({"filepath": filepath, "filename": filename, "novo_nome": novo_nome, "status": status, "linhas_excel": [base_excel]})
            self.app.root.after(0, lambda f=filename, a=apelido, c=info.get("Competencia",""), v=info.get("Valor",""), act=msg_acao, s=status: self.tree_pdf.insert("", "end", values=(f, a, c, v, act, s)))

        self.app.root.after(0, lambda: self.btn_ler_pdf.config(state=tk.NORMAL, text="🔍 INICIAR LEITURA DARF"))
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
            pd.DataFrame(dados_excel).to_excel(os.path.join(self.pasta_pdfs, "Relatorio_DARF.xlsx"), index=False)
        self.app.root.after(0, lambda: messagebox.showinfo("Concluído", f"{sucesso} PDFs processados!"))
        self.app.root.after(0, lambda: self.tree_pdf.delete(*self.tree_pdf.get_children())) 
        self.dados_pre_pdf = []
        self.app.root.after(0, lambda: self.btn_exec_pdf.config(text="✅ APROVAR E EXECUTAR ROTINA", state=tk.DISABLED))