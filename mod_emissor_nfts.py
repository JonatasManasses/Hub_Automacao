import os, threading, time, re
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from utils import limpar_cnpj, obter_certificado_padrao, ROOT_DIR_NFTS

try:
    import pyautogui
    import pyperclip
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    HAS_RPA = True
except ImportError:
    HAS_RPA = False

class AbaEmissorNFTS:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.caminho_planilha_nfts = ""
        self.df_nfts_completo = None
        self.stop_nfts_event = threading.Event()
        self.pause_nfts_event = threading.Event()
        self.construir_tela()

    def garantir_pastas_nfts(self, mes_ano_str):
        base = os.path.join(ROOT_DIR_NFTS, f"Execucao_{mes_ano_str}")
        pastas = {"base": base, "pdf": os.path.join(base, "NOTAS_FISCAIS_PDF"), "csv": os.path.join(base, "RELATORIO_CSV")}
        for p in pastas.values(): os.makedirs(p, exist_ok=True)
        return pastas

    def compilar_csvs_nfts(self, pasta_csv, mes_ano):
        arquivos = [f for f in os.listdir(pasta_csv) if f.lower().endswith('.csv')]
        if not arquivos: return
        lista_dfs = []
        for arquivo in arquivos:
            try:
                df = pd.read_csv(os.path.join(pasta_csv, arquivo), sep=';', encoding='latin-1', on_bad_lines='skip')
                if not df.empty: lista_dfs.append(df)
            except: continue
        if lista_dfs:
            df_final = pd.concat(lista_dfs, ignore_index=True)
            caminho_final = os.path.join(pasta_csv, f"RELATORIO_CONSOLIDADO_{mes_ano}.xlsx")
            df_final.to_excel(caminho_final, index=False)

    def construir_tela(self):
        tk.Label(self.parent, text="Robô Emissor: NFTS (São Paulo)", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#2c3e50").pack(anchor="w", pady=(20, 10), padx=20)

        frame_config = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_config.pack(fill=tk.X, padx=20, pady=(0, 15))
        tk.Label(frame_config, text="Configurações da Emissão", font=("Segoe UI", 10, "bold"), bg="white", fg="#2980b9").pack(anchor="w", pady=(0, 10))

        row1 = tk.Frame(frame_config, bg="white")
        row1.pack(fill=tk.X, pady=5)
        tk.Label(row1, text="CNPJ do Prestador:", bg="white", width=20, anchor="w").pack(side=tk.LEFT)
        self.ent_nfts_cnpj_prestador = ttk.Entry(row1, width=25)
        self.ent_nfts_cnpj_prestador.pack(side=tk.LEFT, padx=(0, 20))

        lista_codigos = [
            "3476 - Contabilidade, inclusive serviços técnicos e auxiliares.",
            "3158 - Consultoria e assessoria econômica ou financeira.",
            "3131 - Auditoria.",
            "3140 - Perícia.",
            "2801 - Administração de bens e negócios de terceiros.",
            "3212 - Serviços de RH, seleção, treinamento.",
            "2887 - Assessoria ou consultoria de qualquer natureza."
        ]
        tk.Label(row1, text="Cód. Atividade:", bg="white", width=15, anchor="w").pack(side=tk.LEFT)
        self.combo_nfts_cod_atividade = ttk.Combobox(row1, values=lista_codigos, width=65)
        self.combo_nfts_cod_atividade.set(lista_codigos[0])
        self.combo_nfts_cod_atividade.pack(side=tk.LEFT)
        
        def check_input(event):
            value = event.widget.get()
            if value == '': self.combo_nfts_cod_atividade['values'] = lista_codigos
            else:
                data = [item for item in lista_codigos if value.lower() in item.lower()]
                self.combo_nfts_cod_atividade['values'] = data
        self.combo_nfts_cod_atividade.bind('<KeyRelease>', check_input)

        frame_top = tk.Frame(self.parent, bg="white", pady=15, padx=20, relief=tk.RIDGE, borderwidth=1)
        frame_top.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        frame_input = tk.Frame(frame_top, bg="white")
        frame_input.pack(fill=tk.X)
        tk.Button(frame_input, text="📁 Procurar Planilha", command=self.selecionar_planilha_nfts, bg="#ecf0f1", fg="#2c3e50", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10).pack(side=tk.LEFT)
        self.lbl_pasta_nfts = tk.Label(frame_input, text="Nenhuma planilha selecionada", fg="#7f8c8d", bg="white", font=("Segoe UI", 10))
        self.lbl_pasta_nfts.pack(side=tk.LEFT, padx=15)
        
        self.btn_toggle_nfts = tk.Button(frame_input, text="☑ Marcar/Desmarcar Todos", command=self.toggle_todas_nfts, bg="#95a5a6", fg="white", font=("Segoe UI", 8, "bold"), relief=tk.FLAT, state=tk.DISABLED)
        self.btn_toggle_nfts.pack(side=tk.RIGHT)

        frame_table = tk.Frame(self.parent, bg="#f4f6f9")
        frame_table.pack(fill=tk.BOTH, expand=True, padx=20)
        
        colunas_nfts = ("Seleção", "Data", "Nº NF", "Cliente", "Valor da NF", "CNPJ", "Status")
        self.tree_nfts = ttk.Treeview(frame_table, columns=colunas_nfts, show="headings", height=8)
        for col in colunas_nfts:
            self.tree_nfts.heading(col, text=col)
            if col == "Seleção": w = 60
            elif col == "Data": w = 80
            elif col == "Nº NF": w = 80
            elif col == "Cliente": w = 200
            else: w = 120
            self.tree_nfts.column(col, minwidth=60, width=w, anchor="center" if col=="Seleção" else "w")
        self.tree_nfts.pack(fill=tk.BOTH, expand=True, pady=5)
        
        def on_tree_click(event):
            region = self.tree_nfts.identify("region", event.x, event.y)
            if region == "cell":
                column = self.tree_nfts.identify_column(event.x)
                if column == '#1': 
                    item = self.tree_nfts.identify_row(event.y)
                    vals = list(self.tree_nfts.item(item, "values"))
                    vals[0] = "[ ]" if vals[0] == "[X]" else "[X]"
                    self.tree_nfts.item(item, values=vals)
        self.tree_nfts.bind('<ButtonRelease-1>', on_tree_click)

        frame_btns = tk.Frame(self.parent, bg="#f4f6f9")
        frame_btns.pack(pady=15, fill=tk.X, padx=20)

        frame_controles = tk.Frame(frame_btns, bg="#f4f6f9")
        frame_controles.pack(side=tk.LEFT, fill=tk.X)
        
        self.btn_exec_nfts = tk.Button(frame_controles, text="🚀 INICIAR EMISSÃO", command=self.iniciar_nfts_wrapper, bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_exec_nfts.pack(side=tk.LEFT, padx=(0,5))
        self.btn_pause_nfts = tk.Button(frame_controles, text="⏸ Pausar", command=self.pausar_nfts, bg="#f39c12", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_pause_nfts.pack(side=tk.LEFT, padx=5)
        self.btn_stop_nfts = tk.Button(frame_controles, text="⏹ Parar", command=self.parar_nfts, bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_stop_nfts.pack(side=tk.LEFT, padx=5)

        frame_rel = tk.Frame(frame_btns, bg="#f4f6f9")
        frame_rel.pack(side=tk.RIGHT, fill=tk.X)
        self.nfts_d_ini = tk.StringVar(value=datetime.now().replace(day=1).strftime("%d/%m/%Y"))
        self.nfts_d_fim = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        tk.Label(frame_rel, text="Início:", bg="#f4f6f9").pack(side=tk.LEFT)
        ttk.Entry(frame_rel, textvariable=self.nfts_d_ini, width=12).pack(side=tk.LEFT, padx=5)
        tk.Label(frame_rel, text="Fim:", bg="#f4f6f9").pack(side=tk.LEFT)
        ttk.Entry(frame_rel, textvariable=self.nfts_d_fim, width=12).pack(side=tk.LEFT, padx=5)
        self.btn_rel_nfts = tk.Button(frame_rel, text="📊 GERAR RELATÓRIOS", command=self.iniciar_relatorio_nfts_wrapper, bg="#2980b9", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, state=tk.DISABLED)
        self.btn_rel_nfts.pack(side=tk.LEFT, padx=(10,0))

    def selecionar_planilha_nfts(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")])
        if path:
            self.caminho_planilha_nfts = path
            self.lbl_pasta_nfts.config(text=os.path.basename(path), fg="#2c3e50")
            self.carregar_planilha_nfts()

    def toggle_todas_nfts(self):
        items = self.tree_nfts.get_children()
        if not items: return
        estado_atual = self.tree_nfts.item(items[0], "values")[0]
        novo_estado = "[ ]" if estado_atual == "[X]" else "[X]"
        for item in items:
            vals = list(self.tree_nfts.item(item, "values"))
            vals[0] = novo_estado
            self.tree_nfts.item(item, values=vals)

    def carregar_planilha_nfts(self):
        try:
            try:
                if self.caminho_planilha_nfts.endswith('.csv'): df = pd.read_csv(self.caminho_planilha_nfts, sep=';', encoding='latin-1', on_bad_lines='skip')
                else: df = pd.read_excel(self.caminho_planilha_nfts)
            except: df = pd.read_excel(self.caminho_planilha_nfts, engine='openpyxl' if self.caminho_planilha_nfts.endswith('.xlsx') else 'xlrd')
                
            def extrair_data_local(v):
                if pd.isna(v) or not str(v).strip(): return None
                v_str = str(v).strip().split(' ')[0]
                if re.match(r'^\d{2}-\d{2}-\d{4}$', v_str): v_str = v_str.replace('-', '/')
                try: 
                    d = pd.to_datetime(v_str, dayfirst=True)
                    if pd.notnull(d): return d
                except: pass
                return None

            dt_obj = extrair_data_local(df["Data de Emissão"].iloc[0]) if not df.empty and "Data de Emissão" in df.columns else None
            mes_ano_ref = dt_obj.strftime('%m-%Y') if dt_obj else datetime.now().strftime('%m-%Y')
            caminho_resultado = os.path.join(ROOT_DIR_NFTS, f"Execucao_{mes_ano_ref}", "RESULTADO_EMISSAO.xlsx")
            
            if os.path.exists(caminho_resultado):
                df_res = pd.read_excel(caminho_resultado)
                if "Status" in df_res.columns:
                    status_map = dict(zip(df_res["CNPJ"], df_res["Status"]))
                    df["Status"] = df["CNPJ"].map(status_map).fillna("Pendente")
            
            if "Status" not in df.columns: df["Status"] = "Pendente"
            else: df["Status"] = df["Status"].fillna("Pendente")
                
            self.df_nfts_completo = df
            self.tree_nfts.delete(*self.tree_nfts.get_children())
            
            for i, row in df.iterrows():
                dt = extrair_data_local(row.get("Data de Emissão"))
                dt_str = dt.strftime('%d/%m/%Y') if dt else str(row.get("Data de Emissão", ""))
                status = str(row.get("Status", ""))
                chk = "[ ]" if status == "Emitida" else "[X]"
                
                self.tree_nfts.insert("", "end", values=(
                    chk, dt_str, str(row.get("Numero da NF", row.get("Número da NF", ""))),
                    str(row.get("Cliente", "")), str(row.get("Valor da NF", "")), str(row.get("CNPJ", "")), status
                ))

            self.btn_exec_nfts.config(state=tk.NORMAL)
            self.btn_toggle_nfts.config(state=tk.NORMAL)
            self.btn_rel_nfts.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Não foi possível carregar a planilha.\nVerifique se o formato está correto (Instale 'xlrd' se for XLS antigo).\nErro: {str(e)}")

    def pausar_nfts(self):
        if self.pause_nfts_event.is_set():
            self.pause_nfts_event.clear()
            self.btn_pause_nfts.config(text="⏸ Pausar", bg="#f39c12")
        else:
            self.pause_nfts_event.set()
            self.btn_pause_nfts.config(text="▶ Retomar", bg="#27ae60")

    def parar_nfts(self):
        self.stop_nfts_event.set()
        self.btn_stop_nfts.config(state=tk.DISABLED, text="Parando...")

    def iniciar_nfts_wrapper(self):
        if not HAS_RPA: return messagebox.showerror("Bibliotecas Ausentes", "As bibliotecas de automação não foram encontradas.\nAbra o terminal e digite:\npip install selenium pyautogui pyperclip")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "O Airtable ainda não foi sincronizado.")
        if not self.ent_nfts_cnpj_prestador.get().strip(): return messagebox.showwarning("Atenção", "Preencha o CNPJ do Prestador.")
        if not obter_certificado_padrao(): return messagebox.showwarning("Atenção", "Certificado Digital não configurado.\nVá na aba Configurações.")

        self.stop_nfts_event.clear()
        self.pause_nfts_event.clear()
        self.btn_exec_nfts.config(state=tk.DISABLED, text="PROCESSANDO...")
        self.btn_pause_nfts.config(state=tk.NORMAL, text="⏸ Pausar", bg="#f39c12")
        self.btn_stop_nfts.config(state=tk.NORMAL, text="⏹ Parar")

        threading.Thread(target=self.emissao_nfts_thread, daemon=True).start()

    def emissao_nfts_thread(self):
        driver = None
        caminho_resultado = None
        try:
            cnpjs_marcados = []
            mapa_itens_tree = {} 
            for item in self.tree_nfts.get_children():
                vals = self.tree_nfts.item(item, "values")
                if vals[0] == "[X]":
                    cnpj_raw = str(vals[5])
                    cnpj_limpo_v = limpar_cnpj(cnpj_raw)
                    cnpjs_marcados.append(cnpj_limpo_v)
                    mapa_itens_tree[cnpj_limpo_v] = item
            
            if not cnpjs_marcados: return self.app.root.after(0, lambda: messagebox.showinfo("Aviso", "Nenhuma nota marcada [X] para emissão."))

            df_completo = self.df_nfts_completo
            df_completo['CNPJ_Limpo_Temp'] = df_completo['CNPJ'].apply(limpar_cnpj)
            df_proc = df_completo[df_completo['CNPJ_Limpo_Temp'].isin(cnpjs_marcados)].copy()
            df_completo = df_completo.drop('CNPJ_Limpo_Temp', axis=1)

            def ext_dt(v):
                if pd.isna(v) or not str(v).strip(): return None
                v_str = str(v).strip().split(' ')[0]
                if re.match(r'^\d{2}-\d{2}-\d{4}$', v_str): v_str = v_str.replace('-', '/')
                try: 
                    d = pd.to_datetime(v_str, dayfirst=True)
                    if pd.notnull(d): return d
                except: pass
                return None
                
            dt_obj_ref = ext_dt(df_proc["Data de Emissão"].iloc[0]) if not df_proc.empty else None
            mes_ano_geral = dt_obj_ref.strftime('%m-%Y') if dt_obj_ref else datetime.now().strftime('%m-%Y')
            pastas_gerais = self.garantir_pastas_nfts(mes_ano_geral)
            caminho_resultado = os.path.join(pastas_gerais['base'], "RESULTADO_EMISSAO.xlsx")

            cnpj_prestador = self.ent_nfts_cnpj_prestador.get().strip()
            cod_atividade_str = self.combo_nfts_cod_atividade.get()
            cod_atividade = str(cod_atividade_str).split(" - ")[0].strip() 
            nome_certificado = obter_certificado_padrao()

            opts = Options()
            opts.add_argument("--start-maximized")
            driver = webdriver.Chrome(options=opts)
            driver.set_page_load_timeout(100)
            
            driver.get("https://nfe.prefeitura.sp.gov.br/login.aspx")
            wait = WebDriverWait(driver, 30)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Login')]"))).click()
            
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-target='#collapseOutrasFormas']"))).click()
                time.sleep(1)
            except: pass
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-target='#collapseCertificado']"))).click()
                time.sleep(1)
            except: pass
            
            wait.until(EC.element_to_be_clickable((By.ID, "btnCertificado"))).click()
            time.sleep(2)
            if nome_certificado:
                pyautogui.write(nome_certificado)
                time.sleep(1)
            pyautogui.press('enter')

            WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.ID, "ctl00_body_btAcesso"))).click()
            WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.LINK_TEXT, "Emissão de NFTS")))
            
            for i, (idx_orig, row) in enumerate(df_proc.iterrows()):
                if self.stop_nfts_event.is_set():
                    self.app.log_backend("Processo Parado pelo usuário.")
                    break
                while self.pause_nfts_event.is_set(): time.sleep(1)

                cnpj_tomador = limpar_cnpj(str(row.get('CNPJ', '')))
                tree_item_id = mapa_itens_tree.get(cnpj_tomador)
                def update_tree(status_text, remove_x=False):
                    if tree_item_id:
                        vals = list(self.tree_nfts.item(tree_item_id, "values"))
                        vals[6] = status_text
                        if remove_x: vals[0] = "[ ]"
                        self.app.root.after(0, lambda v=vals, i=tree_item_id: self.tree_nfts.item(i, values=v))
                        
                update_tree("Processando...")

                if cnpj_tomador in self.app.mapping_cnpj:
                    ccm_raw = self.app.mapping_cnpj[cnpj_tomador].get("CCM", "")
                else:
                    df_completo.at[idx_orig, "Status"] = "Erro: CNPJ não mapeado no Airtable"
                    update_tree("Erro: CNPJ não mapeado")
                    try: df_completo.to_excel(caminho_resultado, index=False)
                    except: pass
                    continue
                
                if not ccm_raw:
                    df_completo.at[idx_orig, "Status"] = "Erro: Cliente sem CCM no Airtable"
                    update_tree("Erro: Cliente sem CCM")
                    continue

                ccm_limpo = str(ccm_raw).replace(".", "").replace("-", "").strip().split(',')[0].lstrip('0')
                dt_obj = ext_dt(row.get("Data de Emissão"))
                if not dt_obj:
                    df_completo.at[idx_orig, "Status"] = "Erro: Data Inválida no Excel"
                    update_tree("Erro: Data Inválida")
                    continue
                
                mes_ano_ref = dt_obj.strftime('%m-%Y')
                pastas = self.garantir_pastas_nfts(mes_ano_ref)
                data_formatada = dt_obj.strftime('%d/%m/%Y')
                
                try:
                    driver.get("https://nfe.prefeitura.sp.gov.br/contribuinte/notatomador.aspx?tipo=t")
                    Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_ddlTomador")))).select_by_value(ccm_limpo)
                    driver.find_element(By.ID, "ctl00_body_tbCPFCNPJPrestador").send_keys(cnpj_prestador)
                    
                    f_data = wait.until(EC.presence_of_element_located((By.ID, "ctl00_body_tbDataFatoGerador")))
                    driver.execute_script(f"arguments[0].value = '{data_formatada}';", f_data)
                    f_data.send_keys(Keys.TAB)
                    time.sleep(2.5)
                    
                    try:
                        radio_iss = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_rbVersaoISS")))
                        radio_iss.click()
                    except: pass

                    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_btAvancar"))).click()
                    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_rblComEmissao"))).click()
                    driver.find_element(By.ID, "ctl00_body_tbNumDocFiscal").send_keys(str(row.get("Número da NF", row.get("Numero da NF", ""))))
                    Select(driver.find_element(By.ID, "ctl00_body_ddlItemSubItem")).select_by_visible_text("17.19 - Contabilidade, inclusive serviços técnicos e auxiliares.")
                    time.sleep(2.5) 
                    Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_ddlAtividade")))).select_by_value(cod_atividade)
                    time.sleep(1.5) 
                    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_body_rblSimplesNacional_0"))).click()
                    
                    val_raw = str(row["Valor da NF"]).strip().replace("R$", "").strip()
                    if "," in val_raw: val_raw = val_raw.replace(".", "").replace(",", ".")
                    val_f = f"{float(val_raw):.2f}".replace(".", ",")
                    driver.execute_script(f"document.getElementById('ctl00_body_tbValor').value = '{val_f}';")
                    
                    driver.find_element(By.ID, "ctl00_body_btEmitir").click()
                    try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
                    except: pass

                    wait.until(lambda d: "notatomadorprint.aspx" in d.current_url.lower())
                    time.sleep(2); pyautogui.press('enter'); time.sleep(2); pyautogui.press('enter'); time.sleep(3)
                    
                    nome_limpo_arq = re.sub(r'[\\/*?:"<>|]', "", str(row.get('Cliente', '')).strip())
                    caminho_final_pdf = os.path.abspath(os.path.join(pastas['pdf'], f"NFTS - {nome_limpo_arq}.pdf"))
                    pyperclip.copy(caminho_final_pdf); pyautogui.hotkey('ctrl', 'v'); time.sleep(1.5); pyautogui.press('enter')
                    
                    df_completo.at[idx_orig, "Status"] = "Emitida"
                    update_tree("Emitida", remove_x=True)
                    try: df_completo.to_excel(caminho_resultado, index=False)
                    except: pass
                    
                except Exception as e:
                    df_completo.at[idx_orig, "Status"] = f"Erro no portal SP"
                    update_tree("Erro no portal SP", remove_x=False) 
                    try: df_completo.to_excel(caminho_resultado, index=False)
                    except: pass

        finally:
            if df_completo is not None and caminho_resultado is not None:
                try: df_completo.to_excel(caminho_resultado, index=False)
                except: df_completo.to_excel(os.path.join(ROOT_DIR_NFTS, "RESULTADO_EMISSAO_BACKUP.xlsx"), index=False)
                
            if driver: driver.quit()
            self.app.root.after(0, lambda: self.btn_exec_nfts.config(text="🚀 EMITIR NFTS PENDENTES", state=tk.NORMAL))
            self.app.root.after(0, lambda: self.btn_pause_nfts.config(state=tk.DISABLED))
            self.app.root.after(0, lambda: self.btn_stop_nfts.config(state=tk.DISABLED, text="⏹ Parar"))
            
            if self.stop_nfts_event.is_set(): self.app.root.after(0, lambda: messagebox.showwarning("Parado", "O Processo de Emissão foi interrompido pelo usuário."))
            else: self.app.root.after(0, lambda: messagebox.showinfo("Concluído", "Processo de Emissão Finalizado!"))

    def iniciar_relatorio_nfts_wrapper(self):
        if not HAS_RPA: return messagebox.showerror("Bibliotecas Ausentes", "Instale selenium pyautogui pyperclip")
        if not self.app.mapping_cnpj: return messagebox.showwarning("Atenção", "O Airtable ainda não foi sincronizado.")
        if not obter_certificado_padrao(): return messagebox.showwarning("Atenção", "Certificado Digital não configurado.\nVá na aba Configurações.")
            
        self.btn_rel_nfts.config(state=tk.DISABLED, text="PROCESSANDO...")
        threading.Thread(target=self.relatorio_nfts_thread, daemon=True).start()

    def relatorio_nfts_thread(self):
        driver = None
        try:
            d_ini = self.nfts_d_ini.get()
            d_fim = self.nfts_d_fim.get()
            mes_ano_rel = d_ini[3:].replace('/', '-') 
            pastas = self.garantir_pastas_nfts(mes_ano_rel)
            nome_certificado = obter_certificado_padrao()

            opts = Options()
            opts.add_argument("--start-maximized")
            opts.add_experimental_option("prefs", {"download.default_directory": os.path.abspath(pastas['csv']), "download.prompt_for_download": False})
            driver = webdriver.Chrome(options=opts)
            
            driver.get("https://nfe.prefeitura.sp.gov.br/login.aspx")
            wait = WebDriverWait(driver, 30)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Login')]"))).click()
            
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-target='#collapseOutrasFormas']"))).click(); time.sleep(1)
            except: pass
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-target='#collapseCertificado']"))).click(); time.sleep(1)
            except: pass

            wait.until(EC.element_to_be_clickable((By.ID, "btnCertificado"))).click()
            time.sleep(2)
            if nome_certificado:
                pyautogui.write(nome_certificado)
                time.sleep(1)
            pyautogui.press('enter')

            WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.ID, "ctl00_body_btAcesso"))).click()
            WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.LINK_TEXT, "Emissão de NFTS")))
            
            driver.get("https://nfe.prefeitura.sp.gov.br/contribuinte/exportaarquivoNFTS.aspx")
            driver.execute_script(f"document.getElementById('ctl00_body_tbInicio').value = '{d_ini}';")
            driver.execute_script(f"document.getElementById('ctl00_body_tbFim').value = '{d_fim}';")
            
            ccms = []
            for _, row in self.df_nfts_completo.iterrows():
                cnpj_limpo = limpar_cnpj(row.get('CNPJ', ''))
                if cnpj_limpo in self.app.mapping_cnpj:
                    c_raw = self.app.mapping_cnpj[cnpj_limpo].get("CCM", "")
                    if c_raw: ccms.append(str(c_raw).replace(".", "").replace("-", "").strip().split(',')[0].lstrip('0'))

            for ccm_limpo in list(set(ccms)):
                try:
                    Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_body_ddlPrestador")))).select_by_value(ccm_limpo)
                    driver.find_element(By.ID, "ctl00_body_btGerar").click(); time.sleep(4)
                except: continue
            
            self.compilar_csvs_nfts(pastas['csv'], mes_ano_rel)
        finally:
            if driver: driver.quit()
            self.app.root.after(0, lambda: self.btn_rel_nfts.config(text="📊 GERAR RELATÓRIOS", state=tk.NORMAL))
            self.app.root.after(0, lambda: messagebox.showinfo("Concluído", "Downloads e Compilação Concluídos!"))