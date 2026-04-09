import customtkinter as ctk
import pandas as pd
import oracledb
import json
import os
import re
import pyperclip
import xml.etree.ElementTree as ET
from tkinter import messagebox, filedialog
from datetime import datetime

# Configurações de Aparência
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class OracleSeniorApp(ctk.CTk):
    def __init__(self):  
        super().__init__()
        self.title("Flash Tool - v1.9.3") 
        self.geometry("1300x900")
        
       
        self.withdraw()
        self.after(10, self.deiconify)
        self.after(100, self.lift)
        
        self.config_file = "db_config.json"
        self.df_atual = None
        self.colunas_inputs = {}
        self.widgets_colunas = [] 
        self.snippets_data = {} 
        self.widgets_snippets = []

        self.setup_ui()
        self.carregar_config()

    def setup_ui(self):
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab_conn = self.tabview.add("Conexão")
        self.tab_main = self.tabview.add("Gerador de Update")
        self.tab_critica = self.tabview.add("Crítica SIPPES") # NOVA ABA
        self.tab_snippets = self.tabview.add("Snippets / Biblioteca")

        self.setup_aba_conexao()
        self.setup_aba_main()
        self.setup_aba_critica() 
        self.setup_aba_snippets()

    def setup_aba_conexao(self):
        frame = ctk.CTkFrame(self.tab_conn)
        frame.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(frame, text="Configurações de Acesso Oracle", font=("Arial", 18, "bold")).pack(pady=10)
        
        # Usuário travado
        self.ent_user = self.criar_campo(frame, "Usuário:")
        self.ent_user.insert(0, "CDS_FLASHTOOLS_APP") 
        self.ent_user.configure(state="readonly") 

        # Senha em foco
        self.ent_pass = self.criar_campo(frame, "Senha:", show="*")
        self.ent_pass.focus() 
        
        # Host padrão
        self.ent_host = self.criar_campo(frame, "Host / IP:")
        self.ent_host.insert(0, "exa12-scan1.eb.mil.br")
        
        # Porta padrão
        self.ent_port = self.criar_campo(frame, "Porta:")
        self.ent_port.insert(0, "1521")
        
        # Serviço padrão
        self.ent_service = self.criar_campo(frame, "Service Name / SID:")
        self.ent_service.insert(0, "SIPPES_HOMOL")
        
        ctk.CTkLabel(frame, text="Arquivo de Snippets (XML):", font=("Arial", 12)).pack(pady=(20, 0))
        f_xml = ctk.CTkFrame(frame, fg_color="transparent")
        f_xml.pack(fill="x", padx=100)
        self.ent_xml_path = ctk.CTkEntry(f_xml)
        self.ent_xml_path.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_xml, text="Selecionar XML", width=120, command=self.selecionar_xml).pack(side="right")

        btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        btn_row.pack(pady=30)
        ctk.CTkButton(btn_row, text="Testar Conexão", command=self.testar_conexao, fg_color="#2b5e91").pack(side="left", padx=10)
        ctk.CTkButton(btn_row, text="Salvar Configurações", command=self.salvar_config, fg_color="#28a745").pack(side="left", padx=10)

    # NOVA FUNCIONALIDADE: CRÍTICA SIPPES -- CHAMAAAAA!!!
    def setup_aba_critica(self):
        self.tab_critica.grid_columnconfigure(0, weight=1)
        self.tab_critica.grid_columnconfigure(1, weight=2)
        self.tab_critica.grid_rowconfigure(0, weight=1)

        # Frame da Esquerda (Parâmetros e Importação)
        f_params = ctk.CTkFrame(self.tab_critica)
        f_params.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        ctk.CTkLabel(f_params, text="1. Importar Planilha", font=("Arial", 13, "bold")).pack(pady=(10,5))
        
        f_planilha = ctk.CTkFrame(f_params, fg_color="transparent")
        f_planilha.pack(fill="x", padx=10, pady=5)
        self.ent_excel_path = ctk.CTkEntry(f_planilha, placeholder_text="Caminho do Excel...")
        self.ent_excel_path.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkButton(f_planilha, text="Procurar", width=80, command=self.selecionar_excel_critica).pack(side="right")
        
        ctk.CTkButton(f_params, text="Carregar Tabela CODIGO_SIPPES", fg_color="#d35400", command=self.importar_codigos_sippes).pack(pady=10, fill="x", padx=10)
        
        ctk.CTkLabel(f_params, text="2. Parâmetros dos Scripts", font=("Arial", 13, "bold")).pack(pady=(20,5))
        
        self.ent_folha_pgto = self.criar_campo_simples(f_params, "FOLHA_PGTO_CODIGO:")
        self.ent_sit_lanc = self.criar_campo_simples(f_params, "TIPO_SIT_LANC_CODIGO:")
        self.ent_status_alt = self.criar_campo_simples(f_params, "TIPO_STATUS_ALT_CODIGO:")
        self.ent_motivo_cpex = self.criar_campo_simples(f_params, "MOTIVO_CPEX:")
        
        ctk.CTkLabel(f_params, text="3. Gerar Comandos", font=("Arial", 13, "bold")).pack(pady=(20,5))
        ctk.CTkButton(f_params, text="Gerar LANCAMENTO (3.1)", command=self.gerar_script_lancamento).pack(pady=5, fill="x", padx=10)
        ctk.CTkButton(f_params, text="Gerar LANCAMENTO_ALT (3.2)", command=self.gerar_script_lancamento_alt).pack(pady=5, fill="x", padx=10)

        # Frame da Direita (Resultados)
        f_result = ctk.CTkFrame(self.tab_critica)
        f_result.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        ctk.CTkLabel(f_result, text="Scripts Gerados", font=("Arial", 13, "bold")).pack(pady=5)
        self.txt_out_critica = ctk.CTkTextbox(f_result)
        self.txt_out_critica.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkButton(f_result, text="Copiar Scripts", fg_color="#28a745", command=lambda: self.copiar_texto(self.txt_out_critica)).pack(pady=10)

    def criar_campo_simples(self, master, texto):
        f = ctk.CTkFrame(master, fg_color="transparent")
        f.pack(fill="x", padx=10, pady=2)
        ctk.CTkLabel(f, text=texto, width=170, anchor="w").pack(side="left")
        ent = ctk.CTkEntry(f)
        ent.pack(side="right", fill="x", expand=True)
        return ent

    def selecionar_excel_critica(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if path:
            self.ent_excel_path.delete(0, 'end')
            self.ent_excel_path.insert(0, path)

    def obter_conexao(self):
        dsn = f"{self.ent_host.get()}:{self.ent_port.get()}/{self.ent_service.get()}"
        return oracledb.connect(user=self.ent_user.get(), password=self.ent_pass.get(), dsn=dsn)

    def importar_codigos_sippes(self):
        path = self.ent_excel_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Selecione um arquivo de planilha válido.")
            return
            
        try:
            # Lê o arquivo
            if path.endswith('.csv'): df = pd.read_csv(path)
            else: df = pd.read_excel(path)
            
            # Tenta achar a coluna CODIGO ignorando maiúsculas/minúsculas
            coluna_codigo = None
            for col in df.columns:
                if str(col).upper().strip() == 'CODIGO':
                    coluna_codigo = col
                    break
            
            if not coluna_codigo:
                coluna_codigo = df.columns[0] # Fallback para a primeira coluna
            
            # Extrai os números brutos da planilha como lista
            codigos_brutos = df[coluna_codigo].dropna().astype(str).str.extract(r'(\d+)')[0].dropna().tolist()
            total_lidos = len(codigos_brutos)
            
            if total_lidos == 0:
                messagebox.showwarning("Aviso", "Nenhum código numérico encontrado na planilha.")
                return

            # Auditoria e Duplicatas
            from collections import Counter
            contagem = Counter(codigos_brutos)
            
            # Filtra os códigos que apareceram mais de uma vez
            duplicados = [codigo for codigo, qtd in contagem.items() if qtd > 1]
            
            # Transforma em 'set' para remover duplicatas e depois volta para lista para o INSERT
            codigos_unicos = list(set(codigos_brutos))
            total_unicos = len(codigos_unicos)
            
            # Prepara a tupla para o executemany
            codigos_para_inserir = [(int(c),) for c in codigos_unicos]

            # Conexão e Inserção
            with self.obter_conexao() as conn:
                cursor = conn.cursor()
                cursor.execute("TRUNCATE TABLE CDS_FLASHTOOLS_APP.CODIGO_SIPPES")
                cursor.executemany("INSERT INTO CDS_FLASHTOOLS_APP.CODIGO_SIPPES (CODIGO) VALUES (:1)", codigos_para_inserir)
                conn.commit()
                
            # Chama a nova tela de resumo com opção de copiar
            self.mostrar_resumo_importacao(total_lidos, total_unicos, duplicados)
            
        except Exception as e:
            messagebox.showerror("Erro na Importação", str(e))

    def mostrar_resumo_importacao(self, total_lidos, total_unicos, duplicados):
        # Cria uma nova janelinha sobreposta
        janela = ctk.CTkToplevel(self)
        janela.title("Resumo da Importação")
        janela.geometry("450x450")
        janela.attributes("-topmost", True) # Mantém a janela sempre na frente
        janela.transient(self) # Foca nela
        
        # Textos de cabeçalho
        msg_header = f"Importação concluída!\n\n• Linhas lidas: {total_lidos}\n• Inseridos: {total_unicos}"
        ctk.CTkLabel(janela, text=msg_header, font=("Arial", 14, "bold"), justify="left").pack(pady=(15, 5))
        
        if duplicados:
            msg_alerta = f"⚠️ Foram encontrados {len(duplicados)} código(s) duplicado(s):"
            ctk.CTkLabel(janela, text=msg_alerta, font=("Arial", 12), text_color="#f39c12").pack(pady=5)
            
            # Caixa de texto rolável com os duplicados (permite seleção do mouse)
            txt_dups = ctk.CTkTextbox(janela, height=200)
            txt_dups.pack(fill="both", expand=True, padx=20, pady=5)
            
            # Formata a lista um embaixo do outro e insere na caixa
            lista_formatada = "\n".join(duplicados)
            txt_dups.insert("1.0", lista_formatada)
            txt_dups.configure(state="disabled") # Desabilita digitação, mas permite cópia
            
            # Botão de copiar direto
            def acao_copiar():
                pyperclip.copy(lista_formatada)
                messagebox.showinfo("Copiado", "Códigos duplicados copiados!", parent=janela)
                
            btn_row = ctk.CTkFrame(janela, fg_color="transparent")
            btn_row.pack(pady=15)
            ctk.CTkButton(btn_row, text="Copiar Duplicados", fg_color="#d35400", command=acao_copiar).pack(side="left", padx=10)
            ctk.CTkButton(btn_row, text="Fechar", fg_color="#555555", command=janela.destroy).pack(side="left", padx=10)
        else:
            ctk.CTkButton(janela, text="OK", command=janela.destroy).pack(pady=20)

    def mostrar_resumo_importacao(self, total_lidos, total_unicos, duplicados):
        
        janela = ctk.CTkToplevel(self)
        janela.title("Resumo da Importação")
        janela.geometry("450x450")
        janela.attributes("-topmost", True) 
        janela.transient(self) 
        
        
        msg_header = f"Importação concluída!\n\n• Linhas lidas: {total_lidos}\n• Inseridos: {total_unicos}"
        ctk.CTkLabel(janela, text=msg_header, font=("Arial", 14, "bold"), justify="left").pack(pady=(15, 5))
        
        if duplicados:
            msg_alerta = f"⚠️ Foram encontrados {len(duplicados)} código(s) duplicado(s):"
            ctk.CTkLabel(janela, text=msg_alerta, font=("Arial", 12), text_color="#f39c12").pack(pady=5)
            
            
            txt_dups = ctk.CTkTextbox(janela, height=200)
            txt_dups.pack(fill="both", expand=True, padx=20, pady=5)
            
          
            lista_formatada = "\n".join(duplicados)
            txt_dups.insert("1.0", lista_formatada)
            txt_dups.configure(state="disabled") # Desabilita digitação, mas permite cópia
            
           
            def acao_copiar():
                pyperclip.copy(lista_formatada)
                messagebox.showinfo("Copiado", "Códigos duplicados copiados!", parent=janela)
                
            btn_row = ctk.CTkFrame(janela, fg_color="transparent")
            btn_row.pack(pady=15)
            ctk.CTkButton(btn_row, text="Copiar Duplicados", fg_color="#d35400", command=acao_copiar).pack(side="left", padx=10)
            ctk.CTkButton(btn_row, text="Fechar", fg_color="#555555", command=janela.destroy).pack(side="left", padx=10)
        else:
            ctk.CTkButton(janela, text="OK", command=janela.destroy).pack(pady=20)

    def gerar_script_lancamento(self):
        folha = self.ent_folha_pgto.get().strip()
        sit_lanc = self.ent_sit_lanc.get().strip()
        
        if not folha or not sit_lanc:
            messagebox.showwarning("Aviso", "Preencha os campos FOLHA_PGTO_CODIGO e TIPO_SIT_LANC_CODIGO.")
            return
            
        
        query = f"""
        SELECT 
        '-- ATUALIZANDO A TABELA SIPPES.LANCAMENTO'||CHR(13)||
        '-- CAMPO TIPO_SIT_LANC_CODIGO DE '||L.TIPO_SIT_LANC_CODIGO||' PARA {sit_lanc}'||CHR(13)||
        'UPDATE SIPPES.LANCAMENTO SET TIPO_SIT_LANC_CODIGO = {sit_lanc} WHERE FOLHA_PGTO_CODIGO = {folha} AND CODIGO = '||CODIGO||';'||CHR(13) AS CMDO 
        FROM SIPPES.LANCAMENTO L
        WHERE CODIGO IN (SELECT CODIGO FROM CDS_FLASHTOOLS_APP.CODIGO_SIPPES)
        """
        self.executar_query_critica(query)

    def gerar_script_lancamento_alt(self):
        folha = self.ent_folha_pgto.get().strip()
        sit_lanc = self.ent_sit_lanc.get().strip()
        status_alt = self.ent_status_alt.get().strip()
        motivo_cpex = self.ent_motivo_cpex.get().strip()
        
        if not all([folha, sit_lanc, status_alt, motivo_cpex]):
            messagebox.showwarning("Aviso", "Preencha FOLHA_PGTO_CODIGO, TIPO_SIT_LANC_CODIGO, TIPO_STATUS_ALT_CODIGO e MOTIVO_CPEX.")
            return
            
    
        query = f"""
        SELECT 
        '-- ATUALIZANDO A TABELA SIPPES.LANCAMENTO_ALT'||CHR(13)||
        '-- CAMPO TIPO_STATUS_ALT_CODIGO DE '||L.TIPO_STATUS_ALT_CODIGO||' PARA {status_alt}'||CHR(13)||
        '-- CAMPO TIPO_SIT_LANC_CODIGO DE '||L.TIPO_SIT_LANC_CODIGO||' PARA {sit_lanc}'||CHR(13)||
        '-- CAMPO MOTIVO_CPEX DE '||L.MOTIVO_CPEX||' PARA ABAIXO'||CHR(13)||
        'UPDATE SIPPES.LANCAMENTO_ALT SET TIPO_STATUS_ALT_CODIGO = {status_alt}, TIPO_SIT_LANC_CODIGO = {sit_lanc}, MOTIVO = '||CHR(39)||'O motivo da exclusão está disponível no relatório Lançamentos Excluídos pelo CPEx na ferramenta SINFOPPES (http://sippes.cpex.eb.mil.br/)'||CHR(39)||', MOTIVO_CPEX = '||CHR(39)||'{motivo_cpex}'||CHR(39)||' WHERE FOLHA_PGTO_CODIGO = {folha} AND CODIGO = '||CODIGO||';'||CHR(13) AS CMDO 
        FROM SIPPES.LANCAMENTO_ALT L
        WHERE CODIGO IN (SELECT CODIGO FROM CDS_FLASHTOOLS_APP.CODIGO_SIPPES)
        """
        self.executar_query_critica(query)

    def executar_query_critica(self, query):
        self.txt_out_critica.delete("1.0", "end")
        self.txt_out_critica.insert("end", "-- Consultando banco de dados, aguarde...\n")
        self.update() # Força a interface a atualizar a mensagem
        
        try:
            with self.obter_conexao() as conn:
                df = pd.read_sql(query, conn)
                
            if df.empty:
                self.txt_out_critica.delete("1.0", "end")
                self.txt_out_critica.insert("end", "-- Nenhum registro encontrado com os códigos informados e a folha de pagamento.")
                return
                
            # Concatena todos os comandos gerados
            comandos = df['CMDO'].tolist()
            script_final = "\n".join(comandos)
            
            self.txt_out_critica.delete("1.0", "end")
            self.txt_out_critica.insert("end", "-- ==========================================\n")
            self.txt_out_critica.insert("end", f"-- SCRIPTS GERADOS ({len(comandos)} REGISTROS)\n")
            self.txt_out_critica.insert("end", "-- ==========================================\n\n")
            self.txt_out_critica.insert("end", script_final)
            self.txt_out_critica.insert("end", "\nCOMMIT;\n")
            
        except Exception as e:
            messagebox.showerror("Erro SQL", str(e))
            self.txt_out_critica.delete("1.0", "end")
            self.txt_out_critica.insert("end", f"ERRO:\n{str(e)}")

    # Nao mudei nada do teu codigo lucao, chamaaaa!!!
    def setup_aba_main(self):
        self.tab_main.grid_columnconfigure(0, weight=1)
        self.tab_main.grid_columnconfigure(1, weight=1)
        self.tab_main.grid_columnconfigure(2, weight=1)
        self.tab_main.grid_rowconfigure(0, weight=1)
        
        f1 = ctk.CTkFrame(self.tab_main); f1.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(f1, text="1. SQL Select Original", font=("Arial", 13, "bold")).pack(pady=5)
        self.txt_sql = ctk.CTkTextbox(f1, height=300); self.txt_sql.pack(fill="both", expand=True, padx=5)
        ctk.CTkButton(f1, text="Buscar Dados no Banco", command=self.buscar_dados_reais).pack(pady=10)
        
        f2 = ctk.CTkFrame(self.tab_main); f2.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(f2, text="2. Valores para Alterar", font=("Arial", 13, "bold")).pack(pady=5)
        self.ent_busca = ctk.CTkEntry(f2, placeholder_text="🔍 Filtrar coluna...")
        self.ent_busca.pack(fill="x", padx=10, pady=5)
        self.ent_busca.bind("<KeyRelease>", self.filtrar_colunas) 
        self.frame_campos = ctk.CTkScrollableFrame(f2)
        self.frame_campos.pack(fill="both", expand=True, padx=5, pady=5)
        
        f3 = ctk.CTkFrame(self.tab_main); f3.grid(row=0, column=2, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(f3, text="3. Auditoria e Script", font=("Arial", 13, "bold")).pack(pady=5)
        self.txt_out = ctk.CTkTextbox(f3); self.txt_out.pack(fill="both", expand=True, padx=5)
        btn_f3 = ctk.CTkFrame(f3, fg_color="transparent"); btn_f3.pack(fill="x", pady=10)
        ctk.CTkButton(btn_f3, text="Gerar Relatório e SQL", fg_color="darkgreen", command=self.gerar_log_e_sql).pack(side="left", padx=5, expand=True)
        ctk.CTkButton(btn_f3, text="Copiar", fg_color="#d35400", command=lambda: self.copiar_texto(self.txt_out)).pack(side="left", padx=5)

    def setup_aba_snippets(self):
        self.tab_snippets.grid_columnconfigure(0, weight=1)
        self.tab_snippets.grid_columnconfigure(1, weight=3)
        self.tab_snippets.grid_rowconfigure(0, weight=0) 
        self.tab_snippets.grid_rowconfigure(1, weight=1) 

        f_busca_snip = ctk.CTkFrame(self.tab_snippets, fg_color="transparent")
        f_busca_snip.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))
        self.ent_busca_snip = ctk.CTkEntry(f_busca_snip, placeholder_text="🔍 Buscar snippet...")
        self.ent_busca_snip.pack(fill="x")
        self.ent_busca_snip.bind("<KeyRelease>", self.filtrar_snippets)

        self.scroll_snippets = ctk.CTkScrollableFrame(self.tab_snippets, label_text="Snippets Disponíveis")
        self.scroll_snippets.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        f_edit = ctk.CTkFrame(self.tab_snippets)
        f_edit.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(f_edit, text="Editor de Snippet", font=("Arial", 13, "bold")).pack(pady=5)
        self.txt_edit_snippet = ctk.CTkTextbox(f_edit)
        self.txt_edit_snippet.pack(fill="both", expand=True, padx=5)
        ctk.CTkButton(f_edit, text="Injetar no SQL Original", fg_color="#d35400", command=self.injetar_snippet).pack(pady=10)

    def filtrar_snippets(self, event=None):
        termo = self.ent_busca_snip.get().lower()
        for widget, nome_snip in self.widgets_snippets:
            if termo in nome_snip.lower(): 
                widget.pack(fill="x", pady=1)
            else: 
                widget.pack_forget()

    def selecionar_xml(self):
        path = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
        if path:
            self.ent_xml_path.delete(0, 'end')
            self.ent_xml_path.insert(0, path)
            self.carregar_snippets_xml(path)

    def carregar_snippets_xml(self, path):
        if not os.path.exists(path): return
        try:
            for w in self.scroll_snippets.winfo_children(): w.destroy()
            self.snippets_data = {}
            self.widgets_snippets = []
            
            tree = ET.parse(path)
            root = tree.getroot()
            
            for snip in root.iter('snippet'):
                nome = snip.get('name') or "Sem Nome"
                code_elem = snip.find('code')
                if code_elem is None: code_elem = snip.find('codeSnippet')
                
                codigo = code_elem.text if code_elem is not None else ""
                if codigo:
                    self.snippets_data[nome] = codigo.strip()
                    btn = ctk.CTkButton(self.scroll_snippets, text=nome, anchor="w", fg_color="transparent", 
                                        text_color="white", hover_color="#3d3d3d",
                                        command=lambda n=nome: self.visualizar_snippet(n))
                    btn.pack(fill="x", pady=1)
                    self.widgets_snippets.append((btn, nome))
        except Exception as e: messagebox.showerror("Erro XML", f"Falha ao ler snippets: {e}")

    def visualizar_snippet(self, nome):
        self.txt_edit_snippet.delete("1.0", "end")
        self.txt_edit_snippet.insert("end", self.snippets_data.get(nome, ""))

    def injetar_snippet(self):
        codigo = self.txt_edit_snippet.get("1.0", "end-1c")
        if not codigo.strip(): return
        self.tabview.set("Gerador de Update")
        self.txt_sql.delete("1.0", "end")
        self.txt_sql.insert("end", codigo)

    def criar_campo(self, master, texto, **kwargs):
        f = ctk.CTkFrame(master, fg_color="transparent"); f.pack(fill="x", padx=100, pady=5)
        ctk.CTkLabel(f, text=texto, width=150, anchor="w").pack(side="left")
        ent = ctk.CTkEntry(f, **kwargs); ent.pack(side="right", fill="x", expand=True)
        return ent

    def filtrar_colunas(self, event=None):
        termo = self.ent_busca.get().lower()
        for widget, nome_col in self.widgets_colunas:
            if termo in nome_col.lower(): widget.pack(fill="x", pady=2)
            else: widget.pack_forget()

    # Modificado esse krai pra receber o widget Textbox dinamicamente
    def copiar_texto(self, widget_textbox):
        texto = widget_textbox.get("1.0", "end-1c")
        if texto.strip(): 
            pyperclip.copy(texto)
            messagebox.showinfo("Copiado", "Copiado!")

    def testar_conexao(self):
        try:
            with self.obter_conexao() as conn:
                messagebox.showinfo("Sucesso", "Conexão OK!")
        except Exception as e: messagebox.showerror("Erro", str(e))

    def salvar_config(self):
        data = {
            "user": self.ent_user.get(), "host": self.ent_host.get(), "port": self.ent_port.get(), 
            "service": self.ent_service.get(), "xml_path": self.ent_xml_path.get()
        }
        with open(self.config_file, "w") as f: json.dump(data, f)
        messagebox.showinfo("Sucesso", "Salvo!")

    def carregar_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                data = json.load(f)
                
                self.ent_user.configure(state="normal")
                self.ent_user.delete(0, 'end')
                self.ent_user.insert(0, data.get("user", "CDS_FLASHTOOLS_APP"))
                self.ent_user.configure(state="readonly")
                
                self.ent_host.delete(0, 'end')
                self.ent_host.insert(0, data.get("host", "exa12-scan1.eb.mil.br")) 
                
                self.ent_port.delete(0, 'end')
                self.ent_port.insert(0, data.get("port", "1521"))
                
                self.ent_service.delete(0, 'end')
                self.ent_service.insert(0, data.get("service", "SIPPES_HOMOL"))
                
                path_xml = data.get("xml_path", "")
                if path_xml:
                    self.ent_xml_path.delete(0, 'end')
                    self.ent_xml_path.insert(0, path_xml)
                    self.carregar_snippets_xml(path_xml)

    def tratar_query_data(self, query):
        padrao_data = r"=\s*'(\d{2}/\d{2}/\d{4})'"
        def substituir(match): return f"= TO_DATE('{match.group(1)}', 'DD/MM/YYYY')"
        return re.sub(padrao_data, substituir, query, flags=re.IGNORECASE)

    def buscar_dados_reais(self):
        self.txt_out.delete("1.0", "end") 
        query = self.txt_sql.get("1.0", "end-1c").strip()
        if not query.lower().startswith("select"):
            messagebox.showerror("Segurança", "Apenas SELECT é permitido."); return
        query = self.tratar_query_data(query)
        if query.endswith(';'): query = query[:-1]
        try:
            with self.obter_conexao() as conn:
                self.df_atual = pd.read_sql(query, conn)
            for w in self.frame_campos.winfo_children(): w.destroy()
            self.colunas_inputs = {}; self.widgets_colunas = []
            self.ent_busca.delete(0, 'end') 
            for col in self.df_atual.columns:
                row_f = ctk.CTkFrame(self.frame_campos); row_f.pack(fill="x", pady=2)
                ctk.CTkLabel(row_f, text=col, width=150, anchor="w").pack(side="left", padx=5)
                ent = ctk.CTkEntry(row_f, placeholder_text="Novo valor..."); ent.pack(side="right", fill="x", expand=True, padx=5)
                self.colunas_inputs[col] = ent
                self.widgets_colunas.append((row_f, col))
            messagebox.showinfo("Sucesso", f"{len(self.df_atual)} linhas carregadas.")
        except Exception as e: messagebox.showerror("Erro SQL", str(e))

    def formatar_sql_oracle(self, col, val):
        val = str(val).strip()
        if val.upper() == "NULL" or not val: return "NULL"
        if any(x in col.upper() for x in ["DATA", "DATE", "DT_"]):
            data_limpa = val.split(' ')[0]
            if '-' in data_limpa:
                try: data_limpa = datetime.strptime(data_limpa, '%Y-%m-%d').strftime('%d/%m/%Y')
                except: pass
            return f"TO_DATE('{data_limpa}', 'DD/MM/YYYY')"
        if val.isdigit() and not val.startswith('0'): return val
        return f"'{val.replace("'", "''")}'"

    def gerar_log_e_sql(self):
        if self.df_atual is None or self.df_atual.empty: return
        query_orig = self.txt_sql.get("1.0", "end-1c")
        try:
            tabela_match = re.search(r'FROM\s+([^\s\n,;]+)', query_orig, re.IGNORECASE)
            tabela = tabela_match.group(1) if tabela_match else "TABELA"
            
            where_match = re.search(r'WHERE\s+(.*)', query_orig, re.IGNORECASE | re.DOTALL)
            if where_match:
                where_raw = where_match.group(1)
                where_orig = re.split(r'\s+ORDER\s+BY\s+', where_raw, flags=re.IGNORECASE)[0].strip()
            else:
                where_orig = "1=1"
                
            if where_orig.endswith(';'): where_orig = where_orig[:-1]
        except: 
            messagebox.showerror("Erro", "Falha no SQL."); return

        relatorio = [f"--- RELATÓRIO DE ALTERAÇÕES ({datetime.now().strftime('%H:%M:%S')}) ---\n"]
        sets_consolidado = []
        sets_rollback_consolidado = [] 

        pk_col = self.df_atual.columns[0]

        for _, row in self.df_atual.iterrows():
            sets_linha = []
            sets_linha_rollback = []
            log_antes = []
            log_atual = []
            id_registro = str(row[pk_col])
            
            for col, entry in self.colunas_inputs.items():
                val_novo = entry.get().strip()
                val_antigo = row[col]
                
                if pd.api.types.is_datetime64_any_dtype(type(val_antigo)) or any(x in col.upper() for x in ["DATA", "DATE", "DT_"]):
                    if pd.notnull(val_antigo):
                        if hasattr(val_antigo, 'strftime'): val_ant_str = val_antigo.strftime('%d/%m/%Y')
                        else:
                            temp_dt = str(val_antigo).split(' ')[0]
                            try: val_ant_str = datetime.strptime(temp_dt, '%Y-%m-%d').strftime('%d/%m/%Y')
                            except: val_ant_str = temp_dt
                    else: val_ant_str = "NULL"
                else: val_ant_str = str(val_antigo) if pd.notnull(val_antigo) else "NULL"
                
                if val_novo and val_novo != val_ant_str:
                    log_antes.append(f"{col}: {val_ant_str}")
                    log_atual.append(f"{col}: {val_novo}")
                    
                    sets_linha.append(f"{col} = {self.formatar_sql_oracle(col, val_novo)}")
                    sets_linha_rollback.append(f"{col} = {self.formatar_sql_oracle(col, val_ant_str)}")
            
            if sets_linha:
                relatorio.append(f"TABELA -> {tabela}\n{pk_col} -> {id_registro}")
                relatorio.append(f"ANTES  -> {' | '.join(log_antes)}\nATUAL  -> {' | '.join(log_atual)}")
                relatorio.append("-" * 40)
                
                if not sets_consolidado: 
                    sets_consolidado = sets_linha
                if not sets_rollback_consolidado: 
                    sets_rollback_consolidado = sets_linha_rollback

        if sets_consolidado:
            res = "\n".join(relatorio) 
            res += f"\n\n-- SCRIPT DE UPDATE --\nUPDATE {tabela}\nSET {', '.join(sets_consolidado)}\nWHERE {where_orig};\n\nCOMMIT;"
            res += f"\n\n-- SCRIPT DE ROLLBACK (SE NECESSÁRIO) --\nUPDATE {tabela}\nSET {', '.join(sets_rollback_consolidado)}\nWHERE {where_orig};\n\nCOMMIT;"
            
            self.txt_out.delete("1.0", "end")
            self.txt_out.insert("end", res)
        else: 
            messagebox.showinfo("Aviso", "Sem alterações.")

if __name__ == "__main__":
    app = OracleSeniorApp()
    app.mainloop()