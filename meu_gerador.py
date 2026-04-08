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
        self.title("Flash Tool - v1.9.2")
        self.geometry("1300x900")
        
        # --- SOLUÇÃO PARA TELA BRANCA NO WINDOWS ---
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
        self.tab_snippets = self.tabview.add("Snippets / Biblioteca")

        self.setup_aba_conexao()
        self.setup_aba_main()
        self.setup_aba_snippets()

    def setup_aba_conexao(self):
        frame = ctk.CTkFrame(self.tab_conn)
        frame.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(frame, text="Configurações de Acesso Oracle", font=("Arial", 18, "bold")).pack(pady=10)
        
        self.ent_user = self.criar_campo(frame, "Usuário:")
        self.ent_pass = self.criar_campo(frame, "Senha:", show="*")
        self.ent_host = self.criar_campo(frame, "Host / IP:")
        self.ent_port = self.criar_campo(frame, "Porta:", placeholder_text="1521")
        self.ent_service = self.criar_campo(frame, "Service Name / SID:")
        
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
        # --- BOTÃO ALTERADO AQUI ---
        ctk.CTkButton(btn_f3, text="Copiar", fg_color="#d35400", command=self.copiar_texto).pack(side="left", padx=5)

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

    def copiar_texto(self):
        texto = self.txt_out.get("1.0", "end-1c")
        if texto.strip(): pyperclip.copy(texto); messagebox.showinfo("Copiado", "Copiado!")

    def testar_conexao(self):
        try:
            dsn = f"{self.ent_host.get()}:{self.ent_port.get()}/{self.ent_service.get()}"
            with oracledb.connect(user=self.ent_user.get(), password=self.ent_pass.get(), dsn=dsn) as conn:
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
                self.ent_user.insert(0, data.get("user", ""))
                self.ent_host.insert(0, data.get("host", ""))
                self.ent_port.insert(0, data.get("port", "1521"))
                self.ent_service.insert(0, data.get("service", ""))
                path_xml = data.get("xml_path", "")
                if path_xml:
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
            dsn = f"{self.ent_host.get()}:{self.ent_port.get()}/{self.ent_service.get()}"
            with oracledb.connect(user=self.ent_user.get(), password=self.ent_pass.get(), dsn=dsn) as conn:
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