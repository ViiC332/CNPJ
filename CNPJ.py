import tkinter as tk
from tkinter import ttk, messagebox
import requests
import os
import json
from PIL import Image, ImageTk
from datetime import datetime
from num2words import num2words
from babel.numbers import format_currency

# Configuração de estilo moderno
BG_COLOR = "#f0f0f0"
FG_COLOR = "#333333"
ACCENT_COLOR = "#4a6fa5"
ENTRY_COLOR = "#ffffff"
TEXT_BG = "#ffffff"

class CNPJApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta de CNPJ")
        self.root.geometry("1000x700")
        self.root.configure(bg=BG_COLOR)
        
        # Carregar ícone
        self.load_icon()
        
        # Carregar preferências
        self.preferences = self.load_preferences()
        
        # Adicionando MenuStrip
        self.create_menu()
        
        self.create_widgets()
        
        # Definir foco na textbox ao iniciar
        self.cnpj_entry.focus_set()
    
    def load_icon(self):
        try:
            ico_path = os.path.join(os.path.dirname(__file__), "Assets", "icon_multi.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
                # Para Windows
                if os.name == 'nt':
                    import ctypes
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"cnpj.app.id")
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")
    
    def get_preferences_path(self):
        # Salvar em Documentos (1)
        documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        return os.path.join(documents_path, 'CNPJConsult_preferences.json')
    
    def get_historico_path(self):
        documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        return os.path.join(documents_path, 'HistoricoConsult_preferences.json')
    
    def load_preferences(self):
        try:
            pref_path = self.get_preferences_path()
            if os.path.exists(pref_path):
                with open(pref_path, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Erro ao carregar preferências: {e}")
        return {
            "copiar_nome_empresarial": True,
            "copiar_cnpj": True,
            "copiar_telefone": True,
            "copiar_endereco": True,
            "alertas_situacao": False
        }
    
    def save_preferences(self):
        try:
            pref_path = self.get_preferences_path()
            with open(pref_path, 'w') as f:
                json.dump(self.preferences, f)
            return pref_path
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar preferências: {e}")
            return None
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # Menu Preferências
        pref_menu = tk.Menu(menubar, tearoff=0)
        pref_menu.add_command(label="Salvar preferências", command=self.salvar_preferencias)
        menubar.add_cascade(label="Preferências", menu=pref_menu)
        
        # Menu Histórico
        historico_menu = tk.Menu(menubar, tearoff=0)
        historico_menu.add_command(label="Ver Histórico", command=self.mostrar_historico)
        menubar.add_cascade(label="Histórico", menu=historico_menu)
        
        self.root.config(menu=menubar)
    
    def salvar_preferencias(self):
        top = tk.Toplevel(self.root)
        top.title("Salvar Preferências")
        top.geometry("300x250")
        top.resizable(False, False)
        
        try:
            ico_path = os.path.join(os.path.dirname(__file__), "Assets", "icon_multi.ico")
            if os.path.exists(ico_path):
                top.iconbitmap(ico_path)
        except:
            pass
        
        self.copiar_nome_var = tk.BooleanVar(value=self.preferences.get("copiar_nome_empresarial", True))
        self.copiar_cnpj_var = tk.BooleanVar(value=self.preferences.get("copiar_cnpj", True))
        self.copiar_telefone_var = tk.BooleanVar(value=self.preferences.get("copiar_telefone", True))
        self.copiar_endereco_var = tk.BooleanVar(value=self.preferences.get("copiar_endereco", True))
        self.alertas_var = tk.BooleanVar(value=self.preferences.get("alertas_situacao", False))
        
        tk.Checkbutton(top, text="Copiar Nome empresarial", variable=self.copiar_nome_var, font=('Arial', 11)).pack(anchor=tk.W, pady=(10, 0), padx=20)
        tk.Checkbutton(top, text="Copiar CNPJ", variable=self.copiar_cnpj_var, font=('Arial', 11)).pack(anchor=tk.W, pady=(10, 0), padx=20)
        tk.Checkbutton(top, text="Copiar Telefone", variable=self.copiar_telefone_var, font=('Arial', 11)).pack(anchor=tk.W, pady=(10, 0), padx=20)
        tk.Checkbutton(top, text="Copiar Endereço", variable=self.copiar_endereco_var, font=('Arial', 11)).pack(anchor=tk.W, pady=(10, 0), padx=20)
        tk.Checkbutton(top, text="Alertas de Situação Cadastral", variable=self.alertas_var, font=('Arial', 11)).pack(anchor=tk.W, pady=(10, 0), padx=20)
        
        tk.Button(top, text="Salvar", command=lambda: self.save_prefs_and_close(top), 
                 bg=ACCENT_COLOR, fg="white", font=('Arial', 11), padx=20).pack(pady=20)
    
    def save_prefs_and_close(self, top):
        self.preferences = {
            "copiar_nome_empresarial": self.copiar_nome_var.get(),
            "copiar_cnpj": self.copiar_cnpj_var.get(),
            "copiar_telefone": self.copiar_telefone_var.get(),
            "copiar_endereco": self.copiar_endereco_var.get(),
            "alertas_situacao": self.alertas_var.get()
        }
        saved_path = self.save_preferences()
        top.destroy()
        if saved_path:
            messagebox.showinfo("Sucesso", f"Preferências salvas com sucesso em:\n{saved_path}")

    def copiar_informacoes(self):
        if not hasattr(self, 'empresa_data'):
            messagebox.showwarning("Aviso", "Nenhum CNPJ consultado ainda.")
            return
        texto_para_copiar = ""

        if self.preferences.get("copiar_nome_empresarial", True):
            nome = self.get_nested_value(self.empresa_data, "company.name")
            fantasia = self.get_nested_value(self.empresa_data, "alias")
            texto_para_copiar += f"Nome Empresarial: {nome}\nNome Fantasia: {fantasia}\n\n"
        
        if self.preferences.get("copiar_cnpj", True):
            cnpj = self.get_nested_value(self.empresa_data, "taxId")
            texto_para_copiar += f"CNPJ: {self.format_cnpj(cnpj)}\n\n"
        
        if self.preferences.get("copiar_telefone", True):
            telefone = self.format_phone(self.empresa_data.get("phones", []))
            texto_para_copiar += f"Telefone: {telefone}\n\n"
        
        if self.preferences.get("copiar_endereco", True):
            endereco = self.format_address(self.empresa_data.get("address", {}))
            texto_para_copiar += f"Endereço:\n{endereco}\n"
        
        if texto_para_copiar:
            self.root.clipboard_clear()
            self.root.clipboard_append(texto_para_copiar)
            messagebox.showinfo("Copiado", "Informações copiadas para a área de transferência!")
        else:
            messagebox.showwarning("Aviso", "Nenhuma opção de cópia selecionada nas preferências.")
    
    def salvar_no_historico(self, nome, cnpj):
        try:
            cnpj_limpo = ''.join(filter(str.isdigit, cnpj))
            cnpj_formatado = cnpj_limpo.zfill(14)[:14]
            historico_path = self.get_historico_path()
            historico = []
            
            if os.path.exists(historico_path):
                with open(historico_path, 'r') as f:
                    historico = json.load(f)
            
            existe = any(item['cnpj'] == cnpj_formatado for item in historico)
            if not existe:
                historico.append({
                    'nome': nome,
                    'cnpj': cnpj_formatado,
                    'data': datetime.now().strftime("%d/%m/%Y %H:%M")
                })
                with open(historico_path, 'w') as f:
                    json.dump(historico, f, indent=4)
                    
        except Exception as e:
            print(f"Erro ao salvar histórico: {e}")

    def sort_historico(self, tree, historico, sort_type):
        tree.delete(*tree.get_children())
        
        if sort_type == "name_asc":
            sorted_items = sorted(historico, key=lambda x: x['nome'].upper())
        elif sort_type == "date_desc":
            sorted_items = sorted(historico, key=lambda x: datetime.strptime(x['data'], "%d/%m/%Y %H:%M"), reverse=True)
        elif sort_type == "date_asc":
            sorted_items = sorted(historico, key=lambda x: datetime.strptime(x['data'], "%d/%m/%Y %H:%M"))
        else:
            sorted_items = historico
        
        for item in sorted_items:
            tree.insert("", tk.END, values=(item['nome'], item['cnpj'], item['data']))

    def mostrar_historico(self):
        try:
            historico_path = self.get_historico_path()
            if not os.path.exists(historico_path):
                messagebox.showinfo("Histórico", "Nenhuma consulta realizada ainda.")
                return
                
            with open(historico_path, 'r') as f:
                historico = json.load(f)
                
            top = tk.Toplevel(self.root)
            top.title("Histórico de Consultas")
            top.geometry("600x500")
            top.resizable(False, False)
            
            filter_frame = ttk.Frame(top)
            filter_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Button(filter_frame, text="Ordenar por Nome (A-Z)", 
                      command=lambda: self.sort_historico(tree, historico, "name_asc")).pack(side=tk.LEFT, padx=2)
            ttk.Button(filter_frame, text="Ordenar por Data (Recente)", 
                      command=lambda: self.sort_historico(tree, historico, "date_desc")).pack(side=tk.LEFT, padx=2)
            ttk.Button(filter_frame, text="Ordenar por Data (Antiga)", 
                      command=lambda: self.sort_historico(tree, historico, "date_asc")).pack(side=tk.LEFT, padx=2)
            
            tree = ttk.Treeview(top, columns=("Nome", "CNPJ", "Data"), show="headings")
            tree.heading("Nome", text="Nome")
            tree.heading("CNPJ", text="CNPJ")
            tree.heading("Data", text="Data")
            
            tree.column("Nome", width=250)
            tree.column("CNPJ", width=150)
            tree.column("Data", width=150)
            
            self.sort_historico(tree, historico, "date_desc")
                
            scroll_y = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
            scroll_x = ttk.Scrollbar(top, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            
            btn_frame = ttk.Frame(top)
            
            ttk.Label(btn_frame, text="Selecione um CNPJ e clique em OK para continuar", 
                     font=('Arial', 9)).pack(pady=(0, 5))
            
            btn_ok = ttk.Button(btn_frame, text="OK", command=lambda: self.selecionar_do_historico(tree, top))
            
            tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
            scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
            btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
            btn_ok.pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar o histórico: {str(e)}")

    def selecionar_do_historico(self, tree, top):
        selecionado = tree.focus()
        if selecionado:
            item = tree.item(selecionado)
            cnpj_bruto = item['values'][1]
            cnpj = str(cnpj_bruto).zfill(14)
            self.cnpj_entry.delete(0, tk.END)
            self.cnpj_entry.insert(0, self.format_cnpj(cnpj))
            top.destroy()
            self.consultar_cnpj()

    def format_cnpj(self, cnpj):
        cnpj = ''.join(filter(str.isdigit, str(cnpj)))
        if len(cnpj) == 14:
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
        return cnpj

    def create_widgets(self):
        self.configure_styles()
        
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(self.input_frame, text="Digite o CNPJ :", font=('Arial', 11)).pack(side=tk.LEFT, padx=(0, 10))
        
        self.cnpj_entry = ttk.Entry(self.input_frame, width=25, font=('Arial', 12))
        self.cnpj_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.cnpj_entry.insert(0, "Digite o CNPJ da empresa")
        self.cnpj_entry.config(foreground='grey')
        self.cnpj_entry.bind("<FocusIn>", self.clear_placeholder)
        self.cnpj_entry.bind("<FocusOut>", self.restore_placeholder)
        
        self.btn_colar = ttk.Button(self.input_frame, text="Colar", command=self.colar_cnpj, style='Accent.TButton')
        self.btn_colar.pack(side=tk.LEFT, padx=5)

        # Adicionando botão Copiar na interface principal
        self.btn_copiar = ttk.Button(self.input_frame, text="Copiar", command=self.copiar_informacoes, style='Accent.TButton')
        self.btn_copiar.pack(side=tk.LEFT, padx=5)

        self.btn_consultar = ttk.Button(self.main_frame, text="Consultar CNPJ", command=self.consultar_cnpj, style='Accent.TButton')
        self.btn_consultar.pack(pady=(0, 10))
        self.cnpj_entry.bind('<Return>', lambda event: self.consultar_cnpj())
        
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.create_info_tab()
        self.create_socios_tab()
        self.create_atividades_tab()
        self.create_registrations_tab()
        
        ttk.Label(self.main_frame, 
                 text="© Feito por Barba", 
                 font=('Arial', 8), foreground="#777777").pack(side=tk.BOTTOM, pady=(10, 0))
        
    def clear_placeholder(self, event):
        if self.cnpj_entry.get() == "Digite o CNPJ da empresa":
            self.cnpj_entry.delete(0, tk.END)
            self.cnpj_entry.config(foreground='black')
    
    def restore_placeholder(self, event):
        if not self.cnpj_entry.get():
            self.cnpj_entry.insert(0, "Digite o CNPJ da empresa")
            self.cnpj_entry.config(foreground='grey')
    
    def configure_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('.', background=BG_COLOR, foreground=FG_COLOR)
        style.configure('TFrame', background=BG_COLOR)
        style.configure('TLabel', background=BG_COLOR, font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10), borderwidth=1)
        style.configure('Accent.TButton', background=ACCENT_COLOR, foreground='white')
        
        style.map('Accent.TButton',
                 background=[('active', '#3a5a80'), ('pressed', '#2c4764')],
                 foreground=[('active', 'white'), ('pressed', 'white')])
        
        style.configure('TEntry', fieldbackground=ENTRY_COLOR, font=('Arial', 12))
        style.configure('TText', background=TEXT_BG, font=('Arial', 10))
        style.configure('Treeview', background=TEXT_BG, fieldbackground=TEXT_BG, font=('Arial', 11))
        style.configure('Treeview.Heading', background=ACCENT_COLOR, foreground='white', font=('Arial', 11, 'bold'))
        style.map('Treeview', background=[('selected', '#3a5a80')])
    
    def create_info_tab(self):
        self.info_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.info_frame, text="Informações Básicas")
        
        self.info_text = tk.Text(
            self.info_frame, 
            wrap='word',
            bg=TEXT_BG,
            fg=FG_COLOR,
            font=('Arial', 12),
            padx=15,
            pady=15,
            bd=0,
            highlightthickness=1,
            highlightbackground="#cccccc",
            highlightcolor=ACCENT_COLOR,
            state='disabled'
        )
        scrollbar = ttk.Scrollbar(self.info_frame, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.info_text.tag_configure("bold", font=('Arial', 12, 'bold'))
    
    def create_socios_tab(self):
        self.socios_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.socios_frame, text="Sócios")
        
        self.socios_tree = ttk.Treeview(
            self.socios_frame, 
            columns=("Nome", "CPF", "Tipo", "Cargo", "Desde", "Idade"), 
            show="headings"
        )
        
        colunas = [
            ("Nome", 200),
            ("CPF", 120),
            ("Tipo", 80),
            ("Cargo", 150),
            ("Desde", 100),
            ("Idade", 80)
        ]
        
        for col, width in colunas:
            self.socios_tree.heading(col, text=col)
            self.socios_tree.column(col, width=width, anchor=tk.W)
        
        scroll_y = ttk.Scrollbar(self.socios_frame, orient="vertical", command=self.socios_tree.yview)
        scroll_x = ttk.Scrollbar(self.socios_frame, orient="horizontal", command=self.socios_tree.xview)
        self.socios_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.socios_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_atividades_tab(self):
        self.atividades_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.atividades_frame, text="Atividades")
        
        ttk.Label(self.atividades_frame, text="Atividade Principal:", font=('Arial', 11, 'bold')).pack(anchor=tk.W, pady=(5,0))
        
        self.atividade_principal = ttk.Entry(
            self.atividades_frame, 
            width=80, 
            state='readonly',
            font=('Arial', 10)
        )
        self.atividade_principal.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.atividades_frame, text="Atividades Secundárias:", font=('Arial', 11, 'bold')).pack(anchor=tk.W, pady=(10,0))
        
        self.atividades_secundarias = tk.Text(
            self.atividades_frame, 
            height=10, 
            wrap=tk.WORD, 
            state='disabled',
            bg=TEXT_BG,
            fg=FG_COLOR,
            font=('Arial', 10),
            padx=5,
            pady=5
        )
        scrollbar = ttk.Scrollbar(self.atividades_frame, command=self.atividades_secundarias.yview)
        self.atividades_secundarias.configure(yscrollcommand=scrollbar.set)
        
        self.atividades_secundarias.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_registrations_tab(self):
        self.registrations_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.registrations_frame, text="Inscrições")
        
        self.registrations_tree = ttk.Treeview(
            self.registrations_frame, 
            columns=("Número", "UF", "Situação", "Tipo", "Data Status"), 
            show="headings"
        )
        
        colunas = [
            ("Número", 150),
            ("UF", 50),
            ("Situação", 150),
            ("Tipo", 100),
            ("Data Status", 100)
        ]
        
        for col, width in colunas:
            self.registrations_tree.heading(col, text=col)
            self.registrations_tree.column(col, width=width, anchor=tk.W)
        
        scroll_y = ttk.Scrollbar(self.registrations_frame, orient="vertical", command=self.registrations_tree.yview)
        scroll_x = ttk.Scrollbar(self.registrations_frame, orient="horizontal", command=self.registrations_tree.xview)
        self.registrations_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.registrations_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    def consultar_cnpj(self):
        cnpj = self.cnpj_entry.get().strip()
        cnpj = ''.join(filter(str.isdigit, cnpj))
        
        if len(cnpj) != 14:
            messagebox.showwarning("CNPJ inválido", "O CNPJ deve conter 14 dígitos numéricos.")
            return
        
        try:
            self.limpar_dados()
            self.root.update_idletasks()
            
            url = f"https://open.cnpja.com/office/{cnpj}"
            headers = {"Accept": "application/json"}
            response = requests.get(url, headers=headers)

            if response.status_code == 404:
                messagebox.showerror("Não encontrado", "CNPJ não encontrado na base de dados.")
            elif response.status_code == 400:
                messagebox.showerror("Erro de Requisição", "CNPJ inválido. Verifique e tente novamente.")
                return
                
            response.raise_for_status()

            empresa = response.json()
            self.empresa_data = empresa  # Salvar os dados para cópia
            
            nome_empresa = self.get_nested_value(empresa, "company.name")
            if nome_empresa:
                self.salvar_no_historico(nome_empresa, cnpj)
            
            if self.preferences.get("alertas_situacao", False):
                self.verificar_situacao_cadastral(empresa)
            
            self.preencher_info_tab(empresa)
            self.preencher_socios_tab(empresa)
            self.preencher_atividades_tab(empresa)
            self.preencher_registrations_tab(empresa)
            
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Erro", f"Falha na consulta: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {str(e)}")
    
    def verificar_situacao_cadastral(self, empresa):
        situacao = self.get_nested_value(empresa, "status.text")
        
        if situacao == "Ativa":
            messagebox.showinfo("Situação Cadastral", "Empresa ATIVA - Regular")
        elif situacao == "Baixada":
            messagebox.showwarning("Situação Cadastral", "Empresa BAIXADA - Verifique os detalhes")
        elif situacao == "SUSPENSA":
            messagebox.showwarning("Situação Cadastral", "Empresa SUSPENSA - Pode haver restrições")
        elif situacao == "INAPTA":
            messagebox.showwarning("Situação Cadastral", "Empresa INAPTA - Verifique os detalhes")
        elif situacao == "BAIXADA POR INEXISTÊNCIA DE FATO":
            messagebox.showwarning("Situação Cadastral", "Empresa BAIXADA POR INEXISTÊNCIA DE FATO")
        else:
            messagebox.showinfo("Situação Cadastral", f"Situação: {situacao}")

    def preencher_info_tab(self, empresa):
        self.info_text.config(state='normal')
        self.info_text.delete(1.0, tk.END)
        
        capital_social = float(self.get_nested_value(empresa, 'company.equity') or 0)
        capital_formatado = format_currency(capital_social, 'BRL', locale='pt_BR')
        
        campos = [
            ("Nome Empresarial", self.get_nested_value(empresa, "company.name")),
            ("Nome Fantasia", self.get_nested_value(empresa, "alias")),
            ("Data de Abertura", datetime.strptime(self.get_nested_value(empresa, "founded"), "%Y-%m-%d").strftime("%d-%m-%Y")),
            ("Situação Cadastral", self.get_nested_value(empresa, "status.text")),
            ("Natureza Jurídica", self.get_nested_value(empresa, "company.nature.text")),
            ("Porte da Empresa", self.get_nested_value(empresa, "company.size.text")),
            ("Capital Social", capital_formatado),
            ("Telefone", self.format_phone(empresa.get("phones", []))),
            ("Endereço", self.format_address(empresa.get("address", {}))),
            ("Atividade Principal", self.get_nested_value(empresa, "mainActivity.text"))
        ]
        
        for label, valor in campos:
            self.info_text.insert(tk.END, f"{label}: ", "bold")
            
            if label == "Capital Social":
                self.info_text.insert(tk.END, valor)
                self.info_text.insert(tk.END, "    [Clique aqui]", "capital_link")
                self.info_text.tag_config("capital_link", foreground="blue", underline=1)
                self.info_text.tag_bind("capital_link", "<Button-1>", 
                                      lambda e, cs=capital_social: self.mostrar_capital_extenso(cs))
            else:
                self.info_text.insert(tk.END, f"{valor or 'Não informado'}")
            
            self.info_text.insert(tk.END, "\n")
        
        self.info_text.config(state='disabled')
    
    def mostrar_capital_extenso(self, valor):
        try:
            valor_extenso = num2words(valor, lang='pt_BR', to='currency')
            valor_extenso = valor_extenso.capitalize() + ""
            messagebox.showinfo("Capital Social por Extenso", valor_extenso)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível converter o valor: {str(e)}")

    def preencher_socios_tab(self, empresa):
        for member in empresa.get("company", {}).get("members", []):
            self.socios_tree.insert("", tk.END, values=(
                member.get("person", {}).get("name", "N/A"),
                member.get("person", {}).get("taxId", "N/A"),
                member.get("person", {}).get("type", "N/A"),
                member.get("role", {}).get("text", "N/A"),
                member.get("since", "N/A"),
                member.get("person", {}).get("age", "N/A")
            ))
    
    def preencher_atividades_tab(self, empresa):
        self.atividade_principal.config(state='normal')
        self.atividade_principal.delete(0, tk.END)
        self.atividade_principal.insert(0, self.get_nested_value(empresa, "mainActivity.text") or "Não informado")
        self.atividade_principal.config(state='readonly')
        
        self.atividades_secundarias.config(state='normal')
        self.atividades_secundarias.delete(1.0, tk.END)
        
        side_activities = empresa.get("sideActivities", [])
        if side_activities:
            for atividade in side_activities:
                self.atividades_secundarias.insert(tk.END, f"• {atividade.get('text', 'N/A')}\n")
        else:
            self.atividades_secundarias.insert(tk.END, "Nenhuma atividade secundária registrada")
        
        self.atividades_secundarias.config(state='disabled')
    
    def preencher_registrations_tab(self, empresa):
        for reg in empresa.get("registrations", []):
            self.registrations_tree.insert("", tk.END, values=(
                reg.get("number", "N/A"),
                reg.get("state", "N/A"),
                self.get_nested_value(reg, "status.text"),
                self.get_nested_value(reg, "type.text"),
                reg.get("statusDate", "N/A")
            ))
    
    def limpar_dados(self):
        self.info_text.config(state='normal')
        self.info_text.delete(1.0, tk.END)
        self.info_text.config(state='disabled')
        
        for item in self.socios_tree.get_children():
            self.socios_tree.delete(item)
        
        self.atividade_principal.config(state='normal')
        self.atividade_principal.delete(0, tk.END)
        self.atividade_principal.config(state='readonly')
        
        self.atividades_secundarias.config(state='normal')
        self.atividades_secundarias.delete(1.0, tk.END)
        self.atividades_secundarias.config(state='disabled')
        
        for item in self.registrations_tree.get_children():
            self.registrations_tree.delete(item)
    
    def colar_cnpj(self):
        self.cnpj_entry.delete(0, tk.END)
        try:
            texto = self.root.clipboard_get()
            self.cnpj_entry.insert(0, texto)
        except:
            messagebox.showinfo("Erro", "Nada para colar na área de transferência.")
        finally:
            self.cnpj_entry.focus_set()
    
    @staticmethod
    def get_nested_value(data, key):
        keys = key.split('.')
        for k in keys:
            if isinstance(data, dict) and k in data:
                data = data[k]
            else:
                return None
        return data
    
    @staticmethod
    def format_phone(phones):
        if not phones:
            return "Não informado"
        phone = phones[0]
        return f"({phone.get('area', '')}) {phone.get('number', '')}"
    
    @staticmethod
    def format_address(address):
        if not address:
            return "Não informado"
        return (
            f"{address.get('street', '')}, {address.get('number', '')} - "
            f"{address.get('district', '')}, {address.get('city', '')} - "
            f"{address.get('state', '')}, CEP: {address.get('zip', '')}"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = CNPJApp(root)
    root.mainloop()