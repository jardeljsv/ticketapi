import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import csv
import threading
from queue import Queue
from datetime import datetime

class GLPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Criador de Chamados em Lote - GLPI")
        self.session_token = None
        self.csv_data = []
        self.log_queue = Queue()
        self.running = False
        
        # Configuração de estilo
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10), padding=5, foreground='black', relief='flat')
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))

        # Frame principal
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Cabeçalho
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        ttk.Label(header_frame, text="GLPI - Criador de Chamados em Lote via CSV", style='Header.TLabel').pack()

        # Configuração da API
        config_frame = ttk.LabelFrame(self.main_frame, text="Configuração da API", padding=10)
        config_frame.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(config_frame, text="URL GLPI API:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.url_entry = ttk.Entry(config_frame, width=50)
        self.url_entry.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(config_frame, text="App Token:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.app_token_entry = ttk.Entry(config_frame, width=50, show='*')
        self.app_token_entry.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(config_frame, text="User Token:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.user_token_entry = ttk.Entry(config_frame, width=50, show='*')
        self.user_token_entry.grid(row=2, column=1, padx=5, pady=2)

        self.ssl_verify_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(config_frame, text="Verificar SSL", variable=self.ssl_verify_var).grid(row=3, column=1, sticky=tk.W, pady=2)

        btn_frame = ttk.Frame(config_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(btn_frame, text="Iniciar Sessão", command=self.iniciar_sessao).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Testar Conexão", command=self.testar_conexao).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Encerrar Sessão", command=self.encerrar_sessao).pack(side=tk.LEFT, padx=5)

        # Importação CSV
        csv_frame = ttk.LabelFrame(self.main_frame, text="Importação de CSV", padding=10)
        csv_frame.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(csv_frame, text="Arquivo CSV com os dados dos chamados:").pack(pady=5)
        ttk.Button(csv_frame, text="Selecionar Arquivo CSV", command=self.carregar_csv).pack(pady=5)
        ttk.Button(csv_frame, text="Baixar Modelo CSV", command=self.gerar_modelo_csv).pack(pady=5)

        self.status_label = ttk.Label(csv_frame, text="Nenhum arquivo carregado")
        self.status_label.pack(pady=5)

        # Progresso
        self.progress_frame = ttk.Frame(self.main_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        self.progress_label = ttk.Label(self.progress_frame, text="Pronto")
        self.progress_label.pack()
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X)

        # Botões de ação
        btn_action_frame = ttk.Frame(self.main_frame)
        btn_action_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_action_frame, text="Abrir Logs", command=self.abrir_logs).pack(side=tk.LEFT, padx=5)
        self.create_button = ttk.Button(btn_action_frame, text="Criar Chamados em Lote", 
                                      command=self.iniciar_criacao_chamados)
        self.create_button.pack(side=tk.RIGHT, padx=5)
        self.stop_button = ttk.Button(btn_action_frame, text="Parar", command=self.parar_processamento, 
                                    state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=5)

        # Atualização periódica da interface
        self.root.after(100, self.atualizar_interface)

    def iniciar_sessao(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Aviso", "URL da API não pode estar vazia")
            return
            
        url = url.rstrip('/') + "/initSession"
        headers = {
            'App-Token': self.app_token_entry.get().strip(),
            'Authorization': 'user_token ' + self.user_token_entry.get().strip()
        }
        
        try:
            self.log(f"Tentando iniciar sessão em {url}")
            r = requests.post(
                url, 
                headers=headers, 
                verify=self.ssl_verify_var.get(),
                timeout=30
            )
            
            if r.status_code == 200:
                self.session_token = r.json()['session_token']
                messagebox.showinfo("Sucesso", "Sessão iniciada com sucesso!")
                self.log("Sessão iniciada com sucesso")
            else:
                error_msg = f"Erro {r.status_code}: {r.text}"
                messagebox.showerror("Erro", error_msg)
                self.log(error_msg)
                
        except requests.exceptions.RequestException as e:
            error_msg = f"Falha na conexão: {str(e)}"
            messagebox.showerror("Erro", error_msg)
            self.log(error_msg)

    def testar_conexao(self):
        self.iniciar_sessao()

    def encerrar_sessao(self):
        if not self.session_token:
            return
            
        url = self.url_entry.get().strip().rstrip('/') + "/killSession"
        headers = {
            'App-Token': self.app_token_entry.get().strip(),
            'Session-Token': self.session_token
        }
        
        try:
            requests.post(
                url, 
                headers=headers, 
                verify=self.ssl_verify_var.get(),
                timeout=10
            )
        except:
            pass
            
        self.session_token = None
        self.log("Sessão encerrada")

    def gerar_modelo_csv(self):
        modelo = """id_requerente;titulo;descricao;id_categoria;urgencia;tipo
12;Problema com impressora;A impressora não está respondendo;15;3;1
8;Computador lento;O computador está muito lento;12;2;1"""
        caminho = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv")],
            title="Salvar modelo CSV"
        )
        if caminho:
            try:
                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(modelo)
                messagebox.showinfo("Sucesso", "Modelo CSV salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o arquivo:\n{str(e)}")

    def carregar_csv(self):
        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos CSV", "*.csv")],
            title="Selecionar arquivo CSV"
        )
        if caminho:
            try:
                with open(caminho, newline='', encoding='utf-8') as csvfile:
                    leitor = csv.DictReader(csvfile, delimiter=';')
                    self.csv_data = []
                    linhas_invalidas = 0
                    
                    for idx, linha in enumerate(leitor, 1):
                        campos_obrigatorios = ['id_requerente', 'titulo', 'descricao', 'id_categoria']
                        if not all(campo in linha for campo in campos_obrigatorios):
                            self.log(f"Linha {idx} ignorada: campos obrigatórios faltando")
                            linhas_invalidas += 1
                            continue
                        
                        try:
                            linha['id_requerente'] = int(linha['id_requerente'])
                            linha['id_categoria'] = int(linha['id_categoria'])
                        except ValueError:
                            self.log(f"Linha {idx} ignorada: IDs devem ser números")
                            linhas_invalidas += 1
                            continue
                            
                        self.csv_data.append(linha)
                    
                    self.status_label.config(text=f"{len(self.csv_data)} chamados válidos carregados")
                    if linhas_invalidas > 0:
                        messagebox.showwarning("Aviso", 
                            f"{linhas_invalidas} linhas foram ignoradas por dados inválidos")
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o arquivo CSV:\n{str(e)}")
                self.csv_data = []
                self.status_label.config(text="Erro ao carregar arquivo")

    def log(self, mensagem):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {mensagem}"
        print(log_entry)
        self.log_queue.put(log_entry)

    def abrir_logs(self):
        if hasattr(self, 'log_window') and self.log_window.winfo_exists():
            self.log_window.lift()
            return
            
        self.log_window = tk.Toplevel(self.root)
        self.log_window.title("Logs de Execução")
        self.log_window.protocol("WM_DELETE_WINDOW", self.fechar_log_window)
        
        frame = ttk.Frame(self.log_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(
            frame, 
            wrap=tk.WORD, 
            height=20, 
            width=80,
            yscrollcommand=scrollbar.set
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        while not self.log_queue.empty():
            self.log_text.insert(tk.END, self.log_queue.get() + "\n")
        
        self.log_text.see(tk.END)
        
        clear_button = ttk.Button(
            self.log_window, 
            text="Limpar Logs", 
            command=self.limpar_logs
        )
        clear_button.pack(pady=5)

    def fechar_log_window(self):
        if hasattr(self, 'log_text'):
            self.log_text = None
        self.log_window.destroy()

    def limpar_logs(self):
        if hasattr(self, 'log_text'):
            self.log_text.delete(1.0, tk.END)

    def iniciar_criacao_chamados(self):
        if not self.session_token:
            messagebox.showwarning("Aviso", "Inicie a sessão antes de criar chamados")
            return
            
        if not self.csv_data:
            messagebox.showwarning("Aviso", "Importe um arquivo CSV primeiro")
            return
            
        self.running = True
        self.create_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar['maximum'] = len(self.csv_data)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Preparando...")
        
        threading.Thread(target=self.criar_chamados, daemon=True).start()

    def parar_processamento(self):
        self.running = False
        self.log("Processamento interrompido pelo usuário")
        self.progress_label.config(text="Interrompido")

    def criar_chamados(self):
        url = self.url_entry.get().strip().rstrip('/') + "/Ticket"
        headers = {
            'App-Token': self.app_token_entry.get().strip(),
            'Session-Token': self.session_token,
            'Content-Type': 'application/json'
        }
        
        sucesso = erros = 0
        detalhes_erros = []
        
        for idx, linha in enumerate(self.csv_data, 1):
            if not self.running:
                break
                
            try:
                payload = {
                    "input": {
                        "name": linha['titulo'],
                        "content": linha['descricao'],
                        "itilcategories_id": linha['id_categoria'],
                        "_users_id_requester": linha['id_requerente'],
                        "urgency": int(linha.get('urgencia', 3)),
                        "type": int(linha.get('tipo', 1))
                    }
                }
                
                self.log(f"Criando chamado {idx}/{len(self.csv_data)}: {linha['titulo']}")
                
                try:
                    r = requests.post(
                        url, 
                        headers=headers, 
                        json=payload, 
                        verify=self.ssl_verify_var.get(),
                        timeout=30
                    )
                    
                    self.log(f"Resposta: {r.status_code} - {r.text}")
                    
                    if r.status_code == 201:
                        sucesso += 1
                    else:
                        erros += 1
                        detalhes_erros.append(
                            f"Chamado '{linha['titulo']}': {r.status_code} - {r.text}"
                        )
                        
                except requests.exceptions.RequestException as e:
                    erros += 1
                    detalhes_erros.append(
                        f"Chamado '{linha['titulo']}': Erro de conexão - {str(e)}"
                    )
                    self.log(f"Erro de conexão: {str(e)}")
                    
            except Exception as e:
                erros += 1
                detalhes_erros.append(
                    f"Chamado '{linha.get('titulo', 'Sem título')}': Erro ao processar - {str(e)}"
                )
                self.log(f"Erro ao processar linha: {str(e)}")
                
            finally:
                self.progress_bar['value'] = idx
                self.progress_label.config(
                    text=f"Processando {idx}/{len(self.csv_data)} - Sucesso: {sucesso}, Erros: {erros}"
                )
                
        self.running = False
        self.root.after(100, lambda: self.finalizar_processamento(sucesso, erros, detalhes_erros))

    def finalizar_processamento(self, sucesso, erros, detalhes_erros):
        self.create_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        resultado = f"Processamento concluído!\n\nChamados criados: {sucesso}\nFalhas: {erros}"
        
        if detalhes_erros:
            resultado += "\n\nDetalhes dos erros:\n" + "\n".join(detalhes_erros[:5])
            if len(detalhes_erros) > 5:
                resultado += f"\n... e mais {len(detalhes_erros) - 5} erros"
                
        messagebox.showinfo("Resultado", resultado)
        self.progress_label.config(text=f"Concluído - Sucesso: {sucesso}, Erros: {erros}")

    def atualizar_interface(self):
        if hasattr(self, 'log_text') and self.log_text:
            while not self.log_queue.empty():
                self.log_text.insert(tk.END, self.log_queue.get() + "\n")
                self.log_text.see(tk.END)
        
        self.root.after(100, self.atualizar_interface)

if __name__ == "__main__":
    root = tk.Tk()
    app = GLPIApp(root)
    root.mainloop()
