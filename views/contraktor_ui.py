import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import queue
import sys
from datetime import timedelta
from tkinter import BooleanVar
import threading

class TextRedirector:
    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue

    def write(self, string):
        self.queue.put(string)

    def flush(self):
        pass

class ContraktorBotUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bot Consultor")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Cores baseadas na imagem enviada
        self.COR_FUNDO = "#2D2D2D"  # Fundo escuro
        self.COR_BOTAO_PESQUISAR = "#00C2C7"  # Azul turquesa
        self.COR_BOTAO_LIMPAR = "#D0FF00"  # Verde limão
        self.COR_TEXTO = "white"
        self.COR_ENTRADA = "#E0E0E0"  # Cinza claro para campos de entrada
        
        self.criar_bot = lambda: None
        # Configurar cores da janela principal
        self.root.configure(bg=self.COR_FUNDO)
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use("default")
        
        # Configurar estilos personalizados
        self.style.configure("TFrame", background=self.COR_FUNDO)
        self.style.configure("TLabelframe", background=self.COR_FUNDO, foreground=self.COR_TEXTO)
        self.style.configure("TLabelframe.Label", background=self.COR_FUNDO, foreground=self.COR_TEXTO)
        self.style.configure("TLabel", background=self.COR_FUNDO, foreground=self.COR_TEXTO)
        self.style.configure("Custom.Horizontal.TProgressbar", background=self.COR_BOTAO_LIMPAR)
        
        # Botões personalizados
        self.style.configure("Pesquisar.TButton", background=self.COR_BOTAO_PESQUISAR, foreground="black")
        
        # Variáveis
        self.limite_var = tk.StringVar(value="")
        self.is_running = False
        self.log_queue = queue.Queue()
        self.bot = None
        
        # Criar interface
        self.criar_widgets()
        
        # Configurar atualização periódica do log
        self.root.after(100, self.atualizar_log)
    
    def criar_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg=self.COR_FUNDO, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame de configurações
        config_frame = tk.LabelFrame(main_frame, text="Configurações", bg=self.COR_FUNDO, fg=self.COR_TEXTO, padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=5)
        
        # Limite de contratos
        limite_label = tk.Label(config_frame, text="Digite o limite de contratos:", bg=self.COR_FUNDO, fg=self.COR_TEXTO)
        limite_label.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        
        limite_entry = tk.Entry(config_frame, width=20, textvariable=self.limite_var, bg=self.COR_ENTRADA)
        limite_entry.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        
        # Botões + Check de exceção
        controle_frame = tk.Frame(config_frame, bg=self.COR_FUNDO)
        controle_frame.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)

        self.botao_iniciar = tk.Button(controle_frame, text="Iniciar", bg=self.COR_BOTAO_PESQUISAR, fg="black",
                                    command=self.iniciar_processamento, width=15)
        self.botao_iniciar.pack(side=tk.LEFT, padx=5)
        
        # Frame de progresso
        progress_frame = tk.LabelFrame(main_frame, text="Andamento", bg=self.COR_FUNDO, fg=self.COR_TEXTO, padx=10, pady=10)
        progress_frame.pack(fill=tk.X, pady=5)
        
        tempo_label = tk.Label(progress_frame, text="Previsão de tempo: ", bg=self.COR_FUNDO, fg=self.COR_TEXTO)
        tempo_label.grid(column=0, row=0, sticky=tk.W, pady=5)
        
        total_label = tk.Label(progress_frame, text="Contratos processados:", bg=self.COR_FUNDO, fg=self.COR_TEXTO)
        total_label.grid(column=1, row=0, sticky=tk.W, pady=5, padx=(50, 0))
        
        self.tempo_field = tk.Entry(progress_frame, width=20, bg=self.COR_ENTRADA)
        self.tempo_field.grid(column=0, row=1, sticky=tk.W, pady=5)
        
        self.total_field = tk.Entry(progress_frame, width=20, bg=self.COR_ENTRADA)
        self.total_field.grid(column=1, row=1, sticky=tk.W, pady=5, padx=(50, 0))
        
        # Barra de progresso
        self.progressbar = ttk.Progressbar(progress_frame, length=100, mode='determinate', style="Custom.Horizontal.TProgressbar")
        self.progressbar.grid(column=0, row=2, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Log de execução
        log_frame = tk.LabelFrame(main_frame, text="Log de Execução", bg=self.COR_FUNDO, fg=self.COR_TEXTO, padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=15, bg=self.COR_ENTRADA)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
    
    def atualizar_log(self):
        """Atualiza o widget de log com as mensagens da fila"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, message)
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                self.log_queue.task_done()
        except queue.Empty:
            pass
        
        # Agendar próxima atualização
        self.root.after(100, self.atualizar_log)
    
    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso e o campo total"""
        if total > 0:
            self.progressbar["value"] = (atual / total) * 100
            self.total_field.delete(0, tk.END)
            self.total_field.insert(0, f"{atual}/{total}")
    
    def atualizar_tempo_estimado(self, segundos_restantes):
        """Atualiza o tempo estimado restante"""
        if segundos_restantes > 0:
            tempo_formatado = str(timedelta(seconds=int(segundos_restantes)))
            self.tempo_field.delete(0, tk.END)
            self.tempo_field.insert(0, tempo_formatado)
        else:
            self.tempo_field.delete(0, tk.END)
            self.tempo_field.insert(0, "Calculando...")
    
    def iniciar_processamento(self):
        """Inicia o processamento em uma thread separada"""
        if self.is_running:
            messagebox.showwarning("Aviso", "Processamento já está em andamento!")
            return
        
        # Obter valores da interface
        limite_texto = self.limite_var.get().strip()
        limite = int(limite_texto) if limite_texto else None
        
        # Limpar log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Resetar progresso
        self.progressbar["value"] = 0
        self.tempo_field.delete(0, tk.END)
        self.total_field.delete(0, tk.END)

        
        # Desabilitar botões durante o processamento
        self.botao_iniciar.config(state=tk.DISABLED)
        
        # Indicar que está executando
        self.is_running = True
        
        # Iniciar thread de processamento
        threading.Thread(target=self.executar_bot, args=(limite,), daemon=True).start()
    
    def executar_bot(self, limite):
        """Executa o bot em uma thread separada"""
        try:
            if self.criar_bot is None:
                self.log_queue.put("Erro: A função criar_bot não foi definida\n")
                return

            # Redirecionar stdout e stderr para o widget de log
            original_stdout = sys.stdout
            original_stderr = sys.stderr
            
            sys.stdout = TextRedirector(self.log_text, self.log_queue)
            sys.stderr = TextRedirector(self.log_text, self.log_queue)
            
            self.bot = self.criar_bot()

            if self.bot is None:
                self.log_queue.put("Erro: Falha ao criar o bot\n")
                return

            self.bot.executar(limite=limite)
            
        except Exception as e:
            self.log_queue.put(f"Erro crítico: {str(e)}\n")
            
        finally:
            # Restaurar stdout e stderr
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            
            # Atualizar interface
            self.root.after(0, self.finalizar_processamento)
    
    def finalizar_processamento(self):
        """Atualiza a interface ao finalizar o processamento"""
        self.is_running = False
        self.botao_iniciar.config(state=tk.NORMAL)
        
        self.log_queue.put("\n--- Processamento finalizado ---\n")
