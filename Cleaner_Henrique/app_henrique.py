import customtkinter as ctk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import sys
import os
import threading

# Importa as *funções* dos nossos outros scripts
try:
    from carregar_base_compras import carregar_base_sqlite
    from rodar_conciliacao import rodar_conciliacao
except ImportError:
    print("ERRO: Verifique se os arquivos 'carregar_base_compras.py' e 'rodar_conciliacao.py' estão na mesma pasta.")
    sys.exit()


# --- Classe para redirecionar o PRINT para a caixa de texto ---
class StdoutRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, string):
        # Garante que a atualização da GUI seja feita na thread principal
        self.textbox.after(0, self._insert_text, string)

    def _insert_text(self, string):
        self.textbox.insert(ctk.END, string)
        self.textbox.see(ctk.END) # Auto-scroll

    def flush(self):
        pass # Necessário para o sys.stdout

# --- Classe Principal da Aplicação ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela ---
        self.title("Conciliador de Títulos TOTVS")
        self.geometry("700x650") # Um pouco mais de altura para o checkbox
        ctk.set_appearance_mode("dark") # Fundo Cinza Escuro

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Faz o log crescer

        # --- Cores Personalizadas ---
        self.COR_BOTAO = "#FFB900"       # Amarelo
        self.COR_BOTAO_HOVER = "#FFa000"  # Amarelo mais escuro
        self.COR_TEXTO_BOTAO = "#000000" # Texto preto para legibilidade

        # --- Variáveis de estado ---
        self.arquivo_b_path = ""
        self.arquivo_a_path = ""
        
        # (NOVO) Variável para controlar o checkbox
        self.var_modo_replace = ctk.BooleanVar(value=False) 

        # --- Widgets ---
        
        # --- PARTE 1: Carregar Base (Arquivo B) ---
        self.frame_b = ctk.CTkFrame(self)
        self.frame_b.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.frame_b.grid_columnconfigure(1, weight=1)

        self.btn_b = ctk.CTkButton(
            self.frame_b, 
            text="1. Selecionar Base de Compras (XLSX)", 
            command=self.selecionar_arquivo_b
        )
        self.btn_b.grid(row=0, column=0, padx=10, pady=10)

        self.label_b = ctk.CTkLabel(self.frame_b, text="Nenhum arquivo selecionado.", text_color="gray")
        self.label_b.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_run1 = ctk.CTkButton(
            self.frame_b, 
            text="2. Carregar Base no SQLite", 
            command=self.executar_parte_1,
            fg_color=self.COR_BOTAO,
            hover_color=self.COR_BOTAO_HOVER,
            text_color=self.COR_TEXTO_BOTAO,
            font=ctk.CTkFont(weight="bold")
        )
        self.btn_run1.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")
        
        # --- (NOVO) Checkbox de Modo ---
        self.check_modo = ctk.CTkCheckBox(
            self.frame_b,
            text="Substituir base de dados existente (Modo Replace)",
            variable=self.var_modo_replace,
            onvalue=True,
            offvalue=False
        )
        self.check_modo.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")
        # --- FIM DO NOVO WIDGET ---

        # --- PARTE 2: Rodar Conciliação (Arquivo A) ---
        self.frame_a = ctk.CTkFrame(self)
        self.frame_a.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.frame_a.grid_columnconfigure(1, weight=1)
        
        self.btn_a = ctk.CTkButton(
            self.frame_a, 
            text="3. Selecionar XML TOTVS (Arquivo A)", 
            command=self.selecionar_arquivo_a
        )
        self.btn_a.grid(row=0, column=0, padx=10, pady=10)
        
        self.label_a = ctk.CTkLabel(self.frame_a, text="Nenhum arquivo selecionado.", text_color="gray")
        self.label_a.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_run2 = ctk.CTkButton(
            self.frame_a, 
            text="4. RODAR CONCILIAÇÃO", 
            command=self.executar_parte_2,
            fg_color=self.COR_BOTAO,
            hover_color=self.COR_BOTAO_HOVER,
            text_color=self.COR_TEXTO_BOTAO,
            font=ctk.CTkFont(weight="bold")
        )
        self.btn_run2.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")

        # --- Log/Status Textbox ---
        self.log_textbox = ctk.CTkTextbox(self, state="normal", wrap="word")
        self.log_textbox.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")

        # Redireciona o stdout (prints) para a caixa de texto
        sys.stdout = StdoutRedirector(self.log_textbox)

        print("Pronto para começar.\n")
        print("COMO USAR:")
        print("1. Selecione a Base de Compras (XLSX).")
        print("2. (Opcional) Marque 'Substituir' se quiser apagar a base antiga.")
        print("3. Clique em 'Carregar Base' (Se desmarcado, adiciona os dados novos).\n")
        print("4. Selecione o XML TOTVS (Arquivo A).")
        print("5. Clique em 'RODAR CONCILIAÇÃO'.\n")

    # --- Funções de Callback dos Botões ---

    def selecionar_arquivo_b(self):
        path = fd.askopenfilename(
            title="Selecione a Base de Compras",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if path:
            self.arquivo_b_path = path
            self.label_b.configure(text=os.path.basename(path), text_color="white")
            print(f"Arquivo B selecionado: {path}\n")

    def selecionar_arquivo_a(self):
        path = fd.askopenfilename(
            title="Selecione o XML do TOTVS",
            filetypes=[("Arquivos XML", "*.xml")]
        )
        if path:
            self.arquivo_a_path = path
            self.label_a.configure(text=os.path.basename(path), text_color="white")
            print(f"Arquivo A selecionado: {path}\n")

    def executar_parte_1(self):
        if not self.arquivo_b_path:
            mb.showerror("Erro", "Por favor, selecione a 'Base de Compras (XLSX)' primeiro.")
            return
            
        # --- (ATUALIZADO) Lê o estado do checkbox ---
        if self.var_modo_replace.get() == True:
            modo_execucao = 'replace'
            print("--- MODO REPLACE SELECIONADO ---")
            print("A base de dados antiga será APAGADA e substituída por esta.")
        else:
            modo_execucao = 'append'
            print("--- MODO APPEND (ADICIONAR) SELECIONADO ---")
            print("Os dados deste arquivo serão ADICIONADOS ao banco de dados existente.")
        # --- FIM DA ATUALIZAÇÃO ---
            
        self.btn_run1.configure(state="disabled", text="Carregando... Aguarde...")
        self.btn_run2.configure(state="disabled")

        # Roda o script pesado em uma thread separada para não travar a GUI
        # (ATUALIZADO) Passa o 'modo_execucao' para a thread
        threading.Thread(target=self._thread_parte_1, args=(modo_execucao,), daemon=True).start()

    def _thread_parte_1(self, modo_execucao): # <-- Recebe o modo
        try:
            # (ATUALIZADO) Usa a variável 'modo_execucao'
            sucesso = carregar_base_sqlite(self.arquivo_b_path, modo_execucao=modo_execucao)
            if sucesso:
                print("Base de dados carregada. Você já pode rodar a Parte 2.")
        except Exception as e:
            print(f"\n--- ERRO INESPERADO (Parte 1) ---")
            print(f"Erro: {e}\n")
        finally:
            self.after(0, self._habilitar_botoes_parte_1)

    def _habilitar_botoes_parte_1(self):
        self.btn_run1.configure(state="normal", text="2. Carregar Base no SQLite")
        self.btn_run2.configure(state="normal")


    def executar_parte_2(self):
        if not self.arquivo_a_path:
            mb.showerror("Erro", "Por favor, selecione o 'XML TOTVS (Arquivo A)' primeiro.")
            return
        
        caminho_saida = fd.asksaveasfilename(
            title="Salvar Relatório Desmembrado Como...",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            defaultextension=".xlsx",
            initialfile="Relatorio_Final_Desmembrado.xlsx"
        )
        
        if not caminho_saida: 
            print("Salvamento cancelado.\n")
            return

        self.btn_run1.configure(state="disabled")
        self.btn_run2.configure(state="disabled", text="Processando... Aguarde...")
        
        threading.Thread(target=self._thread_parte_2, args=(caminho_saida,), daemon=True).start()

    def _thread_parte_2(self, caminho_saida):
        try:
            sucesso = rodar_conciliacao(self.arquivo_a_path, caminho_saida)
            if sucesso:
                print(f"Relatório salvo em: {caminho_saida}")
                mb.showinfo("Sucesso", f"Relatório final gerado com sucesso em:\n{caminho_saida}")
        except Exception as e:
            print(f"\n--- ERRO INESPERADO (Parte 2) ---")
            print(f"Erro: {e}\n")
            mb.showerror("Erro Inesperado", f"Ocorreu um erro:\n{e}")
        finally:
            self.after(0, self._habilitar_botoes_parte_2)

    def _habilitar_botoes_parte_2(self):
        self.btn_run1.configure(state="normal")
        self.btn_run2.configure(state="normal", text="4. RODAR CONCILIAÇÃO")


# --- Execução da Aplicação ---
if __name__ == "__main__":
    app = App()
    app.mainloop()