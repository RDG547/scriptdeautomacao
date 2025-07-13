import tkinter as tk
from tkinter import messagebox, ttk, Frame
from elevate import elevate
from script import (executar_script_completo, AtivadorWindows, 
                    InstaladorSoftware, ConfiguradorWindows, ativar_office, 
                    CONFIGURACAO_DE_SOFTWARE)

# Elevar permissões para o script
elevate()

# Funções para integração com o Tkinter
def atualizar_barra(etapa, total):
    progresso_percentual = int((etapa / total) * 100)
    progress_bar["value"] = progresso_percentual
    janela.update_idletasks()  # Atualiza a interface

def executar_completo():
    try:
        executar_script_completo()
        messagebox.showinfo("Concluído", "Script completo executado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao executar o script completo: {e}")

def ativar_windows():
    try:
        AtivadorWindows.verificar_e_ativar_windows()
        messagebox.showinfo("Concluído", "Windows verificado/ativado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ativar o Windows: {e}")

def instalar_software_individual(software):
    try:
        if InstaladorSoftware.instalar_software(software):
            messagebox.showinfo("Instalação", f"{software} instalado com sucesso!")
        else:
            messagebox.showwarning("Aviso", f"Falha ao instalar {software}.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao instalar {software}: {e}")

def instalar_todos_softwares():
    for software in CONFIGURACAO_DE_SOFTWARE:
        if software in ['Word', 'Excel', 'PowerPoint']:
            continue
        if software in ['Word', 'Excel', 'PowerPoint']:
            continue
        if software in ['Word', 'Excel', 'PowerPoint']:
            continue
        if software in ['Word', 'Excel', 'PowerPoint']:
            continue
        if software in ['Word', 'Excel', 'PowerPoint']:
            continue
        instalar_software_individual(software)
    messagebox.showinfo("Concluído", "Todos os softwares foram instalados.")

def configurar_desempenho():
    try:
        ConfiguradorWindows.configurar_aparencia_desempenho()
        messagebox.showinfo("Concluído", "Configuração de desempenho concluída!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao configurar desempenho: {e}")

def ativar_office_interface():
    try:
        ativar_office()
        messagebox.showinfo("Concluído", "Microsoft Office ativado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ativar o Microsoft Office: {e}")

def desativar_servicos():
    try:
        ConfiguradorWindows.desativar_servicos_do_windows()
        messagebox.showinfo("Concluído", "Serviços desativados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao desativar serviços: {e}")

def instalar_atualizacoes():
    try:
        ConfiguradorWindows.verificar_e_instalar_atualizacoes()
        messagebox.showinfo("Concluído", "Atualizações instaladas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao instalar atualizações: {e}")

# Interface expandida para menus internos

class MenuExpandido(Frame):
    def __init__(self, master, titulo, conteudo_func):
        super().__init__(master)
        self.titulo = titulo
        self.conteudo_func = conteudo_func
        self.expanded = False
        self.toggle_button = tk.Button(self, text=f"{self.titulo} ▼", command=self.toggle, width=40)
        self.toggle_button.pack()
        self.conteudo_frame = Frame(self)
        self.conteudo_frame.pack(fill="x", expand=True)

    def toggle(self):
        if self.expanded:
            for widget in self.conteudo_frame.winfo_children():
                widget.destroy()
            self.conteudo_frame.pack_forget()
            self.toggle_button.config(text=f"{self.titulo} ▼")
        else:
            self.conteudo_func(self.conteudo_frame)
            self.conteudo_frame.pack(fill="x", expand=True)
            self.toggle_button.config(text=f"{self.titulo} ▲")
        self.expanded = not self.expanded

# Conteúdo das seções expandíveis
def conteudo_softwares(frame):
    tk.Button(frame, text="Instalar Todos os Softwares", command=instalar_todos_softwares, width=30).pack(pady=5)
    for software in CONFIGURACAO_DE_SOFTWARE:
        tk.Button(frame, text=f"Instalar {software}", command=lambda s=software: instalar_software_individual(s), width=30).pack(pady=2)

def conteudo_desempenho(frame):
    tk.Button(frame, text="Desativar Serviços", command=desativar_servicos, width=30).pack(pady=5)
    tk.Button(frame, text="Configurar Aparência", command=configurar_desempenho, width=30).pack(pady=5)

# Criação da janela principal
janela = tk.Tk()
janela.title("Gerenciador de Sistema")
janela.geometry("500x600")

# Barra de progresso
progress_bar = ttk.Progressbar(janela, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)

# Botões principais
btn_executar = tk.Button(janela, text="Executar Completo", command=executar_completo, width=40)
btn_executar.pack(pady=5)

btn_ativar = tk.Button(janela, text="Ativar Windows", command=ativar_windows, width=40)
btn_ativar.pack(pady=5)

btn_office = tk.Button(janela, text="Ativar Office", command=ativar_office_interface, width=40)
btn_office.pack(pady=5)

# Menus expansíveis
menu_softwares = MenuExpandido(janela, "Softwares", conteudo_softwares)
menu_softwares.pack(fill="x", pady=5)

menu_desempenho = MenuExpandido(janela, "Desempenho", conteudo_desempenho)
menu_desempenho.pack(fill="x", pady=5)

btn_sair = tk.Button(janela, text="Sair", command=janela.quit, width=40)
btn_sair.pack(pady=5)

# Iniciar o loop da GUI
janela.mainloop()
