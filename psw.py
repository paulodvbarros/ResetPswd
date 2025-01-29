import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
import sys
import os


class AlterarSenhaDialog:
    def __init__(self, parent):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Alterar senha")
        self.dialog.geometry("300x200")
        self.dialog.resizable(False, False)

        # Configurar ícone da janela
        try:
            self.dialog.iconbitmap('icone.ico')
        except:
            pass

        # Torna a janela modal
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Centraliza a janela
        self.dialog.grid_columnconfigure(0, weight=1)
        self.dialog.grid_columnconfigure(1, weight=1)

        # Campos de senha
        tk.Label(self.dialog, text="Senha atual:").grid(row=0, column=0, pady=5, padx=5, sticky='e')
        self.senha_atual = ttk.Entry(self.dialog, show="*")
        self.senha_atual.grid(row=0, column=1, pady=5, padx=5, sticky='ew')

        tk.Label(self.dialog, text="Nova senha:").grid(row=1, column=0, pady=5, padx=5, sticky='e')
        self.nova_senha = ttk.Entry(self.dialog, show="*")
        self.nova_senha.grid(row=1, column=1, pady=5, padx=5, sticky='ew')

        tk.Label(self.dialog, text="Confirme a nova senha:").grid(row=2, column=0, pady=5, padx=5, sticky='e')
        self.confirma_senha = ttk.Entry(self.dialog, show="*")
        self.confirma_senha.grid(row=2, column=1, pady=5, padx=5, sticky='ew')

        # Botões
        self.btn_frame = tk.Frame(self.dialog)
        self.btn_frame.grid(row=3, column=0, columnspan=2, pady=20)

        ttk.Button(self.btn_frame, text="Confirmar", command=self.alterar_senha).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Cancelar", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)

    def bloquear_windows(self):
        ctypes.windll.user32.LockWorkStation()

    def alterar_senha(self):
        if not all([self.senha_atual.get(), self.nova_senha.get(), self.confirma_senha.get()]):
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return

        if self.nova_senha.get() != self.confirma_senha.get():
            messagebox.showerror("Erro", "As senhas não coincidem!")
            return

        ps_script = f'''
        $ErrorActionPreference = "Stop"
        Add-Type -AssemblyName System.DirectoryServices
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        try {{
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $domainContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $domain.Name)

            if ($domainContext.ValidateCredentials($currentUser.Split('\\')[1], '{self.senha_atual.get()}')) {{
                $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($domainContext, $currentUser)
                $userPrincipal.ChangePassword('{self.senha_atual.get()}', '{self.nova_senha.get()}')
                Write-Output "SUCCESS"
            }} else {{
                Write-Output "INVALID_PASSWORD"
            }}
        }} catch {{
            Write-Output "ERROR: $($_.Exception.Message)"
        }}
        '''

        try:
            # Executa PowerShell sem mostrar janela
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            resultado = subprocess.run(["powershell", "-Command", ps_script],
                                       capture_output=True,
                                       text=True,
                                       encoding='utf-8',
                                       startupinfo=startupinfo)

            if "SUCCESS" in resultado.stdout:
                response = messagebox.showinfo("Sucesso",
                                               "Senha alterada com sucesso!\n\nPor favor, faça logon novamente com suas novas credenciais.",
                                               type=messagebox.OK)
                self.dialog.destroy()
                if response == 'ok':
                    self.bloquear_windows()
            elif "INVALID_PASSWORD" in resultado.stdout:
                messagebox.showerror("Erro", "Senha atual incorreta!")
            else:
                messagebox.showerror("Erro", f"Erro ao alterar senha: {resultado.stdout}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar alteração: {str(e)}")


def verificar_expiracao():
    try:
        ps_script = '''
        $Resultado = net user $env:USERNAME /domain
        $DataExpiracao = $Resultado | Select-String "A senha expira"
        $DataExpiracao = $DataExpiracao -replace ".*A senha expira  ", ""
        $DataExpiracao
        '''

        # Executa PowerShell sem mostrar janela
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        comando = ["powershell", "-Command", ps_script]
        resultado = subprocess.run(comando, capture_output=True, text=True,
                                   encoding='cp850', startupinfo=startupinfo)
        data_expiracao = resultado.stdout.strip()
        resultado_label.config(text=f"Sua senha expira em: {data_expiracao}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao verificar expiração: {str(e)}")


def abrir_alterar_senha():
    AlterarSenhaDialog(janela)


# Cria a janela principal
janela = tk.Tk()
janela.title("Gerenciador de Senha")
janela.geometry("400x220")  # Aumentei um pouco a altura para acomodar a marca d'água
janela.resizable(False, False)

# Configurar ícone da janela
try:
    janela.iconbitmap('icone.ico')
except:
    pass

# Centraliza os elementos
janela.grid_rowconfigure(0, weight=1)
janela.grid_rowconfigure(4, weight=1)
janela.grid_columnconfigure(0, weight=1)
janela.grid_columnconfigure(1, weight=1)

# Estilo para os botões
style = ttk.Style()
style.configure('TButton', font=('Arial', 10))

# Cria e posiciona os elementos
resultado_label = tk.Label(janela, text="Clique no botão para verificar", font=('Arial', 10))
resultado_label.grid(row=0, column=0, columnspan=2, pady=10)

verificar_button = ttk.Button(janela, text="Verificar expiração", command=verificar_expiracao)
verificar_button.grid(row=1, column=0, columnspan=2, pady=10, padx=20, sticky='ew')

alterar_button = ttk.Button(janela, text="Alterar senha", command=abrir_alterar_senha)
alterar_button.grid(row=2, column=0, pady=10, padx=20, sticky='ew')

fechar_button = ttk.Button(janela, text="Fechar", command=janela.destroy)
fechar_button.grid(row=2, column=1, pady=10, padx=20, sticky='ew')

# Adiciona marca d'água
marca_dagua = tk.Label(janela, text="PAULOTB - TECH CCR",
                       font=('Arial', 8), fg='gray')
marca_dagua.grid(row=3, column=0, columnspan=2, pady=10)

# Inicia a janela
janela.mainloop()