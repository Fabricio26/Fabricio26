import time
import os
from os import path
import tkinter as tk
from tkinter import END, messagebox
import win32com.client as win32
import zipfile as zip
from pywintypes import com_error
from getmac import get_mac_address as gma

outlook = win32.Dispatch('outlook.application')

pasta = path.join(path.expanduser("~"), "Desktop\\Arquivos\\")

limite = 20971520

class EnviarEmail:
    def __init__(self):
        self.janela = tk.Tk()

        self.janela.title("ENVIAR EMAIL")
        self.janela.iconbitmap("Programa de envio de email\\icone_Z41_icon.ico")

        self.janela.resizable(0, 0)

        self.label_descricao = tk.Label(text="EMAIL DO DESTINATÁRIO", )
        self.label_descricao.grid(row=1, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=4)

        self.entry_descricao = tk.Entry()
        self.entry_descricao.grid(row=2, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=1)

        self.label_assunto = tk.Label(text="ASSUNTO")
        self.label_assunto.grid(row=3, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=4)

        self.entry_assunto = tk.Entry()
        self.entry_assunto.grid(row=4, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=1)

        self.label_texto = tk.Label(text="DESCRIÇÃO")
        self.label_texto.grid(row=5, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=4)

        self.entry_texto = tk.Text(width=60)
        self.entry_texto.grid(row=6, column=0, padx=10, pady=10, sticky=tk.NSEW, columnspan=1)

        self.botao_ok = tk.Button(text="Enviar", command=self.zip)
        self.botao_ok.grid(row=1, column=1, padx=10, pady=10, sticky=tk.NSEW, columnspan=4)

        self.janela.mainloop()

    def zip(self):
        
        EnviarEmail.verificar(self)
        
        try:
                
            for arquivos in os.listdir(pasta):

                size = os.path.getsize(os.path.join(pasta, arquivos))
            
                if size <= limite:

                    if not os.path.exists("anexos.zip"):
                        z = zip.ZipFile("anexos.zip", 'w', zip.ZIP_DEFLATED)
                    
                    anexos = os.path.getsize(os.path.join(pasta, "anexos.zip"))

                    total = size + anexos

                    if (total) <= limite:
                                
                        z.write(arquivos)

                        os.remove(arquivos)

                    else:
                        
                        z.close()

                        EnviarEmail.escrever(self)
                                                    
                        z = zip.ZipFile("anexos.zip", 'w', zip.ZIP_DEFLATED)

                        z.write(arquivos)

                        os.remove(arquivos)

            if os.path.exists("anexos.zip"):
 
                z.close()

                EnviarEmail.escrever(self)
    

        except:

            EnviarEmail.erros()

            exit()

        finally:

            messagebox.showinfo(message="ENVIADO")


    def escrever(self):

        self.email = outlook.CreateItem(0)

        destino = self.entry_descricao.get()
        assunto = self.entry_assunto.get()      
        texto = self.entry_texto.get("1.0", END)
    
        self.email.To = f"{destino}"
        self.email.Subject = f"{assunto}"
        self.email.HTMLBody = f"""
        <p>{texto}</p>
        """

        EnviarEmail.enviar(self)

    def enviar(self):

        try:

            self.email.Attachments.Add(str(os.getcwd()  + "\\anexos.zip"))

            time.sleep(3)      
                
            self.email.Send()

            os.remove(str(os.getcwd()  + "\\anexos.zip"))      
            
        except com_error as error:

            messagebox.showerror(message= error)

            exit()

    def verificar(self):

        if not os.path.exists(pasta):

            messagebox.showwarning(message="Crie uma pasta na área de trabalho com nome (Arquivos)")

            exit()

        else:

            os.chdir(str(pasta))

        if not gma() == "0c:d2:92:b5:06:08":

            messagebox.showwarning(message="EXECUTÁVEL NÃO PERMITIDO NESSA MÁQUINA")

            exit()
    
        if len(os.listdir(pasta)) == 0:                 

            messagebox.showinfo(message="Pasta vazia")

            exit()
        
        if os.path.exists("anexos.zip"):

            if os.path.getsize(os.path.join(pasta, "anexos.zip")) > limite:

                messagebox.showerror(message= """Arquivo nomeado (anexos) com o limite maior do permitido para:
                envio por email""")

                exit()

            else:

                EnviarEmail.escrever(self)

    def erros(self):
    
        if UnboundLocalError:

            messagebox.showerror(message="Arquivo com mesmo nome de anexo.zip ou .rar")

            exit()

        if ValueError:

            messagebox.showerror(message= "Tentativa de anexar em Zip arquivo que já estava fechado")

            exit()

EnviarEmail()