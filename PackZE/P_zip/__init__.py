import time
import os
from os import path
from tkinter import END, messagebox
from PackZE import P_erros
import win32com.client as win32
import zipfile as zip
from pywintypes import com_error
from getmac import get_mac_address as gma

outlook = win32.Dispatch('outlook.application')

pasta = path.join(path.expanduser("~"), "Desktop\\Arquivos\\")

limite = 20971520

class EnviarEmail:

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
                                                    

            if os.path.exists("anexos.zip"):
 
                z.close()

                EnviarEmail.escrever(self)

        except:

            P_erros()

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
        
        finally:

            exit()

    def verificar(self):

        if self.label_descricao == "":

            messagebox.showerror(message="SEM DESTINATÁRIO")

            exit()

        if not gma() == "0c:d2:92:b5:06:08":

            messagebox.showwarning(message="EXECUTÁVEL NÃO PERMITIDO NESSA MÁQUINA")

            exit()

        if not os.path.exists(pasta):

            messagebox.showwarning(message="Crie uma pasta na área de trabalho com nome (Arquivos)")

            exit()

        else:

            os.chdir(str(pasta))

        if len(os.listdir(pasta)) == 0:                 

            messagebox.showinfo(message="Pasta vazia")

            exit()
        
        if os.path.exists("anexos.zip"):

            if os.path.getsize(os.path.join(pasta, "anexos.zip")) > limite:

                messagebox.showerror(message= """
                Arquivo nomeado (anexos) com o limite maior do permitido para:
                envio por email""")

                exit()

            else:

                EnviarEmail.escrever(self)