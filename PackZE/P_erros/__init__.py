from tkinter import messagebox


def erros(self):

        if UnboundLocalError:

            messagebox.showerror(message="Arquivo com mesmo nome de anexo.zip ou .rar")

            exit()

        if ValueError:

            messagebox.showerror(message= "Tentativa de anexar em Zip arquivo que já estava fechado")

            exit()