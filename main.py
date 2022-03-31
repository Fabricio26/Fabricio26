import tkinter as tk
from PackZE import P_zip

class InterFace:

    Z = P_zip.EnviarEmail.zip

    def __init__(self):
        
        self.janela = tk.Tk()

        self.janela.title("ENVIAR EMAIL")
        self.janela.iconbitmap("icone_Z41_icon.ico")

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

        self.botao_ok = tk.Button(text="Enviar", command=self.Z)
        self.botao_ok.grid(row=1, column=1, padx=10, pady=10, sticky=tk.NSEW, columnspan=4)

        self.janela.mainloop()

InterFace()