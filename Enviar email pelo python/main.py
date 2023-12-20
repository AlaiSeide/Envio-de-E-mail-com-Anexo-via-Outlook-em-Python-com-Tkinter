import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext

# função para enviar o e-mail
def send_email():
    try:
        # criar uma instância do Outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um novo e-mail
        mail = outlook.CreateItem(0)

        # definir o destinatário, assunto e corpo da mensagem
        mail.To = to_entry.get()
        mail.Subject = subject_entry.get()
        mail.Body = message_entry.get('1.0', 'end-1c')  # obter todo o texto do campo de entrada de texto

        # anexar arquivo, se especificado
        if file_path.get() != '':
            mail.Attachments.Add(file_path.get())

        # enviar o e-mail
        mail.Send()

        # exibir uma mensagem de sucesso
        messagebox.showinfo('Sucesso', 'E-mail enviado com sucesso!')

    except Exception as e:
        # exibir uma mensagem de erro
        messagebox.showerror('Erro', f'Ocorreu um erro ao enviar o e-mail: {str(e)}')


# função para abrir o seletor de arquivos
def browse_files():
    filename = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo")
    file_path.set(filename)


# criar a interface gráfica
root = tk.Tk()
root.title('Enviar E-mail pelo Outlook')
root.geometry('700x600')  # definir o tamanho da janela
root.resizable(False, False)

# adicionar os widgets
to_label = tk.Label(root, text='Para:')
to_label.pack()
to_entry = tk.Entry(root)
to_entry.pack()

subject_label = tk.Label(root, text='Assunto:')
subject_label.pack()
subject_entry = tk.Entry(root)
subject_entry.pack()

message_label = tk.Label(root, text='Mensagem:')
message_label.pack()
message_entry = scrolledtext.ScrolledText(root)  # usar ScrolledText para permitir rolagem
message_entry.pack()

file_path = tk.StringVar()
file_label = tk.Label(root, text='Anexo:')
file_label.pack()
file_entry = tk.Entry(root, textvariable=file_path)
file_entry.pack()
file_button = tk.Button(root, text='Procurar', command=browse_files)
file_button.pack()

send_button = tk.Button(root, text='Enviar', command=send_email)
send_button.pack()

root.mainloop()
