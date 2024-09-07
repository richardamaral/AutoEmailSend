import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from PIL import Image, ImageTk

imagem_anexada = None

def enviar_email_suporte():
    destinatario = "suporte@exemplo.com"  
    email_copia = email_copia_entry.get()
    anydesk = anydesk_entry.get()
    motivo = motivo_text.get("1.0", 'end-1c')[:250]  
    enviar_email(destinatario, "PRO OCUPACIONAL - SUPORTE T.I", "assinatura.png", suporte_ti_html(email_copia, anydesk, motivo), email_copia=email_copia)
    limpar_campos_suporte()  

def enviar_email_senha():
    destinatario = destinatario_entry.get()
    enviar_email(destinatario, "PRO OCUPACIONAL - REDEFINIÇÃO DE SENHA", "esquecisenha.png", redefinicao_senha_html())
    limpar_campos_email()  

def enviar_email_analise():
    destinatario = destinatario_entry.get()
    enviar_email(destinatario, "PRO OCUPACIONAL - ANÁLISE DO ACESSO DO CLIENTE", "assinatura.png", analise_acesso_html())
    limpar_campos_email()  

def enviar_email_credenciais():
    destinatario = destinatario_entry.get()
    email_login = email_entry.get()
    senha_login = senha_entry.get()
    enviar_email(destinatario, "PRO OCUPACIONAL - ACESSO CENTRAL DO CLIENTE", "assinatura.png", acesso_central_html(email_login, senha_login))
    limpar_campos_credenciais()  

def enviar_email_ajuda():
    destinatario = destinatario_entry.get()
    enviar_email(destinatario, "PRO OCUPACIONAL - MANUAL DE AJUDA", "assinatura.png", ajuda_html())
    limpar_campos_email()  

def enviar_email(destinatario, assunto, assinatura_imagem, html, email_copia=None):
    smtp_server = 'smtp.office365.com'
    porta = 587
    remetente = 'exemplo@exemplo.com'
    remetente_senha = 'senha_ficticia'

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    if email_copia:
        msg['Cc'] = email_copia
    msg['Subject'] = assunto

    msg.attach(MIMEText(html, 'html'))

    try:
        with open(assinatura_imagem, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-ID', f'<{assinatura_imagem.split(".")[0]}>')
            img.add_header('Content-Disposition', 'inline', filename=assinatura_imagem)
            msg.attach(img)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Imagem {assinatura_imagem} não encontrada.")

    if imagem_anexada:
        try:
            with open(imagem_anexada, 'rb') as f:
                img_anexada = MIMEImage(f.read())
                img_anexada.add_header('Content-ID', f'<{imagem_anexada.split(".")[0]}>')
                img_anexada.add_header('Content-Disposition', 'attachment', filename=imagem_anexada.split('/')[-1])
                msg.attach(img_anexada)
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Imagem {imagem_anexada} não encontrada.")

    try:
        with smtplib.SMTP(smtp_server, porta) as server:
            server.starttls()
            server.login(remetente, remetente_senha)
            recipients = [destinatario] + ([email_copia] if email_copia else [])
            server.sendmail(remetente, recipients, msg.as_string())
        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
    except smtplib.SMTPException as e:
        messagebox.showerror("Erro", f"Erro ao enviar e-mail: {str(e)}")

def selecionar_imagem():
    global imagem_anexada
    imagem_anexada = filedialog.askopenfilename(title="Selecione uma imagem", filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.gif")])
    if imagem_anexada:
        anexar_imagem_label.config(text=f"Imagem anexada: {imagem_anexada.split('/')[-1]}")

def limpar_campos_suporte():
    email_copia_entry.delete(0, 'end')
    anydesk_entry.delete(0, 'end')
    motivo_text.delete("1.0", 'end')
    global imagem_anexada
    imagem_anexada = None
    anexar_imagem_label.config(text="Nenhuma imagem anexada.")

def limpar_campos_credenciais():
    destinatario_entry.delete(0, 'end')
    email_entry.delete(0, 'end')
    senha_entry.delete(0, 'end')

def limpar_campos_email():
    destinatario_entry.delete(0, 'end')

def suporte_ti_html(email_copia, anydesk, motivo):
    return f"""
    <html>
    <body>
    <p>Olá,</p>
    <p>Através desse e-mail foi solicitado um chamado para:</p>
    <p>E-mail: <b>{email_copia}</b></p>
    <p>Anydesk: <b>{anydesk}</b></p>
    <p>Motivo: <b>{motivo}</b></p>
    <br><br>
    <img src="cid:assinatura">
    </body>
    </html>
    """

def redefinicao_senha_html():
    return """
    <html>
    <body>
    <p><b>Para recuperar a sua senha, o processo é bem simples...</b></p>
    <br>
    <p>- Basta clicar no botão <b>"Esqueci minha senha"</b> e digitar o seu e-mail de acesso.</p>
    <p>- Consulte o e-mail inserido para utilizar o <b>link de redefinição de senha</b> enviado para ele.</p>
    <img src="cid:esquecisenha">
    <br>
    <br>
    <p><b>&darr; [LINK DIRETO] &darr;</b></p>
    <p>Caso já esteja logado: <a href="https://proocupacional.tawk.help/article/alterar-senha-na-central-do-cliente" style="color: #cc3333;">Alterar minha senha da Central do Cliente</a></p>
    <br>
    <a href="https://centralcliente.exemplo.com.br/CentralCliente/" style="color: green;">Central do Cliente</a><br>
    <a href="https://www.youtube.com" style="color: red;">Vídeo Explicativo</a>
    <br>
    <img src="cid:assinatura">
    </body>
    </html>
    """

def analise_acesso_html():
    return """
       <html>
         <body>
         <p>Prezado(a) Cliente,</p>
         <p>Estamos analisando a situação atual do acesso do cliente e investigando qualquer possível problema.</p>
         <p><b>Para podermos auxiliá-lo da melhor maneira possível, solicitamos gentilmente as seguintes informações:</b></p>
         <p><b>- Email de acesso utilizado para login:</b></p>
         <p><b>- CNPJ da sua empresa:</b></p>
         <p><b>- Nome da empresa:</b></p>
         <p><b>- Telefone para contato:</b></p>
         <p>Manteremos você atualizado conforme avançamos na resolução dessa questão. Agradecemos pela sua paciência e compreensão.</p>
       <br><br>
             <img src="cid:assinatura">
         </body>
       </html>
       """

def acesso_central_html(email_login, senha_login):
    return f"""
    <html>
    <body>
    <p>Olá,</p>
    <p>Segue abaixo as credenciais para o acesso a nossa <a href="https://centralcliente.exemplo.com.br/CentralCliente/" style="color: black;"><b>Central do Cliente</b></a>:</p>
    <p>E-mail: <b>{email_login}</b></p>
    <p>Senha: <b>{senha_login}</b></p>
    <br><br>
    <p><b>&darr; [LINK DIRETO] &darr;</b></p>
    <a href="https://www.youtube.com" style="color: #0077b6;">Vídeo Demonstrativo Central do Cliente</a>
    <br><br>
    <img src="cid:assinatura">
    </body>
    </html>
    """

def ajuda_html():
    return """
       <html>
         <body>
         <p>Olá,</p>
         <p>Segue abaixo alguns manuais de ajuda para nossa <a href="https://centralcliente.exemplo.com.br/CentralCliente/" style="color: black;"><b>Central do Cliente</b></a>:</p>
       <p esd-text="true" class="esd-text"><b>↓ [LINK DIRETO] ↓</b></p>
             <a href="https://www.youtube.com" style="color: #0077b6; text-decoration: none;"><b>Vídeo Demonstrativo Central do Cliente</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/atualiza%C3%A7%C3%A3o-de-base-funcionarios" style="color: #ff5733; text-decoration: none;"><b>Atualização de Funcionários</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/emitir-guia-de-autoriza%C3%A7%C3%A3o-m%C3%A9dica" style="color: #33aaff; text-decoration: none;"><b>Emitir Guia de Autorização de Exame</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/como-visualizar-os-asos-dos-colaboradores" style="color: #33aa66; text-decoration: none;"><b>Visualizar o ASO do colaborador</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/visualiza%C3%A7%C3%A3o-dos-documentos-legais" style="color: #ff9933; text-decoration: none;"><b>Visualizar os Laudos / Documentos legais</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/consulta-de-faturamento-hist%C3%B3rico-financeiro" style="color: #6699ff; text-decoration: none;"><b>Consulta de Faturamento / Histórico Financeiro</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/como-visualizar-os-recibos-do-esocial" style="color: #ff6699; text-decoration: none;"><b>Visualizar recibos do eSocial</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/alterar-senha-na-central-do-cliente" style="color: #cc3333; text-decoration: none;"><b>Alterar minha senha da Central do Cliente</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/abertura-de-atividades" style="color: #663399; text-decoration: none;"><b>Solicitações por Atividade</b></a>
    <br><br>

    <a href="https://proocupacional.tawk.help/article/abertura-de-cat-central-do-cliente" style="color: #ff3366; text-decoration: none;"><b>Cadastro Comunicação de Acidente de Trabalho</b></a>
    <br><br>
             <img src="cid:assinatura">
         </body>
       </html>
       """

def abrir_tela_email():
    abrir_tela("PRO OCUPACIONAL - ACESSO CENTRAL DO CLIENTE", "#4DB86F", enviar_email_credenciais)

def abrir_tela_analise():
    abrir_tela("PRO OCUPACIONAL - ANÁLISE CENTRAL DO CLIENTE", "#015FBE", enviar_email_analise)

def abrir_tela_senha():
    abrir_tela("PRO OCUPACIONAL - REDEFINIÇÃO DE SENHA", "#8E0D0D", enviar_email_senha)

def abrir_tela_ajuda():
    abrir_tela("PRO OCUPACIONAL - MANUAL DE AJUDA", "#8F9116", enviar_email_ajuda)

def abrir_tela_suporte():
    abrir_tela_suporte_ti("PRO OCUPACIONAL - SUPORTE T.I", "#FF5733", enviar_email_suporte)

def abrir_tela_suporte_ti(titulo, cor_fundo, funcao_envio):
    tela = tk.Toplevel(root)
    tela.title(titulo)
    tela.configure(bg=cor_fundo)

    destinatario_label = tk.Label(tela, text="E-mail Cópia:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
    global email_copia_entry
    email_copia_entry = tk.Entry(tela, width=30, font=("monospace", 12, "bold"), bg="black", fg="white")
    destinatario_label.grid(row=0, column=0, padx=10, pady=5)
    email_copia_entry.grid(row=0, column=1, padx=10, pady=5)

    anydesk_label = tk.Label(tela, text="Anydesk:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
    global anydesk_entry
    anydesk_entry = tk.Entry(tela, width=30, font=("monospace", 12, "bold"), bg="black", fg="white")
    anydesk_label.grid(row=1, column=0, padx=10, pady=5)
    anydesk_entry.grid(row=1, column=1, padx=10, pady=5)

    motivo_label = tk.Label(tela, text="Motivo:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
    global motivo_text
    motivo_text = tk.Text(tela, width=30, height=5, font=("monospace", 12, "bold"), bg="black", fg="white")
    motivo_label.grid(row=2, column=0, padx=10, pady=5)
    motivo_text.grid(row=2, column=1, padx=10, pady=5)

    anexar_imagem_button = tk.Button(tela, text="Anexar Imagem", command=selecionar_imagem, font=("monospace", 12, "bold"), bg='black', fg='white', width='20')
    anexar_imagem_button.grid(row=3, column=0, columnspan=2, pady=10)

    global anexar_imagem_label
    anexar_imagem_label = tk.Label(tela, text="Nenhuma imagem anexada.", font=("monospace", 10, "bold"), bg=cor_fundo, fg="white")
    anexar_imagem_label.grid(row=4, column=0, columnspan=2)

    enviar_button = tk.Button(tela, text="Enviar E-mail", command=funcao_envio, font=("monospace", 12, "bold"), bg='black', fg='white', width='20')
    enviar_button.grid(row=5, column=0, columnspan=2, pady=10)

def abrir_tela(titulo, cor_fundo, funcao_envio):
    tela = tk.Toplevel(root)
    tela.title(titulo)
    tela.configure(bg=cor_fundo)

    destinatario_label = tk.Label(tela, text="E-mail Destinatário:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
    global destinatario_entry
    destinatario_entry = tk.Entry(tela, width=30, font=("monospace", 12, "bold"), bg="black", fg="white")
    destinatario_label.grid(row=0, column=0, padx=10, pady=5)
    destinatario_entry.grid(row=0, column=1, padx=10, pady=5)

    if funcao_envio == enviar_email_credenciais:
        email_label = tk.Label(tela, text="E-mail do Cliente:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
        global email_entry
        email_entry = tk.Entry(tela, width=30, font=("monospace", 12, "bold"), bg="black", fg="white")
        senha_label = tk.Label(tela, text="Senha do Cliente:", font=("monospace", 12, "bold"), bg=cor_fundo, fg="white")
        global senha_entry
        senha_entry = tk.Entry(tela, width=30, font=("monospace", 12, "bold"), bg="black", fg="white", show="*")
        email_label.grid(row=1, column=0, padx=10, pady=5)
        email_entry.grid(row=1, column=1, padx=10, pady=5)
        senha_label.grid(row=2, column=0, padx=10, pady=5)
        senha_entry.grid(row=2, column=1, padx=10, pady=5)

    enviar_button = tk.Button(tela, text="Enviar E-mail", command=funcao_envio, font=("monospace", 12, "bold"), bg='black', fg='white', width='20')
    enviar_button.grid(row=3, column=0, columnspan=2, pady=10)

root = tk.Tk()
root.title("PRO OCUPACIONAL - PAINEL DE ATENDIMENTO")
root.geometry("300x600")

imagem_principal = Image.open("fractal.png")
imagem_principal = imagem_principal.resize((350, 190), Image.Resampling.LANCZOS)
img = ImageTk.PhotoImage(imagem_principal)
root.iconbitmap('system.ico')
img_label = tk.Label(root, image=img, bd=0)
root.configure(bg='black')
img_label.pack(pady=10)

botao_email = tk.Button(root, text="ENVIO DE CREDENCIAIS", command=abrir_tela_email, font=("monospace", 12, "bold"), bg='black', fg='white', width='25')
botao_email.pack(pady=10)

botao_analise = tk.Button(root, text="ANÁLISE DO ACESSO", command=abrir_tela_analise, font=("monospace", 12, "bold"), bg='black', fg='white', width='25')
botao_analise.pack(pady=10)

botao_senha = tk.Button(root, text="REDEFINIÇÃO DE SENHA", command=abrir_tela_senha, font=("monospace", 12, "bold"), bg='black', fg='white', width='25')
botao_senha.pack(pady=10)

botao_ajuda = tk.Button(root, text="MANUAL DE AJUDA", command=abrir_tela_ajuda, font=("monospace", 12, "bold"), bg='black', fg='white', width='25')
botao_ajuda.pack(pady=10)

botao_suporte = tk.Button(root, text="SUPORTE T.I", command=abrir_tela_suporte, font=("monospace", 12, "bold"), bg='black', fg='white', width='25')
botao_suporte.pack(pady=10)

root.mainloop()
