# #GMAIL
# import yagmail

# # - Você precisa liberar o seu e-mail para esse tipo de atividade.
# # https://myaccount.google.com/lesssecureapps

# # fazer loguin conta gmail
# usuario = yagmail.SMTP(user= 'leandrolopescfal@gmail.com', password= 'Leo171020!')
# # enviar email
# usuario.send(to= 'leandroporocachoka@gmail.com', subject= 'Mr. Robot',contents= 'Control is an illusion')
# # enviar email com anexo
# usuario.send(to='joaoprlira@gmail.com', subject='Relatório Financeiro',  \
#              contents='Prezado Lira,\nSegue em anexo o Relatório Financeiro\nAtt.,', attachments='Financeiro.xlsx')
# # enviar mesmo email para mais de um destinatário
# usuario.send(to='joaoprlira@gmail.com', bcc='hashtagtreinamentos@gmail.com', \
#              subject='Meu primeiro Email no Python', contents='Fala ai, Lira, meu 1º email')

# # personalizar texto do email com python
# #primeira forma: -> lista de frases
# nome = 'Lira Hashtag'
# corpo_email = [
#     'Fala {}, tranquilo?'.format(nome),
#     'Envio esse e-mail para te passar o relatório de vendas do ano passado',
#     'Att.,',
#     'João'
# ]
# corpo_email = '\n'.join(corpo_email)

# #segunda forma: -> string de várias linhas
# #corpo_email = '''
# #Fala Lira, tranquilo?
# #Envio esse e-mail para te passar o relatório de vendas do ano passado.
# #Att.,
# #João
# #'''

# # personalizar texto do email com HTML
# corpo_email = '''
# <p>Fala <b>Lira</b>, tranquilo?</p>
# <p>Envio esse e-mail para te passar o relatório de vendas do ano passado.</p>
# <p>Att.,</p>
# <p>João</p>
# '''

# usuario.send(to='joaoprlira@gmail.com', subject='Meu segundo Email no Python', contents=corpo_email)


# # OUTLOOK
# # precisa ter o app instalado

# import win32com.client as win32
# outlook = win32.Dispatch('outlook.application')  #code padrão

# mail = outlook.CreateItem(0)  #code padrão
# mail.To = 'joaoprlira@gmail.com'
# mail.CC = 'email@gmail.com'  #para criar copia do email
# mail.BCC = 'email@gmail.com'  #para enviar email oculto
# mail.Subject = 'Email vindo do Outlook'  #titulo email
# mail.Body = 'Texto do E-mail'
# #ou mail.HTMLBody = '<p>Corpo do Email em HTML</p>'

# # Anexos (pode colocar quantos quiser):
# attachment  = r'C:\Users\joaop\Google Drive\Python Impressionador\Financeiro.xlsx'
# mail.Attachments.Add(attachment)
# mail.Send()


# # acessar conteúdo dentro do email:

# from imap_tools import MailBox, AND

# # pegar emails de um remetente para um destinatário
# username = "seu_email"
# password = "senha"

# # lista de imaps: https://www.systoolsgroup.com/imap/
# meu_email = MailBox('imap.gmail.com').login(username, password)

# # criterios: https://github.com/ikvk/imap_tools#search-criteria
# lista_emails = meu_email.fetch(AND(from_="remetente", to="destinatario")) 
# for email in lista_emails:
#     print(email.subject)
#     print(email.text)

# # pegar emails com um anexo específico
# lista_emails = meu_email.fetch(AND(from_="remetente"))
# for email in lista_emails:
#     if len(email.attachments) > 0:
#         for anexo in email.attachments:
#             if "TituloAnexo" in anexo.filename:
#                 print(anexo.content_type)
#                 print(anexo.payload)
#                 with open("Teste.xlsx", 'wb') as arquivo_excel:
#                     arquivo_excel.write(anexo.payload)

import yagmail

# fazer loguin conta gmail
usuario = yagmail.SMTP(user= 'leandrolopescfal@gmail.com', password= 'Leo171020!')
# enviar email
usuario.send(to= 'leandroporocachoka@gmail.com', subject= 'Mr. Robot',contents= 'Control is an illusion')
# enviar email com anexo
# usuario.send(to='joaoprlira@gmail.com', subject='Relatório Financeiro',  \
#              contents='Prezado Lira,\nSegue em anexo o Relatório Financeiro\nAtt.,', attachments='Financeiro.xlsx')
# enviar mesmo email para mais de um destinatário
# usuario.send(to='joaoprlira@gmail.com', bcc='hashtagtreinamentos@gmail.com', \
#              subject='Meu primeiro Email no Python', contents='Fala ai, Lira, meu 1º email')
