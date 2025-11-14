import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

html_content = None

for msg in messages:
    if "Planilha faturamento Unipix-SMS".lower() in msg.Subject.lower():
        html_content = msg.HTMLBody
        break

if html_content:
    with open("html_email.txt", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("HTML extraído com sucesso para o arquivo 'html_email.txt'")
else:
    print("Nenhum e-mail com o assunto esperado foi encontrado.")
