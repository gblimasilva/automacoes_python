import win32com.client
import requests
import os
from datetime import datetime
import re

# 📨 Configurações
EMAIL_CONTA = "gabriel.silva@vonex.com.br"
ASSUNTO = "Planilha faturamento Unipix-SMS"
PASTA_DESTINO = r"C:\Users\User\Desktop\arquivo_bruto_"

if not os.path.exists(PASTA_DESTINO):
    os.makedirs(PASTA_DESTINO)

# === 1. Conectar ao Outlook ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

caixa_entrada = outlook.Folders(EMAIL_CONTA).Folders("Caixa de Entrada")

emails = caixa_entrada.Items
emails.Sort("ReceivedTime", True)

html = None

# === 2. Localizar e-mail correto ===
for msg in emails:
    if msg.Class == 43:  # Email
        if ASSUNTO in msg.Subject:
            html = msg.HTMLBody
            break

if not html:
    print("❌ Nenhum e-mail com o assunto encontrado.")
    exit()

# === 3. Extrair link do HTML ===
match = re.search(r'href="(https://[^"]+\.xlsx)"', html)
if not match:
    print("❌ Nenhum link .xlsx encontrado no e-mail!")
    exit()

link = match.group(1)
print("🔗 Link encontrado:", link)

# === 4. Baixar arquivo ===
# 🔥 Aqui está a alteração: formato de data desejado (dd-mm-aaaa)
data_str = datetime.now().strftime("%d-%m-%Y")

arquivo_destino = os.path.join(PASTA_DESTINO, f"{data_str}.xlsx")

print("⬇️ Baixando arquivo...")

response = requests.get(link)

if response.status_code == 200:
    with open(arquivo_destino, "wb") as f:
        f.write(response.content)
    print(f"✅ Download concluído: {arquivo_destino}")
else:
    print(f"❌ Erro no download: HTTP {response.status_code}")
