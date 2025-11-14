import win32com.client
import requests
import os
import subprocess
from datetime import datetime
import re

# ========================
# CONFIGURAÇÕES DO SISTEMA
# ========================

EMAIL_CONTA = "gabriel.silva@vonex.com.br"
ASSUNTO = "Planilha faturamento Unipix-SMS"

# Pasta destino do download
PASTA_DESTINO = r"C:\Users\User\Desktop\arquivo_bruto"

# Caminho do script principal que trata o Excel
SCRIPT_AUTOMACAO = r"C:\Users\User\Desktop\automacoes_python\automacao_diaria.py"

# Criar pasta destino, caso não exista
os.makedirs(PASTA_DESTINO, exist_ok=True)

# ========================
# 1. Conectar ao Outlook
# ========================

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
caixa_entrada = outlook.Folders(EMAIL_CONTA).Folders("Caixa de Entrada")

emails = caixa_entrada.Items
emails.Sort("ReceivedTime", True)

html = None

# ========================
# 2. Achar o e-mail correto
# ========================

for msg in emails:
    if msg.Class == 43:  # Email
        if ASSUNTO in msg.Subject:
            html = msg.HTMLBody
            break

if not html:
    print("❌ Nenhum e-mail encontrado com o assunto alvo.")
    exit()

# ========================
# 3. Extrair link do HTML
# ========================

match = re.search(r'href="(https://[^"]+\.xlsx)"', html)

if not match:
    print("❌ Nenhum link .xlsx encontrado dentro do e-mail!")
    exit()

link = match.group(1)
print("🔗 Link encontrado:", link)

# ========================
# 4. Baixar arquivo com nome baseado na data
# ========================

data_str = datetime.now().strftime("%d-%m-%Y")

arquivo_destino = os.path.join(PASTA_DESTINO, f"{data_str}.xlsx")

# Evitar sobrescrição
contador = 1
while os.path.exists(arquivo_destino):
    arquivo_destino = os.path.join(PASTA_DESTINO, f"{data_str} ({contador}).xlsx")
    contador += 1

print("⬇️ Baixando arquivo...")

response = requests.get(link)

if response.status_code == 200:
    with open(arquivo_destino, "wb") as f:
        f.write(response.content)
    print(f"✅ Download concluído: {arquivo_destino}")
else:
    print(f"❌ Erro no download: HTTP {response.status_code}")
    exit()

# ========================
# 5. Executar automacao_diaria.py
# ========================

print("⚙️ Iniciando processamento do Excel pela automação principal...")

try:
    subprocess.run(["python", SCRIPT_AUTOMACAO], check=True)
    print("✅ automacao_diaria concluído com sucesso!")
except subprocess.CalledProcessError:
    print("❌ Erro ao executar automacao_diaria.py")
    exit()

# ========================
# FINAL
# ========================

print("\n🎉 PROCESSO COMPLETO FINALIZADO COM SUCESSO!")
print("✔ Download do arquivo")
print("✔ Tratamento no Excel")
print("✔ Automação Python executada")
