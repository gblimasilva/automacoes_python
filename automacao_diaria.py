import time
import win32com.client as win32
import pythoncom
import subprocess

# Caminho do arquivo Excel com a macro
caminho_excel = r"C:\Users\User\Desktop\teste_automacao\Painel_automacao.xlsm"

# Inicia COM
pythoncom.CoInitialize()

print("🟢 Abrindo Excel...")

# Abre Excel (invisível para estabilidade, opcional)
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Pode deixar False se quiser rodar totalmente oculto

# Abre seu arquivo principal
wb = excel.Workbooks.Open(caminho_excel)
print("🟢 Arquivo carregado.")

# Aguarda carregamento
time.sleep(3)

# Executa a macro silenciosa
print("🟡 Executando macro...")
excel.Application.Run("Painel_automacao.xlsm!AtualizarAutomacao")

# 🔄 Atualizar o arquivo final automaticamente
print("🔄 Atualizando arquivo final...")

caminho_final = r"C:\Users\User\Desktop\Atualizacoes_diarias\SMS_Minutagem_dia.xlsx"

# Abre o arquivo final
wb_final = excel.Workbooks.Open(caminho_final)

# Atualiza tudo (Power Query, conexões, tabelas, etc.)
wb_final.RefreshAll()

wb_final.RefreshAll()
time.sleep(1)  # pequena pausa inicial

# ⏳ Aguarda todas as conexões terminarem
def consultas_ativas(wb):
    """Retorna True se alguma conexão ainda estiver atualizando."""
    for conn in wb.Connections:
        try:
            if conn.OLEDBConnection.Refreshing:
                return True
        except:
            pass
        try:
            if conn.ODBCConnection.Refreshing:
                return True
        except:
            pass
    return False

while consultas_ativas(wb_final):
    time.sleep(1)


# Salva o arquivo já atualizado
wb_final.Close(SaveChanges=True)

print("✅ Arquivo final atualizado e salvo com sucesso!")

