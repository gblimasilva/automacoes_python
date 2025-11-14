import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print("\n=== CONTAS ENCONTRADAS NO OUTLOOK ===\n")

for i in range(outlook.Folders.Count):
    pasta = outlook.Folders.Item(i + 1)
    print(f"{i+1} → {pasta.Name}")
