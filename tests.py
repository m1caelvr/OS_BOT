import pandas as pd

sheet = pd.read_excel("PREVENTIVAS.xlsx")

# line = 108

# valor1 = sheet.at[line - 2, "N_OS"]
# valor2 = sheet.at[line - 2, "Status"]

# print(f"Valor: {valor1} em {valor2}")

sheet.loc[:32, "Status"] = "Finalizada"

sheet.to_excel("PREVENTIVAS.xlsx", index=False)
print("Planilha atualizada com sucesso.")

# page.window.always_on_top = True
