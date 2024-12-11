import pandas as pd
from src.bot.constants import SOURCE_FILE

sheet = pd.read_excel(SOURCE_FILE)

# line = 108

# valor1 = sheet.at[line - 2, "N_OS"]
# valor2 = sheet.at[line - 2, "Status"]

# print(f"Valor: {valor1} em {valor2}")

sheet.loc[:70, "Status"] = "Finalizada"

sheet.to_excel(SOURCE_FILE, index=False)
print("Planilha atualizada com sucesso.")

# page.window.always_on_top = True


# file = "PREVENTIVAS.xlsx"
# df = pd.read_excel(file)

# for_finalize = df[df["Status"] != "Finalizada"]

# print(len(for_finalize))
