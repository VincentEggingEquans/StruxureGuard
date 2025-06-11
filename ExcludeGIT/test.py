import xlwings as xw

bestand = r"VERSIE 202504 PANELEN - Onderhoudsrapportage KLANTNAAM CHECKBOXFIX.xlsm"

app = xw.App(visible=False)
wb = app.books.open(bestand)

if "Gegevens" in [s.name for s in wb.sheets]:
    sheet = wb.sheets["Gegevens"]
    data = sheet.range("A1").expand().value
    for rij in data:
        print(rij)
else:
    print("Tabblad 'Gegevens' niet gevonden.")

wb.close()
app.quit()