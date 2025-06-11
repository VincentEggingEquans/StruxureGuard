import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import xlwings as xw
import os

# Hardcoded data uit het Excel-bestand
CONTRACTNIVEAU_KOLOMMEN = ["Omschrijving", "Basis", "Standaard", "Totaal"]
CONTRACTNIVEAU_TABEL = [
    ["HARDWARE REGELPANEEL", "", "", ""],
    ["Voor aanvang LMRA EQUANS uitvoeren.", "x", "x", "x"],
    ["Visuele NEN inspectie van bedrading, relais etc. in het regelpaneel", "x", "x", "x"],
    ["Warmtebeeldcontrole van het regelpaneel", "x", "x", "x"],
    ["Kastventilatie werkt naar behoren en is stofvrij", "x", "x", "x"],
    ["Bedienbeeldscherm regelkast werkt naar behoren", "x", "x", "x"],
    ["Regeltechnische omschrijving aanwezig en up-to-date, datum en versie", "x", "x", "x"],
    ["Regelkastschema aanwezig en up-to-date, datum en versie", "x", "x", "x"],
    ["DDC HARDWARE", "", "", ""],
    ["Meting voedingsspanning regelaar(s)", "x", "x", "x"],
    ["Interventieschakelaars op automatisch ingesteld", "x", "x", "x"],
    ["Werking UPS en standtijd controleren, indien aanwezig", "x", "x", "x"],
    ["Tekstlabels regelaar(s) en IO-modulen aanwezig", "x", "x", "x"],
    ["DDC FIRMWARE", "", "", ""],
    ["Benodigde licenties aanwezig", "x", "x", "x"],
    ["Trendstorage geheugen", "x", "x", "x"],
    ["Systeemdatum en -tijd correct", "x", "x", "x"],
    ["Controle softwareversie op regelaar, modules en naregelingen", "x", "x", "x"],
    ["Systeemload CPU/geheugen controleren", "x", "x", "x"],
    ["Map Opmerkingen nakijken in regelaar (logboek)", "x", "x", "x"],
    ["Back-up maken van controler", "x", "x", "x"],
    ["Controleren aanwezigheid automatische back-up ", "x", "x", "x"],
    ["Integriteit automatische back-up", "x", "x", "x"],
    ["Systeem events (100) controleren", "x", "x", "x"],
    ["FUNCTIONELE SOFTWARE (FYSIEK)", "", "", ""],
    ["Map Programma's controleren", "x", "x", "x"],
    ["Binding Diagnostics controleren", "x", "x", "x"],
    ["Standaard loggen aanwezig", "", "x", "x"],
    ["Extended loggen aanwezig", "", "x", "x"],
    ["Standen softwareschakelaars op automatisch", "x", "x", "x"],
    ["I/O softwarematig geforceerd", "x", "x", "x"],
    ["Values geforceerd", "x", "x", "x"],
    ["Alarmen disabled", "x", "x", "x"],
    ["Aantal urgente en niet urgente alarmen noteren", "x", "x", "x"],
    ["Storingen overzicht en historie vastleggen", "x", "x", "x"],
    ["Instellingen prioriteiten storingen", "", "x", "x"],
    ["Doormeldingen verzamelstoringen controleren", "x", "x", "x"],
    ["Controleren werking overwerkschakelaars (software)", "", "", "x"],
    ["Controleren klok- en vakantieprogramma's", "", "", "x"],
    ["Controleren aanwezigheid van zomer-/winterblokkeringen", "", "x", "x"],
    ["Werking snelkoppelingen grafische interface", "", "", "x"],
    ["Regelkringen berekend setpoint/actuele waarde", "", "", "x"],
    ["Werking brandschakelaars/schakelingen, volgens PVE en RTO", "", "", "x"],
    ["Functietest scenario's: thermostaten vorst, minimaal, maximaal, druk", "", "", "x"],
    ["Functietest schakelende servomotoren", "", "", "x"],
    ["Warmtebeeldcontrole van aangesloten naregelingen", "", "", "x"],
]

class RapportageGenerator(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("StruxureGuard - Rapportage Generator")
        self.resizable(True, True)
        self._create_widgets()
        self.after(100, self._adjust_window_size)

    def _adjust_window_size(self):
        self.update_idletasks()
        req_width = self.winfo_reqwidth()
        req_height = self.winfo_reqheight()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        max_width = int(screen_width * 0.8)
        max_height = int(screen_height * 0.8)
        final_width = min(req_width, max_width)
        final_height = min(req_height, max_height)
        self.geometry(f"{final_width}x{final_height}")

    def _create_widgets(self):
        self.entries = {}

        combobox_velden = {
            "Contractniveau:": ["Basis", "Standaard", "Totaal"],
            "Type gebouwgebruik:": ["Kantoor (gehuurd)", "Kantoor (verhuurd)", "Kantoor eigen gebruik", "School", "Ziekenhuis", "Overig"],
            "Aantal onderhoudsrapportages in contractjaar": ["1", "2", "3", "4"],
            "Meet- en regel onderhoudsrapportage": ["1", "2", "3", "4"],
            "Merk regelinstallatie": ["Schneider Electric"],
            "Type regelinstallatie": ["Ecostruxure Building Operation"],
            "Naregelingen op het GBS aangesloten": ["geen naregelingen aanwezig", "middels BACnet IP", "middels Modbus IP", "middels BACnet MS/tp", "middels Modbus RTU", "middels diverse protocollen"],
            "Type naregelingen": ["B3 serie", "RP-C serie", "MP-C serie", "6-wegventielen", "diverse"],
            "Hoofd gasmeter uitgelezen op GBS": ["Nee", "Ja"],
            "Hoofd elektrameter uitgelezen op GBS": ["Nee", "Ja"],
            "Warmte wordt opgewekt middels": ["nvt", "ketel(s)", "warmtepomp(en)", "bron(nen)", "WKO (bron en WP)"],
            "Koude wordt opgewekt middels": ["nvt", "ketel(s)", "warmtepomp(en)", "bron(nen)", "WKO (bron en WP)"],
        }

        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(top_frame, text="Templatepad:").grid(row=0, column=0, sticky="w", pady=(6, 0))
        self.template_path_var = tk.StringVar()
        self.template_entry = ttk.Entry(top_frame, textvariable=self.template_path_var, width=50, state="readonly")
        self.template_entry.grid(row=0, column=1, sticky="w", pady=(6, 0))
        ttk.Button(top_frame, text="Selecteer Template", command=self._select_template).grid(row=0, column=2, padx=(10, 0), pady=(6, 0))

        ttk.Label(top_frame, text="Wachtwoord:").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.password_entry = ttk.Entry(top_frame, show="*", width=30)
        self.password_entry.grid(row=1, column=1, sticky="w")

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10)

        left_column = ttk.Frame(container)
        right_column = ttk.Frame(container)
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        right_column.grid(row=0, column=1, sticky="nsew")

        secties_links = {
            "KLANT- EN CONTRACTINFORMATIE": [
                "Klantnaam/Gebouwnaam:", "Locatie:", "Adres:", "Type gebouwgebruik:",
                "__SEPARATOR__",
                "Contactpersoon technische dienst:", "Telefoonnummer contactpersoon:", "Email contactpersoon:",
                "__SEPARATOR__",
                "Contactpersoon contract:", "Telefoonnummer contactpersoon:", "Email contactpersoon:",
                "__SEPARATOR__",
                "Contractjaar:", "Aantal onderhoudsrapportages in contractjaar", "Contractniveau:"
            ],
            "INFORMATIE EQUANS SERVICES": [
                "Onderhoud uitgevoerd door:", "Rapportage opgesteld door:", "Contractmanager Services:"
            ],
            
        }

        secties_rechts = {
            "INFORMATIE ONDERHOUDSBEURT": [
                "Meet- en regel onderhoudsrapportage", "Datum of periode uitgevoerde onderhoud:"
            ],
            "INFORMATIE REGELINSTALLATIE": [
                "Merk regelinstallatie", "Type regelinstallatie", "Versie GBS software:",
                "Aantal centrale regelpanelen aanwezig:", "Naregelingen op het GBS aangesloten", "Type naregelingen",
                "Aantal floormanagerpanelen aanwezig:", "Aantal naregelingen aanwezig:", "Aantal ruimtebedieningen aanwezig:"
            ],
            "INFORMATIE KLIMAATINSTALLATIES EN ENERGIE": [
                "Hoofd gasmeter uitgelezen op GBS", "Hoofd elektrameter uitgelezen op GBS",
                "Warmte wordt opgewekt middels", "Koude wordt opgewekt middels",
                "Aantal aanwezige luchtbehandelingskasten:"
            ]
        }

        def plaats_secties(parent, secties):
            entry_width = 40
            combobox_width = 37
            for titel, velden in secties.items():
                frame = ttk.LabelFrame(parent, text=titel)
                frame.pack(fill="x", pady=4)
                frame.columnconfigure(0, minsize=220)
                frame.columnconfigure(1, minsize=320, weight=1)

                for idx, veld in enumerate(velden):
                    if veld == "__SEPARATOR__":
                        ttk.Separator(frame, orient="horizontal").grid(row=idx, column=0, columnspan=2, sticky="ew", pady=4)
                        continue

                    label = ttk.Label(frame, text=veld, anchor="w")
                    label.grid(row=idx, column=0, sticky="w", padx=(10, 8), pady=2)

                    if veld in combobox_velden:
                        values = combobox_velden[veld]
                        if veld == "Contractniveau:":
                            help_frame = ttk.Frame(frame)
                            help_frame.grid(row=idx, column=1, padx=(0, 11), pady=2, sticky="e")
                            combobox = ttk.Combobox(help_frame, values=values, width=32)
                            combobox.grid(row=0, column=0, sticky="ew")
                            help_button = ttk.Button(help_frame, text="?", width=3, command=self._toon_contractniveau_popup)
                            help_button.grid(row=0, column=1, padx=(3, 0))
                            self.entries[veld] = combobox
                        else:
                            entry = ttk.Combobox(frame, values=values, width=combobox_width)
                            entry.grid(row=idx, column=1, padx=(0, 11), pady=2, sticky="e")
                            self.entries[veld] = entry
                    else:
                        entry = ttk.Entry(frame, width=entry_width)
                        entry.grid(row=idx, column=1, padx=(0, 10), pady=2, sticky="e")
                        self.entries[veld] = entry

        plaats_secties(left_column, secties_links)
        plaats_secties(right_column, secties_rechts)

        actions_frame = ttk.Frame(self)
        actions_frame.pack(pady=12)
        ttk.Button(actions_frame, text="Genereer Rapportage", command=self._generate_report).grid(row=0, column=0, padx=10)
        ttk.Button(actions_frame, text="Exporteer naar XML", command=self._export_to_xml).grid(row=0, column=1, padx=10)
        ttk.Button(actions_frame, text="Importeer vanuit XML", command=self._import_from_xml).grid(row=0, column=2, padx=10)

    def _toon_contractniveau_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Toelichting Contractniveau")

        style = ttk.Style(popup)
        style.configure("Bold.TREEVIEW", font=("TkDefaultFont", 10, "bold"))

        # Treeview met scrollbar in apart frame
        frame = ttk.Frame(popup)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(frame, columns=CONTRACTNIVEAU_KOLOMMEN, show="headings", height=20)
        for col in CONTRACTNIVEAU_KOLOMMEN:
            tree.heading(col, text=col)
            if col == "Omschrijving":
                tree.column(col, width=450, anchor="w")  # links uitgelijnd
            else:
                tree.column(col, width=70, anchor="center")  # gecentreerd


        tree.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))

        for row in CONTRACTNIVEAU_TABEL:
            if all(cell == "" for cell in row[1:]):
                tree.insert("", "end", values=row, tags=("bold",))
            else:
                tree.insert("", "end", values=row)

        # Scrollbar toevoegen
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # Sluitknop
        ttk.Button(popup, text="Sluiten", command=popup.destroy).pack(pady=10)

        # Laat tkinter de juiste grootte bepalen
        popup.update_idletasks()
        w = min(popup.winfo_reqwidth(), 1000)
        popup.minsize(width=w, height=900)
        popup.geometry(f"{w}x600")


    def _select_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Templates", "*.docx *.dotx *.xlsx *.xml *.xlsm"), ("Alle bestanden", "*.*")])
        if file_path:
            self.template_path_var.set(file_path)
            self._adjust_window_size()
            if file_path.lower().endswith(".xlsm"):
                self._laad_templategegevens(file_path)

    def _laad_templategegevens(self, excel_pad):
        import xlwings as xw

        if not os.path.exists(excel_pad):
            messagebox.showerror("Fout", f"Bestand niet gevonden: {excel_pad}")
            return

        try:
            app = xw.App(visible=False)
            wb = app.books.open(excel_pad)

            if "Gegevens" not in [s.name for s in wb.sheets]:
                wb.close()
                app.quit()
                messagebox.showerror("Fout", "Tabblad 'Gegevens' niet gevonden in Excelbestand.")
                return

            sheet = wb.sheets["Gegevens"]

            # Laad kolom A en B tot maximaal 200 rijen (voorkomt grote selecties)
            kolom_a = sheet.range("A1:A200").value
            kolom_b = sheet.range("B1:B200").value

            wb.close()
            app.quit()
        except Exception as e:
            try:
                app.quit()
            except:
                pass
            messagebox.showerror("Fout", f"Fout bij uitlezen Excelbestand: {str(e)}")
            return

        # Zorg dat we altijd met lijsten werken
        if not isinstance(kolom_a, list):
            kolom_a = [kolom_a]
        if not isinstance(kolom_b, list):
            kolom_b = [kolom_b]

        # Zet sleutel/waarde-paren in dictionary
        data_dict = {}
        for sleutel, waarde in zip(kolom_a, kolom_b):
            if isinstance(sleutel, str) and sleutel.strip():
                sleutel_clean = sleutel.strip().rstrip(":")
                waarde_str = str(waarde).strip() if waarde else ""
                if waarde_str:
                    data_dict[sleutel_clean] = waarde_str

        label_mapping = {
        "Klantnaam": "Klantnaam/Gebouwnaam",
        "Locatie": "Locatie",
        "Adres": "Adres",
        "Type gebouwgebruik door klant": "Type gebouwgebruik",
        "Contactpersoon technische dienst, indien van toepassing:": "Contactpersoon technische dienst",
        "Telefoonnummer contactpersoon:": "Telefoonnummer contactpersoon",
        "Email contactpersoon:": "Email contactpersoon",
        "Contactpersoon contract": "Contactpersoon contract",
        "Telefoonnummer contactpersoon:": "Telefoonnummer contactpersoon",
        "Email contactpersoon:": "Email contactpersoon",
        "Contractjaar": "Contractjaar",
        "Aantal onderhoudsrapportages in contractjaar": "Aantal onderhoudsrapportages in contractjaar",
        "Contractniveau": "Contractniveau",
        "Onderhoud uitgevoerd door": "Onderhoud uitgevoerd door",
        "Rapportage opgesteld door": "Rapportage opgesteld door",
        "Contractmanager Services": "Contractmanager Services",
        "Meet- en regel onderhoudsrapportage": "Meet- en regel onderhoudsrapportage",
        "Datum of periode uitgevoerde onderhoud:": "Datum of periode uitgevoerde onderhoud",
        "Merk regelinstallatie": "Merk regelinstallatie",
        "Type regelinstallatie": "Type regelinstallatie",
        "Versie GBS software, indien van toepassing": "Versie GBS software",
        "Aantal centrale regelpanelen aanwezig": "Aantal centrale regelpanelen aanwezig",
        "Naregelingen aangesloten op GBS": "Naregelingen op het GBS aangesloten",
        "Type naregelingen": "Type naregelingen",
        "Aantal floormanagerpanelen aanwezig": "Aantal floormanagerpanelen aanwezig",
        "Aantal naregelingen aanwezig": "Aantal naregelingen aanwezig",
        "Aantal ruimtebedieningen aanwezig": "Aantal ruimtebedieningen aanwezig",
        "Gasmeter op GBS": "Hoofd gasmeter uitgelezen op GBS",
        "Elektrameter op GBS": "Hoofd elektrameter uitgelezen op GBS",
        "Warmteopwekking": "Warmte wordt opgewekt middels",
        "Koudeopwekking": "Koude wordt opgewekt middels",
        "Aantal aanwezige luchtbehandelingskasten": "Aantal aanwezige luchtbehandelingskasten"
    }
        for label, entry in self.entries.items():
            gui_label = label.strip(":")
            # Zoek sleutel in mapping, anders neem de gui_label zelf
            excel_key = next((k for k, v in label_mapping.items() if v == gui_label), gui_label)
            if excel_key in data_dict:
                try:
                    if isinstance(entry, ttk.Combobox):
                        # Kijk of de waarde in de lijst staat; zo niet, voeg tijdelijk toe
                        combowaarde = data_dict[excel_key]
                        if combowaarde not in entry['values']:
                            entry['values'] = list(entry['values']) + [combowaarde]
                        entry.set(combowaarde)
                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, data_dict[excel_key])
                except Exception:
                    pass

    def _generate_report(self):
        print("Rapportage wordt gegenereerd met de volgende gegevens:")
        print(f"Wachtwoord: {self.password_entry.get()}")
        print(f"Templatepad: {self.template_path_var.get()}")
        for label, entry in self.entries.items():
            print(f"{label} {entry.get()}")

    def _export_to_xml(self):
        data = {
            "Wachtwoord": self.password_entry.get(),
            "Templatepad": self.template_path_var.get(),
            **{label: entry.get() for label, entry in self.entries.items()}
        }
        root = ET.Element("Rapportage")
        for key, value in data.items():
            child = ET.SubElement(root, "Veld", naam=key)
            child.text = value
        file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
        if file_path:
            tree = ET.ElementTree(root)
            tree.write(file_path, encoding="utf-8", xml_declaration=True)
            messagebox.showinfo("Succes", "Gegevens succesvol geëxporteerd naar XML.")

    def _import_from_xml(self):
        file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        if file_path:
            tree = ET.parse(file_path)
            root = tree.getroot()
            for veld in root.findall("Veld"):
                naam = veld.attrib.get("naam")
                waarde = veld.text or ""
                if naam == "Wachtwoord":
                    self.password_entry.delete(0, tk.END)
                    self.password_entry.insert(0, waarde)
                elif naam == "Templatepad":
                    self.template_path_var.set(waarde)
                elif naam in self.entries:
                    self.entries[naam].delete(0, tk.END)
                    self.entries[naam].insert(0, waarde)
            self._adjust_window_size()
            messagebox.showinfo("Succes", "Gegevens succesvol geïmporteerd uit XML.")


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    app = RapportageGenerator(master=root)
    app.mainloop()
