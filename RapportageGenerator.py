import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET

class RapportageGenerator(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("StruxureGuard - Rapportage Generator")
        self.geometry("600x800")
        self.resizable(False, False)

        self._create_widgets()

    def _create_widgets(self):
        self.entries = {}

        veldstructuur = [
            ("KLANT- EN CONTRACTINFORMATIE", [
                "Klantnaam:", "Locatie:", "Adres:", "Type gebouwgebruik:",
                "__SEPARATOR__",
                "Contactpersoon technische dienst:",
                "Telefoonnummer contactpersoon:", "Email contactpersoon:",
                "__SEPARATOR__",
                "Contactpersoon contract:", "Telefoonnummer contactpersoon:",
                "Email contactpersoon:",
                "__SEPARATOR__",
                "Contractjaar:", "Contractniveau:"
            ]),
            ("INFORMATIE EQUANS SERVICES", [
                "Onderhoud uitgevoerd door:", "Rapportage opgesteld door:",
                "Contractmanager Services:"
            ]),
            ("INFORMATIE ONDERHOUDSBEURT", [
                "Datum of periode uitgevoerde onderhoud:"
            ]),
            ("INFORMATIE REGELINSTALLATIE", [
                "Versie GBS software:",
                "Aantal centrale regelpanelen aanwezig:",
                "Aantal floormanagerpanelen aanwezig:",
                "Aantal naregelingen aanwezig:",
                "Aantal ruimtebedieningen aanwezig:"
            ]),
            ("INFORMATIE KLIMAATINSTALLATIES EN ENERGIE", [
                "Aantal aanwezige luchtbehandelingskasten:"
            ])
        ]

        entry_width = 40
        combobox_width = 37

        for titel, velden in veldstructuur:
            frame = ttk.LabelFrame(self, text=titel)
            frame.pack(fill="x", padx=10, pady=4)
            frame.columnconfigure(0, minsize=220)
            frame.columnconfigure(1, minsize=320, weight=1)

            for idx, veld in enumerate(velden):
                if veld == "__SEPARATOR__":
                    ttk.Separator(frame, orient="horizontal").grid(row=idx, column=0, columnspan=2, sticky="ew", pady=4)
                    continue

                label = ttk.Label(frame, text=veld, anchor="w")
                label.grid(row=idx, column=0, sticky="w", padx=(10, 8), pady=2)

                if veld == "Contractniveau:":
                    entry = ttk.Combobox(frame, values=["Basis", "Standaard", "Totaal"], width=combobox_width)
                    entry.grid(row=idx, column=1, padx=(0, 11), pady=2, sticky="e")
                    self.entries[veld] = entry

                elif veld == "Type gebouwgebruik:":
                    type_frame = ttk.Frame(frame)
                    type_frame.grid(row=idx, column=1, padx=(0, 11), pady=2, sticky="e")

                    combobox = ttk.Combobox(type_frame, values=[
                        "Kantoor (gehuurd)", "Kantoor (verhuurd)", "Kantoor eigen gebruik",
                        "School", "Ziekenhuis", "Overig"
                    ], width=combobox_width)
                    combobox.grid(row=0, column=0, sticky="ew")

                    other_label = ttk.Label(type_frame, text="Vul type in")
                    other_entry = ttk.Entry(type_frame, width=entry_width)
                    other_label.grid(row=1, column=0, sticky="w", pady=(6, 2))
                    other_entry.grid(row=2, column=0, sticky="ew")
                    other_label.grid_remove()
                    other_entry.grid_remove()

                    self.entries[veld] = combobox
                    self.entries[veld + " (specificeer)"] = other_entry

                    def toggle_other_field(event):
                        if combobox.get() == "Overig":
                            other_label.grid()
                            other_entry.grid()
                        else:
                            other_label.grid_remove()
                            other_entry.grid_remove()

                    combobox.bind("<<ComboboxSelected>>", toggle_other_field)

                else:
                    entry = ttk.Entry(frame, width=entry_width)
                    entry.grid(row=idx, column=1, padx=(0, 10), pady=2, sticky="e")
                    self.entries[veld] = entry

        # Actieknoppen
        actions_frame = ttk.Frame(self)
        actions_frame.pack(pady=12)

        ttk.Button(actions_frame, text="Genereer Rapportage", command=self._generate_report).grid(row=0, column=0, padx=10)
        ttk.Button(actions_frame, text="Exporteer naar XML", command=self._export_to_xml).grid(row=0, column=1, padx=10)
        ttk.Button(actions_frame, text="Importeer vanuit XML", command=self._import_from_xml).grid(row=0, column=2, padx=10)

    def _generate_report(self):
        print("Rapportage wordt gegenereerd met de volgende gegevens:")
        for label, entry in self.entries.items():
            print(f"{label} {entry.get()}")

    def _export_to_xml(self):
        data = {label: entry.get() for label, entry in self.entries.items()}
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
                if naam in self.entries:
                    self.entries[naam].delete(0, tk.END)
                    self.entries[naam].insert(0, waarde)
            messagebox.showinfo("Succes", "Gegevens succesvol geïmporteerd uit XML.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    app = RapportageGenerator(master=root)
    app.mainloop()
