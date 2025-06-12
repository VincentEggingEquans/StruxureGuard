import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import xlwings as xw
import os
import logging
import threading
import win32com.client as win32
import time
from functools import partial
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

# Attach handler only once (recommended in main or top-level window)
if not any(isinstance(h, TkinterLogHandler) for h in logging.getLogger().handlers):
    handler = TkinterLogHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logging.getLogger().addHandler(handler)
    logging.getLogger().setLevel(logging.INFO)

def lees_formulierbesturingselementen(excel_pad):
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        time.sleep(1)
        try:
            excel.Visible = False
        except Exception:
            logger.warning("Excel.Visible kon niet worden ingesteld (wordt overgeslagen)")

        wb = excel.Workbooks.Open(excel_pad)
        resultaten = {}

        try:
            sheet = wb.Sheets("Gegevens")
        except Exception as e:
            logger.error(f"Tabblad 'Gegevens' niet gevonden: {e}")
            wb.Close(False)
            return {}

        for shape in sheet.Shapes:
            if shape.Type == 8:  # Form Control
                try:
                    control = shape.ControlFormat
                    if not control.ListFillRange:
                        continue  # Alleen elementen met lijst: dus dropdowns

                    naam = shape.Name
                    index = int(control.Value)
                    range_str = control.ListFillRange

                    if index > 0 and '!' in range_str:
                        sheet_name, cell_range = range_str.split('!')
                        sheet_name = sheet_name.replace("'", "")
                        try:
                            data_sheet = wb.Sheets(sheet_name)
                            excel_range = data_sheet.Range(cell_range)
                            lijst = [cell.Value for cell in excel_range]
                            if 0 < index <= len(lijst):
                                resultaten[naam] = str(lijst[index - 1]).strip()
                            else:
                                resultaten[naam] = ""
                        except Exception as e:
                            logger.warning(f"Kon bereik '{range_str}' niet uitlezen voor '{naam}': {e}")
                            resultaten[naam] = ""
                except Exception as e:
                    logger.warning(f"Dropdown '{shape.Name}' niet uitgelezen: {e}")

        wb.Close(SaveChanges=False)
        return resultaten

    except Exception as outer_e:
        logger.error(f"Dropdowns niet geladen: {outer_e}")
        return {}

    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception as quit_e:
                logger.warning(f"Kon Excel niet afsluiten: {quit_e}")

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
    """
    Rapportage Generator window for StruxureGuard.
    Allows user to fill in contract and installation data, import/export XML, and load Excel template data.
    """

    def __init__(self, master=None):
        logger.info("Initialiseren RapportageGenerator venster")
        super().__init__(master)
        style = ttk.Style(self)
        style.configure("InvulRood.TEntry", fieldbackground="#ffcccc")
        style.configure("InvulRood.TCombobox", fieldbackground="#ffcccc")
        self.entry_wrappers = {}  # veldnaam → wrapper-frame (voor styling)
        self.title("StruxureGuard - Rapportage Generator")
        self.resizable(True, True)
        self._create_widgets()
        self.after(100, self._adjust_window_size)
        self.after(150, self._controleer_alle_velden)
        logger.info("RapportageGenerator venster aangemaakt")

    def _markeer_invullen_veld(self, widget, actief=True):
        try:
            veld = next((k for k, v in self.entries.items() if v == widget), None)
            warning_label = self.entry_wrappers.get(f"{veld}_warning")
            if warning_label:
                if actief:
                    warning_label.grid()
                else:
                    warning_label.grid_remove()
        except Exception as e:
            logger.warning(f"Kan veldmarkering niet aanpassen: {e}")

    def _controleer_alle_velden(self):
        """Controleer alle velden op leegte of 'INVULLEN' en toon waarschuwingen."""
        for entry in self.entries.values():
            self._controleer_en_markeer(entry)

    def _controleer_en_markeer(self, widget):
        try:
            waarde = widget.get().strip()
            actief = (not waarde) or ("INVULLEN" in waarde.upper())
            self._markeer_invullen_veld(widget, actief=actief)
        except Exception as e:
            logger.warning(f"Fout bij controleren veld op leegte of 'INVULLEN': {e}")

    def _adjust_window_size(self):
        """Resize window to fit content, but not more than 80% of screen."""
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
        logger.debug(f"Venstergrootte aangepast naar {final_width}x{final_height}")
    
    def _bind_combobox_selectie(self, entry, veldnaam):
        def controleer_waarde(event=None):
            self._controleer_en_markeer(entry)

            if veldnaam.strip().startswith("Aantal onderhoudsrapportages in contractjaar"):
                waarde = entry.get().strip()
                try:
                    waarde_clean = waarde.split(" ")[0]
                    aantal = int(float(waarde_clean))
                    target_entry = self.entries.get("Meet- en regel onderhoudsrapportage:")
                    if target_entry:
                        nieuwe_waarden = [f"{i} van {aantal}" for i in range(1, aantal + 1)]
                        target_entry['values'] = nieuwe_waarden
                        target_entry.set(nieuwe_waarden[0])
                        self._markeer_invullen_veld(target_entry, actief=False)
                        logger.info(f"Aangepaste opties voor Meet- en regel onderhoudsrapportage: {nieuwe_waarden}")
                except Exception as ex:
                    logger.warning(f"Fout bij bijwerken Meet- en regel onderhoudsrapportage: {ex}")

        entry.bind("<<ComboboxSelected>>", controleer_waarde)
        entry.bind("<FocusOut>", controleer_waarde)
        entry.bind("<KeyRelease>", controleer_waarde)

    def _aantal_rapportages_aangepast(self, widget):
        try:
            aantal = int(widget.get())
            target_entry = self.entries.get("Meet- en regel onderhoudsrapportage:")
            if target_entry:
                nieuwe_waarden = [f"{i} van {aantal}" for i in range(1, aantal + 1)]
                target_entry['values'] = nieuwe_waarden
                target_entry.set("")  # wis eventuele oude waarde
                self._controleer_en_markeer(target_entry)
        except Exception as ex:
            logger.warning(f"Kon Meet- en regel onderhoudsrapportage bijwerken: {ex}")

    def _on_combobox_selected(self, veld, widget):
        self._controleer_en_markeer(widget)
        logger.debug(f"→ veld geselecteerd: {veld} | waarde: {widget.get()}")
        logger.debug(f"Combobox gewijzigd: veld='{veld}', waarde='{widget.get()}'")
        logger.debug(f"ON SELECT: veld='{veld}'")
        logger.debug(f"Alle veldnamen: {list(self.entries.keys())}")

        if veld.strip().startswith("Aantal onderhoudsrapportages in contractjaar"):
            waarde = widget.get().strip()
            try:
                waarde_clean = waarde.split(" ")[0]  # "3 van 4" → "3"
                aantal = int(float(waarde_clean))    # "3.0" → 3
                target_entry = self.entries.get("Meet- en regel onderhoudsrapportage:")
                if target_entry:
                    nieuwe_waarden = [f"{i} van {aantal}" for i in range(1, aantal + 1)]
                    target_entry['values'] = nieuwe_waarden
                    target_entry.set(nieuwe_waarden[0])  # kies eerste optie
                    self._markeer_invullen_veld(target_entry, actief=False)  # ❗ verbergen
                    logger.info(f"Aangepaste opties voor Meet- en regel onderhoudsrapportage: {nieuwe_waarden}")
            except Exception as ex:
                logger.warning(f"Fout bij bijwerken van Meet- en regel onderhoudsrapportage: {ex}")



    def _create_widgets(self):
        """Create all UI widgets and layout."""
        logger.info("UI-widgets aanmaken")
        self.entries = {}

        combobox_velden = {
            "Contractniveau:": ["Basis", "Standaard", "Totaal"],
            "Type gebouwgebruik:": ["Kantoor (gehuurd)", "Kantoor (verhuurd)", "Kantoor eigen gebruik", "School", "Ziekenhuis", "Overig"],
            "Aantal onderhoudsrapportages in contractjaar:": ["1", "2", "3", "4"],
            "Meet- en regel onderhoudsrapportage:": ["1 van 4", "2 van 4", "3 van 4", "4 van 4"],
            "Merk regelinstallatie:": ["Schneider Electric"],
            "Type regelinstallatie:": ["Ecostruxure Building Operation"],
            "Naregelingen op het GBS aangesloten:": ["geen naregelingen aanwezig", "middels BACnet IP", "middels Modbus IP", "middels BACnet MS/tp", "middels Modbus RTU", "middels diverse protocollen"],
            "Type naregelingen:": ["B3 serie", "RP-C serie", "MP-C serie", "6-wegventielen", "diverse"],
            "Hoofd gasmeter uitgelezen op GBS:": ["Nee", "Ja"],
            "Hoofd elektrameter uitgelezen op GBS:": ["Nee", "Ja"],
            "Warmte wordt opgewekt middels:": ["nvt", "ketel(s)", "warmtepomp(en)", "bron(nen)", "WKO (bron en WP)"],
            "Koude wordt opgewekt middels:": ["nvt", "ketel(s)", "warmtepomp(en)", "bron(nen)", "WKO (bron en WP)"],
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
            ],
            "GEGEVENS CONTACTPERSOON TECHNISCHE DIENST": [
                "Contactpersoon TD:", "Telefoonnummer contactpersoon TD:", "Email contactpersoon TD:",
            ],
            "GEGEVENS CONTACTPERSOON CONTRACT (KLANT)": [
                "Contactpersoon contract:", "Telefoonnummer contactpersoon contract:", "Email contactpersoon contract:",
            ],
            "CONTRACTGEGEVENS": [
                "Contractjaar:", "Aantal onderhoudsrapportages in contractjaar:", "Contractniveau:"
            ],
            "INFORMATIE EQUANS SERVICES": [
                "Onderhoud uitgevoerd door:", "Rapportage opgesteld door:", "Contractmanager Services:"
            ],
        }

        secties_rechts = {
            "INFORMATIE ONDERHOUDSBEURT": [
                "Meet- en regel onderhoudsrapportage:", "Datum of periode uitgevoerde onderhoud:"
            ],
            "INFORMATIE REGELINSTALLATIE": [
                "Merk regelinstallatie:", "Type regelinstallatie:", "Versie GBS software:",
                "Aantal centrale regelpanelen aanwezig:", "Naregelingen op het GBS aangesloten:", "Type naregelingen:",
                "Aantal floormanagerpanelen aanwezig:", "Aantal naregelingen aanwezig:", "Aantal ruimtebedieningen aanwezig:"
            ],
            "INFORMATIE KLIMAATINSTALLATIES EN ENERGIE": [
                "Hoofd gasmeter uitgelezen op GBS:", "Hoofd elektrameter uitgelezen op GBS:",
                "Warmte wordt opgewekt middels:", "Koude wordt opgewekt middels:",
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
                        values = combobox_velden[veld]

                        input_container = ttk.Frame(frame)
                        input_container.grid(row=idx, column=1, padx=(0, 10), pady=2, sticky="e")
                        input_container.columnconfigure(0, weight=1)     # Combobox
                        input_container.columnconfigure(1, minsize=24)   # ? knop
                        input_container.columnconfigure(2, minsize=20)   # ❗ waarschuwing

                        entry = ttk.Combobox(input_container, values=values, width=32)
                        entry.grid(row=0, column=0, sticky="ew")

                        help_button = ttk.Button(input_container, text="?", width=3, command=self._toon_contractniveau_popup)
                        help_button.grid(row=0, column=1, padx=(1, 1))

                        warning = ttk.Label(input_container, text="❗", foreground="red")
                        warning.grid(row=0, column=2, sticky="e")
                        warning.grid_remove()

                        self.entry_wrappers[veld] = input_container
                        self.entries[veld] = entry
                        self.entry_wrappers[f"{veld}_warning"] = warning
                        self._bind_combobox_selectie(entry, veld)

                    elif veld in combobox_velden:
                        values = combobox_velden[veld]
                        input_container = ttk.Frame(frame)
                        input_container.grid(row=idx, column=1, padx=(0, 10), pady=2, sticky="e")
                        input_container.columnconfigure(0, weight=1)
                        input_container.columnconfigure(1, minsize=20)

                        entry = ttk.Combobox(input_container, values=values, width=37)
                        entry.grid(row=0, column=0, sticky="ew")

                        warning = ttk.Label(input_container, text="❗", foreground="red")
                        warning.grid(row=0, column=1, sticky="e")
                        warning.grid_remove()

                        self.entry_wrappers[veld] = input_container
                        self.entries[veld] = entry
                        self.entry_wrappers[f"{veld}_warning"] = warning
                        self._bind_combobox_selectie(entry, veld)
                        # 👇 Handmatige binding voor Aantal onderhoudsrapportages
                        if veld == "Aantal onderhoudsrapportages in contractjaar:":
                            logger.debug("✅ Binding toegevoegd aan combobox 'Aantal onderhoudsrapportages in contractjaar:'")
                    else:
                        input_container = ttk.Frame(frame)
                        input_container.grid(row=idx, column=1, padx=(0, 10), pady=2, sticky="e")
                        input_container.columnconfigure(0, weight=1)
                        input_container.columnconfigure(1, minsize=20)

                        entry = ttk.Entry(input_container, width=entry_width)
                        entry.grid(row=0, column=0, sticky="ew")

                        warning = ttk.Label(input_container, text="❗", foreground="red")
                        warning.grid(row=0, column=1, sticky="e")
                        warning.grid_remove()

                        self.entry_wrappers[veld] = input_container
                        self.entries[veld] = entry
                        self.entry_wrappers[f"{veld}_warning"] = warning

        plaats_secties(left_column, secties_links)
        plaats_secties(right_column, secties_rechts)

        actions_frame = ttk.Frame(self)
        actions_frame.pack(pady=12)
        ttk.Button(actions_frame, text="Genereer Rapportage", command=self._generate_report).grid(row=0, column=0, padx=10)
        ttk.Button(actions_frame, text="Exporteer naar XML", command=self._export_to_xml).grid(row=0, column=1, padx=10)
        ttk.Button(actions_frame, text="Importeer vanuit XML", command=self._import_from_xml).grid(row=0, column=2, padx=10)
        
        # for veldnaam, entry in self.entries.items():
        #     if isinstance(entry, ttk.Entry):
        #         entry.bind("<KeyRelease>", lambda e, w=entry: self._controleer_en_markeer(w))
        #     elif isinstance(entry, ttk.Combobox):
        #         def maak_combobox_handler(widget):
        #             return lambda e: self._controleer_en_markeer(widget)

        #         entry.bind("<KeyRelease>", maak_combobox_handler(entry))
        #         entry.bind("<FocusOut>", maak_combobox_handler(entry))
        #         entry.bind("<<ComboboxSelected>>", maak_combobox_handler(entry))

        
        self._controleer_alle_velden() 
        logger.info("UI-widgets aangemaakt")



        # Algemene binding voor alle comboboxen (ook deze komt na handmatige binding)
        for veldnaam, entry in self.entries.items():
            if isinstance(entry, ttk.Entry):
                entry.bind("<KeyRelease>", lambda e, w=entry: self._controleer_en_markeer(w))
            elif isinstance(entry, ttk.Combobox):
                self._bind_combobox_selectie(entry, veldnaam)


    def _toon_contractniveau_popup(self):
        """Show popup with contract level explanation table."""
        logger.info("Popup contractniveau toelichting geopend")
        popup = tk.Toplevel(self)
        popup.title("Toelichting Contractniveau")

        style = ttk.Style(popup)
        style.configure("Bold.TREEVIEW", font=("TkDefaultFont", 10, "bold"))

        frame = ttk.Frame(popup)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(frame, columns=CONTRACTNIVEAU_KOLOMMEN, show="headings", height=20)
        for col in CONTRACTNIVEAU_KOLOMMEN:
            tree.heading(col, text=col)
            if col == "Omschrijving":
                tree.column(col, width=450, anchor="w")
            else:
                tree.column(col, width=70, anchor="center")

        tree.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))

        for row in CONTRACTNIVEAU_TABEL:
            if all(cell == "" for cell in row[1:]):
                tree.insert("", "end", values=row, tags=("bold",))
            else:
                tree.insert("", "end", values=row)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        ttk.Button(popup, text="Sluiten", command=popup.destroy).pack(pady=10)
        popup.update_idletasks()
        w = min(popup.winfo_reqwidth(), 1000)
        popup.minsize(width=w, height=900)
        popup.geometry(f"{w}x600")
        logger.debug("Popup contractniveau toelichting getoond")

    def _select_template(self):
        """Let user select a template file and load Excel data if needed."""
        logger.info("Selecteer Template dialoog geopend")
        file_path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Templates", "*.xlsm")]
        )
        if file_path:
            logger.info(f"Template geselecteerd: {file_path}")
            self.template_path_var.set(file_path)
            self._adjust_window_size()
            self.lift()
            if file_path.lower().endswith(".xlsm"):
                logger.info("Excel template geselecteerd, gegevens worden geladen")
                loading_window = self._toon_loading_venster("Template wordt geladen...")

                def laad_data_en_sluit():
                    try:
                        self._laad_templategegevens(file_path)
                    finally:
                        loading_window.destroy()

                threading.Thread(target=laad_data_en_sluit, daemon=True).start()
        else:
            logger.info("Geen template geselecteerd")

    def _laad_templategegevens(self, excel_pad):
        logger.info(f"Probeer gegevens te laden uit Excelbestand: {excel_pad}")
        if not os.path.exists(excel_pad):
            logger.error(f"Bestand niet gevonden: {excel_pad}")
            messagebox.showerror("Fout", f"Bestand niet gevonden: {excel_pad}")
            return

        # Standaarddata uit Excel uitlezen (via xlwings)
        try:
            app = xw.App(visible=False)
            wb = app.books.open(excel_pad)
            if "Gegevens" not in [s.name for s in wb.sheets]:
                wb.close()
                app.quit()
                logger.error("Tabblad 'Gegevens' niet gevonden in Excelbestand.")
                messagebox.showerror("Fout", "Tabblad 'Gegevens' niet gevonden in Excelbestand.")
                return
            sheet = wb.sheets["Gegevens"]
            kolom_a = sheet.range("A1:A200").value
            kolom_b = sheet.range("B1:B200").value
            wb.close()
            app.quit()
            logger.info("Excelgegevens succesvol uitgelezen via xlwings")
        except Exception as e:
            try:
                app.quit()
            except Exception:
                pass
            logger.error(f"Fout bij uitlezen Excelbestand: {str(e)}")
            messagebox.showerror("Fout", f"Fout bij uitlezen Excelbestand: {str(e)}")
            return

        if not isinstance(kolom_a, list):
            kolom_a = [kolom_a]
        if not isinstance(kolom_b, list):
            kolom_b = [kolom_b]

        data_dict = {}
        for sleutel, waarde in zip(kolom_a, kolom_b):
            if isinstance(sleutel, str) and sleutel.strip():
                sleutel_clean = sleutel.strip().rstrip(":")
                waarde_str = str(waarde).strip() if waarde else ""
                if waarde_str:
                    data_dict[sleutel_clean] = waarde_str

        # Formulier dropdowns uitlezen en mappen naar labels
        dropdown_mapping = {
            "Drop Down 1": "Contractniveau:",
            "Drop Down 2": "Aantal onderhoudsrapportages in contractjaar:",
            "Drop Down 3": "Meet- en regel onderhoudsrapportage:",
            "Drop Down 7": "Hoofd gasmeter uitgelezen op GBS:",
            "Drop Down 9": "Hoofd elektrameter uitgelezen op GBS:",
            "Drop Down 10": "Naregelingen op het GBS aangesloten:",
            "Drop Down 14": "Type naregelingen:",
            "Drop Down 15": "Warmte wordt opgewekt middels:",
            "Drop Down 16": "Koude wordt opgewekt middels:",
            "Drop Down 12": "Merk regelinstallatie:",
            "Drop Down 13": "Type regelinstallatie:",
            "Drop Down 20": "Type gebouwgebruik:",
        }

        try:
            formulierdata = lees_formulierbesturingselementen(excel_pad)
            logger.info(f"Form Controls uitgelezen: {formulierdata}")

            for dropdown, waarde in formulierdata.items():
                if dropdown not in dropdown_mapping:
                    foutmelding = f"Dropdown '{dropdown}' is niet bekend in dropdown_mapping."
                    logger.error(foutmelding)
                    messagebox.showerror("Fout in Excel-besturingselement", foutmelding)
                    raise ValueError(foutmelding)

                veld = dropdown_mapping[dropdown]
                if veld not in self.entries:
                    foutmelding = f"Dropdown '{dropdown}' verwijst naar onbekend veld '{veld}' in de GUI."
                    logger.error(foutmelding)
                    messagebox.showerror("Fout in veldmapping", foutmelding)
                    raise ValueError(foutmelding)

                data_dict[veld] = waarde

        except Exception as e:
            logger.warning(f"Formulierbesturingselementen niet geladen: {e}")

        # Mapping naar GUI labels (uit xlwings-gegevens)
        label_mapping = {
            "Klantnaam": "Klantnaam/Gebouwnaam:",
            "Locatie": "Locatie:",
            "Adres": "Adres:",
            "Contactpersoon TD, indien van toepassing": "Contactpersoon TD:",
            "Telefoonnummer contactpersoon TD": "Telefoonnummer contactpersoon TD:",
            "Email contactpersoon TD": "Email contactpersoon TD:",
            "Contactpersoon contract": "Contactpersoon contract:",
            "Telefoonnummer contactpersoon contract": "Telefoonnummer contactpersoon contract:",
            "Email contactpersoon contract": "Email contactpersoon contract:",
            "Contractjaar": "Contractjaar:",
            "Datum of periode uitgevoerde onderhoud": "Datum of periode uitgevoerde onderhoud:",
            "Onderhoud uitgevoerd door": "Onderhoud uitgevoerd door:",
            "Rapportage opgesteld door": "Rapportage opgesteld door:",
            "Contractmanager Services": "Contractmanager Services:",
            "Versie GBS software, indien van toepassing": "Versie GBS software:",
            "Aantal centrale regelpanelen aanwezig": "Aantal centrale regelpanelen aanwezig:",
            "Aantal floormanagerpanelen aanwezig": "Aantal floormanagerpanelen aanwezig:",
            "Aantal naregelingen aanwezig": "Aantal naregelingen aanwezig:",
            "Aantal ruimtebedieningen aanwezig": "Aantal ruimtebedieningen aanwezig:",
            "Aantal aanwezige luchtbehandelingskasten": "Aantal aanwezige luchtbehandelingskasten:"
        }
        # Vul velden in de GUI op basis van data_dict
        ingevulde = 0
        for label, entry in self.entries.items():
            gui_label = label.strip()
            excel_key = next((k for k, v in label_mapping.items() if v.strip() == gui_label), gui_label)

            if excel_key in data_dict:
                try:
                    waarde = data_dict[excel_key]

        # Stap 1: verwerk aantal rapportages
                    if label == "Aantal onderhoudsrapportages in contractjaar:":
                        if waarde in ["1.0", "2.0", "3.0", "4.0"]:
                            waarde = str(int(float(waarde)))
                        aantal_rapportages_totaal = int(waarde)

                    # Stap 2: verwerk meet- en regel rapportage
                    elif label == "Meet- en regel onderhoudsrapportage:":
                        if aantal_rapportages_totaal is not None:
                            opties = [f"{i} van {aantal_rapportages_totaal}" for i in range(1, aantal_rapportages_totaal + 1)]
                            entry['values'] = opties

                            if waarde not in opties:
                                # Probeer alleen nummer uit tekst te halen (bijv. "2 van 3" → "2")
                                nummer = waarde.split("van")[0].strip()
                                if nummer.isdigit() and 1 <= int(nummer) <= aantal_rapportages_totaal:
                                    waarde = f"{nummer} van {aantal_rapportages_totaal}"
                                else:
                                    logger.warning(f"Ongeldige waarde '{waarde}' voor veld '{label}' bij totaal={aantal_rapportages_totaal}")
                                    continue

                    # Veld invullen
                    if isinstance(entry, ttk.Combobox):
                        bestaande_waarden = list(entry['values'])
                        matches = [opt for opt in bestaande_waarden if opt.lower() == str(waarde).lower()]
                        if matches:
                            waarde = matches[0]
                        elif waarde not in bestaande_waarden:
                            foutmelding = f"Waarde '{waarde}' is ongeldig voor veld '{label}'. Toegestane waarden: {bestaande_waarden}"
                            logger.error(foutmelding)
                            messagebox.showerror("Ongeldige waarde in Excel", foutmelding)
                            continue

                        entry.set(waarde)

                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, waarde)

                    waarde_upper = str(waarde).strip().upper()
                    actief = not waarde_upper or "INVULLEN" in waarde_upper
                    self._markeer_invullen_veld(entry, actief=actief)
                    ingevulde += 1

                except Exception as ex:
                    logger.warning(f"Kon veld '{label}' niet invullen: {ex}")
            else:
                logger.debug(f"Geen waarde gevonden voor veld: {label}")

        logger.info(f"{ingevulde} velden automatisch ingevuld vanuit Excel en formulierbesturingselementen.")

    def _generate_report(self):
        """Print all current values to the console (placeholder for actual report generation)."""
        logger.info("Genereer Rapportage gestart")
        print("Rapportage wordt gegenereerd met de volgende gegevens:")
        print(f"Wachtwoord: {self.password_entry.get()}")
        print(f"Templatepad: {self.template_path_var.get()}")
        for label, entry in self.entries.items():
            print(f"{label} {entry.get()}")
        logger.info("Rapportage gegenereerd (console output)")

    def _export_to_xml(self):
        """Export all current values to an XML file."""
        logger.info("Exporteren naar XML gestart")
        data = {
            "Wachtwoord": self.password_entry.get(),
            "Templatepad": self.template_path_var.get(),
            **{label: entry.get() for label, entry in self.entries.items()}
        }
        file_path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml")]
        )
        if file_path:
            root = ET.Element("Rapportage")
            for key, value in data.items():
                child = ET.SubElement(root, "Veld", naam=key)
                child.text = value
            tree = ET.ElementTree(root)
            tree.write(file_path, encoding="utf-8", xml_declaration=True)
            logger.info(f"Gegevens succesvol geëxporteerd naar XML: {file_path}")
            messagebox.showinfo("Succes", "Gegevens succesvol geëxporteerd naar XML.")
        else:
            logger.info("Exporteren naar XML geannuleerd door gebruiker")

    def _import_from_xml(self):
        """Import values from an XML file and fill in the fields."""
        logger.info("Importeren vanuit XML gestart")
        file_path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("XML files", "*.xml")]
        )
        if file_path:
            logger.info(f"XML-bestand geselecteerd voor import: {file_path}")
            tree = ET.parse(file_path)
            root = tree.getroot()
            ingevulde = 0
            for veld in root.findall("Veld"):
                naam = veld.attrib.get("naam")
                waarde = veld.text or ""
                if naam == "Wachtwoord":
                    self.password_entry.delete(0, tk.END)
                    self.password_entry.insert(0, waarde)
                    ingevulde += 1
                elif naam == "Templatepad":
                    self.template_path_var.set(waarde)
                    ingevulde += 1
                elif naam in self.entries:
                    entry = self.entries[naam]
                    
                    if isinstance(entry, ttk.Combobox):
                        bestaande_waarden = list(entry['values'])
                        if waarde not in bestaande_waarden:
                            matches = [opt for opt in bestaande_waarden if opt.lower() == waarde.lower()]
                            if matches:
                                waarde = matches[0]
                            else:
                                foutmelding = f"Waarde '{waarde}' is ongeldig voor veld '{naam}' (XML). Toegestane waarden: {bestaande_waarden}"
                                logger.error(foutmelding)
                                messagebox.showerror("Ongeldige waarde bij XML-import", foutmelding)
                                continue
                        entry.set(waarde)
                        waarde_upper = waarde.strip().upper()
                        actief = not waarde_upper or "INVULLEN" in waarde_upper
                        self._markeer_invullen_veld(entry, actief=actief)

                        # 👇 Dynamische aanpassing voor rapportage-combobox
                        if naam == "Aantal onderhoudsrapportages in contractjaar:":
                            self._aantal_rapportages_aangepast(entry)

                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, waarde)
                        waarde_upper = waarde.strip().upper()
                        actief = not waarde_upper or "INVULLEN" in waarde_upper
                        self._markeer_invullen_veld(entry, actief=actief)
                    if isinstance(entry, ttk.Combobox):
                        bestaande_waarden = list(entry['values'])
                        if waarde not in bestaande_waarden:
                            matches = [opt for opt in bestaande_waarden if opt.lower() == waarde.lower()]
                            if matches:
                                waarde = matches[0]
                            else:
                                foutmelding = f"Waarde '{waarde}' is ongeldig voor veld '{naam}' (XML). Toegestane waarden: {bestaande_waarden}"
                                logger.error(foutmelding)
                                messagebox.showerror("Ongeldige waarde bij XML-import", foutmelding)
                                continue
                        entry.set(waarde)
                        waarde_upper = waarde.strip().upper()
                        actief = not waarde_upper or "INVULLEN" in waarde_upper
                        self._markeer_invullen_veld(entry, actief=actief)
                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, waarde)
                        waarde_upper = waarde.strip().upper()
                        actief = not waarde_upper or "INVULLEN" in waarde_upper
                        self._markeer_invullen_veld(entry, actief=actief)
                    ingevulde += 1
            self._adjust_window_size()
            logger.info(f"Gegevens succesvol geïmporteerd uit XML ({ingevulde} velden ingevuld)")
            messagebox.showinfo("Succes", "Gegevens succesvol geïmporteerd uit XML.")
        else:
            logger.info("Importeren vanuit XML geannuleerd door gebruiker")

    def _toon_loading_venster(self, bericht="Template wordt geladen..."):
        loading_win = tk.Toplevel(self)
        loading_win.title("Even geduld...")
        loading_win.geometry("350x120")
        loading_win.resizable(False, False)
        loading_win.transient(self)
        loading_win.configure(padx=20, pady=20)

        ttk.Label(loading_win, text=bericht, font=("Segoe UI", 11)).pack(pady=(0, 15))

        pb = ttk.Progressbar(loading_win, mode='indeterminate', length=280)
        pb.pack()
        pb.start(10)

        # Centraal op het scherm positioneren
        loading_win.update_idletasks()
        screen_width = loading_win.winfo_screenwidth()
        screen_height = loading_win.winfo_screenheight()
        window_width = loading_win.winfo_width()
        window_height = loading_win.winfo_height()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        loading_win.geometry(f"{window_width}x{window_height}+{x}+{y}")

        return loading_win