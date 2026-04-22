"""
Gestione Contratti v7.2 — Comune di Sanremo, Servizio Demanio Marittimo
Interfaccia coerente: tutte le maschere con Modifica/Salva/Elimina,
campi in sola lettura di default, calcoli automatici.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import os, sys, subprocess
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import database as db
from genera_documenti import genera_determina, genera_verbale, fmtE, fmtD
import ai_helper


def format_importo(v):
    return fmtE(v).replace("€ ", "") if v else "0,00"


# ═════════════════════════════════════════════════════════   ═════
# FORM BASE — tutte le maschere derivano da qui
# ═══════════════════════════════════════════════════════════════

class FormBase(tk.Toplevel):
    """
    Maschera base con comportamento standard:
    - Campi in sola lettura all'apertura (in edit mode)
    - Pulsanti Modifica / Salva / Elimina / Annulla sempre presenti
    - Calcoli automatici su campi numerici
    - Conferma prima di chiudere se ci sono modifiche non salvate
    - Validazione campi obbligatori
    """

    TITLE = "Form"
    WIDTH = 650
    HEIGHT = 550

    def __init__(self, master, data=None, on_save=None, **kw):
        super().__init__(master)
        self.data = data
        self.is_edit = data is not None and data.get("id") is not None
        self.on_save_callback = on_save
        self._modified = False
        self._editing = not self.is_edit  # nuovo record o bozza AI: subito editabile
        self.fields = {}
        self._calc_bindings = []

        self.title(self.TITLE)
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.transient(master)
        self.grab_set()

        # Contenitore scrollabile
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(outer)
        self.scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.main = ttk.Frame(self.canvas, padding=14)
        self.main.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.main, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        # Mousewheel scroll
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Barra pulsanti in basso (fissa, fuori dallo scroll)
        self.btn_bar = ttk.Frame(self)
        self.btn_bar.pack(fill=tk.X, padx=10, pady=8)

        self.btn_modifica = ttk.Button(self.btn_bar, text="✏ Modifica", command=self._start_edit)
        self.btn_salva = ttk.Button(self.btn_bar, text="💾 Salva", command=self._do_save)
        self.btn_elimina = ttk.Button(self.btn_bar, text="🗑 Elimina", command=self._do_delete)
        self.btn_annulla = ttk.Button(self.btn_bar, text="Annulla", command=self._do_close)

        # Layout pulsanti
        self.btn_salva.pack(side=tk.LEFT, padx=5)
        self.btn_annulla.pack(side=tk.LEFT, padx=5)
        if self.is_edit:
            self.btn_modifica.pack(side=tk.LEFT, padx=15)
            self.btn_elimina.pack(side=tk.RIGHT, padx=5)

        # Hook per contenuto specifico
        self.build_form()

        # Stato iniziale
        self._apply_lock()

        # Intercetta chiusura finestra
        self.protocol("WM_DELETE_WINDOW", self._do_close)

    # ── API per le sottoclassi ──

    def add_field(self, key, label, width=40, row=None, values=None, readonly_always=False, required=False):
        """Aggiunge un campo testo o combo. Ritorna la StringVar."""
        if row is None:
            row = len(self.fields)
        ttk.Label(self.main, text=("* " if required else "") + label).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 5))
        var = tk.StringVar(value=str((self.data or {}).get(key, "") or ""))
        if values:
            w = ttk.Combobox(self.main, textvariable=var, values=values, width=width, state="readonly" if readonly_always else "normal")
        else:
            w = ttk.Entry(self.main, textvariable=var, width=width)
        w.grid(row=row, column=1, sticky="w", padx=5, pady=3)
        var.trace_add("write", lambda *a: self._mark_modified())
        self.fields[key] = {"var": var, "widget": w, "readonly_always": readonly_always, "required": required}
        return var

    def add_separator(self, text, row=None):
        if row is None:
            row = len(self.fields) + self._sep_count if hasattr(self, "_sep_count") else len(self.fields)
        self._sep_count = getattr(self, "_sep_count", 0) + 1
        ttk.Label(self.main, text=f"── {text} ──", style="SubHeader.TLabel").grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(12, 4))

    def add_label(self, text, row=None, style="Info.TLabel"):
        if row is None:
            row = len(self.fields) + getattr(self, "_sep_count", 0)
        ttk.Label(self.main, text=text, style=style).grid(row=row, column=0, columnspan=2, sticky="w", pady=2)

    def add_calc(self, source_keys, target_key, calc_fn):
        """Collega campi sorgente a un calcolo automatico."""
        def do_calc(*args):
            try:
                vals = {}
                for k in source_keys:
                    v = self.fields[k]["var"].get().replace(",", ".")
                    vals[k] = float(v) if v else 0
                result = calc_fn(vals)
                self.fields[target_key]["var"].set(str(round(result, 2)))
            except (ValueError, KeyError):
                pass
        for k in source_keys:
            if k in self.fields:
                self.fields[k]["var"].trace_add("write", do_calc)

    def get_values(self):
        """Raccoglie tutti i valori come dict."""
        return {k: f["var"].get().strip() for k, f in self.fields.items()}

    def build_form(self):
        """Override nelle sottoclassi per costruire il form."""
        pass

    def validate(self):
        """Override per validazione custom. Ritorna (bool, messaggio)."""
        for key, f in self.fields.items():
            if f.get("required") and not f["var"].get().strip():
                label = key.replace("_", " ").title()
                return False, f"Il campo '{label}' è obbligatorio."
        return True, ""

    def save(self, values):
        """Override per logica di salvataggio. Ritorna True se OK."""
        return True

    def delete(self):
        """Override per logica di eliminazione. Ritorna True se OK."""
        return True

    def get_title_detail(self):
        """Override per dettaglio nel titolo (es. nome record)."""
        return ""

    # ── Logica interna ──

    def _mark_modified(self):
        self._modified = True
        t = self.TITLE
        d = self.get_title_detail()
        if d:
            t += f" — {d}"
        if self._modified and self._editing:
            t += " *"
        self.title(t)

    def _start_edit(self):
        self._editing = True
        self._apply_lock()
        self.btn_modifica.config(state="disabled")
        # Focus sul primo campo editabile
        for f in self.fields.values():
            w = f["widget"]
            if str(w.cget("state")) != "disabled":
                w.focus_set()
                break

    def _apply_lock(self):
        state = "normal" if self._editing else "disabled"
        for f in self.fields.values():
            if f.get("readonly_always"):
                f["widget"].config(state="disabled")
            else:
                try:
                    f["widget"].config(state=state)
                except tk.TclError:
                    pass
        self.btn_salva.config(state="normal" if self._editing else "disabled")

    def _do_save(self):
        ok, msg = self.validate()
        if not ok:
            messagebox.showwarning("Validazione", msg)
            return
        values = self.get_values()
        try:
            if self.save(values):
                self._modified = False
                if self.on_save_callback:
                    self.on_save_callback()
                self.destroy()
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def _do_delete(self):
        if not self.is_edit:
            return
        if messagebox.askyesno("Conferma eliminazione",
                               "Eliminare questo record e tutti i dati collegati?\nL'operazione non è reversibile."):
            try:
                if self.delete():
                    if self.on_save_callback:
                        self.on_save_callback()
                    self.destroy()
            except Exception as e:
                messagebox.showerror("Errore", str(e))

    def _do_close(self):
        if self._modified and self._editing:
            r = messagebox.askyesnocancel("Modifiche non salvate",
                                         "Ci sono modifiche non salvate. Salvare prima di chiudere?")
            if r is True:
                self._do_save()
                return
            elif r is None:
                return  # Annulla la chiusura
        self.destroy()


# ═══════════════════════════════════════════════════════════════
# FORM: OPERATORE ECONOMICO
# ═══════════════════════════════════════════════════════════════

class FormOperatore(FormBase):
    TITLE = "Operatore Economico"
    WIDTH = 580
    HEIGHT = 520

    def build_form(self):
        r = 0
        self.add_field("ragione_sociale", "Ragione Sociale", 50, row=r, required=True); r += 1
        self.add_field("tipo_soggetto", "Tipo Soggetto", 25, row=r, values=["Impresa","Professionista","Società","Ente","Persona fisica"]); r += 1
        self.add_field("codice_fiscale", "Codice Fiscale", 20, row=r); r += 1
        self.add_field("partita_iva", "P.IVA", 15, row=r); r += 1
        self.add_field("legale_rappresentante", "Legale Rappresentante", 40, row=r); r += 1
        self.add_separator("Sede", row=r); r += 1
        self.add_field("indirizzo", "Indirizzo", 50, row=r); r += 1
        self.add_field("cap", "CAP", 8, row=r); r += 1
        self.add_field("citta", "Città", 30, row=r); r += 1
        self.add_field("provincia", "Prov.", 5, row=r); r += 1
        self.add_separator("Contatti", row=r); r += 1
        self.add_field("telefono", "Telefono", 18, row=r); r += 1
        self.add_field("email", "Email", 40, row=r); r += 1
        self.add_field("pec", "PEC", 40, row=r); r += 1
        self.add_field("iban_dedicato", "IBAN Dedicato (L.136/2010)", 35, row=r); r += 1

    def save(self, v):
        v["fid_excel"] = (self.data or {}).get("fid_excel", "") or ""
        if self.is_edit:
            db.aggiorna_operatore(self.data["id"], v)
        else:
            db.inserisci_operatore(v)
        return True

    def delete(self):
        db.aggiorna_operatore(self.data["id"], {**dict(self.data), "attivo": 0})
        return True

    def get_title_detail(self):
        return self.fields.get("ragione_sociale", {}).get("var", tk.StringVar()).get()[:40]


# ═══════════════════════════════════════════════════════════════
# FORM: FATTURA
# ═══════════════════════════════════════════════════════════════

class FormFattura(FormBase):
    TITLE = "Fattura"
    WIDTH = 650
    HEIGHT = 580

    def build_form(self):
        r = 0
        # Affidamento (combobox)
        al = db.cerca_affidamenti()
        self._aff_map = {"": None}
        self._aff_map.update({f"{a['id']} - {a['oggetto'][:40]}": a["id"] for a in al})
        self.add_field("_aff", "Affidamento", 52, row=r, values=list(self._aff_map.keys())); r += 1
        if self.data and self.data.get("id_affidamento"):
            for k, v in self._aff_map.items():
                if v == self.data["id_affidamento"]:
                    self.fields["_aff"]["var"].set(k); break

        # Operatore (opzionale se c'è affidamento, ma utile per filtri)
        ops = db.cerca_operatori()
        self._op_map = {"": None}
        self._op_map.update({f"{o['id']} - {o['ragione_sociale']}": o["id"] for o in ops})
        self.add_field("_op", "Operatore Economico", 52, row=r, values=list(self._op_map.keys())); r += 1
        if self.data and self.data.get("id_operatore"):
            for k, v in self._op_map.items():
                if v == self.data["id_operatore"]:
                    self.fields["_op"]["var"].set(k); break

        self.add_separator("Dati Fattura", row=r); r += 1
        self.add_field("numero_fattura", "N. Fattura", 20, row=r, required=True); r += 1
        self.add_field("data_fattura", "Data (AAAA-MM-GG)", 14, row=r); r += 1

        self.add_separator("Importi (con calcolo automatico IVA)", row=r); r += 1
        self.add_field("importo_netto", "Importo Netto €", 14, row=r); r += 1
        self.add_field("aliquota_iva", "Aliquota IVA %", 7, row=r); r += 1
        if not self.is_edit:
            self.fields["aliquota_iva"]["var"].set("22")
        self.add_field("importo_iva", "IVA € (auto)", 14, row=r); r += 1
        self.add_field("importo_totale", "Totale € (auto)", 14, row=r); r += 1

        # Calcoli automatici IVA e totale
        self.add_calc(["importo_netto", "aliquota_iva"], "importo_iva",
                      lambda v: v["importo_netto"] * v["aliquota_iva"] / 100)
        self.add_calc(["importo_netto", "importo_iva"], "importo_totale",
                      lambda v: v["importo_netto"] + v["importo_iva"])

        self.add_separator("Protocollo e Stato", row=r); r += 1
        self.add_field("protocollo_pec", "Prot. PEC", 14, row=r); r += 1
        self.add_field("data_protocollo_pec", "Data Prot. PEC", 14, row=r); r += 1
        self.add_field("tipo_liquidazione", "Tipo Liquidazione", 14, row=r, values=["Saldo","Acconto"]); r += 1
        self.add_field("stato_pagamento", "Stato Pagamento", 14, row=r, values=["Da liquidare","Liquidato","Pagato"]); r += 1
        if not self.is_edit:
            self.fields["stato_pagamento"]["var"].set("Da liquidare")
            self.fields["tipo_liquidazione"]["var"].set("Saldo")

        # Residuo affidamento (solo edit)
        if self.is_edit and self.data.get("id_affidamento"):
            imp = db.get_importi_affidamento(self.data["id_affidamento"])
            color = "#e74c3c" if imp["residuo"] < 0 else "#1a5276"
            self.add_separator("Situazione Affidamento", row=r); r += 1
            ttk.Label(self.main, text=f"Residuo: {fmtE(imp['residuo'])}",
                      font=("Segoe UI", 10, "bold"), foreground=color).grid(row=r, column=0, columnspan=2, sticky="w", pady=3)

    def save(self, v):
        d = {"id_affidamento": self._aff_map.get(v.pop("_aff", ""))}
        d["id_operatore"] = self._op_map.get(v.pop("_op", ""))
        if d["id_affidamento"] and not d["id_operatore"]:
            aff = db.get_affidamento(d["id_affidamento"])
            d["id_operatore"] = aff["id_operatore"] if aff else None
        for k in ["numero_fattura","data_fattura","protocollo_pec","data_protocollo_pec","tipo_liquidazione","stato_pagamento"]:
            d[k] = v.get(k, "")
        d["stato"] = "Registrata"
        for k in ["importo_netto","aliquota_iva","importo_iva","importo_totale"]:
            try:
                d[k] = float(v.get(k, "0").replace(",", ".") or 0)
            except ValueError:
                d[k] = 0
        if self.is_edit:
            db.aggiorna_fattura(self.data["id"], d)
        else:
            db.inserisci_fattura(d)
        return True

    def delete(self):
        db.elimina_fattura(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: DETERMINA
# ═══════════════════════════════════════════════════════════════

class FormDetermina(FormBase):
    TITLE = "Determina"
    WIDTH = 750
    HEIGHT = 650

    def build_form(self):
        r = 0
        self.add_field("tipo_determina", "Tipo Determina", 35, row=r, values=db.TIPI_DETERMINA, required=True); r += 1

        # N. / Anno / Data su una riga
        self.add_separator("Numerazione", row=r); r += 1
        self.add_field("numero", "N. Determina", 10, row=r); r += 1
        self.add_field("anno", "Anno", 6, row=r); r += 1
        self.add_field("data_determina", "Data (AAAA-MM-GG)", 14, row=r); r += 1
        self.add_field("oggetto", "Oggetto", 55, row=r); r += 1

        # Collegamenti
        self.add_separator("Collegamenti", row=r); r += 1
        al = db.cerca_affidamenti()
        self._aff_map = {"": None}
        self._aff_map.update({f"{a['id']} - {a['oggetto'][:40]}": a["id"] for a in al})
        self.add_field("_aff", "Affidamento", 52, row=r, values=list(self._aff_map.keys())); r += 1

        fl = db.cerca_fatture()
        self._fat_map = {"": None}
        self._fat_map.update({f"{f['id']} - Fatt.{f['numero_fattura']} {f['operatore_nome'] or ''}": f["id"] for f in fl})
        self.add_field("_fat", "Fattura (per liquidazione)", 52, row=r, values=list(self._fat_map.keys())); r += 1

        dl = db.cerca_determine()
        self._pad_map = {"": None}
        self._pad_map.update({f"{d['id']} - {d['tipo_determina']} n.{d['numero'] or '?'}": d["id"] for d in dl})
        self.add_field("_padre", "Richiama Determina", 52, row=r, values=list(self._pad_map.keys())); r += 1

        # Economici
        self.add_separator("Dati Economici", row=r); r += 1
        self.add_field("importo", "Importo €", 14, row=r); r += 1
        self.add_field("capitolo_bilancio", "Capitolo", 15, row=r); r += 1
        self.add_field("impegno_spesa", "Impegno di Spesa", 15, row=r); r += 1
        self.add_field("esercizio", "Esercizio", 6, row=r); r += 1
        self.add_field("rif_normativo", "Rif. Normativo", 50, row=r); r += 1
        
        self.add_separator("Soggetti", row=r); r += 1
        
        pers = db.get_personale(attivi_only=True)
        # Convertiamo le righe in dict per evitare AttributeError su sqlite3.Row.get
        self._pers_map = {p["nome_cognome"]: dict(p) for p in pers}
        nomi = sorted(list(self._pers_map.keys()))

        def _on_pers_change(key_nome, key_qual):
            nome = self.fields[key_nome]["var"].get()
            p = self._pers_map.get(nome)
            if p and p.get("qualifica"):
                self.fields[key_qual]["var"].set(p.get("qualifica"))

        v_rup = self.add_field("rup", "RUP", 40, row=r, values=nomi); r += 1
        v_rup.trace_add("write", lambda *a: _on_pers_change("rup", "qualifica_rup"))
        self.add_field("qualifica_rup", "Qualifica RUP", 50, row=r); r += 1
        
        v_dir = self.add_field("dirigente", "Dirigente", 40, row=r, values=nomi); r += 1
        v_dir.trace_add("write", lambda *a: _on_pers_change("dirigente", "qualifica_dirigente"))
        self.add_field("qualifica_dirigente", "Qualifica Dirigente", 50, row=r); r += 1
        
        v_dl = self.add_field("direttore_lavori", "Direttore Lavori", 40, row=r, values=nomi); r += 1
        v_dl.trace_add("write", lambda *a: _on_pers_change("direttore_lavori", "qualifica_dl"))
        self.add_field("qualifica_dl", "Qualifica DL", 50, row=r); r += 1
        
        # Collaboratori: campo libero + combobox per aggiunta multipla
        self.add_field("collaboratori", "Collaboratori", 60, row=r); r += 1

        # Combobox e pulsante per aggiungere collaboratori multipli (come in FormAffidamento)
        v_add_coll = tk.StringVar()
        cb_coll = ttk.Combobox(self.main, textvariable=v_add_coll, values=nomi, width=40, state="readonly")
        cb_coll.grid(row=r, column=1, sticky="w", padx=5)
        def _add_coll_name(event=None):
            name = v_add_coll.get().strip()
            if not name:
                return
            current = self.fields["collaboratori"]["var"].get().strip()
            # Normalizza lista separata da virgola, evitando duplicati
            existing = [n.strip() for n in current.split(",") if n.strip()]
            if name not in existing:
                existing.append(name)
                self.fields["collaboratori"]["var"].set(", ".join(existing))
            v_add_coll.set("")
        cb_coll.bind("<<ComboboxSelected>>", _add_coll_name)
        ttk.Button(self.main, text="➕ Aggiungi coll.", command=_add_coll_name).grid(row=r, column=0, sticky="e", padx=(0,5))
        r += 1

        self.add_field("stato_iter", "Stato Iter", 14, row=r, values=db.STATI_ITER_DETERMINA); r += 1
        if not self.is_edit:
            self.fields["stato_iter"]["var"].set("Bozza")

        self.add_separator("AI Guidance", row=r); r += 1
        self.add_field("indicazioni_ai", "Istruzioni per l'AI", 60, row=r); r += 1

        # Traces per sincronizzazione automatica (aggiunti alla fine per sicurezza campi)
        self.fields["_aff"]["var"].trace_add("write", lambda *a: self._sync_from_aff())
        self.fields["_padre"]["var"].trace_add("write", lambda *a: self._sync_from_padre())

        # Pre-seleziona se edit o bozza AI (triggers traces)
        if self.data:
            for key, fk, mp in [("_aff","id_affidamento",self._aff_map),("_fat","id_fattura",self._fat_map),("_padre","id_determina_padre",self._pad_map)]:
                if self.data.get(fk):
                    for k, v in mp.items():
                        if v == self.data[fk]:
                            self.fields[key]["var"].set(k); break
        
        # Forza sincronizzazione iniziale se c'è un affidamento
        if self.fields["_aff"]["var"].get():
            self._sync_from_aff()

        ttk.Button(self.main, text="🔄 Sincronizza dati da Affidamento/Padre", 
                   command=lambda: [self._sync_from_aff(force=True), self._sync_from_padre(force=True)]).grid(row=r, column=1, sticky="e", pady=10); r += 1

    def _sync_from_aff(self, force=False):
        """Pulla i dati dall'affidamento se non già presenti."""
        try:
            val = self.fields["_aff"]["var"].get()
            if not val: return
            aid = self._aff_map.get(val)
            if not aid: return
            
            aff = db.get_affidamento(aid)
            if not aff: return

            # Assicuriamoci di lavorare su dict per usare .get senza errori
            aff = dict(aff)

            # Mappa campi: [campo_determina, campo_affidamento]
            mappa = [
                ("oggetto", "oggetto"),
                ("rif_normativo", "rif_normativo"),
                ("esercizio", "esercizio"),
                ("rup", "rup"),
                ("qualifica_rup", "qualifica_rup"),
                ("dirigente", "dirigente"),
                ("qualifica_dirigente", "qualifica_dirigente"),
                ("direttore_lavori", "direttore_lavori"),
                ("qualifica_dl", "qualifica_dl"),
                ("collaboratori", "collaboratori"),
            ]
            
            for det_k, aff_k in mappa:
                if det_k in self.fields:
                    aff_val = str(aff.get(aff_k) or "")
                    if force or not self.fields[det_k]["var"].get().strip():
                        if aff_val:
                            self.fields[det_k]["var"].set(aff_val)
            
            # Se non c'è un importo manuale, usa quello dell'affidamento
            aff_imp = aff.get("importo_affidato") or 0
            if force or not self.fields["importo"]["var"].get().strip() or self.fields["importo"]["var"].get() == "0":
                if aff_imp > 0:
                    self.fields["importo"]["var"].set(str(aff_imp))

            # Capitoli: aggrega tutti i capitoli dalle coperture finanziarie
            coperture = db.get_coperture(aid)
            if coperture:
                caps = sorted(list(set(str(c["capitolo"]) for c in coperture if c.get("capitolo"))))
                if caps:
                    current_caps = self.fields["capitolo_bilancio"]["var"].get().strip()
                    if force or not current_caps:
                        self.fields["capitolo_bilancio"]["var"].set(", ".join(caps))
                
                # Esercizio dal bilancio
                ann = coperture[0].get("anno_bilancio")
                if ann and (force or not self.fields["esercizio"]["var"].get().strip()):
                    self.fields["esercizio"]["var"].set(str(ann))
        except Exception as e:
            print(f"DEBUG sync error: {e}")

    def _sync_from_padre(self, force=False):
        """Pulla i dati dalla determina padre."""
        val = self.fields["_padre"]["var"].get()
        if not val: return
        did = self._pad_map.get(val)
        if not did: return
        
        padre = db.get_determina(did)
        if not padre: return
        
        # Converti in dict per usare .get in modo coerente
        padre = dict(padre)
        
        campi = ["oggetto", "capitolo_bilancio", "impegno_spesa", "esercizio", "rif_normativo", "importo", "rup", "qualifica_rup", "dirigente", "qualifica_dirigente"]
        for k in campi:
            if k in self.fields:
                if force or not self.fields[k]["var"].get().strip():
                    val = str(padre.get(k) or "")
                    if k == "importo" and val == "0": continue
                    self.fields[k]["var"].set(val)

    def save(self, v):
        d = {}
        for k in ["tipo_determina","numero","data_determina","oggetto","capitolo_bilancio","impegno_spesa",
                   "esercizio","rif_normativo","rup","qualifica_rup","dirigente","qualifica_dirigente",
                   "direttore_lavori","qualifica_dl","collaboratori",
                   "stato_iter","indicazioni_ai"]:
            d[k] = v.get(k, "")
        try:
            d["importo"] = float(v.get("importo", "0").replace(",", ".") or 0)
        except ValueError:
            d["importo"] = 0
        try:
            d["anno"] = int(v.get("anno", "")) if v.get("anno") else None
        except ValueError:
            d["anno"] = None
        d["esercizio"] = v.get("esercizio") or None
        d["id_affidamento"] = self._aff_map.get(v.get("_aff", ""))
        d["id_fattura"] = self._fat_map.get(v.get("_fat", ""))
        d["id_determina_padre"] = self._pad_map.get(v.get("_padre", ""))
        
        # Snapshot QE e Coperture se c'è un affidamento
        if d["id_affidamento"]:
            import json
            qe = db.get_quadro_economico(d["id_affidamento"])
            cop = db.get_coperture(d["id_affidamento"])
            d["snapshot_qe"] = json.dumps([dict(v) for v in qe])
            d["snapshot_coperture"] = json.dumps([dict(v) for v in cop])

        if self.is_edit:
            d["file_path"] = self.data.get("file_path")
            db.aggiorna_determina(self.data["id"], d)
        else:
            db.inserisci_determina(d)
        return True

    def delete(self):
        db.elimina_determina(self.data["id"])
        return True

    def get_title_detail(self):
        t = self.fields.get("tipo_determina", {}).get("var", tk.StringVar()).get()
        n = self.fields.get("numero", {}).get("var", tk.StringVar()).get()
        return f"{t} n.{n}" if n else t


class FormPersonale(FormBase):
    TITLE = "Soggetto (Personale Interno)"
    WIDTH = 500
    HEIGHT = 300

    def build_form(self):
        r = 0
        self.add_field("nome_cognome", "Nome e Cognome", 40, row=r, required=True); r += 1
        self.add_field("qualifica", "Qualifica", 50, row=r); r += 1
        self.add_field("attivo", "Attivo", 5, row=r, values=["Sì", "No"]); r += 1
        if not self.is_edit:
            self.fields["attivo"]["var"].set("Sì")

    def save(self, v):
        d = {
            "nome_cognome": v.get("nome_cognome"),
            "qualifica": v.get("qualifica"),
            "attivo": 1 if v.get("attivo") == "Sì" else 0
        }
        if self.is_edit:
            db.aggiorna_personale(self.data["id"], d)
        else:
            db.inserisci_personale(d)
        return True

    def delete(self):
        db.elimina_personale(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: GARA
# ═══════════════════════════════════════════════════════════════

class FormGara(FormBase):
    TITLE = "Gara"
    WIDTH = 720
    HEIGHT = 650

    def build_form(self):
        r = 0
        self.add_field("tipo_gara", "Tipo Gara", 32, row=r, values=db.TIPI_GARA, required=True); r += 1
        self.add_field("oggetto", "Oggetto", 52, row=r, required=True); r += 1
        self.add_field("cig", "CIG", 16, row=r); r += 1
        self.add_field("cup", "CUP", 16, row=r); r += 1
        self.add_separator("Procedura", row=r); r += 1
        self.add_field("importo_base_asta", "Base Asta €", 14, row=r); r += 1
        self.add_field("criterio_aggiudicazione", "Criterio", 35, row=r, values=["Offerta economicamente più vantaggiosa","Prezzo più basso"]); r += 1
        self.add_field("rup", "RUP", 38, row=r); r += 1
        self.add_field("piattaforma", "Piattaforma", 28, row=r); r += 1
        self.add_field("scadenza_offerte", "Scadenza Offerte", 14, row=r); r += 1
        self.add_separator("Esito", row=r); r += 1
        self.add_field("data_aggiudicazione", "Data Aggiudicazione", 14, row=r); r += 1
        self.add_field("importo_aggiudicazione", "Importo Aggiud. €", 14, row=r); r += 1
        self.add_field("ribasso_percentuale", "Ribasso %", 7, row=r); r += 1
        self.add_field("stato", "Stato", 14, row=r, values=["In preparazione","In corso","Aggiudicata","Completata","Annullata"]); r += 1
        self.add_field("fase_corrente", "Fase", 16, row=r, values=["Preparazione","Pubblicazione","Ricezione offerte","Valutazione","Aggiudicazione","Contratto","Completata"]); r += 1
        self.add_field("note", "Note", 52, row=r); r += 1

    def save(self, v):
        d = {k: v.get(k, "") for k in v}
        for k in ["importo_base_asta","importo_aggiudicazione","ribasso_percentuale"]:
            try:
                d[k] = float(d.get(k, "0").replace(",", ".") or 0)
            except ValueError:
                d[k] = 0
        if self.is_edit:
            db.aggiorna_gara(self.data["id"], d)
        else:
            db.inserisci_gara(d)
        return True

    def delete(self):
        db.elimina_gara(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: CONCESSIONE
# ═══════════════════════════════════════════════════════════════

class FormConcessione(FormBase):
    TITLE = "Concessione Demaniale"
    WIDTH = 680
    HEIGHT = 600

    def build_form(self):
        r = 0
        self.add_field("tipo_concessione", "Tipo", 30, row=r, values=["Stabilimento balneare","SLA (art. 45-bis)","Altro"]); r += 1
        self.add_field("oggetto", "Oggetto", 50, row=r, required=True); r += 1
        self.add_field("numero_concessione", "N. Concessione", 20, row=r); r += 1
        # Concessionario
        ops = db.cerca_operatori()
        self._op_map = {"": None}
        self._op_map.update({f"{o['id']} - {o['ragione_sociale']}": o["id"] for o in ops})
        self.add_field("_op", "Concessionario", 45, row=r, values=list(self._op_map.keys())); r += 1
        if self.is_edit and self.data.get("id_operatore"):
            for k, v in self._op_map.items():
                if v == self.data["id_operatore"]:
                    self.fields["_op"]["var"].set(k); break
        self.add_separator("Dettagli", row=r); r += 1
        self.add_field("ubicazione", "Ubicazione", 40, row=r); r += 1
        self.add_field("superficie_mq", "Superficie mq", 10, row=r); r += 1
        self.add_field("canone_annuo", "Canone annuo €", 14, row=r); r += 1
        self.add_field("durata_anni", "Durata (anni)", 5, row=r); r += 1
        self.add_field("data_rilascio", "Data Rilascio", 14, row=r); r += 1
        self.add_field("data_scadenza", "Data Scadenza", 14, row=r); r += 1
        self.add_separator("Riferimenti", row=r); r += 1
        self.add_field("cig", "CIG", 16, row=r); r += 1
        self.add_field("rif_normativo", "Rif. Normativo", 40, row=r); r += 1
        self.add_field("stato", "Stato", 14, row=r, values=["Attiva","Scaduta","Revocata","In rinnovo"]); r += 1
        self.add_field("note", "Note", 50, row=r); r += 1
        if not self.is_edit:
            self.fields["stato"]["var"].set("Attiva")

    def save(self, v):
        d = {k: v.get(k, "") for k in ["tipo_concessione","oggetto","numero_concessione","ubicazione",
             "data_rilascio","data_scadenza","stato","cig","rif_normativo","note"]}
        for k in ["canone_annuo","superficie_mq"]:
            try:
                d[k] = float(v.get(k, "0").replace(",", ".") or 0)
            except ValueError:
                d[k] = 0
        try:
            d["durata_anni"] = int(v.get("durata_anni", "0") or 0)
        except ValueError:
            d["durata_anni"] = 0
        d["id_operatore"] = self._op_map.get(v.get("_op", ""))
        if self.is_edit:
            db.aggiorna_concessione(self.data["id"], d)
        else:
            db.inserisci_concessione(d)
        return True

    def delete(self):
        db.elimina_concessione(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: VOCE QUADRO ECONOMICO (sub-form)
# ═══════════════════════════════════════════════════════════════

class FormVoceQE(FormBase):
    TITLE = "Voce Quadro Economico"
    WIDTH = 500
    HEIGHT = 340

    def __init__(self, master, data=None, on_save=None, lista_voci=None, indice=None):
        self.lista_voci = lista_voci if lista_voci is not None else []
        self.indice = indice
        super().__init__(master, data=data, on_save=on_save)

    def build_form(self):
        r = 0
        self.add_field("sezione", "Sezione", 28, row=r, values=["A - Lavori/Servizi","B - Somme a disposizione"]); r += 1
        if self.is_edit:
            self.fields["sezione"]["var"].set("B - Somme a disposizione" if (self.data or {}).get("sezione") == "B" else "A - Lavori/Servizi")
        else:
            self.fields["sezione"]["var"].set("A - Lavori/Servizi")
        self.add_field("descrizione", "Descrizione", 42, row=r, required=True); r += 1
        self.add_field("importo", "Importo €", 14, row=r); r += 1
        self.add_field("aliquota_iva", "IVA %", 7, row=r); r += 1
        if not self.is_edit:
            self.fields["aliquota_iva"]["var"].set("22")
        self.add_field("soggetto_iva", "Soggetto IVA", 5, row=r, values=["Sì","No"]); r += 1
        if not self.is_edit:
            self.fields["soggetto_iva"]["var"].set("Sì")
        else:
            self.fields["soggetto_iva"]["var"].set("Sì" if self.data.get("soggetto_iva") else "No")
        self.add_field("destinazione", "Destinazione", 18, row=r, values=["Fornitore","Ente"]); r += 1

    def save(self, v):
        try:
            imp = float(v.get("importo", "0").replace(",", ".") or 0)
        except ValueError:
            imp = 0
        try:
            aliq = float(v.get("aliquota_iva", "22") or 22)
        except ValueError:
            aliq = 22
        voce = {
            "sezione": "B" if v.get("sezione", "").startswith("B") else "A",
            "descrizione": v.get("descrizione", ""),
            "importo": imp, "aliquota_iva": aliq,
            "soggetto_iva": 1 if v.get("soggetto_iva") == "Sì" else 0,
            "destinazione": v.get("destinazione", "Fornitore")
        }
        if self.indice is not None:
            self.lista_voci[self.indice] = voce
        else:
            self.lista_voci.append(voce)
        return True

    def delete(self):
        if self.indice is not None:
            self.lista_voci.pop(self.indice)
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: COPERTURA FINANZIARIA (sub-form)
# ═══════════════════════════════════════════════════════════════

class FormCopertura(FormBase):
    TITLE = "Copertura Finanziaria"
    WIDTH = 560
    HEIGHT = 460

    def __init__(self, master, data=None, on_save=None, lista_voci=None, indice=None):
        self.lista_voci = lista_voci if lista_voci is not None else []
        self.indice = indice
        super().__init__(master, data=data, on_save=on_save)

    def build_form(self):
        r = 0
        self.add_field("missione", "Missione", 8, row=r); r += 1
        self.add_field("programma", "Programma", 8, row=r); r += 1
        self.add_field("titolo", "Titolo", 8, row=r); r += 1
        self.add_field("macroaggregato", "Macroaggregato", 8, row=r); r += 1
        self.add_field("capitolo", "Capitolo", 12, row=r); r += 1
        self.add_field("nome_capitolo", "Nome Capitolo", 40, row=r); r += 1
        self.add_field("anno_bilancio", "Bilancio (es. 2026-2028)", 16, row=r); r += 1
        self.add_field("annualita", "Annualità (es. 2026)", 10, row=r); r += 1
        self.add_field("importo", "Importo €", 14, row=r); r += 1
        self.add_field("note", "Note", 40, row=r); r += 1

    def save(self, v):
        c = {k: v.get(k, "") for k in v}
        try:
            c["importo"] = float(c.get("importo", "0").replace(",", ".") or 0)
        except ValueError:
            c["importo"] = 0
        
        # Sincronizza con la lista in memoria del padre
        if self.indice is not None:
            self.lista_voci[self.indice] = c
        else:
            self.lista_voci.append(c)
        
        # IMPORTANTE: Chiamiamo on_save per aggiornare la visualizzazione nel Treeview
        if self.on_save_callback:
            self.on_save_callback()
            
        return True

    def delete(self):
        if self.indice is not None:
            self.lista_voci.pop(self.indice)
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: VERBALE
# ═══════════════════════════════════════════════════════════════

class FormVerbale(FormBase):
    TITLE = "Verbale DL"
    WIDTH = 480
    HEIGHT = 320

    def __init__(self, master, data=None, on_save=None, id_affidamento=None):
        self.id_affidamento = id_affidamento
        super().__init__(master, data=data, on_save=on_save)

    def build_form(self):
        r = 0
        self.add_field("tipo_verbale", "Tipo", 30, row=r, values=db.TIPI_VERBALE, required=True); r += 1
        self.add_field("data_verbale", "Data (AAAA-MM-GG)", 14, row=r); r += 1
        if not self.is_edit:
            self.fields["data_verbale"]["var"].set(date.today().strftime("%Y-%m-%d"))
        self.add_field("redattore", "Redattore / DL", 35, row=r); r += 1
        self.add_field("importo_sal", "Importo SAL €", 14, row=r); r += 1
        self.add_field("note", "Note", 40, row=r); r += 1

    def save(self, v):
        d = {"id_affidamento": self.id_affidamento or (self.data or {}).get("id_affidamento"),
             "tipo_verbale": v["tipo_verbale"], "data_verbale": v["data_verbale"],
             "redattore": v.get("redattore", ""), "note": v.get("note", "")}
        try:
            d["importo_sal"] = float(v.get("importo_sal", "0").replace(",", ".") or 0) or None
        except ValueError:
            d["importo_sal"] = None
        if self.is_edit:
            db.aggiorna_verbale(self.data["id"], d)
        else:
            db.inserisci_verbale(d)
        return True

    def delete(self):
        db.elimina_verbale(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: VERIFICA
# ═══════════════════════════════════════════════════════════════

class FormVerifica(FormBase):
    TITLE = "Verifica"
    WIDTH = 500
    HEIGHT = 370

    def __init__(self, master, data=None, on_save=None, id_affidamento=None):
        self.id_affidamento = id_affidamento
        super().__init__(master, data=data, on_save=on_save)

    def build_form(self):
        r = 0
        self.add_field("tipo_verifica", "Tipo Verifica", 40, row=r, values=db.TIPI_VERIFICA, required=True); r += 1
        self.add_field("esito", "Esito", 8, row=r, values=["","OK","KO"]); r += 1
        self.add_field("data_verifica", "Data Verifica", 14, row=r); r += 1
        self.add_field("protocollo", "Protocollo", 18, row=r); r += 1
        self.add_field("data_scadenza", "Scadenza", 14, row=r); r += 1
        self.add_field("obbligatoria", "Obbligatoria", 5, row=r, values=["Sì","No"]); r += 1
        self.add_field("bloccante", "Bloccante", 5, row=r, values=["Sì","No"]); r += 1
        if not self.is_edit:
            self.fields["obbligatoria"]["var"].set("Sì")
            self.fields["bloccante"]["var"].set("Sì")

    def save(self, v):
        esito_map = {"OK": 1, "KO": 0, "": None}
        d = {"id_affidamento": self.id_affidamento or (self.data or {}).get("id_affidamento"),
             "tipo_verifica": v["tipo_verifica"],
             "esito": esito_map.get(v.get("esito", ""), None),
             "obbligatoria": 1 if v.get("obbligatoria") == "Sì" else 0,
             "bloccante": 1 if v.get("bloccante") == "Sì" else 0,
             "data_verifica": v.get("data_verifica", ""),
             "protocollo": v.get("protocollo", ""),
             "data_scadenza": v.get("data_scadenza", "")}
        if self.is_edit:
            db.aggiorna_verifica(self.data["id"], d)
        else:
            db.inserisci_verifica(d)
        return True

    def delete(self):
        db.elimina_verifica(self.data["id"])
        return True


# ═══════════════════════════════════════════════════════════════
# FORM: AFFIDAMENTO (la più complessa — con sotto-tab)
# ═══════════════════════════════════════════════════════════════

class FormAffidamento(tk.Toplevel):
    """L'affidamento ha sotto-tab (QE, Coperture, Verbali, ecc.), quindi
    non usa FormBase direttamente ma segue le stesse regole."""

    def __init__(self, master, data=None, on_save=None):
        super().__init__(master)
        self.data = data
        self.is_edit = data is not None and data.get("id") is not None
        self.on_save_callback = on_save
        self._modified = False
        self._editing = not self.is_edit
        self.fields = {}

        self.title("Affidamento" + (f" — {data['oggetto'][:40]}" if self.is_edit else " — Nuovo"))
        self.geometry("980x780"); self.transient(master); self.grab_set()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        self._build_tab_dati()
        self._build_tab_qe()
        self._build_tab_coperture()
        if self.is_edit:
            self._build_tab_iter()
            self._build_tab_verbali()
            self._build_tab_verifiche()
            self._build_tab_varianti()

        # Barra pulsanti
        bf = ttk.Frame(self); bf.pack(fill=tk.X, padx=10, pady=8)
        ttk.Button(bf, text="💾 Salva", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Annulla", command=self._close).pack(side=tk.LEFT, padx=5)
        if self.is_edit:
            ttk.Button(bf, text="✏ Modifica", command=self._start_edit).pack(side=tk.LEFT, padx=15)
            ttk.Button(bf, text="Cambia Stato", command=self._cambia_stato).pack(side=tk.LEFT, padx=5)
            ttk.Button(bf, text="🗑 Elimina", command=self._delete).pack(side=tk.RIGHT, padx=5)

        self._apply_lock()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _add_field(self, parent, key, label, width, row, values=None, required=False, readonly=False):
        ttk.Label(parent, text=("* " if required else "") + label).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 5))
        var = tk.StringVar(value=str((self.data or {}).get(key, "") or ""))
        if values:
            w = ttk.Combobox(parent, textvariable=var, values=values, width=width)
        else:
            w = ttk.Entry(parent, textvariable=var, width=width)
        w.grid(row=row, column=1, sticky="w", padx=5, pady=3)
        var.trace_add("write", lambda *a: setattr(self, "_modified", True))
        self.fields[key] = {"var": var, "widget": w, "readonly": readonly, "required": required}
        return var

    def _build_tab_dati(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Dati Generali  ")
        canvas = tk.Canvas(tab); sb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        m = ttk.Frame(canvas, padding=12)
        m.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=m, anchor="nw"); canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self._dati_frame = m

        r = 0
        self._add_field(m, "oggetto", "Oggetto", 60, r, required=True); r += 1
        self._add_field(m, "tipo_prestazione", "Tipo Prestazione", 40, r, values=db.TIPI_PRESTAZIONE, required=True); r += 1
        self._add_field(m, "tipo_procedura", "Tipo Procedura", 55, r, values=db.TIPI_PROCEDURA); r += 1

        ttk.Label(m, text="Classificazione Soglia", style="SubHeader.TLabel").grid(row=r, column=0, sticky="w", pady=3)
        self.fields["classificazione_soglia"] = {"var": tk.StringVar(value=(self.data or {}).get("classificazione_soglia", "")),
                                                  "widget": None, "readonly": True, "required": False}
        ttk.Label(m, textvariable=self.fields["classificazione_soglia"]["var"], style="Totale.TLabel").grid(row=r, column=1, sticky="w", padx=5, pady=3); r += 1

        # Operatore
        ops = db.cerca_operatori()
        self._op_map = {"": None}
        self._op_map.update({f"{o['id']} - {o['ragione_sociale']}": o["id"] for o in ops})
        self._add_field(m, "_op", "Operatore Economico", 50, r, values=list(self._op_map.keys())); r += 1
        if self.data and self.data.get("id_operatore"):
            for k, v in self._op_map.items():
                if v == self.data["id_operatore"]:
                    self.fields["_op"]["var"].set(k); break

        self._add_field(m, "cig", "CIG", 18, r); r += 1
        self._add_field(m, "cup", "CUP", 18, r); r += 1

        # Personale Interno
        pers = db.get_personale(attivi_only=True)
        self._pers_map = {p["nome_cognome"]: p for p in pers}
        nomi = sorted(list(self._pers_map.keys()))

        def _on_pers_change(key_nome, key_qual):
            nome = self.fields[key_nome]["var"].get()
            p = self._pers_map.get(nome)
            if p: self.fields[key_qual]["var"].set(p["qualifica"])

        v_rup = self._add_field(m, "rup", "RUP", 40, r, values=nomi); r += 1
        v_rup.trace_add("write", lambda *a: _on_pers_change("rup", "qualifica_rup"))
        self._add_field(m, "qualifica_rup", "Qualifica RUP", 50, r); r += 1

        v_dir = self._add_field(m, "dirigente", "Dirigente", 40, r, values=nomi); r += 1
        v_dir.trace_add("write", lambda *a: _on_pers_change("dirigente", "qualifica_dirigente"))
        self._add_field(m, "qualifica_dirigente", "Qualifica Dirigente", 60, r); r += 1

        v_dl = self._add_field(m, "direttore_lavori", "Direttore Lavori", 40, r, values=nomi); r += 1
        v_dl.trace_add("write", lambda *a: _on_pers_change("direttore_lavori", "qualifica_dl"))
        self._add_field(m, "qualifica_dl", "Qualifica DL", 60, r); r += 1

        # Collaboratori multiple selection
        ttk.Label(m, text="Collaboratori RUP").grid(row=r, column=0, sticky="w", pady=3, padx=(0, 5))
        v_coll = tk.StringVar(value=(self.data or {}).get("collaboratori", ""))
        ent_coll = ttk.Entry(m, textvariable=v_coll, width=60)
        ent_coll.grid(row=r, column=1, sticky="w", padx=5, pady=3)
        self.fields["collaboratori"] = {"var": v_coll, "widget": ent_coll}
        r += 1
        
        ttk.Label(m, text="Aggiungi collaboratore:", style="Info.TLabel").grid(row=r, column=0, sticky="w", padx=(10, 0))
        v_add_coll = tk.StringVar()
        cb_coll = ttk.Combobox(m, textvariable=v_add_coll, values=nomi, width=40, state="readonly")
        cb_coll.grid(row=r, column=1, sticky="w", padx=5)
        def _add_coll_name(*a):
            name = v_add_coll.get()
            if not name: return
            current = v_coll.get().strip()
            if current:
                if name not in current:
                    v_coll.set(current + ", " + name)
            else:
                v_coll.set(name)
            v_add_coll.set("") 
        cb_coll.bind("<<ComboboxSelected>>", _add_coll_name)
        r += 1

        if not self.is_edit:
            self.fields["dirigente"]["var"].set(db.get_impostazione("dirigente_default", "Arch. Linda Peruggi"))
            self.fields["qualifica_dirigente"]["var"].set(db.get_impostazione("qualifica_dirigente_default", "Dirigente del Settore Sviluppo Economico, Ambientale e Floricoltura"))

        # Preventivo e DURC
        ttk.Label(m, text="── Preventivo e DURC ──", style="SubHeader.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(12, 4)); r += 1
        self._add_field(m, "prot_preventivo", "Prot. Preventivo", 18, r); r += 1
        self._add_field(m, "data_preventivo", "Data Preventivo", 14, r); r += 1
        self._add_field(m, "prot_durc", "Prot. DURC", 18, r); r += 1
        self._add_field(m, "validita_durc", "Validità DURC", 14, r); r += 1
        self._add_field(m, "tempi_esecuzione", "Tempi Esecuzione", 30, r); r += 1
        self._add_field(m, "penali", "Penali", 30, r); r += 1

        # Contabile
        ttk.Label(m, text="── Contabile ──", style="SubHeader.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(12, 4)); r += 1
        self._add_field(m, "rif_normativo", "Rif. Normativo", 45, r, values=db.RIF_NORMATIVI); r += 1
        self._add_field(m, "forma_contratto", "Forma Contratto", 28, r, values=["corrispondenza","scrittura privata","atto pubblico"]); r += 1
        if not self.is_edit:
            self.fields["forma_contratto"]["var"].set("corrispondenza")

        # Stato (readonly)
        ttk.Label(m, text="Stato", style="SubHeader.TLabel").grid(row=r, column=0, sticky="w", pady=3)
        self.fields["stato"] = {"var": tk.StringVar(value=(self.data or {}).get("stato", "Bozza")), "widget": None, "readonly": True, "required": False}
        ttk.Label(m, textvariable=self.fields["stato"]["var"], style="Totale.TLabel").grid(row=r, column=1, sticky="w", padx=5, pady=3); r += 1

        # Premesse
        ttk.Label(m, text="── Premesse (per genera .docx) ──", style="SubHeader.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(12, 4)); r += 1
        for i in range(1, 4):
            self._add_field(m, f"premessa_{i}", f"Premessa {i}", 65, r); r += 1

        # Situazione contabile (solo edit)
        if self.is_edit:
            imp = db.get_importi_affidamento(self.data["id"])
            ttk.Label(m, text="── Situazione contabile ──", style="SubHeader.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(12, 4)); r += 1
            for lbl, val in [("Importo affidato", imp["importo_affidato"]), ("Varianti", imp["importo_varianti"]),
                             ("Tot. autorizzato", imp["importo_totale_autorizzato"]),
                             ("Tot. liquidato", imp["totale_liquidato"]), ("Residuo", imp["residuo"])]:
                color = "#e74c3c" if lbl == "Residuo" and val < 0 else "#1a5276"
                ttk.Label(m, text=lbl).grid(row=r, column=0, sticky="w", pady=2)
                ttk.Label(m, text=fmtE(val), font=("Segoe UI", 10, "bold"), foreground=color).grid(row=r, column=1, sticky="w", padx=5, pady=2); r += 1

    def _build_tab_qe(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Quadro Economico  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        self.tree_qe = ttk.Treeview(f, columns=("sez","desc","imp","iva","sogg","dest"), show="headings", height=10)
        for c, t, w in [("sez","Sez",35),("desc","Descrizione",260),("imp","Importo €",90),("iva","IVA%",50),("sogg","Sogg.",50),("dest","Dest.",80)]:
            self.tree_qe.heading(c, text=t); self.tree_qe.column(c, width=w)
        self.tree_qe.pack(fill=tk.BOTH, expand=True, pady=5)

        if self.is_edit:
            self.qe_voci = [dict(v) for v in db.get_quadro_economico(self.data["id"])]
        elif self.data and self.data.get("importo_affidato"):
            # Bozza AI: crea voce QE iniziale per i lavori/servizi
            self.qe_voci = [{
                "sezione": "A", "descrizione": "Lavori/Servizi (Bozza AI)",
                "importo": self.data["importo_affidato"], "aliquota_iva": 22,
                "soggetto_iva": 1, "destinazione": "Fornitore"
            }]
        else:
            self.qe_voci = []
        
        # Totali
        tf = ttk.LabelFrame(f, text="Totali"); tf.pack(fill=tk.X, pady=5)
        r1 = ttk.Frame(tf); r1.pack(fill=tk.X, padx=8, pady=3)
        self._qe_tots = {}
        for lbl, key in [("A) Lavori:","a"),("B) Somme disp.:","b"),("IVA:","iva"),("TOTALE:","gen")]:
            ttk.Label(r1, text=lbl).pack(side=tk.LEFT, padx=(12, 0))
            self._qe_tots[key] = tk.StringVar(value="€ 0,00")
            ttk.Label(r1, textvariable=self._qe_tots[key], style="Warning.TLabel" if key == "gen" else "Totale.TLabel").pack(side=tk.LEFT, padx=4)
        
        self._refresh_qe()

        bf = ttk.Frame(f); bf.pack(fill=tk.X, pady=3)
        ttk.Button(bf, text="+ Aggiungi", command=self._add_qe).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Modifica", command=self._edit_qe).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Elimina", command=self._del_qe).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Suggerisci Procedura", command=self._suggerisci).pack(side=tk.RIGHT, padx=3)
        ttk.Button(bf, text="📂 Importa QE da Excel", command=self._importa_qe_excel).pack(side=tk.RIGHT, padx=3)

    def _refresh_qe(self):
        for i in self.tree_qe.get_children(): self.tree_qe.delete(i)
        for i, v in enumerate(self.qe_voci):
            self.tree_qe.insert("","end",iid=str(i),values=(v.get("sezione","A"),v.get("descrizione",""),format_importo(v.get("importo",0)),f"{v.get('aliquota_iva',22)}%","Sì" if v.get("soggetto_iva",1) else "No",v.get("destinazione","Fornitore")))
        ta = sum(v.get("importo", 0) or 0 for v in self.qe_voci if v.get("sezione") == "A")
        tb = sum(v.get("importo", 0) or 0 for v in self.qe_voci if v.get("sezione") == "B")
        ti = sum(round((v.get("importo", 0) or 0) * (v.get("aliquota_iva", 0) or 0) / 100, 2) for v in self.qe_voci if v.get("soggetto_iva"))
        self._qe_tots["a"].set(fmtE(ta)); self._qe_tots["b"].set(fmtE(tb))
        self._qe_tots["iva"].set(fmtE(ti)); self._qe_tots["gen"].set(fmtE(ta + tb + ti))

    def _importa_qe_excel(self):
        path = filedialog.askopenfilename(title="Seleziona file Excel con Quadro Economico",
            filetypes=[("Excel","*.xlsx *.xls"),("Tutti","*.*")])
        if not path: return
        try:
            from importa_documenti import importa_qe_da_excel
            voci = importa_qe_da_excel(path)
            if not voci:
                messagebox.showwarning("Importazione", "Nessuna voce trovata nel file.\n"
                    "Assicurati che ci siano colonne 'Descrizione' e 'Importo'.")
                return
            tot = sum(v.get("importo", 0) for v in voci)
            msg = f"Trovate {len(voci)} voci per un totale di {fmtE(tot)}.\n\n"
            for v in voci[:8]:
                msg += f"  {'A' if v['sezione'] == 'A' else 'B'} | {v['descrizione'][:35]} | {fmtE(v['importo'])}\n"
            if len(voci) > 8:
                msg += f"  ... e altre {len(voci) - 8} voci\n"
            msg += "\nSostituire il QE attuale o aggiungere le voci?"
            r = messagebox.askyesnocancel("Importa QE", msg + "\n\nSì = Sostituisci, No = Aggiungi, Annulla = Niente")
            if r is None: return
            elif r:  # Sostituisci
                self.qe_voci.clear()
            self.qe_voci.extend(voci)
            self._refresh_qe()
            messagebox.showinfo("OK", f"Importate {len(voci)} voci.")
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def _add_qe(self):
        FormVoceQE(self, lista_voci=self.qe_voci, on_save=self._refresh_qe)

    def _edit_qe(self):
        sel = self.tree_qe.selection()
        if not sel: return
        idx = int(sel[0])
        FormVoceQE(self, data=self.qe_voci[idx], lista_voci=self.qe_voci, indice=idx, on_save=self._refresh_qe)

    def _del_qe(self):
        sel = self.tree_qe.selection()
        if not sel: return
        if messagebox.askyesno("Conferma", "Eliminare questa voce?"): self.qe_voci.pop(int(sel[0])); self._refresh_qe()

    def _suggerisci(self):
        tp = self.fields["tipo_prestazione"]["var"].get()
        tot = sum(v.get("importo", 0) or 0 for v in self.qe_voci)
        if tp and tot > 0:
            self.fields["tipo_procedura"]["var"].set(db.procedura_suggerita(tp, tot))
            self.fields["classificazione_soglia"]["var"].set(db.classificazione_soglia(tp, tot))

    def _build_tab_coperture(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Coperture Fin.  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        self.tree_cop = ttk.Treeview(f, columns=("miss","prog","tit","macr","cap","nome","anno","ann","imp"), show="headings", height=8)
        for c, t, w in [("miss","Miss.",40),("prog","Prog.",40),("tit","Tit.",35),("macr","Macr.",40),
                        ("cap","Capitolo",70),("nome","Nome",180),("anno","Bilancio",80),("ann","Ann.",60),("imp","Importo €",90)]:
            self.tree_cop.heading(c, text=t); self.tree_cop.column(c, width=w)
        self.tree_cop.pack(fill=tk.BOTH, expand=True, pady=5)

        self.cop_voci = [dict(c) for c in db.get_coperture(self.data["id"])] if self.is_edit else []
        
        # Totale
        self._lbl_tot_cop = ttk.Label(f, text="Totale Coperture: € 0,00", style="Totale.TLabel")
        self._lbl_tot_cop.pack(side=tk.TOP, anchor="e", pady=(0, 5))
        
        self._refresh_cop()

        bf = ttk.Frame(f); bf.pack(fill=tk.X, pady=3)
        ttk.Button(bf, text="+ Aggiungi Capitolo", command=self._add_cop).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Modifica", command=self._edit_cop).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Elimina", command=self._del_cop).pack(side=tk.LEFT, padx=3)

    def _refresh_cop(self):
        for i in self.tree_cop.get_children(): self.tree_cop.delete(i)
        tot = 0
        for i, c in enumerate(self.cop_voci):
            imp = c.get("importo", 0) or 0
            tot += imp
            self.tree_cop.insert("","end",iid=str(i),values=(c.get("missione",""),c.get("programma",""),c.get("titolo",""),c.get("macroaggregato",""),c.get("capitolo",""),c.get("nome_capitolo",""),c.get("anno_bilancio",""),c.get("annualita",""),format_importo(imp)))
        self._lbl_tot_cop.config(text=f"Totale Coperture: {fmtE(tot)}")

    def _add_cop(self):
        FormCopertura(self, lista_voci=self.cop_voci, on_save=self._refresh_cop)

    def _edit_cop(self):
        sel = self.tree_cop.selection()
        if not sel: return
        idx = int(sel[0])
        FormCopertura(self, data=self.cop_voci[idx], lista_voci=self.cop_voci, indice=idx, on_save=self._refresh_cop)

    def _del_cop(self):
        sel = self.tree_cop.selection()
        if not sel: return
        if messagebox.askyesno("Conferma", "Eliminare questa copertura?"): self.cop_voci.pop(int(sel[0])); self._refresh_cop()

    def _build_tab_iter(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Iter  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(f, columns=("id","tipo","num","data","importo","stato"), show="headings", height=10)
        for c, t, w in [("id","ID",35),("tipo","Tipo",150),("num","N.",70),("data","Data",80),("importo","Importo",90),("stato","Stato",80)]:
            tree.heading(c, text=t); tree.column(c, width=w)
        tree.pack(fill=tk.BOTH, expand=True, pady=5)
        for d in db.get_iter_affidamento(self.data["id"]):
            tree.insert("","end",values=(d["id"],d["tipo_determina"],d["numero"] or "",fmtD(d["data_determina"]),fmtE(d["importo"]),d["stato_iter"]))

    def _build_tab_verbali(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Verbali DL  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        self.tree_vb = ttk.Treeview(f, columns=("id","tipo","n","data","importo","note"), show="headings", height=8)
        for c, t, w in [("id","ID",35),("tipo","Tipo",160),("n","N.",40),("data","Data",80),("importo","SAL €",90),("note","Note",200)]:
            self.tree_vb.heading(c, text=t); self.tree_vb.column(c, width=w)
        self.tree_vb.pack(fill=tk.BOTH, expand=True, pady=5)
        self._refresh_vb()

        bf = ttk.Frame(f); bf.pack(fill=tk.X, pady=3)
        ttk.Button(bf, text="+ Nuovo Verbale", command=self._add_vb).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Modifica", command=self._edit_vb).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Elimina", command=self._del_vb).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Genera .docx", command=self._gen_vb_docx).pack(side=tk.LEFT, padx=12)

    def _refresh_vb(self):
        for i in self.tree_vb.get_children(): self.tree_vb.delete(i)
        for v in db.get_verbali(self.data["id"]):
            self.tree_vb.insert("","end",iid=str(v["id"]),values=(v["id"],v["tipo_verbale"],v["numero_progressivo"],fmtD(v["data_verbale"]),fmtE(v["importo_sal"]) if v["importo_sal"] else "",v["note"] or ""))

    def _add_vb(self):
        FormVerbale(self, id_affidamento=self.data["id"], on_save=self._refresh_vb)

    def _edit_vb(self):
        sel = self.tree_vb.selection()
        if not sel: return
        vid = int(sel[0])
        conn = db.get_connection()
        vb = dict(conn.execute("SELECT * FROM verbali_dl WHERE id=?", (vid,)).fetchone()); conn.close()
        FormVerbale(self, data=vb, id_affidamento=self.data["id"], on_save=self._refresh_vb)

    def _del_vb(self):
        sel = self.tree_vb.selection()
        if not sel: return
        if messagebox.askyesno("Conferma", "Eliminare questo verbale?"):
            db.elimina_verbale(int(sel[0])); self._refresh_vb()

    def _gen_vb_docx(self):
        sel = self.tree_vb.selection()
        if not sel: messagebox.showinfo("", "Seleziona un verbale."); return
        vid = int(sel[0])
        conn = db.get_connection()
        vb = dict(conn.execute("SELECT * FROM verbali_dl WHERE id=?", (vid,)).fetchone()); conn.close()
        aff = db.get_affidamento(self.data["id"])
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word","*.docx")],
            initialfile=f"Verbale_{vb['tipo_verbale'].replace(' ','_')}_{vb.get('numero_progressivo',1)}.docx")
        if not path: return
        try:
            genera_verbale(vb, dict(aff), path)
            messagebox.showinfo("OK", f"Verbale generato:\n{path}")
            try:
                if sys.platform == "win32": os.startfile(path)
                elif sys.platform == "darwin": subprocess.call(["open", path])
                else: subprocess.call(["xdg-open", path])
            except Exception: pass
        except Exception as e: messagebox.showerror("Errore", str(e))

    def _build_tab_verifiche(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Verifiche  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        self.tree_ck = ttk.Treeview(f, columns=("id","tipo","esito","obbl","blocc","data","scad"), show="headings", height=8)
        for c, t, w in [("id","ID",35),("tipo","Verifica",200),("esito","Esito",50),("obbl","Obbl.",45),("blocc","Blocc.",45),("data","Data",80),("scad","Scadenza",80)]:
            self.tree_ck.heading(c, text=t); self.tree_ck.column(c, width=w)
        self.tree_ck.pack(fill=tk.BOTH, expand=True, pady=5)
        self._refresh_ck()

        bf = ttk.Frame(f); bf.pack(fill=tk.X, pady=3)
        ttk.Button(bf, text="+ Aggiungi Verifica", command=self._add_ck).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="Elimina", command=self._del_ck).pack(side=tk.LEFT, padx=3)
        ok = db.verifiche_ok(self.data["id"])
        ttk.Label(bf, text="✓ Tutte le verifiche bloccanti OK" if ok else "⚠ Verifiche bloccanti incomplete",
                  style="Totale.TLabel" if ok else "Warning.TLabel").pack(side=tk.RIGHT, padx=10)

    def _refresh_ck(self):
        for i in self.tree_ck.get_children(): self.tree_ck.delete(i)
        for v in db.get_verifiche(self.data["id"]):
            self.tree_ck.insert("","end",iid=str(v["id"]),values=(v["id"],v["tipo_verifica"],"OK" if v["esito"] else ("KO" if v["esito"] == 0 else "—"),"Sì" if v["obbligatoria"] else "No","Sì" if v["bloccante"] else "No",fmtD(v["data_verifica"]),fmtD(v["data_scadenza"])))

    def _add_ck(self):
        FormVerifica(self, id_affidamento=self.data["id"], on_save=self._refresh_ck)

    def _del_ck(self):
        sel = self.tree_ck.selection()
        if not sel: return
        if messagebox.askyesno("Conferma", "Eliminare?"): db.elimina_verifica(int(sel[0])); self._refresh_ck()

    def _build_tab_varianti(self):
        tab = ttk.Frame(self.nb); self.nb.add(tab, text="  Varianti  ")
        f = ttk.Frame(tab, padding=10); f.pack(fill=tk.BOTH, expand=True)
        self.tree_var = ttk.Treeview(f, columns=("id","desc","importo","tipo","stato"), show="headings", height=6)
        for c, t, w in [("id","ID",35),("desc","Descrizione",250),("importo","Importo €",100),("tipo","Tipo",100),("stato","Stato",80)]:
            self.tree_var.heading(c, text=t); self.tree_var.column(c, width=w)
        self.tree_var.pack(fill=tk.BOTH, expand=True, pady=5)
        for v in db.get_varianti(self.data["id"]):
            self.tree_var.insert("","end",values=(v["id"],v["descrizione"],fmtE(v["importo_variante"]),v["tipo_variante"],v["stato"]))
        ttk.Button(f, text="+ Nuova Variante", command=self._add_var).pack(anchor="w", padx=3, pady=3)

    def _add_var(self):
        # Riutilizza un mini FormBase inline per le varianti
        win = tk.Toplevel(self); win.title("Variante"); win.geometry("450x280"); win.grab_set()
        m = ttk.Frame(win, padding=12); m.pack(fill=tk.BOTH, expand=True); fv = {}; r = 0
        for lbl, key, w in [("Descrizione *","desc",40),("Importo €","imp",14),("Tipo","tipo",20),("Motivazione","mot",40)]:
            ttk.Label(m, text=lbl).grid(row=r, column=0, sticky="w", pady=4); fv[key] = tk.StringVar(value="Integrativa" if key == "tipo" else "")
            if key == "tipo": ttk.Combobox(m, textvariable=fv[key], values=["Integrativa","In diminuzione","Suppletiva"], width=w).grid(row=r, column=1, sticky="w", padx=5, pady=4)
            else: ttk.Entry(m, textvariable=fv[key], width=w).grid(row=r, column=1, sticky="w", padx=5, pady=4)
            r += 1
        def salva():
            if not fv["desc"].get().strip(): messagebox.showwarning("", "Descrizione obbligatoria."); return
            try: imp = float(fv["imp"].get().replace(",", ".") or 0)
            except: messagebox.showwarning("", "Importo non valido."); return
            db.inserisci_variante({"id_affidamento": self.data["id"], "descrizione": fv["desc"].get(), "importo_variante": imp,
                "tipo_variante": fv["tipo"].get(), "motivazione": fv["mot"].get(), "stato": "Proposta"})
            win.destroy()
        ttk.Button(m, text="💾 Salva", command=salva).grid(row=r, column=0, columnspan=2, pady=10)

    # ── Lock/Unlock ──
    def _apply_lock(self):
        state = "normal" if self._editing else "disabled"
        for f in self.fields.values():
            if f.get("readonly"):
                continue
            w = f.get("widget")
            if w:
                try:
                    w.config(state=state)
                except tk.TclError:
                    pass

    def _start_edit(self):
        self._editing = True
        self._apply_lock()

    def _cambia_stato(self):
        validi = db.TRANSIZIONI_STATO.get(self.data["stato"], [])
        if not validi: messagebox.showinfo("", "Nessuna transizione disponibile."); return
        w = tk.Toplevel(self); w.title("Cambia stato"); w.geometry("350x150"); w.transient(self); w.grab_set()
        ttk.Label(w, text=f"Stato attuale: {self.data['stato']}", style="SubHeader.TLabel").pack(pady=8)
        sv = tk.StringVar(); ttk.Combobox(w, textvariable=sv, values=validi, width=30).pack(pady=5)
        def applica():
            if not sv.get(): return
            try:
                db.cambia_stato_affidamento(self.data["id"], sv.get())
                self.fields["stato"]["var"].set(sv.get()); self.data["stato"] = sv.get()
                messagebox.showinfo("", f"Stato: {sv.get()}"); w.destroy()
            except ValueError as e: messagebox.showerror("", str(e))
        ttk.Button(w, text="Applica", command=applica).pack(pady=8)

    def _save(self):
        if not self.fields["oggetto"]["var"].get().strip():
            messagebox.showwarning("Validazione", "Oggetto obbligatorio."); return
        d = {}
        # Lista campi effettivamente presenti nella UI
        campi_ui = ["oggetto","tipo_procedura","tipo_prestazione","classificazione_soglia","cig","cup",
                   "rup","qualifica_rup","dirigente","qualifica_dirigente",
                   "direttore_lavori","qualifica_dl","collaboratori",
                   "rif_normativo","forma_contratto","stato",
                   "premessa_1","premessa_2","premessa_3",
                   "prot_preventivo","data_preventivo","prot_durc","validita_durc",
                   "tempi_esecuzione","penali"]
        
        for k in campi_ui:
            if k in self.fields:
                d[k] = self.fields[k]["var"].get().strip()
        
        # Campi legacy o automatici
        d["id_operatore"] = self._op_map.get(self.fields["_op"]["var"].get())
        d["importo_affidato"] = sum(v.get("importo", 0) or 0 for v in self.qe_voci) + sum(
            round((v.get("importo", 0) or 0) * (v.get("aliquota_iva", 0) or 0) / 100, 2) for v in self.qe_voci if v.get("soggetto_iva"))
        
        # Pulla capitolo ed esercizio dalle coperture per mantenere consistenza nel DB
        if self.cop_voci:
            d["capitolo_bilancio"] = ", ".join(sorted(list(set(str(c["capitolo"]) for c in self.cop_voci if c.get("capitolo")))))
            d["esercizio"] = self.cop_voci[0].get("anno_bilancio", "")
        try:
            if self.is_edit:
                db.aggiorna_affidamento(self.data["id"], d); cid = self.data["id"]
            else:
                cid = db.inserisci_affidamento(d)
            for i, v in enumerate(self.qe_voci): v["ordine"] = i
            db.salva_quadro_economico(cid, self.qe_voci)
            for i, c in enumerate(self.cop_voci): c["ordine"] = i
            db.salva_coperture(cid, self.cop_voci)
            self._modified = False
            if self.on_save_callback: self.on_save_callback()
            self.destroy()
        except Exception as e: messagebox.showerror("Errore", str(e))

    def _delete(self):
        if messagebox.askyesno("Conferma", "Eliminare questo affidamento e tutti i dati collegati?"):
            db.elimina_affidamento(self.data["id"])
            if self.on_save_callback: self.on_save_callback()
            self.destroy()

    def _close(self):
        if self._modified and self._editing:
            r = messagebox.askyesnocancel("Modifiche non salvate", "Salvare prima di chiudere?")
            if r is True: self._save(); return
            elif r is None: return
        self.destroy()


# ═══════════════════════════════════════════════════════════════
# APP PRINCIPALE
# ═══════════════════════════════════════════════════════════════

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        servizio = db.get_impostazione("servizio_nome", "Servizio Demanio Marittimo")
        ente = db.get_impostazione("ente_nome", "Comune di Sanremo")
        self.title(f"Gestione Contratti — {servizio} — {ente}")
        self.geometry("1300x800"); self.minsize(1100, 700)

        s = ttk.Style(); s.theme_use("clam")
        s.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#1a5276")
        s.configure("SubHeader.TLabel", font=("Segoe UI", 11, "bold"), foreground="#2c3e50")
        s.configure("Info.TLabel", font=("Segoe UI", 9), foreground="#555")
        s.configure("Totale.TLabel", font=("Segoe UI", 10, "bold"), foreground="#1a5276")
        s.configure("Warning.TLabel", font=("Segoe UI", 10, "bold"), foreground="#e74c3c")
        s.configure("Treeview", rowheight=26, font=("Segoe UI", 9))
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tabs = {}
        for name in ["Dashboard", "Affidamenti", "Operatori", "Fatture", "Determine", "Gare", "Concessioni", "Personale", "Report"]:
            self.tabs[name] = ttk.Frame(self.nb)
            self.nb.add(self.tabs[name], text=f"  {name}  ")

        self._build_dashboard(); self._build_list_tab("Operatori", [("id","ID",35),("nome","Ragione Sociale",250),("tipo","Tipo",80),("piva","P.IVA",120),("cf","CF",130),("citta","Città",100),("pec","PEC",180)], self._refresh_operatori, FormOperatore, self._get_operatore, extra_buttons=self._op_extra_buttons)
        self._build_list_tab("Affidamenti", [("id","ID",35),("oggetto","Oggetto",250),("tipo","Tipo",120),("procedura","Procedura",200),("operatore","Operatore",150),("cig","CIG",100),("stato","Stato",120),("importo","Importo",90)], self._refresh_affidamenti, FormAffidamento, self._get_affidamento)
        self._build_list_tab("Fatture", [("id","ID",35),("num","N.Fattura",90),("data","Data",80),("op","Operatore",170),("aff","Affidamento",200),("netto","Netto",85),("tot","Totale",85),("tipo","Tipo",70),("stato","Stato",90)], self._refresh_fatture, FormFattura, self._get_fattura, extra_buttons=self._fat_extra_buttons)
        self._build_list_tab("Determine", [("id","ID",35),("tipo","Tipo",150),("num","N.",70),("data","Data",80),("aff","Affidamento",200),("op","Operatore",150),("importo","Importo",90),("stato","Stato",80)], self._refresh_determine, FormDetermina, self._get_determina, extra_buttons=self._det_extra_buttons)
        self._build_list_tab("Gare", [("id","ID",35),("tipo","Tipo",120),("oggetto","Oggetto",230),("cig","CIG",100),("fase","Fase",110),("stato","Stato",90),("imp","Base Asta",90)], self._refresh_gare, FormGara, self._get_gara)
        self._build_list_tab("Concessioni", [("id","ID",35),("tipo","Tipo",120),("oggetto","Oggetto",200),("op","Concessionario",150),("ubi","Ubicazione",120),("canone","Canone",90),("scad","Scadenza",80),("stato","Stato",70)], self._refresh_concessioni, FormConcessione, self._get_concessione)
        self._build_list_tab("Personale", [("id","ID",35),("nome","Nome Cognome",250),("qualifica","Qualifica",350),("attivo","Attivo",80)], self._refresh_personale, FormPersonale, self._get_personale)
        self._build_report()
        self._refresh_dashboard()

    def _make_tree(self, parent, cols_config, height=14):
        frame = ttk.Frame(parent); frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = [c[0] for c in cols_config]
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=height)
        for col, text, w in cols_config:
            tree.heading(col, text=text); tree.column(col, width=w, minwidth=30)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sy.pack(side=tk.RIGHT, fill=tk.Y)
        return tree

    # ── Pattern generico per tab lista ──
    def _build_list_tab(self, name, cols, refresh_fn, form_class, get_fn, extra_buttons=None):
        f = self.tabs[name]
        tb = ttk.Frame(f); tb.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(tb, text=name, style="Header.TLabel").pack(side=tk.LEFT)

        def on_save():
            refresh_fn(); self._refresh_dashboard()

        def new_record():
            form_class(self, on_save=on_save)

        def edit_record(event=None):
            sel = tree.selection()
            if not sel: return
            data = get_fn(int(sel[0]))
            if data: form_class(self, data=data, on_save=on_save)

        ttk.Button(tb, text=f"+ Nuovo", command=new_record).pack(side=tk.RIGHT, padx=5)
        if extra_buttons:
            extra_buttons(tb, lambda: tree)

        sf = ttk.Frame(f); sf.pack(fill=tk.X, padx=10, pady=(0, 5))
        search_var = tk.StringVar()
        ttk.Label(sf, text="Cerca:").pack(side=tk.LEFT)
        e = ttk.Entry(sf, textvariable=search_var, width=30); e.pack(side=tk.LEFT, padx=5)
        e.bind("<Return>", lambda e: refresh_fn(search_var.get()))
        ttk.Button(sf, text="Cerca", command=lambda: refresh_fn(search_var.get())).pack(side=tk.LEFT)

        tree = self._make_tree(f, cols)
        tree.bind("<Double-1>", edit_record)

        # Drag & Drop per Determine, Affidamenti e Fatture
        if name in ["Determine", "Affidamenti", "Fatture"]:
            tree.drop_target_register(DND_FILES)
            handler = {
                "Determine": self._handle_det_drop,
                "Affidamenti": self._handle_aff_drop,
                "Fatture": self._handle_fat_drop
            }.get(name)
            tree.dnd_bind("<<Drop>>", handler)
            ttk.Label(f, text=f"💡 Trascina qui documenti (PDF) per bozzare {'una ' + name[:-1] if name != 'Determine' else 'una Determina'} con AI", style="Info.TLabel").pack(pady=2)

        # Salva riferimenti
        setattr(self, f"tree_{name.lower()}", tree)
        setattr(self, f"search_{name.lower()}", search_var)
        refresh_fn()

    def _handle_det_drop(self, event):
        files = self.splitlist(event.data)
        if not files: return
        file_path = files[0]
        
        # Chiedi tipo determina
        win = tk.Toplevel(self)
        win.title("Bozza Determina da Documento")
        win.geometry("400x250")
        win.transient(self); win.grab_set()
        
        ttk.Label(win, text=f"File: {os.path.basename(file_path)}", wraplength=350).pack(pady=10)
        ttk.Label(win, text="Seleziona il tipo di determina da bozzare:").pack(pady=5)
        
        tipo_var = tk.StringVar(value="Determina Dirigenziale")
        combo = ttk.Combobox(win, textvariable=tipo_var, values=["Impegni/Accertamenti", "Liquidazione", "Determina Dirigenziale"], state="readonly", width=30)
        combo.pack(pady=5)
        
        def procedi():
            tipo = tipo_var.get()
            win.destroy()
            
            # Loading indicator (semplice)
            load = tk.Toplevel(self)
            load.title("Analisi AI...")
            load.geometry("300x100")
            ttk.Label(load, text="Analisi documento in corso con Claude AI...\nAttendere prego.").pack(expand=True)
            self.update()
            
            suggerimenti, errore = ai_helper.process_document_ai(file_path)
            load.destroy()
            
            if errore:
                messagebox.showerror("Errore AI", errore)
                return
            
            # Prepara dati per il form
            data = {
                "tipo_determina": tipo,
                "oggetto": suggerimenti.get("oggetto", ""),
                "importo": suggerimenti.get("importo", 0),
                "cig": suggerimenti.get("cig", ""),
                "indicazioni_ai": suggerimenti.get("premesse", "")[:450], 
                "note": suggerimenti.get("note", "")
            }

            # Se è una liquidazione, cerca di collegare la fattura
            if tipo == "Liquidazione":
                search_term = suggerimenti.get("numero_fattura") or suggerimenti.get("oggetto", "")
                if search_term:
                    fatture = db.cerca_fatture(search=search_term)
                    if fatture:
                        data["id_fattura"] = fatture[0]["id"]
                        data["id_affidamento"] = fatture[0]["id_affidamento"]
            
            # Apri il form con i dati suggeriti
            FormDetermina(self, data=data, on_save=lambda: (self._refresh_determine(), self._refresh_dashboard()))

        ttk.Button(win, text="Analizza con AI ✨", command=procedi).pack(pady=20)
        ttk.Button(win, text="Annulla", command=win.destroy).pack()

    def _handle_aff_drop(self, event):
        files = self.splitlist(event.data)
        if not files: return
        file_path = files[0]
        
        # Loading indicator
        load = tk.Toplevel(self)
        load.title("Analisi AI...")
        load.geometry("300x100")
        ttk.Label(load, text="Analisi preventivo/progetto in corso...\nAttendere prego.").pack(expand=True)
        self.update()
        
        suggerimenti, errore = ai_helper.process_document_ai(file_path, mode="affidamento")
        load.destroy()
        
        if errore:
            messagebox.showerror("Errore AI", errore)
            return
            
        data = {
            "oggetto": suggerimenti.get("oggetto", ""),
            "importo_affidato": suggerimenti.get("importo", 0),
            "cig": suggerimenti.get("cig", ""),
            "cup": suggerimenti.get("cup", ""),
            "tempi_esecuzione": suggerimenti.get("durata", ""),
            "premessa_1": suggerimenti.get("note", "")[:450]
        }
        
        # Cerca operatore
        if suggerimenti.get("operatore_nome"):
            ops = db.cerca_operatori(suggerimenti["operatore_nome"])
            if ops: data["id_operatore"] = ops[0]["id"]
            
        FormAffidamento(self, data=data, on_save=lambda: (self._refresh_affidamenti(), self._refresh_dashboard()))

    def _handle_fat_drop(self, event):
        files = self.splitlist(event.data)
        if not files: return
        file_path = files[0]
        
        # Loading indicator
        load = tk.Toplevel(self)
        load.title("Analisi AI...")
        load.geometry("300x100")
        ttk.Label(load, text="Analisi fattura in corso...\nAttendere prego.").pack(expand=True)
        self.update()
        
        suggerimenti, errore = ai_helper.process_document_ai(file_path, mode="fattura")
        load.destroy()
        
        if errore:
            messagebox.showerror("Errore AI", errore)
            return
            
        data = {
            "numero_fattura": suggerimenti.get("numero", ""),
            "data_fattura": suggerimenti.get("data", ""),
            "importo_totale": suggerimenti.get("importo", 0),
            "protocollo_pec": suggerimenti.get("sdi", ""),
            "note": suggerimenti.get("note", "")
        }
        
        # Cerca operatore
        if suggerimenti.get("operatore_nome"):
            ops = db.cerca_operatori(suggerimenti["operatore_nome"])
            if ops:
                data["id_operatore"] = ops[0]["id"]
                # Cerca affidamenti recenti dell'operatore per collegare la fattura
                affs = db.cerca_affidamenti(search=ops[0]["ragione_sociale"])
                for a in affs:
                    if a["id_operatore"] == ops[0]["id"]:
                        data["id_affidamento"] = a["id"]
                        break
            
        FormFattura(self, data=data, on_save=lambda: (self._refresh_fatture(), self._refresh_dashboard()))

    # ── Dashboard ──
    def _build_dashboard(self):
        f = self.tabs["Dashboard"]
        servizio = db.get_impostazione("servizio_nome", "Servizio Demanio Marittimo")
        ente = db.get_impostazione("ente_nome", "Comune di Sanremo")
        ttk.Label(f, text="Dashboard", style="Header.TLabel").pack(pady=(15, 5))
        ttk.Label(f, text=f"{servizio} — {ente}", style="Info.TLabel").pack(pady=(0, 15))
        self.sf = ttk.Frame(f); self.sf.pack(fill=tk.X, padx=20, pady=10)
        self.stat_labels = {}
        for i, (key, label, color) in enumerate([
            ("affidamenti_attivi", "Affidamenti Attivi", "#27ae60"), ("fatture_da_liquidare", "Fatture da Liquidare", "#e67e22"),
            ("importo_da_liquidare", "Importo da Liquidare", "#e74c3c"), ("determine_totali", "Determine", "#8e44ad"),
            ("gare_attive", "Gare", "#3498db"), ("concessioni_attive", "Concessioni", "#1a5276"),
        ]):
            card = ttk.LabelFrame(self.sf, text=label)
            card.grid(row=0, column=i, padx=6, pady=5, sticky="nsew"); self.sf.columnconfigure(i, weight=1)
            lbl = ttk.Label(card, text="0", font=("Segoe UI", 20, "bold"), foreground=color)
            lbl.pack(padx=12, pady=6); self.stat_labels[key] = lbl

        bf = ttk.Frame(f); bf.pack(pady=8)
        ttk.Button(bf, text="Aggiorna", command=self._refresh_dashboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Impostazioni", command=self._impostazioni).pack(side=tk.LEFT, padx=5)

    def _refresh_dashboard(self):
        st = db.get_statistiche()
        for key, lbl in self.stat_labels.items():
            val = st.get(key, 0)
            lbl.config(text=f"€ {format_importo(val)}" if key == "importo_da_liquidare" else str(val))

    def _impostazioni(self):
        win = tk.Toplevel(self); win.title("Impostazioni Generali"); win.geometry("600x480"); win.transient(self); win.grab_set()
        m = ttk.Frame(win, padding=15); m.pack(fill=tk.BOTH, expand=True)
        
        r = 0
        ttk.Label(m, text="Dati Ente e Dirigente", style="SubHeader.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(0, 8)); r += 1
        
        fields = [
            ("ente_nome", "Nome Ente (es. Comune di ...)"),
            ("settore_nome", "Nome Settore"),
            ("servizio_nome", "Nome Servizio"),
            ("dirigente_default", "Dirigente di Default"),
            ("qualifica_dirigente_default", "Qualifica Dirigente"),
            ("anthropic_api_key", "API Key Anthropic (AI)")
        ]
        
        vars = {}
        for key, label in fields:
            ttk.Label(m, text=label + ":").grid(row=r, column=0, sticky="w", pady=5)
            val = db.get_impostazione(key, "")
            v = tk.StringVar(value=val)
            entry = ttk.Entry(m, textvariable=v, width=50, show="*" if "key" in key else "")
            entry.grid(row=r, column=1, sticky="w", padx=5, pady=5)
            vars[key] = v
            r += 1
            
        sv = tk.StringVar(value="")
        ttk.Label(m, textvariable=sv, style="Info.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=3); r += 1
        
        def salva():
            for key, v in vars.items():
                db.set_impostazione(key, v.get().strip())
            sv.set("Impostazioni salvate.")
            
        def testa_ai():
            k = vars["anthropic_api_key"].get().strip()
            if not k: sv.set("Inserisci prima una API key."); return
            sv.set("Test AI in corso..."); win.update()
            try:
                from ai_helper import test_connessione
                ok, msg = test_connessione(k); sv.set(msg)
            except Exception as e: sv.set(f"Errore: {e}")

        def backup_db():
            import shutil
            path = filedialog.asksaveasfilename(defaultextension=".db", 
                initialfile=f"backup_contratti_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                filetypes=[("Database SQLite","*.db")])
            if not path: return
            try:
                shutil.copy2("contratti.db", path)
                messagebox.showinfo("Backup", f"Backup creato con successo:\n{path}")
            except Exception as e: messagebox.showerror("Errore", str(e))
            
        bf = ttk.Frame(m); bf.grid(row=r, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="Salva", command=salva).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Testa AI", command=testa_ai).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Crea Backup", command=backup_db).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Chiudi", command=win.destroy).pack(side=tk.LEFT, padx=5)

    # ── Refresh functions per ogni lista ──
    # ── Operatori: tasto Importa Excel ──
    def _op_extra_buttons(self, toolbar, get_tree):
        ttk.Button(toolbar, text="📂 Importa Excel", command=self._importa_operatori_excel).pack(side=tk.RIGHT, padx=5)

    def _importa_operatori_excel(self):
        excel_path = db.get_percorso_excel_fornitori()
        if excel_path and os.path.exists(excel_path):
            r = messagebox.askyesnocancel("Importa Operatori da Excel",
                f"Ultimo file usato:\n{excel_path}\n\nUsare questo file?")
            if r is None: return
            elif not r: excel_path = None
        if not excel_path or not os.path.exists(excel_path):
            excel_path = filedialog.askopenfilename(title="File Excel operatori",
                filetypes=[("Excel", "*.xlsx *.xls"), ("Tutti", "*.*")])
            if not excel_path: return
        try:
            result = db.sincronizza_fornitori_excel(excel_path)
            db.salva_percorso_excel_fornitori(excel_path)
            messagebox.showinfo("Importazione completata",
                f"Inseriti: {result['inseriti']}\nAggiornati: {result['aggiornati']}\nErrori: {result['errori']}")
            self._refresh_operatori()
        except Exception as e:
            messagebox.showerror("Errore importazione", str(e))

    # ── Fatture: tasto Importa PDF/XML ──
    def _fat_extra_buttons(self, toolbar, get_tree):
        ttk.Button(toolbar, text="📂 Importa Fattura PDF/XML", command=self._importa_fattura).pack(side=tk.RIGHT, padx=5)

    def _importa_fattura(self):
        path = filedialog.askopenfilename(title="Seleziona fattura",
            filetypes=[("Fattura","*.pdf *.xml"),("PDF","*.pdf"),("XML","*.xml"),("Tutti","*.*")])
        if not path: return
        try:
            from importa_documenti import importa_fattura_da_file
            self.config(cursor="watch"); self.update()
            dati = importa_fattura_da_file(path)
            self.config(cursor="")
            if not dati:
                messagebox.showwarning("Importazione", "Non è stato possibile estrarre dati dalla fattura.\nVerifica il file.")
                return
            # Mostra anteprima e conferma
            fonte = dati.get("_fonte", "")
            forn = dati.get("fornitore", {})
            msg = (f"Dati estratti ({fonte}):\n\n"
                   f"N. Fattura: {dati.get('numero_fattura', '—')}\n"
                   f"Data: {dati.get('data_fattura', '—')}\n"
                   f"Netto: € {dati.get('importo_netto', 0):,.2f}\n"
                   f"IVA {dati.get('aliquota_iva', 22)}%: € {dati.get('importo_iva', 0):,.2f}\n"
                   f"Totale: € {dati.get('importo_totale', 0):,.2f}\n"
                   f"Fornitore: {forn.get('ragione_sociale', forn.get('partita_iva', '—'))}\n\n"
                   f"Creare la fattura con questi dati?")
            if not messagebox.askyesno("Conferma importazione", msg):
                return
            # Crea la fattura
            fat_data = {
                "numero_fattura": dati.get("numero_fattura", ""),
                "data_fattura": dati.get("data_fattura", ""),
                "importo_netto": dati.get("importo_netto", 0),
                "aliquota_iva": dati.get("aliquota_iva", 22),
                "importo_iva": dati.get("importo_iva", 0),
                "importo_totale": dati.get("importo_totale", 0),
                "stato": "Registrata",
                "stato_pagamento": "Da liquidare",
                "tipo_liquidazione": "Saldo",
            }
            # Prova a collegare il fornitore
            if forn.get("partita_iva"):
                ops = db.cerca_operatori(forn["partita_iva"])
                if ops:
                    fat_data["id_operatore"] = ops[0]["id"]
            fid = db.inserisci_fattura(fat_data)
            messagebox.showinfo("OK", f"Fattura n.{fat_data['numero_fattura']} importata (ID {fid}).\n"
                                f"Apri la scheda per collegare l'affidamento.")
            self._refresh_fatture(); self._refresh_dashboard()
        except Exception as e:
            self.config(cursor="")
            messagebox.showerror("Errore", str(e))

    def _refresh_operatori(self, search=""):
        tree = self.tree_operatori
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_operatori(search):
            tree.insert("","end",iid=r["id"],values=(r["id"],r["ragione_sociale"],r["tipo_soggetto"],r["partita_iva"] or "",r["codice_fiscale"] or "",r["citta"] or "",r["pec"] or ""))

    def _get_operatore(self, oid):
        r = db.get_operatore(oid); return dict(r) if r else None

    def _refresh_affidamenti(self, search=""):
        tree = self.tree_affidamenti
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_affidamenti(search):
            tree.insert("","end",iid=r["id"],values=(r["id"],r["oggetto"][:40],r["tipo_prestazione"] or "",(r["tipo_procedura"] or "")[:35],r["operatore_nome"] or "",r["cig"] or "",r["stato"],fmtE(r['importo_affidato']) if r["importo_affidato"] else ""))

    def _get_affidamento(self, aid):
        r = db.get_affidamento(aid); return dict(r) if r else None

    def _refresh_fatture(self, search=""):
        tree = self.tree_fatture
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_fatture():
            tree.insert("","end",iid=r["id"],values=(r["id"],r["numero_fattura"],fmtD(r["data_fattura"]),r["operatore_nome"] or "",r["affidamento_oggetto"][:35] if r["affidamento_oggetto"] else "",fmtE(r["importo_netto"]),fmtE(r["importo_totale"]),r["tipo_liquidazione"] or "",r["stato_pagamento"]))

    def _get_fattura(self, fid):
        r = db.get_fattura(fid); return dict(r) if r else None

    def _refresh_determine(self, search=""):
        tree = self.tree_determine
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_determine():
            tree.insert("","end",iid=r["id"],values=(r["id"],r["tipo_determina"],r["numero"] or "",fmtD(r["data_determina"]) if r["data_determina"] else "",r["affidamento_oggetto"][:30] if r["affidamento_oggetto"] else "",r["operatore_nome"] or "",fmtE(r["importo"]) if r["importo"] else "",r["stato_iter"]))

    def _get_determina(self, did):
        r = db.get_determina(did); return dict(r) if r else None

    def _refresh_gare(self, search=""):
        tree = self.tree_gare
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_gare(search):
            tree.insert("","end",iid=r["id"],values=(r["id"],r["tipo_gara"],r["oggetto"][:35],r["cig"] or "",r["fase_corrente"],r["stato"],fmtE(r["importo_base_asta"]) if r["importo_base_asta"] else ""))

    def _get_gara(self, gid):
        r = db.get_gara(gid); return dict(r) if r else None

    def _refresh_concessioni(self, search=""):
        tree = self.tree_concessioni
        for i in tree.get_children(): tree.delete(i)
        for r in db.cerca_concessioni(search):
            tree.insert("","end",iid=r["id"],values=(r["id"],r["tipo_concessione"] or "",r["oggetto"][:30],r["operatore_nome"] or "",r["ubicazione"] or "",fmtE(r["canone_annuo"]) if r["canone_annuo"] else "",fmtD(r["data_scadenza"]),r["stato"]))

    def _get_concessione(self, cid):
        r = db.get_concessione(cid); return dict(r) if r else None

    def _refresh_personale(self, search=""):
        tree = self.tree_personale
        for i in tree.get_children(): tree.delete(i)
        for d in db.get_personale():
            if search and search.lower() not in d["nome_cognome"].lower(): continue
            tree.insert("", "end", iid=str(d["id"]), values=(d["id"],d["nome_cognome"],d["qualifica"],"Sì" if d["attivo"] else "No"))

    def _get_personale(self, pid):
        conn = db.get_connection()
        r = conn.execute("SELECT * FROM personale WHERE id=?", (pid,)).fetchone()
        conn.close()
        return dict(r) if r else None

    # ── Determine: bottoni extra (Genera .docx) ──
    def _det_extra_buttons(self, toolbar, get_tree):
        def gen(use_ai=False):
            tree = get_tree()
            sel = tree.selection()
            if not sel: messagebox.showinfo("", "Seleziona una determina."); return
            self._genera_determina_docx(int(sel[0]), use_ai)
        def gen_da_doc():
            tree = get_tree()
            sel = tree.selection()
            if not sel: messagebox.showinfo("", "Seleziona una determina."); return
            self._genera_determina_da_documento(int(sel[0]))
        ttk.Button(toolbar, text="Genera .docx", command=lambda: gen(False)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="Genera .docx + AI", command=lambda: gen(True)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="📂 AI da Documento", command=gen_da_doc).pack(side=tk.RIGHT, padx=5)

    def _genera_determina_da_documento(self, did):
        """Genera determina con premesse AI basate su un documento esterno (relazione, computo, ecc.)."""
        det = db.get_determina(did)
        if not det: return
        det = dict(det)
        path_doc = filedialog.askopenfilename(title="Seleziona documento (relazione, computo, perizia)",
            filetypes=[("Documenti","*.pdf *.docx *.xlsx *.txt"),("PDF","*.pdf"),("Word","*.docx"),("Excel","*.xlsx"),("Tutti","*.*")])
        if not path_doc: return
        try:
            from importa_documenti import genera_premesse_da_documento
            from ai_helper import is_ai_disponibile
            if not is_ai_disponibile():
                messagebox.showwarning("AI", "API key non configurata.\nDashboard > Impostazioni AI.")
                return
            self.config(cursor="watch"); self.update()
            # Merge affidamento
            aff_data = {}
            if det.get("id_affidamento"):
                aff = db.get_affidamento(det["id_affidamento"])
                if aff:
                    aff_data = dict(aff)
                    for k in ["prot_preventivo","data_preventivo","prot_durc","validita_durc","tempi_esecuzione","forma_contratto","premessa_1","premessa_2","premessa_3"]:
                        if aff_data.get(k) and not det.get(k): det[k] = aff_data[k]
            ai_premesse = genera_premesse_da_documento(path_doc, {**det, **aff_data})
            self.config(cursor="")
            if not ai_premesse:
                messagebox.showinfo("AI", "Non è stato possibile generare premesse dal documento.")
                return
            # Mostra anteprima
            preview = ai_premesse[:500] + ("..." if len(ai_premesse) > 500 else "")
            if not messagebox.askyesno("Premesse generate dall'AI",
                    f"Premesse basate su: {os.path.basename(path_doc)}\n\n{preview}\n\nGenerare la determina con queste premesse?"):
                return
            # Genera
            coperture = [dict(c) for c in db.get_coperture(det["id_affidamento"])] if det.get("id_affidamento") else []
            qe = [dict(v) for v in db.get_quadro_economico(det["id_affidamento"])] if det.get("id_affidamento") else []
            path_out = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word","*.docx")],
                initialfile=f"Det_{det.get('tipo_determina','').replace(' ','_')}_{det.get('numero') or 'BOZZA'}.docx")
            if not path_out: return
            genera_determina(det, path_out, coperture=coperture, qe=qe, ai_premesse=ai_premesse)
            messagebox.showinfo("OK", f"Generata:\n{path_out}")
            try:
                if sys.platform == "win32": os.startfile(path_out)
                elif sys.platform == "darwin": subprocess.call(["open", path_out])
                else: subprocess.call(["xdg-open", path_out])
            except Exception: pass
        except Exception as e:
            self.config(cursor="")
            messagebox.showerror("Errore", str(e))

    def _genera_determina_docx(self, did, use_ai=False):
        det = db.get_determina(did)
        if not det: return
        det = dict(det)
        
        # Merge dati Affidamento
        if det.get("id_affidamento"):
            aff = db.get_affidamento(det["id_affidamento"])
            if aff:
                aff = dict(aff)
                for k in ["prot_preventivo","data_preventivo","prot_durc","validita_durc","tempi_esecuzione","forma_contratto","premessa_1","premessa_2","premessa_3","tipo_prestazione","rup","qualifica_rup","direttore_lavori","qualifica_dl","collaboratori"]:
                    if aff.get(k) and not det.get(k): det[k] = aff[k]
                det["affidamento_oggetto"] = aff["oggetto"]
                det["aff_numero"] = aff.get("numero") # se presente
                det["aff_data"] = aff.get("data")
        
        # Merge dati Fattura e Operatore per LIQUIDAZIONE
        if det.get("tipo_determina") == "Liquidazione" and det.get("id_fattura"):
            conn = db.get_connection()
            f = conn.execute("SELECT f.*, o.ragione_sociale, o.codice_fiscale, o.partita_iva, o.indirizzo as operatore_indirizzo, o.citta as operatore_citta FROM fatture f LEFT JOIN operatori_economici o ON f.id_operatore=o.id WHERE f.id=?", (det["id_fattura"],)).fetchone()
            if f:
                f = dict(f)
                for k, v in f.items():
                    if v is not None: det[k] = v
            conn.close()
            
            # Cerca DURC e Regolare Esecuzione in checklist
            if det.get("id_affidamento"):
                conn = db.get_connection()
                cks = conn.execute("SELECT * FROM checklist_verifiche WHERE id_affidamento=?", (det["id_affidamento"],)).fetchall()
                for ck in cks:
                    ck = dict(ck)
                    t = ck["tipo_verifica"].lower()
                    if "durc" in t or "regolarità contributiva" in t:
                        det["durc_prot"] = ck.get("protocollo")
                        det["durc_scadenza"] = ck.get("data_scadenza")
                    if "regolare esecuzione" in t or "cre" in t:
                        det["data_regolare_esecuzione"] = ck.get("data_verifica")
                conn.close()

        # Merge dati Determina Padre
        if det.get("id_determina_padre"):
            padre = db.get_determina(det["id_determina_padre"])
            if padre:
                det["padre_numero"] = padre["numero"]
                det["padre_data"] = padre["data_determina"]

        # Snapshots QE e Coperture (priorità ai dati salvati nella determina)
        import json
        coperture = []
        if det.get("snapshot_coperture"):
            try: coperture = json.loads(det["snapshot_coperture"])
            except: pass
        if not coperture and det.get("id_affidamento"):
            coperture = [dict(c) for c in db.get_coperture(det["id_affidamento"])]

        qe = []
        if det.get("snapshot_qe"):
            try: qe = json.loads(det["snapshot_qe"])
            except: pass
        if not qe and det.get("id_affidamento"):
            qe = [dict(v) for v in db.get_quadro_economico(det["id_affidamento"])]

        ai_premesse = None
        if use_ai:
            try:
                from ai_helper import is_ai_disponibile, genera_premesse_ai
                if not is_ai_disponibile():
                    messagebox.showwarning("AI", "API key non configurata.\nDashboard > Impostazioni AI.")
                else:
                    self.config(cursor="watch"); self.update()
                    ai_premesse = genera_premesse_ai(det.get("tipo_determina"), det)
                    self.config(cursor="")
                    if not ai_premesse: messagebox.showinfo("AI", "L'AI non ha generato premesse. Uso template standard.")
            except Exception as e:
                self.config(cursor=""); messagebox.showwarning("AI", f"Errore: {e}")
        tipo = det.get("tipo_determina", "Determina")
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word","*.docx")],
            initialfile=f"Det_{tipo.replace(' ','_')}_{det.get('numero') or 'BOZZA'}.docx")
        if not path: return
        try:
            genera_determina(det, path, coperture=coperture, qe=qe, ai_premesse=ai_premesse)
            db.aggiorna_determina(did, {**det, "file_path": path})
            messagebox.showinfo("OK", f"Generata:\n{path}")
            try:
                if sys.platform == "win32": os.startfile(path)
                elif sys.platform == "darwin": subprocess.call(["open", path])
                else: subprocess.call(["xdg-open", path])
            except Exception: pass
            self._refresh_determine()
        except Exception as e: messagebox.showerror("Errore", str(e))

    # ── Report ──
    def _build_report(self):
        f = self.tabs["Report"]
        ttk.Label(f, text="Report e Query", style="Header.TLabel").pack(pady=(8, 5), padx=10, anchor="w")
        tf = ttk.Frame(f); tf.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(tf, text="Report:").pack(side=tk.LEFT)
        self.rep_combo = tk.StringVar(); ttk.Combobox(tf, textvariable=self.rep_combo, values=list(db.REPORT_PREIMPOSTATI.keys()), width=38).pack(side=tk.LEFT, padx=8)
        ttk.Button(tf, text="Esegui", command=self._run_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(tf, text="CSV", command=self._export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Separator(f, orient="horizontal").pack(fill=tk.X, padx=10, pady=8)
        qf = ttk.LabelFrame(f, text="Query SQL (solo SELECT)"); qf.pack(fill=tk.X, padx=10, pady=3)
        self.query_txt = tk.Text(qf, height=3, font=("Consolas", 10)); self.query_txt.pack(fill=tk.X, padx=5, pady=5)
        self.query_txt.insert("1.0", "SELECT * FROM affidamenti LIMIT 20")
        bq = ttk.Frame(qf); bq.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(bq, text="Esegui", command=self._run_query).pack(side=tk.LEFT, padx=5)
        ttk.Label(bq, text=f"Tabelle: {', '.join(db.get_nomi_tabelle())}", style="Info.TLabel").pack(side=tk.LEFT, padx=10)
        rf = ttk.LabelFrame(f, text="Risultati"); rf.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.rep_info = tk.StringVar(value="Nessun report"); ttk.Label(rf, textvariable=self.rep_info, style="Info.TLabel").pack(anchor="w", padx=5, pady=2)
        self.tree_rep = ttk.Treeview(rf, show="headings", height=12)
        sx = ttk.Scrollbar(rf, orient="horizontal", command=self.tree_rep.xview)
        sy = ttk.Scrollbar(rf, orient="vertical", command=self.tree_rep.yview)
        self.tree_rep.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)
        sx.pack(side=tk.BOTTOM, fill=tk.X); self.tree_rep.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sy.pack(side=tk.RIGHT, fill=tk.Y)
        self._rep_cols = []; self._rep_rows = []

    def _show_results(self, cols, rows):
        self._rep_cols = cols; self._rep_rows = rows
        self.tree_rep["columns"] = cols
        for i in self.tree_rep.get_children(): self.tree_rep.delete(i)
        for c in cols: self.tree_rep.heading(c, text=c); self.tree_rep.column(c, width=max(90, len(c) * 9), minwidth=50)
        for r in rows: self.tree_rep.insert("","end",values=[str(v) if v is not None else "" for v in r])
        self.rep_info.set(f"{len(rows)} righe")

    def _run_report(self):
        n = self.rep_combo.get()
        if not n: return
        try: c, r = db.esegui_report(n); self._show_results(c, r)
        except Exception as e: messagebox.showerror("Errore", str(e))

    def _run_query(self):
        sql = self.query_txt.get("1.0", tk.END).strip()
        if not sql: return
        try: c, r = db.esegui_query_libera(sql); self._show_results(c, r)
        except Exception as e: messagebox.showerror("Errore", str(e))

    def _export_csv(self):
        if not self._rep_cols: messagebox.showinfo("", "Esegui prima un report."); return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path: return
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";"); w.writerow(self._rep_cols)
            for r in self._rep_rows: w.writerow(r)
        messagebox.showinfo("OK", f"Esportato: {path}")


if __name__ == "__main__":
    App().mainloop()
