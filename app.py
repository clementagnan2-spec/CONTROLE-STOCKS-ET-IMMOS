# -*- coding: utf-8 -*-
"""Contrôle Stocks & Immobilisations - interface Tkinter."""
import os
import datetime
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import core
from core import to_num, fmt, fmtq, dfmt, parse_date

APP_TITLE = "Contrôle Stocks & Immobilisations"
DEFAULT_DIR = os.path.join(os.path.expanduser("~"), "Documents", "ControleStocksImmos")


SRC = {"article": "Article", "compte": "Moyenne du compte", "classe": "Moyenne de la classe", "aucun": "Aucun (saisir achats)"}


class Grid(ttk.Frame):
    """Tableau avec édition par double-clic sur les colonnes éditables."""

    def __init__(self, master, columns, editable=(), on_edit=None, height=14):
        super().__init__(master)
        self.keys = [c[0] for c in columns]
        self.editable = set(editable)
        self.on_edit = on_edit
        self.tree = ttk.Treeview(self, columns=self.keys, show="headings", selectmode="extended", height=height)
        for key, title, width, anchor in columns:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=False)
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure("warn", background="#fff3cd")
        self.tree.tag_configure("neg", foreground="#b00020")
        self.tree.bind("<Double-1>", self._dbl)

    def reload(self, rows):
        sel = set(self.tree.selection())
        yv = self.tree.yview()
        self.tree.delete(*self.tree.get_children())
        for iid, vals, tags in rows:
            self.tree.insert("", "end", iid=iid, values=vals, tags=tags)
        keep = [i for i in sel if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)
        self.tree.yview_moveto(yv[0])

    def _dbl(self, e):
        if self.tree.identify("region", e.x, e.y) != "cell":
            return
        iid = self.tree.identify_row(e.y)
        col = self.tree.identify_column(e.x)
        key = self.keys[int(col[1:]) - 1]
        if key not in self.editable or not iid:
            return
        x, y, w, h = self.tree.bbox(iid, col)
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=max(w, 60), height=h)
        entry.insert(0, self.tree.set(iid, key))
        entry.focus_set()
        entry.select_range(0, "end")
        done = {"v": False}

        def commit(_=None):
            if done["v"]:
                return
            done["v"] = True
            val = entry.get()
            sel = self.tree.selection()
            entry.destroy()
            if self.on_edit:
                self.on_edit(iid, key, val, sel)

        def cancel(_=None):
            done["v"] = True
            entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("1320x780")
        os.makedirs(DEFAULT_DIR, exist_ok=True)
        self.db = None
        self._open(os.path.join(DEFAULT_DIR, "projet.db"))
        self._menu()
        self._ui()
        self.refresh_all()

    # ------------------------------------------------------------------ projet
    def _open(self, path):
        self.db = core.DB(path)
        self.title(f"{APP_TITLE} - {path}")
        if not self.db.setting("entreprise"):
            self.db.set_setting("entreprise", "Mon entreprise")
            self.db.set_setting("exercice", str(datetime.date.today().year))

    def new_project(self):
        p = filedialog.asksaveasfilename(title="Nouveau projet", defaultextension=".db",
                                         initialdir=DEFAULT_DIR, filetypes=[("Projet", "*.db")])
        if p:
            if os.path.exists(p):
                os.remove(p)
            self._open(p)
            self.refresh_all()

    def open_project(self):
        p = filedialog.askopenfilename(title="Ouvrir un projet", initialdir=DEFAULT_DIR,
                                       filetypes=[("Projet", "*.db")])
        if p:
            self._open(p)
            self.refresh_all()

    def settings(self):
        e = simpledialog.askstring("Paramètres", "Nom de l'entreprise :", initialvalue=self.db.setting("entreprise"), parent=self)
        if e is not None:
            self.db.set_setting("entreprise", e)
        x = simpledialog.askstring("Paramètres", "Exercice (ex. 2025) :", initialvalue=self.db.setting("exercice"), parent=self)
        if x is not None:
            self.db.set_setting("exercice", x)

    def _menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="Nouveau projet...", command=self.new_project)
        f.add_command(label="Ouvrir un projet...", command=self.open_project)
        f.add_command(label="Paramètres (entreprise, exercice)...", command=self.settings)
        f.add_separator()
        f.add_command(label="Quitter", command=self.destroy)
        m.add_cascade(label="Fichier", menu=f)
        self.config(menu=m)

    # ------------------------------------------------------------------ UI
    def _ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)
        self._tab_journal()
        self._tab_lines()
        self._tab_sorties()
        self._tab_inv()
        self._tab_immos()
        self._tab_reg()

    def _tab(self, title):
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text=title)
        return f

    # ---- 1. Journal
    def _tab_journal(self):
        t = self._tab("1. Journal")
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="Importer le journal (Excel / CSV)...", command=self.import_journal).pack(side="left")
        ttk.Button(bar, text="Télécharger le modèle de journal", command=self.download_template).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Générer stocks, sorties (701/702) et immobilisations",
                   command=self.generate).pack(side="left", padx=8)
        self.j_info = ttk.Label(t, text="", font=("TkDefaultFont", 9, "bold"))
        self.j_info.pack(anchor="w", pady=(4, 0))
        ttk.Label(t, text="Colonnes attendues (ligne d'en-tête) : Date, Journal, Pièce, Compte, Libellé, Débit, Crédit. Comptes exploités : 601/602 (achats), 701/702 (ventes), 21 à 24 (immobilisations).",
                  foreground="#555").pack(anchor="w", pady=4)
        cols = [("date", "Date", 90, "center"), ("journal", "Journal", 70, "center"), ("piece", "Pièce", 110, "w"),
                ("compte", "Compte", 90, "w"), ("libelle", "Libellé", 450, "w"),
                ("debit", "Débit", 120, "e"), ("credit", "Crédit", 120, "e")]
        self.j_grid = Grid(t, cols)
        self.j_grid.pack(fill="both", expand=True)

    def download_template(self):
        p = self._save(".xlsx", "modele_journal.xlsx", [("Excel", "*.xlsx")])
        if p:
            try:
                core.export_journal_template(p)
            except Exception as ex:
                messagebox.showerror("Modèle", f"Impossible d'enregistrer le modèle :\n{ex}")
                return
            core.open_file(p)

    def import_journal(self):
        p = filedialog.askopenfilename(title="Journal des opérations comptables",
                                       filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv *.txt"), ("Tous", "*.*")])
        if not p:
            return
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            entries = core.parse_journal(core.read_rows(p))
        except Exception as ex:
            messagebox.showerror("Import impossible", str(ex))
            return
        finally:
            self.config(cursor="")
        if not entries:
            messagebox.showwarning("Import", "Aucune écriture exploitable trouvée.")
            return
        self.db.import_journal(entries)
        self.refresh_journal()
        if messagebox.askyesno("Import terminé", f"{len(entries)} lignes importées.\n\nGénérer maintenant les stocks et les immobilisations ?\n"
                               "(les quantités / références déjà saisies sont conservées)"):
            self.generate()

    def generate(self):
        if self.db.count("journal") == 0:
            messagebox.showinfo("Génération", "Importez d'abord le journal.")
            return
        ns, ni, nn, nso = self.db.generate()
        self.refresh_all()
        messagebox.showinfo(
            "Génération terminée",
            f"Stocks : {ns} lignes d'achats (601/602) rattachées aux comptes 31/32.\n"
            f"Ventes 701/702 : {nso} lignes rattachées aux articles (sorties de stock).\n"
            f"Immobilisations : {ni} lignes (comptes 21 à 24), {nn} nouvelle(s) composante(s) créée(s).\n\n"
            "Étape suivante : renseignez les QUANTITÉS achetées (onglet 2) puis les QUANTITÉS vendues (onglet 3).")
        self.nb.select(1)

    def refresh_journal(self):
        rows = []
        td = tc = 0.0
        for r in self.db.c.execute("SELECT * FROM journal ORDER BY id"):
            td += r["debit"] or 0
            tc += r["credit"] or 0
            if len(rows) < 20000:
                rows.append((str(r["id"]), (dfmt(r["date"]), r["journal"], r["piece"], r["compte"], r["libelle"],
                                            fmt(r["debit"]), fmt(r["credit"])), ()))
        self.j_grid.reload(rows)
        n = self.db.count("journal")
        eq = "équilibré" if abs(td - tc) < 0.01 else f"ÉCART {fmt(td - tc)}"
        self.j_info.config(text=f"{n} lignes | Débit {fmt(td)} | Crédit {fmt(tc)} | {eq}"
                           + (" | (affichage limité à 20 000 lignes)" if n > 20000 else ""))

    # ---- 2. Lignes de stock
    def _tab_lines(self):
        t = self._tab("2. Achats → Stocks (quantités)")
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Label(bar, text="Filtre :").pack(side="left")
        self.l_filter = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.l_filter, width=30)
        e.pack(side="left", padx=4)
        e.bind("<KeyRelease>", lambda _: self.refresh_lines())
        self.l_info = ttk.Label(bar, text="")
        self.l_info.pack(side="left", padx=12)
        ttk.Label(t, foreground="#555", text=(
            "Double-clic sur « Article », « Quantité » ou « Compte stock » pour modifier. Pour regrouper des achats sous un même article : "
            "sélectionnez plusieurs lignes (Ctrl/Maj) puis modifiez « Article » une fois. Lignes jaunes = quantité manquante."
        ), wraplength=1250, justify="left").pack(anchor="w", pady=4)
        cols = [("compte_achat", "Compte achat", 110, "w"), ("compte_stock", "Compte stock", 110, "w"),
                ("date", "Date", 100, "center"), ("piece", "Pièce", 100, "w"), ("libelle", "Libellé", 330, "w"),
                ("article", "Article (éditable)", 260, "w"), ("montant", "Montant", 110, "e"),
                ("quantite", "Quantité (éditable)", 160, "e")]
        self.l_grid = Grid(t, cols, editable=("article", "quantite", "compte_stock"), on_edit=self.edit_line)
        self.l_grid.pack(fill="both", expand=True)

    def edit_line(self, iid, key, val, sel):
        ids = [int(i) for i in sel] if (iid in sel and len(sel) > 1 and key != "quantite") else [int(iid)]
        if key == "quantite":
            self.db.set_line_qty(int(iid), to_num(val))
        elif key == "article":
            self.db.set_line_article(ids, val)
        elif key == "compte_stock" and val.strip():
            self.db.set_line_stock_account(ids, val.strip())
        self.refresh_lines()
        self.refresh_sorties()
        self.refresh_inv()

    def refresh_lines(self):
        q = self.l_filter.get().lower().strip()
        rows = []
        miss = 0
        total = 0
        for r in self.db.c.execute("SELECT * FROM stock_lines ORDER BY compte_stock, article, date, id"):
            total += 1
            if not r["quantite"]:
                miss += 1
            vals = (r["compte_achat"], r["compte_stock"], dfmt(r["date"]), r["piece"], r["libelle"], r["article"],
                    fmt(r["montant"]), fmtq(r["quantite"]))
            if q and q not in " ".join(map(str, vals)).lower():
                continue
            rows.append((str(r["id"]), vals, () if r["quantite"] else ("warn",)))
        self.l_grid.reload(rows)
        self.l_info.config(text=f"{total} lignes | {miss} sans quantité")

    # ---- 3. Ventes -> sorties de stock
    def _tab_sorties(self):
        t = self._tab("3. Ventes → Sorties de stock")
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Label(bar, text="Filtre :").pack(side="left")
        self.s_filter = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.s_filter, width=28)
        e.pack(side="left", padx=4)
        e.bind("<KeyRelease>", lambda _: self.refresh_sorties())
        ttk.Button(bar, text="Associer à un article...", command=self.sort_assoc).pack(side="left", padx=6)
        ttk.Button(bar, text="Rattachement automatique", command=self.sort_auto).pack(side="left")
        self.s_info = ttk.Label(t, text="", font=("TkDefaultFont", 9, "bold"))
        self.s_info.pack(anchor="w", pady=(4, 0))
        ttk.Label(t, foreground="#555", wraplength=1250, justify="left", text=(
            "Les ventes (701 → stock 311, 702 → stock 321) sont rattachées automatiquement à l'article d'achat le plus proche. "
            "Vérifiez le rattachement (« Associer à un article » ou modifier « Article »), puis saisissez la QUANTITÉ vendue. "
            "Coût de sortie = quantité vendue x CMUP de l'article ; si la vente n'est pas rattachée à un article, le CMUP moyen du compte de stock (311, 321...) est utilisé (colonne « Base du CMUP »). Rouge = article non rattaché à un achat ; jaune = quantité manquante."
        )).pack(anchor="w", pady=4)
        cols = [("compte_vente", "Compte vente", 100, "w"), ("compte_stock", "Compte stock", 100, "w"),
                ("date", "Date", 90, "center"), ("piece", "Pièce", 90, "w"), ("libelle", "Libellé", 260, "w"),
                ("article", "Article (éditable)", 230, "w"), ("montant", "Prix de vente", 110, "e"),
                ("quantite", "Qté vendue (éditable)", 140, "e"), ("cmup", "CMUP", 90, "e"), ("src", "Base du CMUP", 130, "w"),
                ("cout", "Coût de sortie", 110, "e"), ("marge", "Marge brute", 110, "e")]
        self.s_grid = Grid(t, cols, editable=("article", "quantite", "compte_stock"), on_edit=self.edit_sortie)
        self.s_grid.tree.tag_configure("bad", foreground="#b00020")
        self.s_grid.pack(fill="both", expand=True)

    def edit_sortie(self, iid, key, val, sel):
        ids = [int(i) for i in sel] if (iid in sel and len(sel) > 1 and key != "quantite") else [int(iid)]
        if key == "quantite":
            self.db.set_sortie_qty(int(iid), to_num(val))
        elif key == "article":
            self.db.set_sortie_article(ids, val)
        elif key == "compte_stock" and val.strip():
            self.db.set_sortie_stock_account(ids, val.strip())
        self.refresh_sorties()
        self.refresh_inv()

    def sort_assoc(self):
        sel = self.s_grid.tree.selection()
        if not sel:
            messagebox.showinfo("Ventes", "Sélectionnez d'abord une ou plusieurs ventes.")
            return
        arts = [(r[0], r[1]) for r in self.db.c.execute(
            "SELECT DISTINCT compte_stock, article FROM stock_lines ORDER BY compte_stock, article")]
        if not arts:
            messagebox.showinfo("Ventes", "Aucun article d'achat disponible.")
            return
        win = tk.Toplevel(self)
        win.title("Associer à un article d'achat")
        win.transient(self)
        win.grab_set()
        lb = tk.Listbox(win, width=80, height=16, exportselection=False)
        for cs, ar in arts:
            lb.insert("end", f"{cs}   |   {ar}")
        lb.pack(padx=8, pady=8, fill="both", expand=True)

        def ok():
            i = lb.curselection()
            if not i:
                return
            cs, ar = arts[i[0]]
            self.db.set_sortie_link([int(x) for x in sel], cs, ar)
            win.destroy()
            self.refresh_sorties()
            self.refresh_inv()

        b = ttk.Frame(win)
        b.pack(pady=(0, 8))
        ttk.Button(b, text="Associer", command=ok).pack(side="left", padx=4)
        ttk.Button(b, text="Annuler", command=win.destroy).pack(side="left", padx=4)
        lb.bind("<Double-1>", lambda _: ok())

    def sort_auto(self):
        n = self.db.auto_match()
        self.refresh_sorties()
        self.refresh_inv()
        messagebox.showinfo("Rattachement", f"{n} vente(s) rattachée(s) automatiquement.")

    def refresh_sorties(self):
        q = self.s_filter.get().lower().strip()
        known = {(r[0], r[1]) for r in self.db.c.execute("SELECT DISTINCT compte_stock, article FROM stock_lines")}
        rows = []
        miss = bad = nocmup = 0
        tv = tc = 0.0
        data = self.db.sorties_rows()
        for r in data:
            tv += r["montant"] or 0
            tc += r["cout"]
            tags = []
            if not r["quantite"]:
                miss += 1
                tags.append("warn")
            if (r["compte_stock"], r["article"]) not in known:
                bad += 1
                tags.append("bad")
            vals = (r["compte_vente"], r["compte_stock"], dfmt(r["date"]), r["piece"], r["libelle"], r["article"],
                    fmt(r["montant"]), fmtq(r["quantite"]), fmt(r["cmup"]), SRC.get(r["cmup_src"], ""),
                    fmt(r["cout"]), fmt(r["marge"]))
            if r["cmup_src"] == "aucun" or not r["cmup"]:
                nocmup += 1
            if q and q not in " ".join(map(str, vals)).lower():
                continue
            rows.append((str(r["id"]), vals, tuple(tags)))
        self.s_grid.reload(rows)
        nq = self.db.c.execute("SELECT COUNT(*) FROM stock_lines WHERE COALESCE(quantite,0)=0").fetchone()[0]
        self.s_info.config(text=f"{len(data)} ventes | {miss} sans quantité | {bad} non rattachées | {nocmup} sans CMUP | "
                                f"Ventes {fmt(tv)} | Coût des sorties {fmt(tc)} | Marge brute {fmt(tv - tc)}"
                                + (f" | ATTENTION : {nq} achat(s) sans quantité (CMUP faussé)" if nq else ""))

    # ---- 4. Inventaire / CMUP
    def _tab_inv(self):
        t = self._tab("4. Inventaire & CMUP")
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="PDF fiches par compte", command=lambda: self.pdf_stock(False)).pack(side="left")
        ttk.Button(bar, text="PDF fiches individuelles (par article)", command=lambda: self.pdf_stock(True)).pack(side="left", padx=6)
        ttk.Button(bar, text="Export Excel", command=self.xlsx_stock).pack(side="left")
        self.i_info = ttk.Label(t, text="", font=("TkDefaultFont", 9, "bold"))
        self.i_info.pack(anchor="w", pady=(4, 0))
        ttk.Label(t, foreground="#555", wraplength=1250, justify="left", text=(
            "CMUP = (valeur du stock initial + achats) / (quantité initiale + quantités achetées). Sorties = quantités vendues x CMUP. "
            "Quantité théorique = initiale + achats - sorties. Saisissez la quantité inventoriée (comptage) : l'écart avec le théorique est calculé, "
            "et le stock final = quantité inventoriée x CMUP (si non saisie, la quantité théorique est retenue). Jaune = quantité théorique négative, sortie sans CMUP ou achat sans quantité (CMUP faussé)."
        )).pack(anchor="w", pady=4)
        cols = [("compte", "Compte stock", 100, "w"), ("article", "Article", 240, "w"), ("unite", "Unité", 55, "center"),
                ("qte_init", "Qté initiale", 85, "e"), ("val_init", "Valeur initiale", 110, "e"),
                ("q_ach", "Qté achats", 85, "e"), ("m_ach", "Montant achats", 115, "e"), ("cmup", "CMUP", 95, "e"),
                ("q_sort", "Qté sorties", 85, "e"), ("v_sort", "Coût sorties", 110, "e"),
                ("q_theo", "Qté théorique", 95, "e"), ("qte_inv", "Qté inventaire", 105, "e"),
                ("ecart", "Écart qté", 80, "e"), ("v_inv", "Valeur stock final", 125, "e")]
        self.i_grid = Grid(t, cols, editable=("unite", "qte_init", "val_init", "qte_inv"), on_edit=self.edit_inv)
        self.i_grid.pack(fill="both", expand=True)

    def edit_inv(self, iid, key, val, sel):
        compte, article = iid.split("|", 1)
        if key == "unite":
            v = val.strip()
        elif key == "qte_inv":
            v = None if not val.strip() else to_num(val)
        else:
            v = to_num(val)
        self.db.set_article_field(compte, article, key, v)
        self.refresh_inv()
        self.refresh_sorties()

    def refresh_inv(self):
        rows = []
        tot = {"vi": 0.0, "vs": 0.0, "vf": 0.0}
        for f in self.db.fiches():
            tot["vi"] += f["val_init"]
            tot["vs"] += f["v_sort"]
            tot["vf"] += f["v_inv"]
            warn = f["q_theo"] < -0.0005 or (f["q_sort"] and not f["cmup"]) or f["sans_qte"]
            rows.append((f"{f['compte']}|{f['article']}",
                         (f["compte"], f["article"], f["unite"], fmtq(f["qte_init"]), fmt(f["val_init"]),
                          fmtq(f["q_ach"]), fmt(f["m_ach"]), fmt(f["cmup"]), fmtq(f["q_sort"]), fmt(f["v_sort"]),
                          fmtq(f["q_theo"]) or "0", fmtq(f["qte_inv"]) if f["counted"] else "",
                          fmtq(f["ecart_q"]), fmt(f["v_inv"])),
                         ("warn",) if warn else ()))
        self.i_grid.reload(rows)
        self.i_info.config(text=f"Stock initial {fmt(tot['vi'])} | Coût des sorties {fmt(tot['vs'])} | "
                                f"Stock final {fmt(tot['vf'])} | Variation (final - initial) {fmt(tot['vf'] - tot['vi'])}")

    def _save(self, ext, name, types):
        return filedialog.asksaveasfilename(defaultextension=ext, initialfile=name, initialdir=DEFAULT_DIR, filetypes=types)

    def pdf_stock(self, individual):
        p = self._save(".pdf", "fiches_individuelles_stock.pdf" if individual else "fiches_inventaire_stock.pdf",
                       [("PDF", "*.pdf")])
        if p:
            core.export_stock_pdf(self.db, p, individual)
            core.open_file(p)

    def xlsx_stock(self):
        p = self._save(".xlsx", "inventaire_stocks.xlsx", [("Excel", "*.xlsx")])
        if p:
            core.export_stock_xlsx(self.db, p)
            core.open_file(p)

    # ---- 4. Immobilisations
    def _tab_immos(self):
        t = self._tab("5. Immobilisations")
        top = ttk.LabelFrame(t, text="Acquisitions (comptes 21, 22, 23, 24) - sélectionnez une ligne", padding=4)
        top.pack(fill="both", expand=True)
        bar = ttk.Frame(top)
        bar.pack(fill="x")
        ttk.Label(bar, text="Filtre :").pack(side="left")
        self.m_filter = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.m_filter, width=30)
        e.pack(side="left", padx=4)
        e.bind("<KeyRelease>", lambda _: self.refresh_immos())
        cols = [("compte", "Compte", 90, "w"), ("date", "Date", 100, "center"), ("piece", "Pièce", 100, "w"),
                ("libelle", "Libellé", 330, "w"), ("montant", "Montant compta", 140, "e"),
                ("reference", "Réf. immo (éditable)", 170, "w"), ("nb", "Nb comp.", 70, "center"),
                ("tot", "Total composantes", 150, "e"), ("ecart", "Écart", 110, "e")]
        self.m_grid = Grid(top, cols, editable=("reference",), on_edit=self.edit_immo, height=8)
        self.m_grid.pack(fill="both", expand=True)
        self.m_grid.tree.bind("<<TreeviewSelect>>", lambda _: self.refresh_comps())

        bot = ttk.LabelFrame(t, text="Fichier d'immobilisation : composantes et sous-références (double-clic pour saisir)", padding=4)
        bot.pack(fill="both", expand=True, pady=(6, 0))
        bar = ttk.Frame(bot)
        bar.pack(fill="x")
        for txt, cmd in (("+ Composante", self.comp_add), ("Dupliquer", self.comp_dup),
                         ("Éclater en unités (sous-réf.)", self.comp_split),
                         ("Ajuster valeur au montant comptable", self.comp_adjust), ("Supprimer", self.comp_del)):
            ttk.Button(bar, text=txt, command=cmd).pack(side="left", padx=2)
        cols = [("ref", "Réf. composante", 150, "w"), ("sous_ref", "Sous-référence", 140, "w"),
                ("designation", "Désignation", 280, "w"), ("quantite", "Quantité", 70, "e"),
                ("pu", "Valeur unitaire", 110, "e"), ("valeur", "Valeur", 110, "e"), ("duree", "Durée (ans)", 80, "e"),
                ("date_service", "Mise en service", 100, "center"), ("localisation", "Localisation", 160, "w")]
        self.c_grid = Grid(bot, cols, editable=("ref", "sous_ref", "designation", "quantite", "pu", "duree",
                                                "date_service", "localisation"), on_edit=self.edit_comp, height=8)
        self.c_grid.pack(fill="both", expand=True)

    def _immo_id(self):
        s = self.m_grid.tree.selection()
        return int(s[0]) if s else None

    def edit_immo(self, iid, key, val, sel):
        self.db.c.execute("UPDATE immos SET reference=? WHERE id=?", (val.strip(), int(iid)))
        self.db.c.commit()
        self.refresh_immos()

    def refresh_immos(self):
        q = self.m_filter.get().lower().strip()
        rows = []
        for r in self.db.c.execute("SELECT * FROM immos ORDER BY compte, date, id"):
            comps = self.db.comps(r["id"])
            tot = sum((c["quantite"] or 0) * (c["valeur_unitaire"] or 0) for c in comps)
            ecart = round(r["montant"] - tot, 2)
            vals = (r["compte"], dfmt(r["date"]), r["piece"], r["libelle"], fmt(r["montant"]), r["reference"],
                    len(comps), fmt(tot), fmt(ecart))
            if q and q not in " ".join(map(str, vals)).lower():
                continue
            rows.append((str(r["id"]), vals, () if (abs(ecart) < 0.01 and comps) else ("warn",)))
        self.m_grid.reload(rows)
        self.refresh_comps()

    def refresh_comps(self):
        iid = self._immo_id()
        rows = []
        if iid:
            for c in self.db.comps(iid):
                v = (c["quantite"] or 0) * (c["valeur_unitaire"] or 0)
                rows.append((str(c["id"]), (c["ref"], c["sous_ref"], c["designation"], fmtq(c["quantite"]),
                                            fmt(c["valeur_unitaire"]), fmt(v), fmtq(c["duree"]),
                                            dfmt(c["date_service"]), c["localisation"]), ()))
        self.c_grid.reload(rows)

    def edit_comp(self, iid, key, val, sel):
        cid = int(iid)
        if key in ("quantite", "pu", "duree"):
            self.db.set_comp(cid, {"pu": "valeur_unitaire"}.get(key, key), to_num(val))
        elif key == "date_service":
            self.db.set_comp(cid, key, parse_date(val) or val.strip())
        else:
            self.db.set_comp(cid, key, val.strip())
        self.refresh_immos()

    def _need_immo(self):
        i = self._immo_id()
        if not i:
            messagebox.showinfo("Immobilisations", "Sélectionnez d'abord une acquisition.")
        return i

    def comp_add(self):
        i = self._need_immo()
        if i:
            self.db.add_comp(i)
            self.refresh_immos()

    def _need_comp(self):
        s = self.c_grid.tree.selection()
        if not s:
            messagebox.showinfo("Composantes", "Sélectionnez une composante.")
            return None
        return [int(x) for x in s]

    def comp_dup(self):
        ids = self._need_comp()
        if ids:
            for i in ids:
                self.db.dup_comp(i)
            self.refresh_immos()

    def comp_split(self):
        ids = self._need_comp()
        if ids:
            n = self.db.split_comp(ids[0])
            if n == 0:
                messagebox.showinfo("Éclatement", "La quantité doit être un entier supérieur ou égal à 2.")
            self.refresh_immos()

    def comp_adjust(self):
        ids = self._need_comp()
        if ids:
            self.db.adjust_comp(ids[0])
            self.refresh_immos()

    def comp_del(self):
        ids = self._need_comp()
        if ids and messagebox.askyesno("Supprimer", f"Supprimer {len(ids)} composante(s) ?"):
            self.db.del_comps(ids)
            self.refresh_immos()

    # ---- 5. Registre
    def _tab_reg(self):
        t = self._tab("6. Registre des immobilisations")
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        self.r_mode = tk.StringVar(value="compte")
        ttk.Radiobutton(bar, text="Par compte comptable", variable=self.r_mode, value="compte",
                        command=self._reg_values).pack(side="left")
        ttk.Radiobutton(bar, text="Par composante", variable=self.r_mode, value="composante",
                        command=self._reg_values).pack(side="left", padx=8)
        ttk.Label(bar, text="Sélection :").pack(side="left")
        self.r_value = ttk.Combobox(bar, width=24, state="readonly")
        self.r_value.pack(side="left", padx=4)
        self.r_value.bind("<<ComboboxSelected>>", lambda _: self.refresh_reg())
        ttk.Label(bar, text="Situation au (jj/mm/aaaa) :").pack(side="left", padx=(12, 2))
        self.r_date = tk.StringVar(value=f"31/12/{datetime.date.today().year}")
        ed = ttk.Entry(bar, textvariable=self.r_date, width=11)
        ed.pack(side="left")
        ed.bind("<Return>", lambda _: self.refresh_reg())
        bar2 = ttk.Frame(t)
        bar2.pack(fill="x", pady=4)
        ttk.Button(bar2, text="Actualiser", command=self.refresh_reg).pack(side="left")
        ttk.Button(bar2, text="Imprimer / PDF", command=lambda: self.reg_pdf(False)).pack(side="left", padx=6)
        ttk.Button(bar2, text="Envoyer à l'imprimante", command=lambda: self.reg_pdf(True)).pack(side="left")
        ttk.Button(bar2, text="Export Excel", command=self.reg_xlsx).pack(side="left", padx=6)
        self.r_info = ttk.Label(bar2, text="")
        self.r_info.pack(side="left", padx=10)
        cols = [(k, h, int(w * 52), "e" if k in core.REG_NUM else "w") for h, k, w in core.REG_COLS]
        self.r_grid = Grid(t, cols)
        self.r_grid.pack(fill="both", expand=True)
        self._reg_values()

    def _reg_values(self):
        comptes, refs = self.db.registre_values()
        vals = ["(Tous)"] + (comptes if self.r_mode.get() == "compte" else refs)
        self.r_value["values"] = vals
        self.r_value.set("(Tous)")
        self.refresh_reg()

    def _reg_params(self):
        d = parse_date(self.r_date.get())
        cloture = datetime.date.fromisoformat(d) if d else datetime.date.today()
        v = self.r_value.get()
        return self.r_mode.get(), ("" if v in ("", "(Tous)") else v), cloture

    def refresh_reg(self):
        mode, value, cl = self._reg_params()
        rows = self.db.registre(mode, value, cl)
        out = []
        for i, r in enumerate(rows):
            vals = []
            for _, k, _ in core.REG_COLS:
                v = r[k]
                if k in ("date_acq", "date_service"):
                    v = dfmt(v)
                elif k in ("quantite", "duree"):
                    v = fmtq(v)
                elif k in core.REG_NUM:
                    v = fmt(v)
                vals.append(v)
            out.append((str(i), tuple(vals), ()))
        self.r_grid.reload(out)
        self.r_info.config(text=f"{len(rows)} composantes | Valeur brute {fmt(sum(r['valeur'] for r in rows))} | "
                                f"VNC {fmt(sum(r['vnc'] for r in rows))}")

    def reg_pdf(self, to_printer):
        mode, value, cl = self._reg_params()
        p = os.path.join(tempfile.gettempdir(), "registre_immobilisations.pdf")
        if not to_printer:
            p = self._save(".pdf", "registre_immobilisations.pdf", [("PDF", "*.pdf")])
            if not p:
                return
        core.export_registre_pdf(self.db, p, mode, value, cl)
        if to_printer:
            try:
                os.startfile(p, "print")  # Windows uniquement
            except Exception:
                core.open_file(p)
        else:
            core.open_file(p)

    def reg_xlsx(self):
        mode, value, cl = self._reg_params()
        p = self._save(".xlsx", "registre_immobilisations.xlsx", [("Excel", "*.xlsx")])
        if p:
            core.export_registre_xlsx(self.db, p, mode, value, cl)
            core.open_file(p)

    # ------------------------------------------------------------------
    def refresh_all(self):
        self.refresh_journal()
        self.refresh_lines()
        self.refresh_sorties()
        self.refresh_inv()
        self.refresh_immos()
        self._reg_values()


if __name__ == "__main__":
    App().mainloop()
