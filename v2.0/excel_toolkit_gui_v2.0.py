import os
import sys
import ast
import customtkinter as ctk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from processors import cleaner, merger, grouper, stats
import pandas as pd

# ---------------- ICON HANDLING ----------------
def get_icon_path():
    if getattr(sys, 'frozen', False):  # Running as exe
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, "myicon.ico")

# ---------------- TRANSLATIONS ----------------
translations = {
    "en": {
        "title": "Excel Toolkit v2.0",
        "action": "Select Action:",
        "clean": "Clean",
        "merge": "Merge",
        "group": "Group & Aggregate",
        "Preview Data": "Preview Data",
        "stats": "Statistics",
        "file1": "File 1:",
        "file2": "File 2 (for Merge):",
        "col1": "Column for Merge / Group / Statistics / Clean:",
        "col2": "Column for Aggregation (Group):",
        "agg_func": "Aggregation Function:",
        "duplicates": "Remove duplicates",
        "trim": "Trim spaces",
        "case": "Case conversion:",
        "uppercase": "Uppercase",
        "lowercase": "Lowercase",
        "capitalize": "Capitalize",
        "lang": "Switch to Greek",
        "run": "Run",
        "about": "About",
        "about_text": "Excel Toolkit v2.0\nAuthor: Kostasp84: gumroad.com",
        "success": "Action completed successfully!",
        "help": "Help",
        "help_text": (
            "Excel Toolkit v2.0 - Help\n\n"
            "1. Choose an action (Clean, Merge, Group, Stats).\n"
            "2. Select File 1 (and File 2 for Merge).\n"
            "3. For Clean:\n"
            "   - Use 'Column' to target specific column(s).\n"
            "   - Case: upper/lower/capitalize.\n"
            "   - Options: remove duplicates, trim spaces.\n"
            "4. For Group: choose grouping and aggregation column.\n"
            "5. Stats: leave column empty for all, or specify one.\n"
            "6. Save As: choose .xlsx or .csv output.\n\n"
            "Tip: For multiple columns, type: Name,City,Country"),
    },
    "el": {
        "title": "Excel Toolkit v2.0",
        "action": "Επιλογή ενέργειας:",
        "clean": "Καθαρισμός",
        "merge": "Συγχώνευση",
        "group": "Ομαδοποίηση & Άθροιση",
        "stats": "Στατιστικά",
        "Preview Data": "Προεπισκόπηση Δεδομένων",
        "file1": "Αρχείο 1:",
        "file2": "Αρχείο 2 (για Συγχώνευση):",
        "col1": "Στήλη για Merge / Group / Statistics / Clean:",
        "col2": "Στήλη για Aggregation (Group):",
        "agg_func": "Συνάρτηση Συγκέντρωσης:",
        "duplicates": "Αφαίρεση διπλότυπων",
        "trim": "Αφαίρεση κενών",
        "case": "Μετατροπή πεζών/κεφαλαίων:",
        "uppercase": "Κεφαλαία",
        "lowercase": "Πεζά",
        "capitalize": "Πρώτο Κεφαλαίο",
        "lang": "Μετάβαση στα Αγγλικά",
        "run": "Εκτέλεση",
        "about": "Σχετικά",
        "about_text": "Excel Toolkit v2.0\nΣυγγραφέας: Kostasp84\nΠερισσότερα: gumroad.com",
        "success": "Η ενέργεια ολοκληρώθηκε!",
        "help": "Βοήθεια",
        "help_text": (
            "Excel Toolkit v2.0 - Οδηγίες\n\n"
            "1. Διάλεξε ενέργεια (Καθαρισμός, Συγχώνευση, Ομαδοποίηση, Στατιστικά).\n"
            "2. Επέλεξε Αρχείο 1 (και Αρχείο 2 για Συγχώνευση).\n"
            "3. Για Καθαρισμό:\n"
            "   - Στήλη: στοχευμένες στήλες για αλλαγές.\n"
            "   - Μετατροπή: κεφαλαία/πεζά/πρώτο κεφαλαίο.\n"
            "   - Επιλογές: αφαίρεση διπλότυπων, αφαίρεση κενών.\n"
            "4. Για Ομαδοποίηση: διάλεξε στήλη ομαδοποίησης & στήλη αθροίσματος.\n"
            "5. Στατιστικά: άφησε κενό για όλες τις στήλες ή γράψε μία.\n"
            "6. Αποθήκευση: διάλεξε .xlsx ή .csv.\n\n"
            "Συμβουλή: Για πολλές στήλες γράψε: Όνομα,Πόλη,Χώρα"
        ),
    }
}

# ---------------- MAIN CLASS ----------------
class ExcelToolkitGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.lang = "en"
        self.trans = translations[self.lang]

        self.title(self.trans["title"])
        self.geometry("600x800")

        try:
            self.iconbitmap(get_icon_path())
        except Exception:
            pass

        ctk.set_appearance_mode("system")  # auto dark/light
        ctk.set_default_color_theme("blue")

        self.bind_shortcuts()
        self.create_widgets()

    def bind_shortcuts(self):
        # Bind shortcuts for Entry, OptionMenu, CheckBox, Button
        def bind_widget_shortcuts(widget):
            # Entry widgets
            if isinstance(widget, ctk.CTkEntry):
                widget.bind('<Control-c>', lambda e: widget.event_generate('<<Copy>>'))
                widget.bind('<Control-x>', lambda e: widget.event_generate('<<Cut>>'))
                widget.bind('<Control-v>', lambda e: widget.event_generate('<<Paste>>'))
            # OptionMenu widgets
            elif isinstance(widget, ctk.CTkOptionMenu):
                widget.bind('<Control-c>', lambda e: self.clipboard_append(widget.get()))
                widget.bind('<Control-v>', lambda e: widget.set(self.clipboard_get()))
            # CheckBox widgets
            elif isinstance(widget, ctk.CTkCheckBox):
                widget.bind('<Control-c>', lambda e: self.clipboard_append(str(widget.get())))
                widget.bind('<Control-v>', lambda e: widget.select() if self.clipboard_get() == '1' else widget.deselect())
            # Button widgets
            elif isinstance(widget, ctk.CTkButton):
                widget.bind('<Control-c>', lambda e: self.clipboard_append(widget.cget('text')))
                widget.bind('<Control-v>', lambda e: widget.invoke())
            
            # Label widgets
            elif isinstance(widget, ctk.CTkLabel):
                widget.bind('<Control-c>', lambda e: self.clipboard_append(widget.cget('text')))
                widget.bind('<Control-v>', lambda e: widget.event_generate('<<Paste>>'))

        self._widget_shortcut_binder = bind_widget_shortcuts

    def bind_tree_shortcuts(self, tree):
        # Enable Ctrl+C for copying selected row(s) in Treeview
        def copy_selected(event=None):
            selected = tree.selection()
            if not selected:
                return
            rows = []
            for item in selected:
                rows.append('\t'.join([str(tree.set(item, col)) for col in tree["columns"]]))
            self.clipboard_clear()
            self.clipboard_append('\n'.join(rows))
        tree.bind('<Control-c>', copy_selected)

    def switch_lang(self):
        self.lang = "el" if self.lang == "en" else "en"
        self.trans = translations[self.lang]
        self.clear_widgets()
        self.create_widgets()

    def clear_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

    def create_widgets(self):
        # Action selection
        action_label = ctk.CTkLabel(self, text=self.trans["action"])
        action_label.pack(pady=5)
        self._widget_shortcut_binder(action_label)
        self.action_var = ctk.StringVar(value="clean")
        for action in ["clean", "merge", "group", "stats","Preview Data"]:
            rb = ctk.CTkRadioButton(
                self, text=self.trans[action], variable=self.action_var, value=action
            )
            rb.pack(anchor="w")
            self._widget_shortcut_binder(rb)

        # Files
        file1_label = ctk.CTkLabel(self, text=self.trans["file1"])
        file1_label.pack(anchor="w")
        self._widget_shortcut_binder(file1_label)
        self.file1_entry = ctk.CTkEntry(self, width=400)
        self.file1_entry.pack()
        self._widget_shortcut_binder(self.file1_entry)
        file1_btn = ctk.CTkButton(self, text="Browse", command=self.browse_file1)
        file1_btn.pack()
        self._widget_shortcut_binder(file1_btn)

        file2_label = ctk.CTkLabel(self, text=self.trans["file2"])
        file2_label.pack(anchor="w")
        self._widget_shortcut_binder(file2_label)
        self.file2_entry = ctk.CTkEntry(self, width=400)
        self.file2_entry.pack()
        self._widget_shortcut_binder(self.file2_entry)
        file2_btn = ctk.CTkButton(self, text="Browse", command=self.browse_file2)
        file2_btn.pack()
        self._widget_shortcut_binder(file2_btn)

        # Columns
        col1_label = ctk.CTkLabel(self, text=self.trans["col1"])
        col1_label.pack(anchor="w")
        self._widget_shortcut_binder(col1_label)
        self.col1_entry = ctk.CTkEntry(self, width=400)
        self.col1_entry.pack()
        self._widget_shortcut_binder(self.col1_entry)

        col2_label = ctk.CTkLabel(self, text=self.trans["col2"])
        col2_label.pack(anchor="w")
        self._widget_shortcut_binder(col2_label)
        self.col2_entry = ctk.CTkEntry(self, width=400)
        self.col2_entry.pack()
        self._widget_shortcut_binder(self.col2_entry)

        # Aggregation dropdown
        agg_label = ctk.CTkLabel(self, text=self.trans["agg_func"])
        agg_label.pack(anchor="w")
        self._widget_shortcut_binder(agg_label)
        self.agg_var = ctk.StringVar(value="sum")
        agg_dropdown = ctk.CTkOptionMenu(self, variable=self.agg_var, values=["sum", "mean", "count", "max", "min"])
        agg_dropdown.pack()
        self._widget_shortcut_binder(agg_dropdown)

        # Cleaner options
        self.dup_var = ctk.BooleanVar(value=False)
        dup_cb = ctk.CTkCheckBox(self, text=self.trans["duplicates"], variable=self.dup_var)
        dup_cb.pack(anchor="w")
        self._widget_shortcut_binder(dup_cb)

        self.trim_var = ctk.BooleanVar(value=False)
        trim_cb = ctk.CTkCheckBox(self, text=self.trans["trim"], variable=self.trim_var)
        trim_cb.pack(anchor="w")
        self._widget_shortcut_binder(trim_cb)

        case_label = ctk.CTkLabel(self, text=self.trans["case"])
        case_label.pack(anchor="w")
        self._widget_shortcut_binder(case_label)
        self.case_var = ctk.StringVar(value="")
        case_menu = ctk.CTkOptionMenu(
            self, variable=self.case_var, values=["", "upper", "lower", "capitalize"]
        )
        case_menu.pack()
        self._widget_shortcut_binder(case_menu)

        # Run + Lang + About
        run_btn = ctk.CTkButton(self, text=self.trans["run"], command=self.run_action)
        run_btn.pack(pady=10)
        self._widget_shortcut_binder(run_btn)
        lang_btn = ctk.CTkButton(self, text=self.trans["lang"], command=self.switch_lang)
        lang_btn.pack(pady=10)
        self._widget_shortcut_binder(lang_btn)
        about_btn = ctk.CTkButton(self, text=self.trans["about"], command=self.show_about)
        about_btn.pack()
        self._widget_shortcut_binder(about_btn)
        help_btn = ctk.CTkButton(self, text=self.trans["help"], command=self.show_help)
        help_btn.pack(pady=5)
        self._widget_shortcut_binder(help_btn)

    def show_help(self):
        messagebox.showinfo(self.trans["help"], self.trans["help_text"])

    def browse_file1(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.csv")])
        if filename:
            self.file1_entry.delete(0, "end")
            self.file1_entry.insert(0, filename)

    def browse_file2(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.csv")])
        if filename:
            self.file2_entry.delete(0, "end")
            self.file2_entry.insert(0, filename)

    def show_about(self):
        messagebox.showinfo(self.trans["about"], self.trans["about_text"])

    def parse_target_columns(self, text: str):
        """Μετατροπή input string σε λίστα στηλών ή None"""
        if not text or text.strip() == "":
            return None
        s = text.strip()
        try:
            if s.startswith("[") or s.startswith("("):
                vals = ast.literal_eval(s)
                if isinstance(vals, (list, tuple, set)):
                    return [str(x).strip() for x in vals if str(x).strip()]
            if "," in s:
                return [c.strip() for c in s.split(",") if c.strip()]
            return s  # single col string
        except Exception:
            return [c.strip() for c in s.split(",") if c.strip()]

    def run_action(self):
        action = self.action_var.get()
        file1 = self.file1_entry.get().strip()
        file2 = self.file2_entry.get().strip()
        col1 = self.col1_entry.get().strip()
        col2 = self.col2_entry.get().strip()

        if not file1:
            messagebox.showerror("Error", "Πρέπει να δώσεις Αρχείο 1.")
            return

        # Save As dialog
        
        if action == "Preview Data":
            output_path = None  # No output needed
        else:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
            )
        if not output_path and action != "Preview Data":
            return
        


        try:
            if action == "clean":
                case_val = None if self.case_var.get() == "" else self.case_var.get()
                target_cols = self.parse_target_columns(col1)
                cleaner.clean_file(
                    file1,
                    output_path,
                    case_option=case_val,
                    drop_duplicates=self.dup_var.get(),
                    trim_spaces=self.trim_var.get(),
                    target_columns=target_cols
                )
            elif action == "merge":
                if not file2 or not col1:
                    messagebox.showerror("Error", "Missing file2 or merge column")
                    return
                merger.merge_files(file1, file2, col1, output_path)
            elif action == "group":
                if not col1 or not col2:
                    messagebox.showerror("Error", "Missing group/aggregation columns")
                    return
                grouper.group_file(file1, col1, col2, self.agg_var.get(), output_path)
            elif action == "stats":
                stats.generate_stats(file1, output_path, column=col1 if col1 else None)
                
                if messagebox.askyesno("Export PDF", "Θέλεις να κάνεις export report σε PDF;"):
                    pdf_path = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("PDF Files", "*.pdf")]
                    )
                    if pdf_path:
                        stats.export_pdf(file1, pdf_path)
            elif action == "Preview Data":
                df = pd.read_excel(file1) if file1.endswith('.xlsx') else pd.read_csv(file1)
                preview_window = ctk.CTkToplevel(self)
                preview_window.title("Data Preview")
                preview_window.geometry("900x600")

                frame = ctk.CTkFrame(preview_window)
                frame.pack(fill="both", expand=True)

                tree = ttk.Treeview(frame, show="headings")
                tree.pack(fill="both", expand=True, side="left")
                self.bind_tree_shortcuts(tree)

                # Add scrollbars
                vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
                vsb.pack(side="right", fill="y")
                hsb = ttk.Scrollbar(preview_window, orient="horizontal", command=tree.xview)
                hsb.pack(side="bottom", fill="x")
                tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

                # Set columns
                cols = list(df.columns)
                tree["columns"] = cols
                for col in cols:
                    tree.heading(col, text=col)
                    tree.column(col, width=120, anchor="center")

                # Insert rows (limit to 200) with tags for alternating row borders
                for i, (_, row) in enumerate(df.head(200).iterrows()):
                    tag = 'row_even' if i % 2 == 0 else 'row_odd'
                    tree.insert("", "end", values=list(row), tags=(tag,))

                # Add style for grid lines and row outlines
                style = ttk.Style()
                style.theme_use("default")
                style.configure("Treeview", rowheight=24)
                style.configure("Treeview.Heading", font=(None, 10, "bold"))
                style.map("Treeview", background=[('selected', '#0078d7')])
                style.layout("Treeview", [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ])
                style.configure("Treeview", bordercolor="#cccccc", borderwidth=1)
                style.configure("Treeview", highlightthickness=1)
                style.configure("Treeview", relief="solid")

                # Custom row tag styles for outlines
                style.configure("row_even.Treeview", background="#f8f8f8", borderwidth=1, relief="solid")
                style.configure("row_odd.Treeview", background="#e0e0e0", borderwidth=1, relief="solid")
                tree.tag_configure('row_even', background="#f8f8f8")
                tree.tag_configure('row_odd', background="#e0e0e0")

                return  # No success message for preview
            

            messagebox.showinfo("Success", self.trans["success"])
        except Exception as e:
            messagebox.showerror("Error", str(e))

# ---------------- RUN ----------------
if __name__ == "__main__":
    app = ExcelToolkitGUI()
    app.mainloop()
