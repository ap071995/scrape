"""Paged worksheet editor. All edits use the same durable workbook writer."""
from tkinter import ttk, messagebox, simpledialog
import customtkinter as ctk


class WorksheetEditor(ctk.CTkToplevel):
    PAGE_SIZE = 100

    def __init__(self, app, workbook):
        super().__init__(app)
        self.book = workbook
        self.page = 0
        self.title('Scraper worksheet · ' + workbook.output_path.name)
        self.geometry(f'{min(1150, self.winfo_screenwidth() - 80)}x580')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(self, text='Your output worksheet', font=ctk.CTkFont(size=22, weight='bold')).grid(
            row=0, column=0, sticky='w', padx=18, pady=(14, 0))
        bar = ctk.CTkFrame(self)
        bar.grid(row=1, column=0, sticky='ew', padx=14, pady=10)
        for text, callback in [('Add column', self.add_column), ('Add rows', self.add_rows),
                               ('Edit selected row', self.edit_row), ('Save', self.save)]:
            ctk.CTkButton(bar, text=text, width=130, command=callback).pack(side='left', padx=5, pady=8)
        frame = ctk.CTkFrame(self)
        frame.grid(row=2, column=0, sticky='nsew', padx=14)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        self.table = ttk.Treeview(frame, show='headings', selectmode='browse')
        self.table.grid(row=0, column=0, sticky='nsew')
        for orient in ('vertical', 'horizontal'):
            scroll = ttk.Scrollbar(frame, orient=orient,
                command=self.table.yview if orient == 'vertical' else self.table.xview)
            scroll.grid(row=0 if orient == 'vertical' else 1,
                        column=1 if orient == 'vertical' else 0,
                        sticky='ns' if orient == 'vertical' else 'ew')
            self.table.configure(**{'yscrollcommand' if orient == 'vertical' else 'xscrollcommand': scroll.set})
        self.table.bind('<Double-1>', lambda event: self.edit_row())
        footer = ctk.CTkFrame(self, fg_color='transparent')
        footer.grid(row=3, column=0, sticky='ew', padx=14, pady=10)
        self.summary = ctk.CTkLabel(footer, text='')
        self.summary.pack(side='left')
        ctk.CTkButton(footer, text='Previous', width=90, command=lambda: self.turn(-1)).pack(side='right', padx=5)
        ctk.CTkButton(footer, text='Next', width=90, command=lambda: self.turn(1)).pack(side='right')
        ctk.CTkLabel(self, text='Custom columns are for your own values. Manual rows are saved here; they are not automatically scraped.').grid(
            row=4, column=0, sticky='w', padx=18, pady=(0, 12))
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.refresh()

    def refresh(self):
        self.table.delete(*self.table.get_children())
        # Stable IDs allow punctuation in user column names.
        columns = ['row'] + [f'c{i}' for i in range(len(self.book.headers))]
        self.table.configure(columns=columns)
        for key, name in zip(columns, ['Excel row'] + self.book.headers):
            self.table.heading(key, text=name)
            self.table.column(key, width=85 if key == 'row' else 180, stretch=False)
        start = self.page * self.PAGE_SIZE + 2
        end = min(start + self.PAGE_SIZE - 1, self.book.ws.max_row)
        for row in range(start, end + 1):
            values = [row] + [self.book.ws.cell(row, col).value or ''
                              for col in range(1, len(self.book.headers) + 1)]
            self.table.insert('', 'end', iid=str(row), values=values)
        self.summary.configure(text=f'{max(0, self.book.ws.max_row - 1):,} rows · {len(self.book.headers)} columns · Page {self.page + 1}')

    def turn(self, delta):
        self.page = min(max(0, self.page + delta), max(0, (self.book.ws.max_row - 2) // self.PAGE_SIZE))
        self.refresh()

    def save(self):
        try:
            self.book.save(force=True)
            self.refresh()
            return True
        except Exception as exc:
            messagebox.showerror('Save failed', str(exc) + '\nClose Excel, then press Save again.', parent=self)
            return False

    def add_column(self):
        name = simpledialog.askstring('Add column', 'New column name:', parent=self)
        if name is None:
            return
        try:
            self.book.add_columns([name.strip()])
            self.save()
            self.refresh()
            self.table.xview_moveto(1)
        except Exception as exc:
            messagebox.showerror('Column not added', str(exc), parent=self)

    def add_rows(self):
        count = simpledialog.askinteger('Add rows', 'Number of manual rows (1–1,000):',
                                       initialvalue=1, minvalue=1, maxvalue=1000, parent=self)
        if count is None:
            return
        try:
            rows = self.book.add_manual_rows(count)
            self.page = (rows[0] - 2) // self.PAGE_SIZE
            self.refresh()
            self.table.selection_set(str(rows[0]))
            if count == 1:
                self.edit_row()
        except Exception as exc:
            messagebox.showerror('Rows could not be saved', str(exc), parent=self)
            self.refresh()

    def edit_row(self):
        selection = self.table.selection()
        if not selection:
            return
        row = int(selection[0])
        dialog = ctk.CTkToplevel(self)
        dialog.title(f'Edit Excel row {row}')
        dialog.geometry('650x540')
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(0, weight=1)
        scroll = ctk.CTkScrollableFrame(dialog)
        scroll.grid(row=0, column=0, sticky='nsew', padx=12, pady=12)
        scroll.grid_columnconfigure(1, weight=1)
        entries = {}
        for index, header in enumerate(self.book.headers):
            ctk.CTkLabel(scroll, text=header, width=180, anchor='w').grid(row=index, column=0, padx=6, pady=4)
            entry = ctk.CTkEntry(scroll)
            entry.grid(row=index, column=1, sticky='ew', padx=6, pady=4)
            entry.insert(0, str(self.book.ws.cell(row, index + 1).value or ''))
            entries[header] = entry
        def submit():
            try:
                self.book.update_record(row, {key: entry.get() for key, entry in entries.items()})
                if self.save():
                    dialog.destroy()
            except Exception as exc:
                messagebox.showerror('Row not saved', str(exc), parent=dialog)
        ctk.CTkButton(dialog, text='Save row', command=submit).grid(row=1, column=0, pady=(0, 12))
        dialog.transient(self)
        dialog.grab_set()

    def close(self):
        if self.save():
            self.destroy()
