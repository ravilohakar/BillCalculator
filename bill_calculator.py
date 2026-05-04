"""
Bill Calculator Desktop Application
Reads item data from items_data.xlsx and provides a billing interface.
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import openpyxl
import os
import sys
import re
import tempfile
from datetime import datetime


class BillCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bill Calculator")
        self.root.geometry("960x660+100+0")
        self.root.minsize(900, 600)
        self.root.configure(bg="#f5f5f5")

        self.items_data = []  # List of dicts from Excel
        self.bill_items = []  # Items added to the bill
        self.filtered_items = []  # Search-filtered items
        self.box_size_map = {}  # Size -> number of bottles per box

        self.load_data()
        self.build_ui()

    # ── Data Loading ──────────────────────────────────────────────

    @staticmethod
    def _size_num(size_str):
        match = re.search(r'[\d.]+', size_str)
        return float(match.group()) if match else 0.0

    def load_data(self):
        data_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "items_data.xlsx")
        if not os.path.exists(data_file):
            messagebox.showerror("Error", f"Data file not found:\n{data_file}\n\nRun create_sample_data.py first.")
            sys.exit(1)

        wb = openpyxl.load_workbook(data_file, read_only=True, data_only=True)
        ws = wb["Items"]
        rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header

        # Load BoxSize sheet (size -> number of bottles per box)
        if "BoxSize" in wb.sheetnames:
            bs_ws = wb["BoxSize"]
            for bs_row in bs_ws.iter_rows(min_row=2, values_only=True):
                if bs_row[0] is not None:
                    self.box_size_map[str(bs_row[0]).strip().lower()] = int(bs_row[1])

        wb.close()

        for row in rows:
            if row[0] is None:
                continue
            self.items_data.append({
                "name": str(row[0]).strip(),
                "size": str(row[1]).strip() if row[1] else "",
                "unit_price": float(row[2]) if row[2] else 0.0,
                "discount": float(row[3]) if row[3] else 0.0,
            })
        self.items_data.sort(key=lambda x: self._size_num(x["size"]), reverse=True)
        self.items_data.sort(key=lambda x: x["name"].lower())
        self.filtered_items = list(self.items_data)

    # ── UI Construction ───────────────────────────────────────────

    def build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), background="#2c3e50", foreground="white")
        style.configure("TFrame", background="#f5f5f5")
        style.configure("Card.TFrame", background="white", relief="solid")
        style.configure("TLabel", background="#f5f5f5", font=("Segoe UI", 10))
        style.configure("White.TLabel", background="white", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Total.TLabel", font=("Segoe UI", 14, "bold"), background="#f5f5f5", foreground="#2c3e50")
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Date Display (top-right corner) ──
        date_frame = ttk.Frame(main)
        date_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(date_frame, text="New Viransh Wine Shop", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        self.date_label_frame = ttk.Frame(date_frame)
        self.date_label_frame.pack(side=tk.RIGHT)
        ttk.Label(self.date_label_frame, text="Date: ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        ttk.Label(self.date_label_frame, textvariable=self.date_var, font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # ── Bill No + Bill To ──
        bill_to_frame = ttk.Frame(main)
        bill_to_frame.pack(fill=tk.X, pady=(0, 5))
        self.bill_no_var = tk.StringVar(value="Bill No: --")
        ttk.Label(bill_to_frame, textvariable=self.bill_no_var, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Separator(bill_to_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        ttk.Label(bill_to_frame, text="Bill To:", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 5))
        self.bill_to_var = tk.StringVar()
        ttk.Entry(bill_to_frame, textvariable=self.bill_to_var, width=40, font=("Segoe UI", 10)).pack(side=tk.LEFT)

        # Separator between Bill To and Search
        ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(5, 5))

        # ── Top Section: Search + Item List + Controls ──
        top = ttk.Frame(main)
        top.pack(fill=tk.X, pady=(0, 10))

        # Search
        search_frame = ttk.Frame(top)
        search_frame.pack(fill=tk.X)
        ttk.Label(search_frame, text="Search Item:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40, font=("Segoe UI", 10))
        search_entry.pack(side=tk.LEFT, padx=(0, 10))

        add_btn = ttk.Button(search_frame, text="  Add Items to Bill  ", command=self.add_to_bill)
        add_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Item list (scrollable grid with inline qty inputs)
        item_frame = ttk.LabelFrame(main, text="List of Items", padding=5)
        item_frame.pack(fill=tk.X, pady=(0, 10))

        # Header row using grid for alignment
        header_frame = tk.Frame(item_frame, bg="#ecf0f1")
        header_frame.pack(fill=tk.X)
        col_widths = [350, 100, 90, 100, 100, 100]
        headers = [("Item Name", "w"), ("Size", "center"), ("Unit Price", "e"),
                   ("Discount", "center"), ("Qty Box", "center"), ("Qty Bottle", "center")]
        for col, ((text, anc), minw) in enumerate(zip(headers, col_widths)):
            header_frame.columnconfigure(col, minsize=minw, weight=0)
            tk.Label(header_frame, text=text, font=("Segoe UI", 9, "bold"), bg="#ecf0f1",
                     anchor=anc).grid(row=0, column=col, padx=4, sticky="ew")

        # Scrollable canvas for item rows
        canvas_frame = tk.Frame(item_frame, height=120)
        canvas_frame.pack(fill=tk.X, expand=False)
        canvas_frame.pack_propagate(False)
        self.item_canvas = tk.Canvas(canvas_frame, highlightthickness=0, bg="white")
        item_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.item_canvas.yview)
        self.item_canvas.configure(yscrollcommand=item_scroll.set)
        self.item_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        item_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.item_inner_frame = tk.Frame(self.item_canvas, bg="white")
        self.item_canvas.create_window((0, 0), window=self.item_inner_frame, anchor="nw")
        self.item_inner_frame.bind("<Configure>", lambda e: self.item_canvas.configure(scrollregion=self.item_canvas.bbox("all")))
        self.item_canvas.bind_all("<MouseWheel>", lambda e: self.item_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        self.item_rows = []  # list of (frame, qty_var, loose_var, item_index)
        self._populate_item_grid()

        # ── Bill Grid ──
        bill_frame = ttk.LabelFrame(main, text="Bill Items", padding=5)
        bill_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        cols_bill = ("sno", "name", "size", "unit_price", "discount", "qty", "loose", "total")
        self.bill_tree = ttk.Treeview(bill_frame, columns=cols_bill, show="headings", height=4, selectmode="browse")
        self.bill_tree.heading("sno", text="#")
        self.bill_tree.heading("name", text="Item Name")
        self.bill_tree.heading("size", text="Size")
        self.bill_tree.heading("unit_price", text="Unit Price (₹)")
        self.bill_tree.heading("discount", text="Discount (₹)")
        self.bill_tree.heading("qty", text="Qty Box")
        self.bill_tree.heading("loose", text="Qty Bottle")
        self.bill_tree.heading("total", text="Total (₹)")
        self.bill_tree.column("sno", width=20, anchor="center")
        self.bill_tree.column("name", width=210, anchor="w")
        self.bill_tree.column("size", width=80, anchor="center")
        self.bill_tree.column("unit_price", width=100, anchor="e")
        self.bill_tree.column("discount", width=90, anchor="e")
        self.bill_tree.column("qty", width=70, anchor="center")
        self.bill_tree.column("loose", width=70, anchor="center")
        self.bill_tree.column("total", width=100, anchor="e")

        bill_scroll = ttk.Scrollbar(bill_frame, orient=tk.VERTICAL, command=self.bill_tree.yview)
        self.bill_tree.configure(yscrollcommand=bill_scroll.set)
        self.bill_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bill_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # ── Bottom: Remove button + Total ──
        bottom = ttk.Frame(main)
        bottom.pack(fill=tk.X, pady=(5, 0))

        remove_btn = ttk.Button(bottom, text="  Remove Selected  ", command=self.remove_item)
        remove_btn.pack(side=tk.LEFT)

        clear_btn = ttk.Button(bottom, text="  Clear All  ", command=self.clear_bill)
        clear_btn.pack(side=tk.LEFT, padx=(10, 0))

        print_btn = ttk.Button(bottom, text="  Print Bill  ", command=self.print_bill)
        print_btn.pack(side=tk.LEFT, padx=(10, 0))

        save_btn = ttk.Button(bottom, text="  Save Bill  ", command=self.save_bill)
        save_btn.pack(side=tk.LEFT, padx=(10, 0))

        fetch_btn = ttk.Button(bottom, text="  Fetch Bill  ", command=self.fetch_bill)
        fetch_btn.pack(side=tk.LEFT, padx=(10, 0))

        self.total_var = tk.StringVar(value="Total Bill:  ₹ 0.00")
        ttk.Label(bottom, textvariable=self.total_var, style="Total.TLabel").pack(side=tk.RIGHT, padx=(0, 10))

        # ── Payment Section: Amount Received + Total Discount + Balance ──
        payment_frame = tk.Frame(main, bg="#dfe6e9", bd=2, relief="groove")
        payment_frame.pack(fill=tk.X, pady=(10, 0), ipady=8)

        # Use grid for consistent horizontal alignment
        payment_frame.columnconfigure(0, weight=1)
        payment_frame.columnconfigure(1, weight=1)
        payment_frame.columnconfigure(2, weight=1)

        # Amount Received (left)
        recv_frame = tk.Frame(payment_frame, bg="#dfe6e9")
        recv_frame.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        tk.Label(recv_frame, text="Amount Received (₹):", font=("Segoe UI", 11, "bold"),
                 bg="#dfe6e9", fg="#2c3e50").pack(side=tk.LEFT, padx=(0, 5))
        self.received_var = tk.StringVar(value="0")
        self.received_var.trace_add("write", self._update_balance)
        received_entry = ttk.Entry(recv_frame, textvariable=self.received_var, width=12, font=("Segoe UI", 11, "bold"), justify="right")
        received_entry.pack(side=tk.LEFT)

        # Total Discount (center)
        self.total_discount_var = tk.StringVar(value="Total Discount:  ₹ 0.00")
        tk.Label(payment_frame, textvariable=self.total_discount_var, font=("Segoe UI", 11, "bold"),
                 bg="#dfe6e9", fg="#e17055").grid(row=0, column=1, padx=10, pady=5)

        # Balance (right)
        self.balance_var = tk.StringVar(value="Balance:  ₹ 0.00")
        tk.Label(payment_frame, textvariable=self.balance_var, font=("Segoe UI", 11, "bold"),
                 bg="#dfe6e9", fg="#d63031").grid(row=0, column=2, padx=10, pady=5, sticky="e")

    # ── Item List Helpers ─────────────────────────────────────────

    def _validate_int(self, value):
        if value == "":
            return True
        try:
            return int(value) >= 0
        except ValueError:
            return False

    def _validate_discount(self, value):
        if value == "":
            return True
        import re
        if re.match(r'^\d*\.?\d{0,2}$', value):
            return True
        return False

    def _populate_item_grid(self):
        for widget in self.item_inner_frame.winfo_children():
            widget.destroy()
        self.item_rows = []

        vcmd = (self.root.register(self._validate_int), '%P')
        vcmd_disc = (self.root.register(self._validate_discount), '%P')

        for i, item in enumerate(self.filtered_items):
            bg = "#ffffff" if i % 2 == 0 else "#f7f9fa"
            row_frame = tk.Frame(self.item_inner_frame, bg=bg)
            row_frame.pack(fill=tk.X, pady=0)

            col_widths = [350, 100, 90, 100, 100, 100]
            for col, minw in enumerate(col_widths):
                row_frame.columnconfigure(col, minsize=minw, weight=0)

            tk.Label(row_frame, text=item["name"], font=("Segoe UI", 9), bg=bg, anchor="w").grid(row=0, column=0, padx=4, sticky="ew")
            tk.Label(row_frame, text=item["size"], font=("Segoe UI", 9), bg=bg, anchor="center").grid(row=0, column=1, padx=4, sticky="ew")
            tk.Label(row_frame, text=f"{item['unit_price']:.2f}", font=("Segoe UI", 9), bg=bg, anchor="e").grid(row=0, column=2, padx=4, sticky="ew")

            discount_var = tk.StringVar(value=f"{item['discount']:.2f}")
            discount_entry = tk.Entry(row_frame, textvariable=discount_var, width=8, font=("Segoe UI", 9), justify="center",
                                      validate="key", validatecommand=vcmd_disc)
            discount_entry.grid(row=0, column=3, padx=4)

            qty_var = tk.StringVar(value="0")
            qty_entry = tk.Entry(row_frame, textvariable=qty_var, width=10, font=("Segoe UI", 9), justify="center",
                                 validate="key", validatecommand=vcmd)
            qty_entry.grid(row=0, column=4, padx=4)

            loose_var = tk.StringVar(value="0")
            loose_entry = tk.Entry(row_frame, textvariable=loose_var, width=10, font=("Segoe UI", 9), justify="center",
                                   validate="key", validatecommand=vcmd)
            loose_entry.grid(row=0, column=5, padx=4)

            def _highlight(row_f, qv, lv, orig_bg):
                def _check(*_):
                    try:
                        has_val = int(qv.get()) > 0 or int(lv.get()) > 0
                    except ValueError:
                        has_val = False
                    new_bg = "#d4edda" if has_val else orig_bg
                    row_f.configure(bg=new_bg)
                    for w in row_f.winfo_children():
                        if isinstance(w, tk.Label):
                            w.configure(bg=new_bg)
                return _check

            highlighter = _highlight(row_frame, qty_var, loose_var, bg)
            qty_var.trace_add("write", highlighter)
            loose_var.trace_add("write", highlighter)

            def _focus_border(row_f):
                def _on_focus_in(e):
                    row_f.configure(highlightbackground="#3498db", highlightthickness=1)
                def _on_focus_out(e):
                    row_f.configure(highlightthickness=0)
                return _on_focus_in, _on_focus_out

            focus_in, focus_out = _focus_border(row_frame)
            qty_entry.bind("<FocusIn>", focus_in)
            qty_entry.bind("<FocusOut>", focus_out)
            loose_entry.bind("<FocusIn>", focus_in)
            loose_entry.bind("<FocusOut>", focus_out)
            discount_entry.bind("<FocusIn>", focus_in)
            discount_entry.bind("<FocusOut>", focus_out)

            self.item_rows.append((row_frame, qty_var, loose_var, i, discount_var))

    def _on_search(self, *_args):
        query = self.search_var.get().strip().lower()
        if not query:
            self.filtered_items = list(self.items_data)
        else:
            self.filtered_items = [
                it for it in self.items_data
                if query in it["name"].lower() or query in it["size"].lower()
            ]
        self._populate_item_grid()

    # ── Bill Actions ──────────────────────────────────────────────

    def add_to_bill(self):
        added = False
        for row_frame, qty_var, loose_var, idx, discount_var in self.item_rows:
            try:
                qty = int(qty_var.get())
            except ValueError:
                qty = 0
            try:
                loose = int(loose_var.get())
            except ValueError:
                loose = 0

            if qty <= 0 and loose <= 0:
                continue

            item = self.filtered_items[idx]
            try:
                discount = float(discount_var.get())
            except ValueError:
                discount = item["discount"]
            bottles_per_box = self.box_size_map.get(item["size"].lower(), 1)

            # Check if item already exists in bill
            existing = None
            for bi in self.bill_items:
                if bi["name"] == item["name"] and bi["size"] == item["size"]:
                    existing = bi
                    break

            if existing:
                existing["qty"] += max(qty, 0)
                existing["loose"] += max(loose, 0)
                existing["discount"] = discount
                new_total_bottles = (existing["qty"] * bottles_per_box) + existing["loose"]
                existing["total"] = (item["unit_price"] * new_total_bottles) - (discount * new_total_bottles)
            else:
                total_bottles = (qty * bottles_per_box) + max(loose, 0)
                total = (item["unit_price"] * total_bottles) - (discount * total_bottles)
                self.bill_items.append({
                    "name": item["name"],
                    "size": item["size"],
                    "qty": max(qty, 0),
                    "loose": max(loose, 0),
                    "unit_price": item["unit_price"],
                    "discount": discount,
                    "total": total,
                })
            # Reset the inputs
            qty_var.set("0")
            loose_var.set("0")
            added = True

        if not added:
            messagebox.showwarning("No Quantity", "Please enter a box quantity or loose count for at least one item.")
            return
        
        # Generate Bill No if not already set
        if self.bill_no_var.get() == "Bill No: --":
            bills_file = self._get_bills_file()
            if os.path.exists(bills_file):
                wb = openpyxl.load_workbook(bills_file)
                ws = wb.active
                bill_no = self._get_next_bill_no(ws)
                wb.close()
            else:
                bill_no = 1
            self.bill_no_var.set(f"Bill No: {bill_no}")
        
        self._refresh_bill()

    def remove_item(self):
        selection = self.bill_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a bill item to remove.")
            return
        idx = int(selection[0])
        self.bill_items.pop(idx)
        self._refresh_bill()

    def clear_bill(self):
        if messagebox.askyesno("Confirm", "Clear all items from the bill?"):
            self.bill_items.clear()
            self.received_var.set("0")
            self.bill_to_var.set("")
            self.bill_no_var.set("Bill No: --")
            self.date_var.set(datetime.now().strftime("%d-%m-%Y"))
            for row_frame, qty_var, loose_var, idx, discount_var in self.item_rows:
                qty_var.set("0")
                loose_var.set("0")
            self.bill_tree.delete(*self.bill_tree.get_children())
            self._refresh_bill()

    def _refresh_bill(self):
        self.bill_tree.delete(*self.bill_tree.get_children())
        # Sort: name ascending, size descending (stable sort - secondary first)
        self.bill_items.sort(key=lambda x: self._size_num(x["size"]), reverse=True)
        self.bill_items.sort(key=lambda x: x["name"].lower())
        grand_total = 0.0
        for i, bi in enumerate(self.bill_items):
            self.bill_tree.insert("", tk.END, iid=str(i), values=(
                i + 1,
                bi["name"],
                bi["size"],
                f"{bi['unit_price']:.2f}",
                f"{bi['discount']:.2f}",
                bi["qty"],
                bi.get("loose", 0),
                f"{bi['total']:.2f}",
            ))
            grand_total += bi["total"]
        total_discount = sum(
            bi["discount"] * ((bi["qty"] * self.box_size_map.get(bi["size"].lower(), 1)) + bi.get("loose", 0))
            for bi in self.bill_items
        )
        self.total_var.set(f"Total Bill:  ₹ {grand_total:,.2f}")
        self.total_discount_var.set(f"Total Discount:  ₹ {total_discount:,.2f}")
        self._update_balance()

    def print_bill(self):
        if not self.bill_items:
            messagebox.showwarning("Empty Bill", "There are no items in the bill to print.")
            return

        # Save bill first
        self.save_bill()

        grand_total = sum(bi["total"] for bi in self.bill_items)
        try:
            received = float(self.received_var.get())
        except ValueError:
            received = 0.0
        balance = grand_total - received

        lines = []
        lines.append("=" * 50)
        lines.append("                   BILL RECEIPT")
        lines.append("=" * 50)
        lines.append(f"Date: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
        bill_no_text = self.bill_no_var.get()
        lines.append(bill_no_text)
        bill_to_text = self.bill_to_var.get().strip()
        if bill_to_text:
            lines.append(f"Bill To: {bill_to_text}")
        lines.append("-" * 58)
        lines.append(f"{'#':<4}{'Item':<18}{'Qty Box':<5}{'Bottles':<6}{'Price':<10}{'Disc':<8}{'Total':<10}")
        lines.append("-" * 58)
        for i, bi in enumerate(self.bill_items, 1):
            lines.append(
                f"{i:<4}{bi['name'][:17]:<18}{bi['qty']:<5}{bi.get('loose', 0):<6}"
                f"{bi['unit_price']:<10.2f}{bi['discount']:<8.2f}{bi['total']:<10.2f}"
            )
        lines.append("-" * 50)
        lines.append(f"{'Grand Total:':<37} ₹ {grand_total:,.2f}")
        lines.append(f"{'Amount Received:':<37} ₹ {received:,.2f}")
        if balance > 0:
            lines.append(f"{'Balance Outstanding:':<37} ₹ {balance:,.2f}")
        elif balance < 0:
            lines.append(f"{'Change to Return:':<37} ₹ {abs(balance):,.2f}")
        else:
            lines.append(f"{'Balance:':<37} ₹ 0.00")
        lines.append("=" * 50)
        lines.append("           Thank you for your purchase!")
        lines.append("")

        receipt_text = "\n".join(lines)

        try:
            tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt", prefix="bill_", delete=False, encoding="utf-8")
            tmp.write(receipt_text)
            tmp.close()
            os.startfile(tmp.name, "print")
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print the bill:\n{e}")

    def _update_balance(self, *_args):
        try:
            received = float(self.received_var.get())
        except ValueError:
            received = 0.0
        # Extract grand total from bill items
        grand_total = sum(bi["total"] for bi in self.bill_items)
        balance = grand_total - received
        if balance > 0:
            self.balance_var.set(f"Balance :  ₹ {balance:,.2f}")
        elif balance < 0:
            self.balance_var.set(f"Change to Return:  ₹ {abs(balance):,.2f}")
        else:
            self.balance_var.set("Balance :  ₹ 0.00")

    # ── Save / Fetch Bill ─────────────────────────────────────────

    def _get_bills_file(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bills_generated.xlsx")

    def _get_next_bill_no(self, ws):
        max_no = 0
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] is not None:
                try:
                    max_no = max(max_no, int(row[0]))
                except (ValueError, TypeError):
                    pass
        return max_no + 1

    def save_bill(self, show_message=True):
        if not self.bill_items:
            messagebox.showwarning("Empty Bill", "There are no items in the bill to save.")
            return

        bills_file = self._get_bills_file()
        
        # Extract bill number if already set, otherwise generate a new one
        bill_no_text = self.bill_no_var.get()
        if bill_no_text != "Bill No: --":
            # Extract the bill number from "Bill No: X"
            try:
                bill_no = int(bill_no_text.split(":")[-1].strip())
            except (ValueError, IndexError):
                bill_no = None
        else:
            bill_no = None
        
        # Generate bill number only if not already set
        if bill_no is None:
            if os.path.exists(bills_file):
                wb = openpyxl.load_workbook(bills_file)
                ws = wb.active
                bill_no = self._get_next_bill_no(ws)
                wb.close()
            else:
                bill_no = 1
        
        if os.path.exists(bills_file):
            wb = openpyxl.load_workbook(bills_file)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Bills"
            ws.append(["Bill No", "Item Name", "Size", "Unit Price", "Discount", "Qty Box", "Qty Bottle", "Total", "Amount Received", "Balance", "Bill To", "Bill Date"])

        # Check if bill number already exists and delete existing entries (in reverse order to preserve row indices)
        rows_to_delete = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
            if row[0].value is not None:
                try:
                    if int(row[0].value) == bill_no:
                        rows_to_delete.append(row_idx)
                except (ValueError, TypeError):
                    pass
        
        # Delete rows in reverse order to avoid index shifting issues
        for row_idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row_idx)

        try:
            received = float(self.received_var.get())
        except ValueError:
            received = 0.0
        grand_total = sum(bi["total"] for bi in self.bill_items)
        balance = grand_total - received

        bill_to = self.bill_to_var.get().strip()
        bill_date = self.date_var.get().strip()

        for bi in self.bill_items:
            ws.append([
                bill_no,
                bi["name"],
                bi["size"],
                bi["unit_price"],
                bi["discount"],
                bi["qty"],
                bi.get("loose", 0),
                bi["total"],
                received,
                balance,
                bill_to,
                bill_date,
            ])

        wb.save(bills_file)
        wb.close()
        self.bill_no_var.set(f"Bill No: {bill_no}")
        if show_message:
            messagebox.showinfo("Bill Saved", f"Bill saved successfully!\nBill No: {bill_no}")

    def fetch_bill(self):
        bills_file = self._get_bills_file()
        if not os.path.exists(bills_file):
            messagebox.showwarning("No Bills", "No saved bills found. Save a bill first.")
            return

        # Prompt user for bill number
        bill_no_str = simpledialog.askstring("Fetch Bill", "Enter Bill No:")
        if not bill_no_str:
            return
        try:
            bill_no = int(bill_no_str.strip())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid numeric bill number.")
            return

        wb = openpyxl.load_workbook(bills_file, read_only=True, data_only=True)
        ws = wb.active
        found_items = []
        fetched_received = 0.0
        fetched_bill_to = ""
        fetched_bill_date = ""
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is not None and int(row[0]) == bill_no:
                found_items.append({
                    "name": str(row[1]),
                    "size": str(row[2]),
                    "unit_price": float(row[3]),
                    "discount": float(row[4]),
                    "qty": int(row[5]),
                    "loose": int(row[6]),
                    "total": float(row[7]),
                })
                if len(row) > 8 and row[8] is not None:
                    fetched_received = float(row[8])
                if len(row) > 10 and row[10] is not None:
                    fetched_bill_to = str(row[10]).strip()
                if len(row) > 11 and row[11] is not None:
                    fetched_bill_date = str(row[11]).strip()
        wb.close()

        if not found_items:
            messagebox.showwarning("Not Found", f"No bill found with Bill No: {bill_no}")
            return

        # Load fetched bill into the bill view
        self.bill_items = found_items
        self.received_var.set(str(fetched_received))
        self.bill_to_var.set(fetched_bill_to)
        self.bill_no_var.set(f"Bill No: {bill_no}")
        if fetched_bill_date:
            self.date_var.set(fetched_bill_date)
        self._refresh_bill()
        messagebox.showinfo("Bill Loaded", f"Bill No {bill_no} loaded successfully ({len(found_items)} items).")


if __name__ == "__main__":
    root = tk.Tk()
    app = BillCalculatorApp(root)
    root.mainloop()
