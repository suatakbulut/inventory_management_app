import tkinter as tk
from tkinter import ttk, messagebox, StringVar
import sqlite3
from datetime import datetime

# Create or connect to the database
conn = sqlite3.connect("inventory.db")
cursor = conn.cursor()

# Create tables if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS incoming_items (
        id INTEGER PRIMARY KEY,
        store TEXT,
        item_name TEXT,
        item_no TEXT,
        quantity INTEGER,
        price REAL,
        entry_date TEXT
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS outgoing_shipments (
        id INTEGER PRIMARY KEY,
        item_name TEXT,
        item_no TEXT,
        quantity INTEGER,
        average_price_at_shipment REAL, 
        shipping_personnel TEXT,
        entry_date TEXT
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS inventory (
        item_no TEXT PRIMARY KEY,
        item_name TEXT,
        total_quantity INTEGER,
        average_price REAL
    )
''')

conn.commit()


class InventoryApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Management App")

        self.tabControl = ttk.Notebook(root)
        self.incoming_tab = ttk.Frame(self.tabControl)
        self.outgoing_tab = ttk.Frame(self.tabControl)
        self.inventory_tab = ttk.Frame(self.tabControl)

        self.tabControl.add(self.incoming_tab, text="Incoming Items")
        self.tabControl.add(self.outgoing_tab, text="Outgoing Shipments")
        self.tabControl.add(self.inventory_tab, text="Inventory")
        self.tabControl.pack(expand=1, fill="both")

        self.incoming_items_window = None
        self.outgoing_shipments_window = None

        # Incoming Items Tab
        self.create_incoming_tab()

        # Outgoing Shipments Tab
        self.create_outgoing_tab()

        # Inventory Tab
        self.create_inventory_tab()

        # reporting Tab
        self.create_reporting_tab()

    def clean_input(self, text):
        return text.strip().upper()

    def validate_input(self, *entries):
        for entry in entries:
            if not entry.get():
                messagebox.showwarning(
                    "Missing Information", "Please fill in all fields.")
                return False
        return True

    def create_incoming_tab(self):
        label_store = tk.Label(self.incoming_tab, text="Store:")
        label_item_name = tk.Label(self.incoming_tab, text="Item Name:")
        label_item_no = tk.Label(self.incoming_tab, text="Item No:")
        label_quantity = tk.Label(self.incoming_tab, text="Quantity:")
        label_price = tk.Label(self.incoming_tab, text="Price:")

        self.entry_store = tk.Entry(self.incoming_tab)
        self.entry_item_name = tk.Entry(self.incoming_tab)
        self.entry_item_no = tk.Entry(self.incoming_tab)
        self.entry_quantity = tk.Entry(self.incoming_tab)
        self.entry_price = tk.Entry(self.incoming_tab)

        btn_enter_incoming = tk.Button(
            self.incoming_tab, text="Enter Incoming Item", command=self.enter_incoming_item)
        btn_display_incoming = tk.Button(
            self.incoming_tab, text="Display Incoming Items", command=self.display_incoming_items)

        label_store.grid(row=0, column=0, padx=10, pady=10)
        label_item_name.grid(row=1, column=0, padx=10, pady=10)
        label_item_no.grid(row=2, column=0, padx=10, pady=10)
        label_quantity.grid(row=3, column=0, padx=10, pady=10)
        label_price.grid(row=4, column=0, padx=10, pady=10)

        self.entry_store.grid(row=0, column=1, padx=10, pady=10)
        self.entry_item_name.grid(row=1, column=1, padx=10, pady=10)
        self.entry_item_no.grid(row=2, column=1, padx=10, pady=10)
        self.entry_quantity.grid(row=3, column=1, padx=10, pady=10)
        self.entry_price.grid(row=4, column=1, padx=10, pady=10)

        btn_enter_incoming.grid(row=5, column=0, columnspan=2, pady=10)
        btn_display_incoming.grid(row=6, column=0, columnspan=2, pady=10)

    def create_outgoing_tab(self):
        label_item_no = tk.Label(self.outgoing_tab, text="Item No:")
        label_item_name = tk.Label(self.outgoing_tab, text="Item Name:")
        label_quantity = tk.Label(self.outgoing_tab, text="Quantity:")
        label_shipping_personnel = tk.Label(
            self.outgoing_tab, text="Shipping Personnel:")

        self.entry_item_no_outgoing = tk.Entry(self.outgoing_tab)
        self.entry_item_name_outgoing = tk.Entry(self.outgoing_tab)
        self.entry_quantity_outgoing = tk.Entry(self.outgoing_tab)

        # Dropdown for Shipping Personnel
        self.shipping_personnel_var = tk.StringVar(self.outgoing_tab)
        self.shipping_personnel_var.set("Bekir")  # Default value
        shipping_personnel_options = ["Bekir", "Celal Fatih"]
        shipping_personnel_menu = tk.OptionMenu(
            self.outgoing_tab, self.shipping_personnel_var, *shipping_personnel_options)

        btn_enter_outgoing = tk.Button(
            self.outgoing_tab, text="Enter Outgoing Shipment", command=self.enter_outgoing_shipment)
        btn_display_outgoing = tk.Button(
            self.outgoing_tab, text="Display Outgoing Shipments", command=self.display_outgoing_shipments)

        label_item_no.grid(row=0, column=0, padx=10, pady=10)
        label_item_name.grid(row=1, column=0, padx=10, pady=10)
        label_quantity.grid(row=2, column=0, padx=10, pady=10)
        label_shipping_personnel.grid(row=3, column=0, padx=10, pady=10)

        self.entry_item_no_outgoing.grid(row=0, column=1, padx=10, pady=10)
        self.entry_item_name_outgoing.grid(row=1, column=1, padx=10, pady=10)
        self.entry_quantity_outgoing.grid(row=2, column=1, padx=10, pady=10)
        shipping_personnel_menu.grid(row=3, column=1, padx=10, pady=10)

        btn_enter_outgoing.grid(row=4, column=0, columnspan=2, pady=10)
        btn_display_outgoing.grid(row=5, column=0, columnspan=2, pady=10)

    def create_inventory_tab(self):
        self.tree_inventory = ttk.Treeview(self.inventory_tab, columns=(
            "Item No", "Item Name", "Total Quantity", "Average Price"))
        self.tree_inventory.heading("#1", text="Item No")
        self.tree_inventory.heading("#2", text="Item Name")
        self.tree_inventory.heading("#3", text="Total Quantity")
        self.tree_inventory.heading("#4", text="Average Cost")

        self.tree_inventory.grid(row=0, column=0, padx=10, pady=10)

        btn_display_inventory = tk.Button(
            self.inventory_tab, text="Display Inventory", command=self.display_inventory)
        btn_display_inventory.grid(row=1, column=0, pady=10)

    def enter_incoming_item(self):
        # Validate input
        if not self.validate_input(self.entry_item_no, self.entry_store, self.entry_item_name, self.entry_quantity, self.entry_price):
            return

        item_no = self.clean_input(self.entry_item_no.get())
        store = self.clean_input(self.entry_store.get())
        item_name = self.clean_input(self.entry_item_name.get())
        quantity = int(self.clean_input(self.entry_quantity.get()))
        price = float(self.clean_input(self.entry_price.get()))
        entry_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute('''
            INSERT INTO incoming_items (item_no, store, item_name, quantity, price, entry_date)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (item_no, store, item_name, quantity, price, entry_date))

        self.update_inventory_incoming(item_no, item_name, quantity, price)
        conn.commit()

        # Display success message
        messagebox.showinfo("Success", "Incoming item entered successfully!")

        # Clear entry fields
        self.entry_store.delete(0, tk.END)
        self.entry_item_name.delete(0, tk.END)
        self.entry_item_no.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        self.entry_price.delete(0, tk.END)

    def enter_outgoing_shipment(self):
        # Validate input
        if not self.validate_input(self.entry_item_no_outgoing, self.entry_quantity_outgoing):
            return

        item_no = self.clean_input(self.entry_item_no_outgoing.get())
        quantity = int(self.clean_input(self.entry_quantity_outgoing.get()))
        item_name = self.clean_input(self.entry_item_name_outgoing.get())
        shipping_personnel = self.clean_input(
            self.shipping_personnel_var.get())
        entry_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Check if there is enough quantity in the inventory
        cursor.execute('''
            SELECT total_quantity, average_price FROM inventory WHERE item_no=?
        ''', (item_no,))
        result = cursor.fetchone()
        if result and result[0] >= quantity:
            average_price_at_shipment = result[1]

            cursor.execute('''
                INSERT INTO outgoing_shipments (item_name, item_no, quantity, shipping_personnel, average_price_at_shipment, entry_date)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (item_name, item_no, quantity, shipping_personnel, average_price_at_shipment, entry_date))

            conn.commit()
            self.update_inventory_outgoing(item_no, item_name, quantity)

            # Display success message
            messagebox.showinfo(
                "Success", "Outgoing shipment entered successfully!")

            # Clear entry fields
            self.entry_item_no_outgoing.delete(0, tk.END)
            self.entry_item_name_outgoing.delete(0, tk.END)
            self.entry_quantity_outgoing.delete(0, tk.END)
        else:
            messagebox.showwarning(
                "Insufficient Quantity", "Not enough quantity in the inventory.")

    def update_inventory_incoming(self, item_no, item_name, quantity, per_unit_price):
        # Check if the item already exists in the inventory
        cursor.execute('SELECT * FROM inventory WHERE item_no = ?', (item_no,))
        existing_item = cursor.fetchone()

        additional_total_price = quantity*per_unit_price
        if existing_item:
            # Update existing item
            cursor.execute('''
                UPDATE inventory 
                SET item_name = ?, 
                    total_quantity = total_quantity + ?, 
                    average_price = ( total_quantity*average_price + ? ) / (total_quantity+?)  
                WHERE item_no = ?
            ''', (item_name, quantity, additional_total_price, quantity, item_no))
        else:
            # Insert new item
            cursor.execute('''
                INSERT INTO inventory (item_no, item_name, total_quantity, average_price)
                VALUES (?, ?, ?, ?)
            ''', (item_no, item_name, quantity, per_unit_price if quantity != 0 else 0))

        conn.commit()

        # Refresh inventory display
        self.display_inventory()

    def update_inventory_outgoing(self, item_no, item_name, quantity):
        # This time the item must already exist in the inventory

        cursor.execute('''
            UPDATE inventory 
            SET total_quantity = total_quantity - ? 
            WHERE item_name = ? AND item_no = ?
        ''', (quantity, item_name, item_no)
        )

        conn.commit()

        # Refresh inventory display
        self.display_inventory()

    def display_inventory(self):
        # Clear previous data
        for row in self.tree_inventory.get_children():
            self.tree_inventory.delete(row)

        # Display updated inventory
        cursor.execute('''
            SELECT item_no, item_name, total_quantity, average_price AS average_cost FROM inventory
        ''')
        result = cursor.fetchall()
        for row in result:
            self.tree_inventory.insert(
                "", "end", values=(row[0], row[1], row[2], row[3]))

    def display_incoming_items(self):
        # Close existing window
        if self.incoming_items_window:
            self.incoming_items_window.destroy()

        # Create a new window for displaying incoming items
        self.incoming_items_window = tk.Toplevel(self.root)
        self.incoming_items_window.title("Incoming Items")

        tree_incoming_items = ttk.Treeview(self.incoming_items_window, columns=(
            "ID", "Item No", "Store", "Item Name", "Quantity", "Price", "Entry Date"))
        tree_incoming_items.heading("#1", text="ID")
        tree_incoming_items.heading("#2", text="Item No")
        tree_incoming_items.heading("#3", text="Store")
        tree_incoming_items.heading("#4", text="Item Name")
        tree_incoming_items.heading("#5", text="Quantity")
        tree_incoming_items.heading("#6", text="Price")
        tree_incoming_items.heading("#7", text="Entry Date")

        tree_incoming_items.grid(row=0, column=0, padx=10, pady=10)

        # Display incoming items
        cursor.execute('''
            SELECT id, item_no, store, item_name, quantity, price, entry_date FROM incoming_items
        ''')
        result = cursor.fetchall()
        for row in result:
            tree_incoming_items.insert("", "end", values=row)

    def display_outgoing_shipments(self):
        # Close existing window
        if self.outgoing_shipments_window:
            self.outgoing_shipments_window.destroy()

        # Create a new window for displaying outgoing shipments
        self.outgoing_shipments_window = tk.Toplevel(self.root)
        self.outgoing_shipments_window.title("Outgoing Shipments")

        tree_outgoing_shipments = ttk.Treeview(self.outgoing_shipments_window, columns=(
            "ID", "Item No", "Quantity", "Item Name", "Average Price at Shipment", "Shipping Personnel", "Entry Date"))
        tree_outgoing_shipments.heading("#1", text="ID")
        tree_outgoing_shipments.heading("#2", text="Item No")
        tree_outgoing_shipments.heading("#3", text="Item Name")
        tree_outgoing_shipments.heading("#4", text="Quantity")
        tree_outgoing_shipments.heading("#5", text="Average Price at Shipment")
        tree_outgoing_shipments.heading("#6", text="Shipping Personnel")
        tree_outgoing_shipments.heading("#7", text="Entry Date")

        tree_outgoing_shipments.grid(row=0, column=0, padx=10, pady=10)

        # Display outgoing shipments
        cursor.execute('''
            SELECT os.id, os.item_no, os.item_name, os.quantity, os.average_price_at_shipment, os.shipping_personnel, os.entry_date
            FROM outgoing_shipments os
            ''')
        result = cursor.fetchall()
        for row in result:
            tree_outgoing_shipments.insert("", "end", values=row)

    def create_reporting_tab(self):
        self.reporting_tab = ttk.Frame(self.tabControl)
        self.tabControl.add(self.reporting_tab, text="Reporting")

        # Dropdown for Shipping Personnel
        self.shipping_personnel_var_report = StringVar(self.reporting_tab)
        self.shipping_personnel_var_report.set("All")  # Default value
        shipping_personnel_options_report = ["All", "Bekir", "Celal Fatih"]
        shipping_personnel_menu_report = tk.OptionMenu(self.reporting_tab, self.shipping_personnel_var_report, *shipping_personnel_options_report)

        # Dropdown for Item No
        self.item_no_var_report = StringVar(self.reporting_tab)
        self.item_no_var_report.set("All")  # Default value
        item_no_options_report = ["All"] + self.get_all_item_nos()
        item_no_menu_report = tk.OptionMenu(self.reporting_tab, self.item_no_var_report, *item_no_options_report)

        # Dropdown for Item Name
        self.item_name_var_report = StringVar(self.reporting_tab)
        self.item_name_var_report.set("All")  # Default value
        item_name_options_report = ["All"] + self.get_all_item_names()
        item_name_menu_report = tk.OptionMenu(self.reporting_tab, self.item_name_var_report, *item_name_options_report)

        btn_generate_report = tk.Button(self.reporting_tab, text="Generate Report", command=self.generate_report)

        shipping_personnel_menu_report.grid(row=0, column=0, padx=10, pady=10)
        item_no_menu_report.grid(row=0, column=1, padx=10, pady=10)
        item_name_menu_report.grid(row=0, column=2, padx=10, pady=10)
        btn_generate_report.grid(row=0, column=3, padx=10, pady=10)

        # Treeview for displaying the report
        self.tree_report = ttk.Treeview(self.reporting_tab, columns=("Item No", "Item Name", "Total Quantity", "Total Cost"))
        self.tree_report.heading("#1", text="Item No")
        self.tree_report.heading("#2", text="Item Name")
        self.tree_report.heading("#3", text="Total Quantity")
        self.tree_report.heading("#4", text="Total Cost")
        self.tree_report.grid(row=1, column=0, columnspan=4, padx=10, pady=10)

    def generate_report(self):
        # Clear previous data
        for row in self.tree_report.get_children():
            self.tree_report.delete(row)

        shipping_personnel = self.clean_input(self.shipping_personnel_var_report.get())
        item_no = self.clean_input(self.item_no_var_report.get())
        item_name = self.clean_input(self.item_name_var_report.get())

        # Modify the SQL query based on user selections
        query_params = []
        where_conditions = []

        if shipping_personnel != "ALL":
            where_conditions.append("os.shipping_personnel = ?")
            query_params.append(shipping_personnel)

        if item_no != "ALL":
            where_conditions.append("os.item_no = ?")
            query_params.append(item_no)

        if item_name != "ALL":
            where_conditions.append("os.item_name = ?")
            query_params.append(item_name)

        where_clause = " AND ".join(where_conditions) if where_conditions else ""

        query = f'''
            SELECT os.item_no, os.item_name, SUM(os.quantity) as total_quantity, SUM(os.quantity * os.average_price_at_shipment) as total_cost
            FROM outgoing_shipments os
            {"WHERE " + where_clause if where_clause else ""}
            GROUP BY os.item_no, os.item_name
        '''

        cursor.execute(query, query_params)
        result = cursor.fetchall()

        for row in result:
            self.tree_report.insert("", "end", values=row)
            
    def get_all_item_nos(self):
        cursor.execute('SELECT DISTINCT item_no FROM inventory')
        item_nos_inventory = [row[0] for row in cursor.fetchall()]

        cursor.execute('SELECT DISTINCT item_no FROM outgoing_shipments')
        item_nos_outgoing_shipments = [row[0] for row in cursor.fetchall()]

        item_nos = list(set(item_nos_inventory + item_nos_outgoing_shipments))
        return [item_no for item_no in item_nos if item_no]

    def get_all_item_names(self):
        cursor.execute('SELECT DISTINCT item_name FROM inventory')
        item_names_inventory = [row[0] for row in cursor.fetchall()]

        cursor.execute('SELECT DISTINCT item_name FROM outgoing_shipments')
        item_names_outgoing_shipments = [row[0] for row in cursor.fetchall()]

        item_names = list(
            set(item_names_inventory + item_names_outgoing_shipments))
        return [item_name for item_name in item_names if item_name]

    def display_report(self):
        personnel = self.clean_input(self.shipping_personnel_var.get())
        item_no = self.clean_input(self.item_no_var_report.get())
        item_name = self.clean_input(self.item_name_var_report.get())

        print("personnel", personnel, type(personnel))
        print("item_no", item_no, type(item_no))
        print("item_name", item_name, type(item_name))

        query = '''
            SELECT SUM(os.quantity * os.average_price_at_shipment) AS total_cost
            FROM outgoing_shipments os
            WHERE (? = 'ALL' OR os.shipping_personnel = ?)
              AND (? = 'ALL' OR os.item_no = ?)
              AND (? = 'ALL' OR os.item_name = ?)
        '''

        cursor.execute(query, (personnel, personnel, item_no, item_no, item_name, item_name))
        result = cursor.fetchone()

        if result and result[0] is not None:
            total_cost = result[0]
            self.result_text.delete(1.0, tk.END)  # Clear previous result
            self.result_text.insert(tk.END, f"Total Cost: {total_cost}")
        else:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(
                tk.END, "No data available for the selected criteria.")


if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

# Close the database connection
conn.close()