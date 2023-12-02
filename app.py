import tkinter as tk
from tkinter import ttk, messagebox, StringVar
import sqlite3
from datetime import datetime
import pandas as pd

# Create or connect to the database
conn = sqlite3.connect("inventory.db")
cursor = conn.cursor()

# Create tables if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS incoming_items (
        id INTEGER PRIMARY KEY,
        store TEXT,
        item_no TEXT,
        item_name TEXT,
        quantity INTEGER,
        price REAL, 
        tax_rate REAL,
        entry_date TEXT
    )
''')


cursor.execute('''
    CREATE TABLE IF NOT EXISTS outgoing_shipments (
        id INTEGER PRIMARY KEY,
        store TEXT,
        item_name TEXT,
        item_no TEXT,
        quantity INTEGER,
        average_price_at_shipment REAL, 
        average_price_at_shipment_after_tax REAL, 
        shipping_to TEXT,
        entry_date TEXT
    )
''')


cursor.execute('''
    CREATE TABLE IF NOT EXISTS inventory (
        store TEXT,
        item_no TEXT,
        item_name TEXT,
        total_quantity INTEGER,
        average_price_before_tax REAL,
        average_price_after_tax REAL,
        PRIMARY KEY (store, item_no, item_name)
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
        # Store
        label_store = tk.Label(self.incoming_tab, text="Store:")
        self.entry_store = tk.Entry(self.incoming_tab)
        label_store.grid(row=0, column=0, padx=10, pady=10)
        self.entry_store.grid(row=0, column=1, padx=10, pady=10)

        # Item no
        label_item_no = tk.Label(self.incoming_tab, text="Item No:")
        self.entry_item_no = tk.Entry(self.incoming_tab)
        label_item_no.grid(row=1, column=0, padx=10, pady=10)
        self.entry_item_no.grid(row=1, column=1, padx=10, pady=10)

        # Item name
        label_item_name = tk.Label(self.incoming_tab, text="Item Name:")
        self.entry_item_name = tk.Entry(self.incoming_tab)
        label_item_name.grid(row=2, column=0, padx=10, pady=10)
        self.entry_item_name.grid(row=2, column=1, padx=10, pady=10)

        # Quantity
        label_quantity = tk.Label(self.incoming_tab, text="Quantity:")
        self.entry_quantity = tk.Entry(self.incoming_tab)
        label_quantity.grid(row=3, column=0, padx=10, pady=10)
        self.entry_quantity.grid(row=3, column=1, padx=10, pady=10)

        # Price
        label_price = tk.Label(self.incoming_tab, text="Price:")
        self.entry_price = tk.Entry(self.incoming_tab)
        label_price.grid(row=4, column=0, padx=10, pady=10)
        self.entry_price.grid(row=4, column=1, padx=10, pady=10)

        # Tax rate
        label_tax_rate = tk.Label(self.incoming_tab, text="Tax Rate (%):")
        self.entry_tax_rate = tk.Entry(self.incoming_tab)
        label_tax_rate.grid(row=5, column=0, padx=10, pady=10)
        self.entry_tax_rate.grid(row=5, column=1, padx=10, pady=10)

        # Entry Date
        label_entry_date = tk.Label(
            self.incoming_tab, text="Entry Date:")
        self.entry_date = tk.Entry(self.incoming_tab)
        label_entry_date.grid(row=6, column=0, padx=10, pady=10)
        self.entry_date.grid(row=6, column=1, padx=10, pady=10)

        # Set today's date as the default value
        self.entry_date.insert(
            0, datetime.now().strftime("%m/%d/%Y"))

        btn_enter_incoming = tk.Button(
            self.incoming_tab, text="Enter Incoming Item", command=self.enter_incoming_item)
        btn_enter_incoming.grid(row=7, column=0, columnspan=2, pady=10)

        btn_display_incoming = tk.Button(
            self.incoming_tab, text="Display Incoming Items", command=self.display_incoming_items)
        btn_display_incoming.grid(row=8, column=0, columnspan=2, pady=10)

    def create_outgoing_tab(self):
        # Store
        label_store = tk.Label(self.outgoing_tab, text="Store:")
        self.entry_store_outgoing = tk.Entry(self.outgoing_tab)
        label_store.grid(row=0, column=0, padx=10, pady=10)
        self.entry_store_outgoing.grid(row=0, column=1, padx=10, pady=10)

        # Item no
        label_item_no = tk.Label(self.outgoing_tab, text="Item No:")
        self.entry_item_no_outgoing = tk.Entry(self.outgoing_tab)
        label_item_no.grid(row=1, column=0, padx=10, pady=10)
        self.entry_item_no_outgoing.grid(row=1, column=1, padx=10, pady=10)

        # Item name
        label_item_name = tk.Label(self.outgoing_tab, text="Item Name:")
        self.entry_item_name_outgoing = tk.Entry(self.outgoing_tab)
        label_item_name.grid(row=2, column=0, padx=10, pady=10)
        self.entry_item_name_outgoing.grid(row=2, column=1, padx=10, pady=10)

        # Quantity
        label_quantity = tk.Label(self.outgoing_tab, text="Quantity:")
        self.entry_quantity_outgoing = tk.Entry(self.outgoing_tab)
        label_quantity.grid(row=3, column=0, padx=10, pady=10)
        self.entry_quantity_outgoing.grid(row=3, column=1, padx=10, pady=10)

        # Shipping to (dropdown menu)
        label_shipping_to = tk.Label(
            self.outgoing_tab, text="Shipping To:")
        self.shipping_to_var = tk.StringVar(self.outgoing_tab)
        self.shipping_to_var.set("USA FBA")  # Default value
        shipping_to_options = [
            "USA FBA", "USA MFN", "CAN FBA", "CAN MFN"]
        shipping_to_menu = tk.OptionMenu(
            self.outgoing_tab, self.shipping_to_var, *shipping_to_options)

        label_shipping_to.grid(row=4, column=0, padx=10, pady=10)
        shipping_to_menu.grid(row=4, column=1, padx=10, pady=10)

        # Entry Date
        label_entry_date_outgoing = tk.Label(
            self.outgoing_tab, text="Entry Date:")
        self.entry_date_outgoing = tk.Entry(self.outgoing_tab)
        label_entry_date_outgoing.grid(row=5, column=0, padx=10, pady=10)
        self.entry_date_outgoing.grid(row=5, column=1, padx=10, pady=10)

        # Set today's date as the default value
        self.entry_date_outgoing.insert(
            0, datetime.now().strftime("%m/%d/%Y"))

        # Enter Outgoing Shipment button
        btn_enter_outgoing = tk.Button(
            self.outgoing_tab, text="Enter Outgoing Shipment", command=self.enter_outgoing_shipment)
        btn_enter_outgoing.grid(row=6, column=0, columnspan=2, pady=10)

        # Display Outgoing Shipments button
        btn_display_outgoing = tk.Button(
            self.outgoing_tab, text="Display Outgoing Shipments", command=self.display_outgoing_shipments)
        btn_display_outgoing.grid(row=7, column=0, columnspan=2, pady=10)

    def create_inventory_tab(self):
        self.tree_inventory = ttk.Treeview(self.inventory_tab, columns=(
            "Store", "Item No", "Item Name", "Total Quantity", "Average Price", "Average Price After Tax"))

        self.tree_inventory.column("#1", width=150, anchor="w")
        self.tree_inventory.column("#2", width=150, anchor="w")
        self.tree_inventory.column("#3", width=150, anchor="w")
        self.tree_inventory.column("#4", width=150, anchor="e")
        self.tree_inventory.column("#5", width=150, anchor="e")
        self.tree_inventory.column("#6", width=150, anchor="e")

        self.tree_inventory.heading("#1", text="Store")
        self.tree_inventory.heading("#2", text="Item No")
        self.tree_inventory.heading("#3", text="Item Name")
        self.tree_inventory.heading("#4", text="Total Quantity")
        self.tree_inventory.heading("#5", text="Average Price")
        self.tree_inventory.heading("#6", text="Average Price After Tax")

        self.tree_inventory.grid(row=0, column=0, padx=10, pady=10)

        btn_display_inventory = tk.Button(
            self.inventory_tab, text="Display Inventory", command=self.display_inventory)
        btn_display_inventory.grid(row=1, column=0, pady=10)

    def enter_incoming_item(self):
        # Validate input
        if not self.validate_input(self.entry_store, self.entry_item_no, self.entry_item_name, self.entry_quantity, self.entry_price, self.entry_tax_rate):
            return

        item_no = self.clean_input(self.entry_item_no.get())
        store = self.clean_input(self.entry_store.get())
        item_name = self.clean_input(self.entry_item_name.get())
        quantity = int(self.clean_input(self.entry_quantity.get()))
        price = float(self.clean_input(self.entry_price.get()))
        tax_rate = float(self.clean_input(self.entry_tax_rate.get()))

        # Get the entry date from the entry field or use today's date if not provided
        entry_date_str = self.clean_input(self.entry_date.get())
        entry_date = datetime.strptime(
            entry_date_str, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S") if entry_date_str else datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute('''
            INSERT INTO incoming_items (store, item_no, item_name, quantity, price, tax_rate, entry_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (store, item_no, item_name, quantity, price, tax_rate, entry_date))

        self.update_inventory_incoming(
            store, item_no, item_name, quantity, price, tax_rate)
        conn.commit()

        # Display success message
        messagebox.showinfo("Success", "Incoming item entered successfully!")

        # Clear entry fields
        self.entry_store.delete(0, tk.END)
        self.entry_item_no.delete(0, tk.END)
        self.entry_item_name.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        self.entry_price.delete(0, tk.END)
        self.entry_tax_rate.delete(0, tk.END)

    def enter_outgoing_shipment(self):
        # Validate input
        if not self.validate_input(self.entry_item_no_outgoing, self.entry_quantity_outgoing):
            return

        store = self.clean_input(self.entry_store_outgoing.get())
        item_no = self.clean_input(self.entry_item_no_outgoing.get())
        item_name = self.clean_input(self.entry_item_name_outgoing.get())
        quantity = int(self.clean_input(self.entry_quantity_outgoing.get()))
        shipping_to = self.clean_input(
            self.shipping_to_var.get())

        # Get the entry date from the entry field or use today's date if not provided
        entry_date_str = self.clean_input(self.entry_date_outgoing.get())
        entry_date = datetime.strptime(
            entry_date_str, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S") if entry_date_str else datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Check if there is enough quantity in the inventory
        cursor.execute('''
            SELECT total_quantity, average_price_before_tax, average_price_after_tax FROM inventory WHERE store=? AND item_no=? AND item_name=?
        ''', (store, item_no, item_name))
        result = cursor.fetchone()
        if result and result[0] >= quantity:
            average_price_at_shipment = result[1]
            average_price_after_tax_at_shipment = result[2]

            cursor.execute('''
                INSERT INTO outgoing_shipments (store, item_no, item_name, quantity, shipping_to, average_price_at_shipment, average_price_at_shipment_after_tax, entry_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (store, item_no, item_name, quantity, shipping_to, average_price_at_shipment, average_price_after_tax_at_shipment, entry_date))

            conn.commit()
            self.update_inventory_outgoing(store, item_no, item_name, quantity)

            # Display success message
            messagebox.showinfo(
                "Success", "Outgoing shipment entered successfully!")

            # Clear entry fields
            self.entry_store_outgoing.delete(0, tk.END)
            self.entry_item_no_outgoing.delete(0, tk.END)
            self.entry_item_name_outgoing.delete(0, tk.END)
            self.entry_quantity_outgoing.delete(0, tk.END)
        else:
            if result:
                available_quantity = result[0]
                messagebox.showwarning(
                    "Insufficient Quantity", f"Not enough quantity in the inventory. Available quantity for ({store}, {item_no}, {item_name}): {available_quantity}.")
            else:
                messagebox.showwarning(
                    "Item Not Found", f"Item ({store}, {item_no}, {item_name}) not found in the inventory.")

    def update_inventory_incoming(self, store, item_no, item_name, quantity, per_unit_price, tax_rate):

        per_unit_price_after_tax = per_unit_price + per_unit_price*tax_rate/100
        # Check if the item already exists in the inventory
        cursor.execute(
            'SELECT * FROM inventory WHERE store = ? AND item_no = ? AND item_name = ?', (store, item_no, item_name))
        existing_item = cursor.fetchone()

        additional_total_price = quantity * per_unit_price
        additional_total_price_after_tax = quantity * per_unit_price_after_tax

        if existing_item:
            # Update existing item
            cursor.execute('''
                UPDATE inventory 
                SET item_name = ?, 
                    total_quantity = total_quantity + ?, 
                    average_price_before_tax = (total_quantity * average_price_before_tax + ?) / (total_quantity + ?), 
                    average_price_after_tax  = (total_quantity * average_price_after_tax + ?) / (total_quantity + ?) 
                WHERE store = ? AND item_no = ? AND item_name = ?
            ''', (item_name, quantity, additional_total_price, quantity, additional_total_price_after_tax, quantity, store, item_no, item_name))
        else:
            # Insert new item
            cursor.execute('''
                INSERT INTO inventory (store, item_no, item_name, total_quantity, average_price_before_tax, average_price_after_tax)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (store, item_no, item_name, quantity, per_unit_price if quantity != 0 else 0, per_unit_price_after_tax if quantity != 0 else 0))

        conn.commit()

        # Refresh inventory display
        self.display_inventory()

    def update_inventory_outgoing(self, store, item_no, item_name, quantity):
        # This time the item must already exist in the inventory

        cursor.execute('''
            UPDATE inventory 
            SET total_quantity = total_quantity - ? 
            WHERE store=? AND item_no = ? AND item_name = ? 
        ''', (quantity, store, item_no, item_name)
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
            SELECT store, item_no, item_name, total_quantity, average_price_before_tax, average_price_after_tax FROM inventory
        ''')
        result = cursor.fetchall()
        for row in result:
            formatted_row = list(row)
            formatted_row[3] = self.format_quantity(
                row[3])  # Format the quantity value
            formatted_row[4] = self.format_price(
                row[4])  # Format the Average Price
            # Format the Average Price After Tax
            formatted_row[5] = self.format_price(row[5])
            self.tree_inventory.insert("", "end", values=formatted_row)

    def format_quantity(self, value):
        return f"{value:,}"

    def format_price(self, value):
        return "{:0,.2f}".format(value)

    def update_inventory_delete_incoming_entry(self, inventory_info, store, item_no, item_name, quantity, unit_price, tax_rate):
        # Current inventory details
        current_quantity = int(inventory_info[0])
        inventory_price_before_tax = float(inventory_info[1])
        inventory_price_after_tax = float(inventory_info[2])
        current_total_cost_before_tax = current_quantity * inventory_price_before_tax
        current_total_cost_after_tax = current_quantity * inventory_price_after_tax

        # Detail of the selected entry
        selected_entry_quantity = int(quantity)
        selected_entry_price = float(unit_price)
        selected_entry_tax_rate = float(tax_rate)
        selected_entry_total_cost_before_tax = selected_entry_quantity * selected_entry_price
        selected_entry_total_cost_after_tax = selected_entry_total_cost_before_tax + \
            selected_entry_total_cost_before_tax * selected_entry_tax_rate

        # Calculate the updated inventory information after deleting the incoming item
        updated_quantity = current_quantity - selected_entry_quantity
        updated_total_cost_before_tax = current_total_cost_before_tax - \
            selected_entry_total_cost_before_tax
        updated_total_cost_after_tax = current_total_cost_after_tax - \
            selected_entry_total_cost_after_tax
        updated_average_price_before_tax = updated_total_cost_before_tax / updated_quantity
        updated_average_price_after_tax = updated_total_cost_after_tax / updated_quantity

        # Update the inventory table with the new information
        cursor.execute('''
            UPDATE inventory
            SET total_quantity = ?,
                average_price_before_tax = ?,
                average_price_after_tax = ?
            WHERE store = ? AND item_no = ? AND item_name = ?
        ''', (updated_quantity, updated_average_price_before_tax if updated_quantity != 0 else 0, updated_average_price_after_tax if updated_quantity != 0 else 0, store, item_no, item_name))
        conn.commit()

    def display_incoming_items(self):
        def delete_selected_incoming_items_entry():
            # Get the selected item's values
            selected_item = tree_incoming_items.selection()
            if not selected_item:
                messagebox.showwarning(
                    "No Selection", "Please select an entry to delete.")
                return

            selected_item_values = tree_incoming_items.item(
                selected_item, 'values')
            entry_date, store, item_no, item_name, quantity, unit_price, tax_rate = selected_item_values

            # Get the inventory details for this item item's values
            cursor.execute('''
                SELECT total_quantity, average_price_before_tax, average_price_after_tax
                FROM inventory
                WHERE store = ? AND item_no = ? AND item_name = ?
            ''', (store, item_no, item_name))

            inventory_info = cursor.fetchone()
            if inventory_info:
                current_quantity = int(inventory_info[0])
                if current_quantity < int(quantity):
                    messagebox.showinfo(
                        "Error", f"Selected entry cannot be deleted because some of them are already shipped. Remaining quantity in the inventory is {current_quantity}.")

                else:
                    # Delete the selected entry from the incoming_items table
                    cursor.execute(f'''
                        DELETE FROM incoming_items
                        WHERE entry_date = ? AND store = ? AND item_no = ? AND item_name = ? AND quantity=? AND price =? AND tax_rate =?
                    ''', (entry_date, store, item_no, item_name, quantity, unit_price, tax_rate))

                    conn.commit()

                    # Update the inventory table / DELETE or UPDATE
                    if current_quantity == int(quantity):
                        # Delete the selected item from the inventory
                        cursor.execute('''
                            DELETE FROM inventory
                            WHERE store = ? AND item_no = ? AND item_name = ?
                        ''', (store, item_no, item_name))
                        conn.commit()
                    else:
                        # Update the selected item in the inventory
                        tax_rate = float(tax_rate) / 100.0
                        self.update_inventory_delete_incoming_entry(
                            inventory_info, store, item_no, item_name, quantity, unit_price, tax_rate)

                    # Refresh the display
                    self.display_incoming_items()
                    self.display_inventory()

                    messagebox.showinfo(
                        "Success", "Selected entry deleted successfully.")
            else:
                messagebox.showinfo(
                    "Error", "Selected item is not in the inventory.")

        # Close existing window
        if self.incoming_items_window:
            self.incoming_items_window.destroy()

        # Create a new window for displaying incoming items
        self.incoming_items_window = tk.Toplevel(self.root)
        self.incoming_items_window.title("Incoming Items")

        # Set the size of the window based on the screen dimensions
        screen_width = self.incoming_items_window.winfo_screenwidth()
        screen_height = self.incoming_items_window.winfo_screenheight()
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.9)
        window_x = int((screen_width - window_width) / 2)
        window_y = int((screen_height - window_height) / 2)

        self.incoming_items_window.geometry(
            f"{window_width}x{window_height}+{window_x}+{window_y}")

        tree_incoming_items = ttk.Treeview(self.incoming_items_window, columns=(
            "Date", "Store", "Item No", "Item Name", "Quantity", "Unit Price", "Tax Rate"))

        # Set column widths
        tree_incoming_items.column("#0", width=20, anchor="w")  # #
        tree_incoming_items.column("#1", width=130, anchor="w")  # Date
        tree_incoming_items.column("#2", width=130, anchor="w")  # Store
        tree_incoming_items.column("#3", width=130, anchor="w")  # Item No
        tree_incoming_items.column("#4", width=130, anchor="w")  # Item Name
        # Quantity (right-justified with commas)
        tree_incoming_items.column("#5", width=130, anchor="e")
        tree_incoming_items.column("#6", width=130, anchor="e")  # Unit Price
        tree_incoming_items.column("#7", width=130, anchor="e")  # Tax Rate

        tree_incoming_items.heading("#0", text="#")
        tree_incoming_items.heading("#1", text="Date")
        tree_incoming_items.heading("#2", text="Store")
        tree_incoming_items.heading("#3", text="Item No")
        tree_incoming_items.heading("#4", text="Item Name")
        tree_incoming_items.heading("#5", text="Quantity")
        tree_incoming_items.heading("#6", text="Unit Price")
        tree_incoming_items.heading("#7", text="Tax Rate")

        tree_incoming_items.grid(row=0, column=0, padx=10, pady=10)

        # Add a button to delete the selected entry
        btn_delete_entry = tk.Button(
            self.incoming_items_window, text="Delete Selected Entry", command=delete_selected_incoming_items_entry)
        btn_delete_entry.grid(row=1, column=0, padx=10, pady=10)

        # Display incoming items
        cursor.execute('''
            SELECT entry_date, store, item_no, item_name, quantity, price, tax_rate FROM incoming_items
        ''')

        result = cursor.fetchall()
        for index, row in enumerate(result):
            tree_incoming_items.insert("", "end", values=row)

    def update_inventory_delete_outgoing_entry(self, inventory_info, store, item_no, item_name, quantity, unit_price, tax_rate):
        new_quantity = str(-1 * int(quantity))
        self.update_inventory_delete_incoming_entry(inventory_info, store, item_no, item_name, new_quantity, unit_price, tax_rate)

    def display_outgoing_shipments(self):
        def delete_selected_outgoing_shipment_entry():
            # Get the selected item's values
            selected_item = tree_outgoing_shipments.selection()
            if not selected_item:
                messagebox.showwarning(
                    "No Selection", "Please select an entry to delete.")
                return

            selected_item_values = tree_outgoing_shipments.item(
                selected_item, 'values')
            entry_date, destination, store, item_no, item_name, quantity, avg_price_before_tax, avg_price_after_tax = selected_item_values

            # Get the inventory details for this item item's values
            cursor.execute('''
                SELECT total_quantity, average_price_before_tax, average_price_after_tax
                FROM inventory
                WHERE store = ? AND item_no = ? AND item_name = ?
            ''', (store, item_no, item_name))

            inventory_info = cursor.fetchone()
            if inventory_info:
                # Update the selected item in the inventory
                tax_rate = float(avg_price_after_tax) / float(avg_price_before_tax) - 1.0
                self.update_inventory_delete_outgoing_entry(
                    inventory_info, store, item_no, item_name, quantity, avg_price_before_tax, tax_rate)
            else:
                # Insert new item
                cursor.execute('''
                    INSERT INTO inventory (store, item_no, item_name, total_quantity, average_price_before_tax, average_price_after_tax)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (store, item_no, item_name, quantity, avg_price_before_tax if quantity != 0 else 0, avg_price_after_tax if quantity != 0 else 0))
                conn.commit()
            
            # Delete the selected entry from the incoming_items table
            cursor.execute(f'''
                DELETE FROM outgoing_shipments
                WHERE entry_date = ? AND store = ? AND item_no = ? AND item_name = ? AND quantity=? AND average_price_at_shipment =? AND average_price_at_shipment_after_tax =? AND shipping_to=?
            ''', (entry_date, store, item_no, item_name, quantity, avg_price_before_tax, avg_price_after_tax, destination))

            conn.commit()
                
            # Refresh the display
            self.display_outgoing_shipments()
            self.display_inventory()

            messagebox.showinfo(
                "Success", "Selected entry deleted successfully.")
            
        # Close existing window
        if self.outgoing_shipments_window:
            self.outgoing_shipments_window.destroy()

        # Create a new window for displaying outgoing shipments
        self.outgoing_shipments_window = tk.Toplevel(self.root)
        self.outgoing_shipments_window.title("Outgoing Shipments")

        # Set the size of the window based on the screen dimensions
        screen_width = self.outgoing_shipments_window.winfo_screenwidth()
        screen_height = self.outgoing_shipments_window.winfo_screenheight()
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.9)
        window_x = int((screen_width - window_width) / 2)
        window_y = int((screen_height - window_height) / 2)

        self.outgoing_shipments_window.geometry(
            f"{window_width}x{window_height}+{window_x}+{window_y}")

        # Prepare outgoing shipment values
        tree_outgoing_shipments = ttk.Treeview(self.outgoing_shipments_window, columns=(
            "Date", "Shipping Destination", "Store", "Item No", "Item Name", "Quantity", "Average Price at Shipment", "Average Price at Shipment After Tax"))

        tree_outgoing_shipments.column("#0", width=20, anchor="w")
        tree_outgoing_shipments.column("#1", width=130, anchor="w")
        tree_outgoing_shipments.column("#2", width=130, anchor="w")
        tree_outgoing_shipments.column("#3", width=130, anchor="w")
        tree_outgoing_shipments.column("#4", width=130, anchor="w")
        tree_outgoing_shipments.column("#5", width=130, anchor="w")
        tree_outgoing_shipments.column("#6", width=130, anchor="e")
        tree_outgoing_shipments.column("#7", width=130, anchor="e")
        tree_outgoing_shipments.column("#8", width=130, anchor="e")

        tree_outgoing_shipments.heading("#0", text="#")
        tree_outgoing_shipments.heading("#1", text="Date")
        tree_outgoing_shipments.heading("#2", text="Shipping Destination")
        tree_outgoing_shipments.heading("#3", text="Store")
        tree_outgoing_shipments.heading("#4", text="Item No")
        tree_outgoing_shipments.heading("#5", text="Item Name")
        tree_outgoing_shipments.heading("#6", text="Quantity")
        tree_outgoing_shipments.heading("#7", text="Average Price at Shipment")
        tree_outgoing_shipments.heading(
            "#8", text="Average Price at Shipment After Tax")

        tree_outgoing_shipments.grid(row=0, column=0, padx=10, pady=10)

        # Add a button to delete the selected entry
        btn_delete_entry = tk.Button(
            self.outgoing_shipments_window, text="Delete Selected Shipment Entry", command=delete_selected_outgoing_shipment_entry)
        btn_delete_entry.grid(row=1, column=0, padx=10, pady=10)

        # Display outgoing shipments
        cursor.execute('''
            SELECT os.entry_date, os.shipping_to, os.store, os.item_no, os.item_name, os.quantity, os.average_price_at_shipment, os.average_price_at_shipment_after_tax
            FROM outgoing_shipments os
            ''')
        result = cursor.fetchall()
        for row in result:
            tree_outgoing_shipments.insert("", "end", values=row)

    def create_reporting_tab(self):
        self.reporting_tab = ttk.Frame(self.tabControl)
        self.tabControl.add(self.reporting_tab, text="Reporting")

        # Dropdown for Shipping Personnel
        self.shipping_to_var_report = StringVar(self.reporting_tab)
        self.shipping_to_var_report.set("ALL")  # Default value
        shipping_to_options_report = [
            "ALL", "USA FBA", "USA MFN", "CAN FBA", "CAN MFN"]
        shipping_to_menu_report = tk.OptionMenu(
            self.reporting_tab, self.shipping_to_var_report, *shipping_to_options_report)

        # Dropdown for Item Details
        self.item_details_var_report = StringVar(self.reporting_tab)
        self.item_details_var_report.set("ALL")  # Default value
        item_details_options_report = ["ALL"] + \
            self.get_all_store_item_no_item_name()

        item_details_menu_report = tk.OptionMenu(
            self.reporting_tab, self.item_details_var_report, *item_details_options_report)

        btn_generate_report = tk.Button(
            self.reporting_tab, text="Generate Report", command=self.generate_report)

        btn_export_to_excel = tk.Button(
            self.reporting_tab, text="Export Database to Excel", command=self.export_to_excel)

        shipping_to_menu_report.grid(row=0, column=0, padx=10, pady=10)
        item_details_menu_report.grid(row=0, column=1, padx=10, pady=10)
        btn_generate_report.grid(row=0, column=3, padx=10, pady=10)
        btn_export_to_excel.grid(row=1, column=4, padx=10, pady=10)

        # Treeview for displaying the report
        self.tree_report = ttk.Treeview(self.reporting_tab, columns=(
            "Store", "Item No", "Item Name", "Total Quantity", "Total Cost", "Total Cost After Tax"))
        self.tree_report.heading("#0", text="Store")
        self.tree_report.heading("#1", text="Item No")
        self.tree_report.heading("#2", text="Item Name")
        self.tree_report.heading("#3", text="Total Quantity")
        self.tree_report.heading("#4", text="Total Cost")
        self.tree_report.heading("#5", text="Total Cost After Tax")
        self.tree_report.grid(row=1, column=0, columnspan=4, padx=10, pady=10)

    def generate_report(self):
        # Clear previous data
        for row in self.tree_report.get_children():
            self.tree_report.delete(row)

        shipping_to = self.clean_input(self.shipping_to_var_report.get())
        item_details = self.item_details_var_report.get()

        # Modify the SQL query based on user selections
        query_params = []
        where_conditions = []

        if shipping_to != "ALL":
            where_conditions.append("os.shipping_to = ?")
            query_params.append(shipping_to)

        if item_details != "ALL":
            where_conditions.append(
                "os.store = ? AND os.item_no = ? AND os.item_name = ?")
            query_params.extend(list(eval(item_details)))

        where_clause = " AND ".join(
            where_conditions) if where_conditions else ""

        query = f'''
            SELECT os.store, os.item_no, os.item_name, SUM(os.quantity) as total_quantity, SUM(os.quantity * os.average_price_at_shipment) as total_cost, SUM(os.quantity * os.average_price_at_shipment_after_tax) as total_cost_after_tax 
            FROM outgoing_shipments os
            {"WHERE " + where_clause if where_clause else ""}
            GROUP BY os.store, os.item_no, os.item_name
        '''

        cursor.execute(query, query_params)
        result = cursor.fetchall()

        for row in result:
            self.tree_report.insert("", "end", values=row)

    def get_all_store_item_no_item_name(self):
        cursor.execute(
            'SELECT DISTINCT store, item_no, item_name FROM inventory')
        inventory = [row for row in cursor.fetchall()]

        cursor.execute(
            'SELECT DISTINCT store, item_no, item_name FROM outgoing_shipments')
        outgoing_shipments = [row for row in cursor.fetchall()]

        all_triplets = list(set(inventory + outgoing_shipments))
        return [triplet for triplet in all_triplets if triplet]

    def export_to_excel(self):
        # Export the data to an Excel file with three sheets

        # Define the file name
        excel_file = "inventory_data.xlsx"

        # Export incoming_items table
        incoming_items_df = pd.read_sql_query(
            'SELECT * FROM incoming_items', conn)
        incoming_items_df.to_excel(
            excel_file, sheet_name='incoming_items', index=False)

        # Export outgoing_shipments table
        outgoing_shipments_df = pd.read_sql_query(
            'SELECT * FROM outgoing_shipments', conn)
        outgoing_shipments_df.to_excel(
            excel_file, sheet_name='outgoing_shipments', index=False)

        # Export inventory table
        inventory_df = pd.read_sql_query(
            'SELECT * FROM inventory', conn)
        inventory_df.to_excel(excel_file, sheet_name='inventory', index=False)

        # Display success message
        messagebox.showinfo(
            "Success", "Data exported to Excel file successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

# Close the database connection
conn.close()
