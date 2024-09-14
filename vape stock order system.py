# -*- coding: utf-8 -*-
"""
Created on Sun Jul 23 19:38:20 2023

@author: user
"""
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd


class StockManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.load_current_stock()

    def load_current_stock(self):
        try:
            self.current_stock = pd.read_excel(self.file_path)
        except FileNotFoundError:
            self.current_stock = pd.DataFrame({
                "Brand": ["Brand A", "Brand B", "Brand C"],
                "Flavour": ["Flavour 1", "Flavour 2", "Flavour 3"],
                "Estimated Quantity": [50, 30, 20]
            })
            self.save_current_stock()

    def save_current_stock(self):
        self.current_stock.to_excel(self.file_path, index=False)

    def add_product(self, brand, flavour, quantity):
        new_product = {"Brand": [brand], "Flavour": [flavour], "Estimated Quantity": [quantity]}
        new_product_df = pd.DataFrame(new_product)
        self.current_stock = pd.concat([self.current_stock, new_product_df], ignore_index=True)
        self.save_current_stock()

    def remove_product(self, index):
        self.current_stock = self.current_stock.drop(index)
        self.save_current_stock()

    def update_quantity(self, index, quantity):
        self.current_stock.at[index, "Estimated Quantity"] = quantity
        self.save_current_stock()

    def get_remaining_stock(self, brand, flavour):
        product = self.current_stock[(self.current_stock["Brand"] == brand) & (self.current_stock["Flavour"] == flavour)]
        if not product.empty:
            return product.iloc[0]["Estimated Quantity"]
        else:
            return None

    def get_total_stock(self):
        return self.current_stock["Estimated Quantity"].sum()


class VapeOrderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vape Order Management System")
        self.root.geometry("700x800")
        self.root.minsize(700, 650)
        self.root.maxsize(1920,675)

        self.current_stock_file = "vape_stock.xlsx"
        self.stock_manager = StockManager(self.current_stock_file)

        self.create_home_page()

    def create_home_page(self):
        self.home_frame = tk.Frame(self.root)
        self.home_frame.pack(fill="both", padx=20, pady=20)

        # Create grid layout
        self.home_frame.grid_columnconfigure(0, weight=1)
        self.home_frame.grid_columnconfigure(1, weight=1)

        self.home_frame.grid_rowconfigure(0, weight=1)
        self.home_frame.grid_rowconfigure(1, weight=1)
        self.home_frame.grid_rowconfigure(2, weight=1)

        # Scrollable table (using Treeview)
        self.table_frame = tk.Frame(self.home_frame)
        self.table_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        self.scrollbar_y = tk.Scrollbar(self.table_frame)
        self.scrollbar_y.pack(side="right", fill="y")

        columns = ["Brand", "Flavour", "Estimated Quantity"]
        self.stock_table = ttk.Treeview(self.table_frame, columns=columns, height=5, show="headings", yscrollcommand=self.scrollbar_y.set)
        self.stock_table.pack(fill="both", expand=True)

        self.scrollbar_y.config(command=self.stock_table.yview)

        # Define column headings
        for col in columns:
            self.stock_table.heading(col, text=col)

        # Set font size for the Treeview
        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica", 16), rowheight=45)

        # Set minimum size for table columns
        self.stock_table.column("#1", minwidth=200)
        self.stock_table.column("#2", minwidth=200)
        self.stock_table.column("#3", minwidth=100)
        self.stock_table.config(height=10)

        # Buttons
        self.order_report_button = tk.Button(self.home_frame, text="Order Report", font=("Helvetica", 16), command=self.open_order_report)
        self.order_report_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.delivery_button = tk.Button(self.home_frame, text="Delivery", font=("Helvetica", 16), command=self.open_delivery)
        self.delivery_button.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.add_product_button = tk.Button(self.home_frame, text="Add Product", font=("Helvetica", 16), command=self.add_product)
        self.add_product_button.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")

        self.remove_entry_button = tk.Button(self.home_frame, text="Remove Entry", font=("Helvetica", 16), command=self.remove_entry_with_confirmation)
        self.remove_entry_button.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")

        # Set minimum size for buttons
        self.order_report_button.config(height=5)
        self.delivery_button.config(height=5)
        self.add_product_button.config(height=3)
        self.remove_entry_button.config(height=3)

        # Define custom style for the quantity column
        self.stock_table.tag_configure("red", background="red", foreground="white", font=("Helvetica", 16))
        self.stock_table.tag_configure("orange", background="orange", foreground="white", font=("Helvetica", 16))
        self.stock_table.tag_configure("green", background="green", foreground="white", font=("Helvetica", 16))

        # Sort stock table by brand and flavour
        self.sort_stock_table()

    def sort_stock_table(self):
        # Sort DataFrame by brand and flavour
        self.stock_manager.current_stock.sort_values(by=["Brand", "Flavour"], inplace=True)

        # Update the stock table
        self.update_stock_table()

    def update_stock_table(self):
        # Clear existing items in the stock table
        self.stock_table.delete(*self.stock_table.get_children())

        # Add vape products to the table
        for index, row in self.stock_manager.current_stock.iterrows():
            brand = row["Brand"]
            flavour = row["Flavour"]
            quantity = row["Estimated Quantity"]
            item = self.stock_table.insert("", tk.END, values=(brand, flavour, quantity))

            # Validate quantity and set entry background color
            bg_color = self.get_entry_bg_color(quantity)

            # Set background color for quantity column (index 2)
            self.stock_table.item(item, values=(brand, flavour, quantity), tags=(bg_color,))

        # Apply the background color to the quantity column for all items
        for bg_color in ["red", "orange", "green"]:
            self.stock_table.tag_configure(bg_color, background=bg_color)

    def get_entry_bg_color(self, quantity):
        if quantity < 2 or quantity == 0:
            return "red"
        elif quantity < 10:
            return "orange"
        else:
            return "green"

    def add_product(self):
        # Create a new product entry form
        self.product_entry_window = tk.Toplevel(self.root)
        self.product_entry_window.title("Add New Product")
        self.product_entry_window.geometry("400x200")

        # Labels and Entry fields
        tk.Label(self.product_entry_window, text="Brand:").grid(row=0, column=0, padx=10, pady=10)
        self.brand_entry = tk.Entry(self.product_entry_window)
        self.brand_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.product_entry_window, text="Flavour:").grid(row=1, column=0, padx=10, pady=10)
        self.flavour_entry = tk.Entry(self.product_entry_window)
        self.flavour_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.product_entry_window, text="Estimated Quantity:").grid(row=2, column=0, padx=10, pady=10)
        self.quantity_entry = tk.Entry(self.product_entry_window)
        self.quantity_entry.grid(row=2, column=1, padx=10, pady=10)

        # Add button
        tk.Button(self.product_entry_window, text="Add", command=self.save_new_product).grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def save_new_product(self):
        # Get the values from the entry fields
        brand = self.brand_entry.get()
        flavour = self.flavour_entry.get()
        quantity = self.quantity_entry.get()

        # Validate quantity as a positive integer
        try:
            quantity = int(quantity)
            if quantity < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Quantity", "Estimated Quantity must be a positive integer.")
            return

        # Add the new product to the current stock
        self.stock_manager.add_product(brand, flavour, quantity)

        # Sort stock table by brand and flavour
        self.sort_stock_table()

        # Close the popup window
        self.product_entry_window.destroy()

    def remove_entry_with_confirmation(self):
        selected_item = self.stock_table.selection()
        if not selected_item:
            return

        # Show a confirmation popup before removing the entry
        confirm = messagebox.askyesno("Confirmation", "Are you sure you want to remove this entry?")
        if confirm:
            self.remove_entry(selected_item)

    def remove_entry(self, selected_item):
        # Get the selected row index
        row_index = int(self.stock_table.index(selected_item[0]))

        # Remove the product from the current stock
        self.stock_manager.remove_product(row_index)

        # Update the stock table
        self.update_stock_table()

    def open_order_report(self):
        # Feature 1: Placeholder for the Order Report screen
        messagebox.showinfo("Feature 1", "This is the Order Report screen.")

    def open_delivery(self):
        # Feature 2: Placeholder for the Delivery screen
        messagebox.showinfo("Feature 2", "This is the Delivery screen.")

    # Feature 6: Show remaining stock by brand and flavour
    def show_remaining_stock(self, brand, flavour):
        remaining_quantity = self.stock_manager.get_remaining_stock(brand, flavour)
        if remaining_quantity is not None:
            messagebox.showinfo("Remaining Stock", f"Remaining stock for {brand} - {flavour}: {remaining_quantity}")
        else:
            messagebox.showinfo("Product Not Found", f"{brand} - {flavour} is not found in the current stock.")

    # Feature 7: Show total stock quantity for all products
    def show_total_stock(self):
        total_stock = self.stock_manager.get_total_stock()
        messagebox.showinfo("Total Stock", f"Total stock quantity for all products: {total_stock}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VapeOrderApp(root)
    root.mainloop()
