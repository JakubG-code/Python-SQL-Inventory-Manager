# ============================================
# SQL Inventory management v.1
# ============================================

import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.ttk import Combobox
import pyodbc
import pandas as pd


SERVER = 'localhost\SQLEXPRESS'  # < Fill this
DATABASE = 'HOMELAB_CorpHR' # < Fill this


def connect_db():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        f'SERVER={SERVER};'
        f'DATABASE={DATABASE};'
        'Trusted_Connection=yes'  
        )
    return conn

# Create database table if it does not exist
def create_table():
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Products')
            CREATE TABLE Products (
                Id INT IDENTITY(1,1) PRIMARY KEY,
                Name NVARCHAR(100) NOT NULL,
                Category NVARCHAR(100),
                Quantity INT,
                Price DECIMAL(10, 2)
            );
        """)
        conn.commit()

# Add a product to the database
def add_product(Name, Category, Quantity, Price):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Products (Name, Category, Quantity, Price) VALUES (?, ?, ?, ?)",
                       (Name, Category, Quantity, Price))
        conn.commit()

# Update an existing product in the database
def update_product_in_db(product_id, Name, Category, Quantity, Price):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE Products 
            SET Name = ?, Category = ?, Quantity = ?, Price = ? 
            WHERE Id = ?
        """, (Name, Category, Quantity, Price, product_id))
        conn.commit()

# Delete a product from the database
def delete_product_from_db(product_id):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Products WHERE Id = ?", (product_id,))
        conn.commit()

# Export database data to Excel
def export_to_excel(file_name="products_export.xlsx"):
    with connect_db() as conn:
        df = pd.read_sql_query("SELECT * FROM Products", conn)
        df.to_excel(file_name, index=False)
        messagebox.showinfo("Success", f"Data exported to {file_name}")

# Import data from Excel
def import_from_excel():
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    if file_path:
        df = pd.read_excel(file_path)
        with connect_db() as conn:
            df.to_sql('Products', conn, if_exists='append', index=False)
            messagebox.showinfo("Success", f"Data imported from {file_path}")

# Retrieve products with optional filtering and sorting
def get_products(sort_by=None, filter_by=None):
    query = "SELECT * FROM Products"
    
    if filter_by:
        query += f" WHERE Name LIKE '%{filter_by}%' OR Category LIKE '%{filter_by}%'"
    
    if sort_by:
        query += f" ORDER BY {sort_by}"

    with connect_db() as conn:
        df = pd.read_sql_query(query, conn)
    return df

# Validate user input data
def validate_product_data(Name, Category, Quantity, Price):
    if not Name or not Category:
        return "Name and category cannot be empty."
    try:
        Quantity = int(Quantity)
        Price = float(Price)
    except ValueError:
        return "Quantity and price must be numbers."
    if Quantity < 0 or Price < 0:
        return "Quantity and price must be positive."
    return None

# Add product using GUI
def add_product_gui():
    Name = entry_name.get()
    Category = entry_category.get()
    Quantity = entry_quantity.get()
    Price = entry_price.get()

    validation_error = validate_product_data(Name, Category, Quantity, Price)
    if validation_error:
        messagebox.showerror("Error", validation_error)
        return
    
    add_product(Name, Category, int(Quantity), float(Price))
    messagebox.showinfo("Success", "Product added successfully!")
    update_product_list()

# Update product using GUI
def update_product_gui():
    try:
        product_id = int(entry_product_id.get())
        Name = entry_name.get()
        Category = entry_category.get()
        Quantity = entry_quantity.get()
        Price = entry_price.get()

        validation_error = validate_product_data(Name, Category, Quantity, Price)
        if validation_error:
            messagebox.showerror("Error", validation_error)
            return
        
        update_product_in_db(product_id, Name, Category, int(Quantity), float(Price))
        messagebox.showinfo("Success", "Product updated successfully!")
        update_product_list()
    except ValueError:
        messagebox.showerror("Error", "Invalid product ID.")

# Delete product using GUI
def delete_product_gui():
    try:
        product_id = int(entry_product_id.get())
        delete_product_from_db(product_id)
        messagebox.showinfo("Success", "Product deleted successfully!")
        update_product_list()
    except ValueError:
        messagebox.showerror("Error", "Invalid product ID.")

# Refresh product list displayed in GUI
def update_product_list():
    filter_text = entry_filter.get()
    sort_by = combo_sort.get()

    products = get_products(sort_by=sort_by, filter_by=filter_text)
    
    for widget in frame_list.winfo_children():
        widget.destroy()

    for index, row in products.iterrows():
        tk.Label(frame_list, text=row['Id'], width=10).grid(row=index, column=0)
        tk.Label(frame_list, text=row['Name'], width=20).grid(row=index, column=1)
        tk.Label(frame_list, text=row['Category'], width=20).grid(row=index, column=2)
        tk.Label(frame_list, text=row['Quantity'], width=10).grid(row=index, column=3)
        tk.Label(frame_list, text=row['Price'], width=10).grid(row=index, column=4)

# Create application GUI
def create_gui():
    global entry_filter, combo_sort, entry_name, entry_category, entry_quantity, entry_price, entry_product_id, frame_list

    root = tk.Tk()
    root.title("Product Management")

    create_table()

    # Product management panel
    frame_add_edit = tk.LabelFrame(root, text="Product Management", padx=10, pady=10)
    frame_add_edit.grid(row=0, column=0, padx=20, pady=20)

    tk.Label(frame_add_edit, text="Product ID (edit/delete):").grid(row=0, column=0)
    entry_product_id = tk.Entry(frame_add_edit)
    entry_product_id.grid(row=0, column=1)

    tk.Label(frame_add_edit, text="Name:").grid(row=1, column=0)
    entry_name = tk.Entry(frame_add_edit)
    entry_name.grid(row=1, column=1)

    tk.Label(frame_add_edit, text="Category:").grid(row=2, column=0)
    entry_category = tk.Entry(frame_add_edit)
    entry_category.grid(row=2, column=1)

    tk.Label(frame_add_edit, text="Quantity:").grid(row=3, column=0)
    entry_quantity = tk.Entry(frame_add_edit)
    entry_quantity.grid(row=3, column=1)

    tk.Label(frame_add_edit, text="Price:").grid(row=4, column=0)
    entry_price = tk.Entry(frame_add_edit)
    entry_price.grid(row=4, column=1)

    btn_add_product = tk.Button(frame_add_edit, text="Add Product", command=add_product_gui)
    btn_add_product.grid(row=5, column=0)

    btn_update_product = tk.Button(frame_add_edit, text="Update Product", command=update_product_gui)
    btn_update_product.grid(row=5, column=1)

    btn_delete_product = tk.Button(frame_add_edit, text="Delete Product", command=delete_product_gui)
    btn_delete_product.grid(row=5, column=2)

    # Product display panel
    frame_list = tk.Frame(root)
    frame_list.grid(row=1, column=0, padx=20, pady=20)

    # Filter and sorting panel
    frame_filter_sort = tk.LabelFrame(root, text="Filter and Sort", padx=10, pady=10)
    frame_filter_sort.grid(row=2, column=0, padx=20, pady=20)

    tk.Label(frame_filter_sort, text="Filter (Name/Category):").grid(row=0, column=0)
    entry_filter = tk.Entry(frame_filter_sort)
    entry_filter.grid(row=0, column=1)

    tk.Label(frame_filter_sort, text="Sorting:").grid(row=1, column=0)
    combo_sort = Combobox(frame_filter_sort, values=["Name", "Category", "Quantity", "Price"])
    combo_sort.grid(row=1, column=1)

    btn_show_all = tk.Button(root, text="Show Products", command=update_product_list)
    btn_show_all.grid(row=3, column=0, pady=10)

    btn_export = tk.Button(root, text="Export to Excel", command=lambda: export_to_excel())
    btn_export.grid(row=4, column=0, pady=10)

    btn_import = tk.Button(root, text="Import from Excel", command=import_from_excel)
    btn_import.grid(row=5, column=0, pady=10)

    update_product_list()

    root.mainloop()

if __name__ == "__main__":
    create_gui()
