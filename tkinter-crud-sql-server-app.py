import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.ttk import Combobox
import pyodbc
import pandas as pd


SERVER = 'localhost\SQLEXPRESS'  
DATABASE = 'HOMELAB' 



def connect_db():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        f'SERVER={SERVER};'
        f'DATABASE={DATABASE};'
        'Trusted_Connection=yes'  
        )
    return conn

# Funkcja tworząca tabelę w bazie danych (jeśli nie istnieje)
def create_table():
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Produkty')
            CREATE TABLE Produkty (
                id INT IDENTITY(1,1) PRIMARY KEY,
                nazwa NVARCHAR(100) NOT NULL,
                kategoria NVARCHAR(100),
                ilosc INT,
                cena DECIMAL(10, 2)
            );
        """)
        conn.commit()

# Funkcja dodająca produkt do bazy danych
def add_product(nazwa, kategoria, ilosc, cena):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Produkty (nazwa, kategoria, ilosc, cena) VALUES (?, ?, ?, ?)",
                       (nazwa, kategoria, ilosc, cena))
        conn.commit()

# Funkcja aktualizująca produkt w bazie danych
def update_product_in_db(product_id, nazwa, kategoria, ilosc, cena):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE Produkty 
            SET nazwa = ?, kategoria = ?, ilosc = ?, cena = ? 
            WHERE id = ?
        """, (nazwa, kategoria, ilosc, cena, product_id))
        conn.commit()

# Funkcja usuwająca produkt z bazy danych
def delete_product_from_db(product_id):
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Produkty WHERE id = ?", (product_id,))
        conn.commit()

# Funkcja eksportująca dane do pliku Excel
def export_to_excel(file_name="produkty_export.xlsx"):
    with connect_db() as conn:
        df = pd.read_sql_query("SELECT * FROM Produkty", conn)
        df.to_excel(file_name, index=False)
        messagebox.showinfo("Sukces", f"Dane wyeksportowane do {file_name}")

# Funkcja importująca dane z pliku Excel
def import_from_excel():
    file_path = filedialog.askopenfilename(title="Wybierz plik Excel", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    if file_path:
        df = pd.read_excel(file_path)
        with connect_db() as conn:
            df.to_sql('Produkty', conn, if_exists='append', index=False)
            messagebox.showinfo("Sukces", f"Dane zaimportowane z {file_path}")

# Funkcja do pobierania wszystkich produktów z bazy danych (z możliwością sortowania i filtrowania)
def get_products(sort_by=None, filter_by=None):
    query = "SELECT * FROM Produkty"
    
    # Filtrowanie
    if filter_by:
        query += f" WHERE nazwa LIKE '%{filter_by}%' OR kategoria LIKE '%{filter_by}%'"
    
    # Sortowanie
    if sort_by:
        query += f" ORDER BY {sort_by}"

    with connect_db() as conn:
        df = pd.read_sql_query(query, conn)
    return df

# Funkcja walidująca dane wejściowe
def validate_product_data(nazwa, kategoria, ilosc, cena):
    if not nazwa or not kategoria:
        return "Nazwa i kategoria nie mogą być puste."
    try:
        ilosc = int(ilosc)
        cena = float(cena)
    except ValueError:
        return "Ilość i cena muszą być liczbami."
    if ilosc < 0 or cena < 0:
        return "Ilość i cena muszą być liczbami dodatnimi."
    return None

# Funkcja dodająca produkt przez GUI
def add_product_gui():
    nazwa = entry_name.get()
    kategoria = entry_category.get()
    ilosc = entry_quantity.get()
    cena = entry_price.get()

    validation_error = validate_product_data(nazwa, kategoria, ilosc, cena)
    if validation_error:
        messagebox.showerror("Błąd", validation_error)
        return
    
    add_product(nazwa, kategoria, int(ilosc), float(cena))
    messagebox.showinfo("Sukces", "Produkt dodany pomyślnie!")
    update_product_list()

# Funkcja aktualizująca produkt w GUI
def update_product_gui():
    try:
        product_id = int(entry_product_id.get())
        nazwa = entry_name.get()
        kategoria = entry_category.get()
        ilosc = entry_quantity.get()
        cena = entry_price.get()

        validation_error = validate_product_data(nazwa, kategoria, ilosc, cena)
        if validation_error:
            messagebox.showerror("Błąd", validation_error)
            return
        
        update_product_in_db(product_id, nazwa, kategoria, int(ilosc), float(cena))
        messagebox.showinfo("Sukces", "Produkt zaktualizowany pomyślnie!")
        update_product_list()
    except ValueError:
        messagebox.showerror("Błąd", "Nieprawidłowy ID produktu.")

# Funkcja usuwająca produkt z GUI
def delete_product_gui():
    try:
        product_id = int(entry_product_id.get())
        delete_product_from_db(product_id)
        messagebox.showinfo("Sukces", "Produkt usunięty pomyślnie!")
        update_product_list()
    except ValueError:
        messagebox.showerror("Błąd", "Nieprawidłowy ID produktu.")

# Funkcja aktualizująca listę produktów w GUI
def update_product_list():
    filter_text = entry_filter.get()
    sort_by = combo_sort.get()

    products = get_products(sort_by=sort_by, filter_by=filter_text)
    
    # Usuwanie starych danych z listy
    for widget in frame_list.winfo_children():
        widget.destroy()

    # Wyświetlanie nowych danych
    for index, row in products.iterrows():
        tk.Label(frame_list, text=row['Id'], width=10).grid(row=index, column=0)
        tk.Label(frame_list, text=row['Name'], width=20).grid(row=index, column=1)
        tk.Label(frame_list, text=row['Category'], width=20).grid(row=index, column=2)
        tk.Label(frame_list, text=row['Quantity'], width=10).grid(row=index, column=3)
        tk.Label(frame_list, text=row['Price'], width=10).grid(row=index, column=4)

# Funkcja tworząca GUI
def create_gui():
    global entry_filter, combo_sort, entry_name, entry_category, entry_quantity, entry_price, entry_product_id, frame_list  # Zadeklarujmy zmienne globalnie

    root = tk.Tk()
    root.title("Zarządzanie Produktami")

    # Tworzenie tabeli w bazie danych, jeśli nie istnieje
    create_table()

    # Panel do dodawania/edycji/usuwania produktów
    frame_add_edit = tk.LabelFrame(root, text="Zarządzanie produktami", padx=10, pady=10)
    frame_add_edit.grid(row=0, column=0, padx=20, pady=20)

    tk.Label(frame_add_edit, text="ID produktu (do edycji/usuwania):").grid(row=0, column=0)
    entry_product_id = tk.Entry(frame_add_edit)
    entry_product_id.grid(row=0, column=1)

    tk.Label(frame_add_edit, text="Nazwa:").grid(row=1, column=0)
    entry_name = tk.Entry(frame_add_edit)
    entry_name.grid(row=1, column=1)

    tk.Label(frame_add_edit, text="Kategoria:").grid(row=2, column=0)
    entry_category = tk.Entry(frame_add_edit)
    entry_category.grid(row=2, column=1)

    tk.Label(frame_add_edit, text="Ilość:").grid(row=3, column=0)
    entry_quantity = tk.Entry(frame_add_edit)
    entry_quantity.grid(row=3, column=1)

    tk.Label(frame_add_edit, text="Cena:").grid(row=4, column=0)
    entry_price = tk.Entry(frame_add_edit)
    entry_price.grid(row=4, column=1)

    btn_add_product = tk.Button(frame_add_edit, text="Dodaj produkt", command=add_product_gui)
    btn_add_product.grid(row=5, column=0)

    btn_update_product = tk.Button(frame_add_edit, text="Zaktualizuj produkt", command=update_product_gui)
    btn_update_product.grid(row=5, column=1)

    btn_delete_product = tk.Button(frame_add_edit, text="Usuń produkt", command=delete_product_gui)
    btn_delete_product.grid(row=5, column=2)

    # Panel do wyświetlania produktów
    frame_list = tk.Frame(root)
    frame_list.grid(row=1, column=0, padx=20, pady=20)

    # Panel filtrów i sortowania
    frame_filter_sort = tk.LabelFrame(root, text="Filtruj i Sortuj", padx=10, pady=10)
    frame_filter_sort.grid(row=2, column=0, padx=20, pady=20)

    tk.Label(frame_filter_sort, text="Filtr (Nazwa/Kategoria):").grid(row=0, column=0)
    entry_filter = tk.Entry(frame_filter_sort)
    entry_filter.grid(row=0, column=1)

    tk.Label(frame_filter_sort, text="Sortowanie:").grid(row=1, column=0)
    combo_sort = Combobox(frame_filter_sort, values=["nazwa", "kategoria", "ilosc", "cena"])
    combo_sort.grid(row=1, column=1)

    # Przycisk do aktualizacji listy produktów
    btn_show_all = tk.Button(root, text="Pokaż produkty", command=update_product_list)
    btn_show_all.grid(row=3, column=0, pady=10)

    # Przycisk eksportu do Excela
    btn_export = tk.Button(root, text="Eksportuj do Excel", command=lambda: export_to_excel())
    btn_export.grid(row=4, column=0, pady=10)

    # Przycisk importu z Excela
    btn_import = tk.Button(root, text="Importuj z Excel", command=import_from_excel)
    btn_import.grid(row=5, column=0, pady=10)

    # Początkowe ładowanie produktów
    update_product_list()

    root.mainloop()

if __name__ == "__main__":
    create_gui()
