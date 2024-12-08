import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # For handling images
import pandas as pd

# Initialize product list
products = []

# Function to add a new product
def add_product():
    product_name = entry_name.get()
    quantity = entry_quantity.get()
    unit_price = entry_price.get()
    unit = unit_var.get()

    if not product_name or not quantity or not unit_price or unit == "Select Unit":
        messagebox.showwarning("Input Error", "Please fill out all fields and select a unit.")
        return
    
    try:
        quantity = int(quantity)
        unit_price = float(unit_price)
        total_price = quantity * unit_price
        products.append({
            "Product Name": product_name,
            "Quantity": quantity,
            "Unit": unit,
            "Unit Price": unit_price,
            "Total Price": total_price
        })
        
        # Show the product in the side panel
        update_side_panel()
        
        messagebox.showinfo("Success", f"Added: {product_name}")
        entry_name.delete(0, tk.END)
        entry_quantity.delete(0, tk.END)
        entry_price.delete(0, tk.END)
        unit_var.set("وحدة")  # Reset the dropdown menu to the default value "وحدة"
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values for Quantity and Unit Price.")

# Function to update the side panel with the latest product list
def update_side_panel():
    # Clear the side panel first
    for widget in side_panel.winfo_children():
        widget.destroy()

    # Add new items to the side panel
    for product in products:
        # Change the separator from '-' to '|'
        product_label = tk.Label(side_panel, text=f"{product['Product Name']} | {product['Quantity']} | {product['Unit']} | {product['Total Price']} DA")
        product_label.pack(padx=5, pady=2)

# Function to save the product list to two Excel workspaces
def save_to_excel():
    if not products:
        messagebox.showwarning("No Products", "No products to save. Add some products first.")
        return
    
    df = pd.DataFrame(products)
    print("DataFrame content:\n", df)  # Debug: Print the DataFrame content to check column names and data
    
    try:
        # Create the 'priced_devis' sheet with all product details
        df_priced = df[["Product Name", "Quantity", "Unit", "Unit Price", "Total Price"]]

        # Create the 'devis' sheet with only quantities (no prices)
        df_devis = df[["Product Name", "Quantity", "Unit"]]

        # Creating a total row for both DataFrames
        total_priced = df_priced['Total Price'].sum()
        total_row_priced = pd.DataFrame([{"Product Name": "Total", "Total Price": total_priced}])

        total_devis = df_devis['Quantity'].sum()
        total_row_devis = pd.DataFrame([{"Product Name": "Total", "Quantity": total_devis}])

        # Concatenate the total rows to the respective DataFrames
        df_priced = pd.concat([df_priced, total_row_priced], ignore_index=True)
        df_devis = pd.concat([df_devis, total_row_devis], ignore_index=True)

        # Add index column with header "N" for both sheets
        df_priced.index = df_priced.index + 1
        df_priced.index.name = 'N'
        df_devis.index = df_devis.index + 1
        df_devis.index.name = 'N'

        # Save the DataFrames to an Excel file with two sheets
        with pd.ExcelWriter('devis.xlsx', engine='openpyxl') as excel_writer:
            df_priced.to_excel(excel_writer, index=True, sheet_name='priced_devis')
            df_devis.to_excel(excel_writer, index=True, sheet_name='devis')

        messagebox.showinfo("Success", "Excel file created successfully with two sheets.")
    except KeyError as e:
        messagebox.showerror("Error", f"Missing column in DataFrame: {e}")

# Initialize the main window
root = tk.Tk()
root.title("Devis bordereau")

# Create a side panel to display added products
side_panel = tk.Frame(root)
side_panel.grid(row=0, column=2, rowspan=7, padx=10, pady=5)

# Create the input fields and labels (start from row 1 now)
tk.Label(root, text="Product Name").grid(row=1, column=0, padx=10, pady=5)
entry_name = tk.Entry(root)
entry_name.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Quantity").grid(row=2, column=0, padx=10, pady=5)
entry_quantity = tk.Entry(root)
entry_quantity.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Unit").grid(row=3, column=0, padx=10, pady=5)

# Dropdown menu for selecting the unit
unit_var = tk.StringVar(root)
unit_var.set("وحدة")  # Default value
unit_options = ["وحدة", "كلم", "كلغ", "متر", "غرام", "سم", "علبة"]
unit_menu = tk.OptionMenu(root, unit_var, *unit_options)
unit_menu.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Unit Price").grid(row=4, column=0, padx=10, pady=5)
entry_price = tk.Entry(root)
entry_price.grid(row=4, column=1, padx=10, pady=5)

# Create the add product button
btn_add = tk.Button(root, text="Add Product", command=add_product)
btn_add.grid(row=5, column=0, columnspan=2, pady=10)

# Create the save to Excel button
btn_save = tk.Button(root, text="Save to Excel", command=save_to_excel)
btn_save.grid(row=6, column=0, columnspan=2, pady=10)

# Add footer text at the bottom
footer_label = tk.Label(root, text="All rights reserved for Amin Abdedaiem and his company LBS Software", font=("Arial", 8))
footer_label.grid(row=7, column=0, columnspan=2, pady=20)

# Start the main loop
root.mainloop()
