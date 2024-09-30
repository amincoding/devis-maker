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
        messagebox.showinfo("Success", f"Added: {product_name}")
        entry_name.delete(0, tk.END)
        entry_quantity.delete(0, tk.END)
        entry_price.delete(0, tk.END)
        unit_var.set("Select Unit")  # Reset the dropdown menu
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values for Quantity and Unit Price.")

# Function to save the product list to an Excel file
def save_to_excel():
    if not products:
        messagebox.showwarning("No Products", "No products to save. Add some products first.")
        return
    
    df = pd.DataFrame(products)
    print("DataFrame content:\n", df)  # Debug: Print the DataFrame content to check column names and data
    try:
        total = df['Total Price'].sum()

        # Creating a total row
        total_row = pd.DataFrame([{"Product Name": "Total", "Total Price": total}])

        # Concatenate the total row to the DataFrame
        df = pd.concat([df, total_row], ignore_index=True)

        # Add index column with header "N"
        df.index = df.index + 1  # Start index from 1 instead of 0
        df.index.name = 'N'

        # Save the DataFrame to an Excel file
        with pd.ExcelWriter('devis.xlsx', engine='openpyxl') as excel_writer:
            df.to_excel(excel_writer, index=True, sheet_name='Devis')

        messagebox.showinfo("Success", "Excel file created successfully.")
    except KeyError as e:
        messagebox.showerror("Error", f"Missing column in DataFrame: {e}")

# Initialize the main window
root = tk.Tk()
root.title("Devis bordereau")

# Add company logo at the top
logo_image = Image.open("logo.png")
logo_image = logo_image.resize((300, 300))  # Adjust size if necessary
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(root, image=logo_photo)
logo_label.grid(row=0, column=0, columnspan=2, pady=10)

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
unit_var.set("Select Unit")  # Default value
unit_options = ["Select Unit", "وحدة", "كلم", "كلغ", "متر", "غرام", "سم", "علبة"]
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
