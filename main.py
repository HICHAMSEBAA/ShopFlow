#!/home/hicham/Hicham/Python/ExcelCrafter/VExcelCrafter/bin/python3
import tkinter as tk
from tkinter import ttk
import openpyxl
import os
from datetime import datetime
from tkinter import messagebox
import uuid
import threading

class ExcelCrafterApp:
    def __init__(self, root):
        # Initialize the main window (root) and set the title
        self.root = root
        self.root.title("Product Management")
        
        # Initialize the styles for the application
        self.style = ttk.Style(root)
        root.tk.call("source", "forest-light.tcl")  # Load the custom theme from a Tcl file
        self.style.theme_use("forest-light")        # Apply the "forest-light" theme
        
        # Initialize product-related dropdown options (ComboBox lists)
        self.combo_list_ProductCategory = ["Adapter", "Cable", "Car Accessory", "Car Charger", "Charger", "Flash USB", "Earphone", "Memory Card", "Shock Glasses", "SIM Card", "Other"] 
        self.combo_list_ProductType = ["Auxiliary", "Bluetooth", "iPhone", "Micro USB", "Simple Phone", "Smart Phone", "Type C", "Wired", "Other"]
        self.combo_list_MemoryType = ["1 GB", "2 GB", "4 GB", "8 GB", "16 GB", "32 GB", "64 GB", "128 GB", "256 GB"]
        self.combo_list_SIMCardType = ["GOLD", "LEGEND", "MOBTASIM", "SAMA", "Normal", "Other"]

        # Initialize a variable to keep track of the selected item in the Treeview
        self.selected_item = None
        
        # Create the main frame that will contain all other widgets
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)  # Expand the frame to fit the window

        # Create a notebook (tabbed interface) within the main frame
        self.create_notebook(self.main_frame)
        
        # Add a title to the first tab (Product Page)
        self.page_title(self.tab1, "The Product Page")
        
        # Add product management widgets (input fields, buttons, etc.) to the first tab
        self.add_product_widgets(self.tab1)
        
        # Create a Treeview widget (a table-like structure) in the first tab to display product data
        self.create_treeview(self.tab1, path="products.xlsx", columns=["Name", "Category", "Type", "Quantity", "Price", "Date", "Time"])
        
        # Create a search bar in the first tab to allow users to search through products
        self.create_search(self.tab1, path="products.xlsx")


    
# Sales part ----------------------------------------------------------------------------

    def open_sales_window(self):
        """
        Opens a new window for processing a sale of the selected product.
        """
        # Get the selected item from the Treeview
        selected_item = self.treeview.selection()
        
        # If no item is selected, show an error message and exit the function
        if not selected_item:
            messagebox.showerror("Error", "Please select a product first.")
            return
        
        # Try to retrieve the details of the selected product
        try:
            product_details = self.treeview.item(selected_item, "values")
            # Unpack the product details
            product_name, category, product_type, available_quantity, price, date, time = product_details
            
            # Check if the product is out of stock
            if int(available_quantity) <= 0:
                messagebox.showinfo("Info", "The product is out of stock.")
                return
        except Exception as e:
            # If there's an error retrieving product details, show an error message
            messagebox.showerror("Error", f"Failed to retrieve product details: {e}")
            return
        
        # Create a new top-level window for the sales operation
        sales_window = tk.Toplevel(self.root)
        sales_window.title("Sales Operation")
        
        # Determine which type of sales window to create based on the product category
        if category == "SIM Card":
            self.create_sales_window(
                sales_window, 
                product_name, 
                category, 
                product_type, 
                available_quantity, 
                price, 
                product_type_list=self.combo_list_SIMCardType, 
                window_title="Sales SIMCards"
            )
        else:
            self.create_sales_window(
                sales_window, 
                product_name, 
                category, 
                product_type, 
                available_quantity, 
                price, 
                product_type_list=self.combo_list_SIMCardType, 
                window_title="Sales Products"
            )

    def create_sales_window(self, sales_window, product_name, category, product_type, available_quantity, price, product_type_list, window_title):
        """
        Creates and configures the sales window with all necessary widgets.
        
        Parameters:
        - sales_window: The Toplevel window where the sales operation will take place.
        - product_name: Name of the product being sold.
        - category: Category of the product.
        - product_type: Type of the product.
        - available_quantity: Current stock quantity of the product.
        - price: Price of the product.
        - product_type_list: List of product types for the Combobox.
        - window_title: Title of the sales window.
        """
        # Create the main frame within the sales window
        main_frame = ttk.Frame(sales_window)
        main_frame.pack(fill="both", expand=True)

        # Create a labeled frame to hold the sales widgets
        frame_widgets = ttk.LabelFrame(main_frame, text=window_title)
        frame_widgets.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        
        # ------------------ Product Name ------------------
        # Label for product name
        name_label = ttk.Label(frame_widgets, text="Name:")
        name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Entry widget for product name (disabled as it's pre-filled)
        product_name_entry = ttk.Entry(frame_widgets)
        product_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        product_name_entry.insert(0, product_name)  # Set the initial value
        product_name_entry.config(state="disabled")  # Disable the entry to prevent editing
        
        # ------------------ Category ------------------
        # Label for category
        category_label = ttk.Label(frame_widgets, text="Category:")
        category_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        # Combobox for category (disabled as it's pre-filled)
        category_combobox = ttk.Combobox(
            frame_widgets, 
            values=self.combo_list_ProductCategory, 
            state="readonly"
        )
        category_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        category_combobox.set(category)  # Set the initial value
        category_combobox.config(state="disabled")  # Disable editing
        
        # ------------------ Type ------------------
        # Label for product type
        type_label = ttk.Label(frame_widgets, text="Type:")
        type_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        # Combobox for product type
        type_combobox = ttk.Combobox(
            frame_widgets, 
            values=product_type_list, 
            state="readonly"
        )
        type_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        # Set the initial value and configure state based on window title
        if window_title != "Sales SIMCards":
            type_combobox.set(product_type)
            type_combobox.config(state="disabled")  # Disable if not SIM Cards
        else:
            type_combobox.set("Normal")  # Default value for SIM Cards
            type_combobox.config(state="normal")  # Enable editing
        
        # ------------------ Quantity ------------------
        # Label for quantity to sell
        quantity_label = ttk.Label(frame_widgets, text="Quantity:")
        quantity_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        
        # Spinbox for selecting quantity to sell (from 0 to available_quantity)
        quantity_spinbox = ttk.Spinbox(
            frame_widgets, 
            from_=0, 
            to=int(available_quantity)
        )
        quantity_spinbox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        # ------------------ Price ------------------
        # Label for price
        price_label = ttk.Label(frame_widgets, text="Price:")
        price_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        
        # Spinbox for setting the price (from 0.0 to 10000.0 with increments of 50)
        price_spinbox = ttk.Spinbox(
            frame_widgets, 
            from_=0.0, 
            to=10000.0, 
            increment=50
        )
        price_spinbox.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        
        # Set the initial price value and configure state based on window title
        if window_title != "Sales SIMCards":
            price_spinbox.insert(0, price)  # Set the initial price
            price_spinbox.config(state="disabled")  # Disable if not SIM Cards
        else:
            price_spinbox.insert(0, 0)  # Default price for SIM Cards
            price_spinbox.config(state="normal")  # Enable editing
        
        # ------------------ Sell Button ------------------
        # Button to execute the sell operation
        sell_button = ttk.Button(
            frame_widgets, 
            text="Sell", 
            command=lambda: self.sell_product_from_window(
                product_name, 
                category_combobox, 
                type_combobox, 
                quantity_spinbox, 
                sales_window
            )
        )
        sell_button.grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="nsew")

    def sell_product_from_window(self, product_name, category_combobox, type_combobox, quantity_spinbox, window):
        """
        Processes the sale of a product by updating the inventory and recording the sale.
        
        Parameters:
        - product_name: Name of the product being sold.
        - category_combobox: Combobox widget for product category.
        - type_combobox: Combobox widget for product type.
        - quantity_spinbox: Spinbox widget for quantity to sell.
        - window: The sales window to be closed after processing the sale.
        """
        # Validate the quantity entered by the user
        try:
            quantity_sold = int(quantity_spinbox.get())  # Get the quantity to sell
            if quantity_sold <= 0:
                # If quantity is not positive, show an error message
                messagebox.showerror("Error", "Please enter a valid quantity.")
                return
        except Exception as e:
            # If there's an error parsing the quantity, show an error message
            messagebox.showerror("Error", f"An error occurred while processing the quantity: {e}")
            return
        
        # Path to the products Excel file
        path = "products.xlsx"
        
        try:
            # Load the Excel workbook and select the active sheet
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active

            # Iterate through the rows to find the matching product
            for row in sheet.iter_rows(min_row=2, values_only=False):
                if row[0].value == product_name:
                    available_quantity = int(row[3].value)  # Current stock
                    
                    # Check if there is sufficient stock to fulfill the sale
                    if quantity_sold > available_quantity:
                        # If not enough stock, show a warning message
                        messagebox.showwarning("Insufficient Stock", f"Only {available_quantity} units available.")
                        return
                    else:
                        # Deduct the sold quantity from available stock
                        row[3].value = available_quantity - quantity_sold
                        # Prepare the updated data for the Treeview
                        new_data = [row[i].value for i in range(len(row))]
                        break
            else:
                # If the product was not found in the Excel sheet, show a warning
                messagebox.showwarning("Product Not Found", "The specified product was not found.")
                return

            # Save the updated workbook back to the file
            workbook.save(path)
            
            # Record the sale in the sales log
            self.record_sale(
                product_name, 
                str(category_combobox.get()), 
                str(type_combobox.get()), 
                quantity_sold
            )
            
            # Close the sales window after successful sale
            window.destroy()

            # Notify the user of the successful sale
            messagebox.showinfo("Success", "Product sold successfully!")
            
            # Update the Treeview to reflect the new quantity
            for item in self.treeview.get_children():
                if self.treeview.item(item, "values")[0] == product_name:
                    self.treeview.item(item, values=new_data)
                    break
            
            # Reset any necessary variables or states
            self.reset()
            
        except Exception as e:
            # If any error occurs during the sale process, show an error message
            messagebox.showerror("Error", f"An error occurred while selling the product: {e}")

    def record_sale(self, product_name, category, type, quantity_sold):
        """
        Records the sale details into the sales Excel file.
        
        Parameters:
        - product_name: Name of the product sold.
        - category: Category of the product.
        - type: Type of the product.
        - quantity_sold: Quantity of the product sold.
        """
        # Path to the sales Excel file
        sales_path = "sales.xlsx"
        
        # Get the current date and time for the sale record
        current_date = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
        
        # Create a list representing the sale record
        sale_record = [product_name, category, type, quantity_sold, current_date, current_time]

        try:
            if not os.path.exists(sales_path):
                # If the sales file doesn't exist, create it and add headers
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                headers = ["Product Name", "Category", "Type", "Quantity Sold", "Date", "Time"]
                sheet.append(headers)
            else:
                # If the sales file exists, load it
                workbook = openpyxl.load_workbook(sales_path)
                sheet = workbook.active

            # Append the sale record to the sales sheet
            sheet.append(sale_record)
            # Save the workbook
            workbook.save(sales_path)
        except Exception as e:
            # If there's an error recording the sale, show an error message
            messagebox.showerror("Error", f"An error occurred while recording the sale: {e}")

# ---------------------------------------------------------------------------------------



# GUI PART ------------------------------------------------------------------------------

    def create_notebook(self, frame):
        """Creates the notebook and adds tabs to it."""
        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True, pady=10, padx=10)

        # Create and add tabs
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Products")


    def add_product_widgets(self, frame):
        """Creates and adds the widgets for the product management section."""
        frame_widgets = ttk.LabelFrame(frame, text="Products")
        frame_widgets.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # Product Name
        name_label = ttk.Label(frame_widgets, text="Name:")
        name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.product_name_entry = ttk.Entry(frame_widgets)
        self.product_name_entry.bind("<FocusIn>", lambda e: self.clear_entry(self.product_name_entry, "Name"))
        self.product_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Category
        category_label = ttk.Label(frame_widgets, text="Category:")
        category_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.category_combobox = ttk.Combobox(frame_widgets, values=self.combo_list_ProductCategory, state="readonly")
        self.category_combobox.current(0)  # Default to the first category
        self.category_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.category_combobox.bind("<<ComboboxSelected>>", self.on_category_selected)  # Bind category selection event

        # Type
        type_label = ttk.Label(frame_widgets, text="Type:")
        type_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.type_combobox = ttk.Combobox(frame_widgets, values=self.combo_list_ProductType, state="readonly")
        self.type_combobox.current(0)
        self.type_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.type_combobox.config(state="disabled")

        # Memory Type
        type_label_ = ttk.Label(frame_widgets, text="Memory Type:")
        type_label_.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.memoryType_combobox = ttk.Combobox(frame_widgets, values=self.combo_list_MemoryType, state="readonly")
        self.memoryType_combobox.current(0)
        self.memoryType_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.memoryType_combobox.config(state="disabled")

        # Quantity
        quantity_label = ttk.Label(frame_widgets, text="Quantity:")
        quantity_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.quantity_spinbox = ttk.Spinbox(frame_widgets, from_=1, to=1000)
        self.quantity_spinbox.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        # Price
        price_label = ttk.Label(frame_widgets, text="Price:")
        price_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.price_spinbox = ttk.Spinbox(frame_widgets, from_=0.0, to=10000.0, increment=50)
        self.price_spinbox.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        # Buttons for product operations
        self.insert_button = ttk.Button(frame_widgets, text="Insert", command=self.insert_product)
        self.insert_button.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")

        self.update_button = ttk.Button(frame_widgets, text="Update", command=self.update_product)
        self.update_button.grid(row=6, column=1, padx=5, pady=5, sticky="nsew")
        self.update_button.config(state="disabled")

        self.delete_button = ttk.Button(frame_widgets, text="Delete", command=self.delete_product)
        self.delete_button.grid(row=7, column=0, padx=5, pady=5, sticky="nsew")
        self.delete_button.config(state="disabled")

        self.cancel_button = ttk.Button(frame_widgets, text="Cancel", command=self.reset)
        self.cancel_button.grid(row=7, column=1, padx=5, pady=5, sticky="nsew")
        self.cancel_button.config(state="disabled")

        # Separator for better UI structure
        separator = ttk.Separator(frame_widgets)
        separator.grid(row=8, column=0, columnspan=2, padx=(20, 10), pady=10, sticky="ew")

        # Sales Button
        self.sales_button = ttk.Button(frame_widgets, text="Sale", command=self.open_sales_window)
        self.sales_button.grid(row=9, column=0, columnspan=2, padx=5, pady=10, sticky="nsew")
        self.sales_button.config(state="disable")

        # Configure grid weights for better resizing behavior
        frame_widgets.columnconfigure(0, weight=1)
        frame_widgets.columnconfigure(1, weight=1)


    def create_treeview(self, frame, path, columns):
        """Creates the treeview widget for displaying product data."""
        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=1, column=2, padx=20, pady=10, sticky="nsew")
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side="right", fill="y")

        # Configure treeview
        self.treeview = ttk.Treeview(tree_frame, show="headings", yscrollcommand=tree_scroll.set, columns=columns, height=13)
        for col in columns:
            self.treeview.column(col, width=100, anchor="center")
            self.treeview.heading(col, text=col, command=lambda _col=col: self.sort_treeview(_col, False))

        self.treeview.pack(fill="both", expand=True)
        tree_scroll.config(command=self.treeview.yview)

        # Bind selection event to Treeview
        self.treeview.bind("<<TreeviewSelect>>", self.on_item_selected)

        # Load data into the treeview
        self.load_data(path=path)


    def page_title(self, frame, title):
        """Creates and displays the page title."""
        # Create a frame for the title to manage layout
        title_frame = ttk.Frame(frame)
        title_frame.grid(row=0, column=1, columnspan=2, padx=20, pady=10, sticky="ew")

        # Create a label with large text for the title
        title_label = ttk.Label(title_frame, text=title, font=("Helvetica", 24, "bold"))
        title_label.grid(row=0, column=0, padx=10, pady=10, sticky="n")
        title_label.config(anchor="center", foreground="#333333")

        # Adjust column weights to center the title
        title_frame.columnconfigure(0, weight=1)


    def create_search(self, frame, path):
        """Creates the search bar for filtering products in the treeview."""
        search_frame = ttk.Frame(frame)
        search_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        search_label = ttk.Label(search_frame, text="Search:")
        search_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.search_data)  # Bind search function to key release events

        # Configure grid weight to ensure proper resizing
        search_frame.columnconfigure(1, weight=1)

    # GUI PART ------------------------------------------------------------------------------


# FUN PART ------------------------------------------------------------------------------

# TABLE FUNCTIONALTY 

    def sort_treeview(self, col, reverse):
        """Sorts the Treeview column when the heading is clicked."""
        try:
            data = [(self.treeview.set(k, col), k) for k in self.treeview.get_children('')]
            data.sort(reverse=reverse)
            for index, (val, k) in enumerate(data):
                self.treeview.move(k, '', index)
            self.treeview.heading(col, command=lambda: self.sort_treeview(col, not reverse))
        except Exception as e:
            print(f"Error sorting column {col}: {e}")
    
    def on_item_selected(self, event):
        selected_item = self.treeview.selection()  # Get the selected item
        if selected_item:
            # Retrieve the values of the selected item
            item_values = self.treeview.item(selected_item, "values")
            
            # Populate the input fields with the selected item's data
            self.product_name_entry.config(state="normal")
            self.product_name_entry.delete(0, "end")
            self.product_name_entry.insert(0, item_values[0])  # Name
            self.product_name_entry.config(state="disabled")
            
            self.category_combobox.set(item_values[1])  # Category
            
            if  item_values[1] == "Flash USB" or item_values[1] == "Memory Card":
                self.memoryType_combobox.set(item_values[2])  # Type
                self.memoryType_combobox.config(state="normal")
                self.type_combobox.config(state="disable")
            else:
                self.type_combobox.set(item_values[2])  # Type
                self.type_combobox.config(state="normal")
                self.memoryType_combobox.config(state="disable")
            
            self.quantity_spinbox.delete(0, "end")
            self.quantity_spinbox.insert(0, item_values[3])  # Quantity
            
            self.price_spinbox.delete(0, "end")
            self.price_spinbox.insert(0, item_values[4])  # Price
            
            # Disable the Insert button and enable the Update and Delete buttons
            self.insert_button.config(state="disabled")
            self.update_button.config(state="normal")
            self.delete_button.config(state="normal")
            self.canceled_button.config(state="normal")
            self.Sales_button.config(state="normal")
            
            # Store the selected item's ID for future updates
            self.selected_item = item_values[0]
    
    def load_data(self, path):
        """Loads data from the Excel file and inserts it into the treeview."""
        try:
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active

            self.all_data = list(sheet.values)  # Store all data for searching

            headers = self.all_data[0]
            self.treeview["columns"] = headers
            for col in headers:
                self.treeview.heading(col, text=col, command=lambda _col=col: self.sort_treeview(_col, False))
                self.treeview.column(col, width=100, anchor="center")

            self.treeview.delete(*self.treeview.get_children())  # Clear existing data

            for value_tuple in self.all_data[1:]:
                self.treeview.insert('', tk.END, values=value_tuple)
        except FileNotFoundError:
            # If the file doesn't exist, create it with headers
            self.create_excel_file(path)
        except Exception as e:
            print(f"Error loading data: {e}")

# TABLE FUNCTIONALTY 

    def create_excel_file(self, path):
        """Creates a new Excel file with headers."""
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            headers = ["Name", "Category", "Type", "Quantity", "Price", "Date", "Time"]
            sheet.append(headers)
            workbook.save(path)
            self.all_data = [headers]
        except Exception as e:
            print(f"Error creating Excel file: {e}")
    # "Flash USB", "Earphone", "Memory Card"
    def on_category_selected(self, event):
        selected_category = self.category_combobox.get()
        if selected_category == "Flash USB" or selected_category == "Memory Card":  # Specify the category you want to check
            self.memoryType_combobox.config(state="normal")  # Enable the type combobox
            self.type_combobox.config(state="disable")  # Disable the type combobox
        else:
            self.type_combobox.config(state="normal")  # normal the type combobox
            self.memoryType_combobox.config(state="disable")  # Enable the type combobox

    def clear_entry(self, entry, default_text):
        if entry.get() == default_text:
            entry.delete(0, "end")
 
# UPDATE FUN
    
    def update_product(self):
        if not self.selected_item:
            messagebox.showwarning("Select Item", "Please select an item to update.")
            return

        # Show confirmation dialog
        response = messagebox.askyesno("Confirm Update", "Are you sure you want to update this product?")
    
        if response:
            validated_data = self.validate_inputs()
            if not validated_data:
                return

            name, category, product_type, quantity, price = validated_data
            current_date = datetime.now().strftime("%Y-%m-%d")
            current_time = datetime.now().strftime("%H:%M:%S")
            
            # Prepare new data
            new_data = [name, category, product_type, quantity, price, current_date, current_time]
            
            # Update Treeview
            for item in self.treeview.get_children():
                if self.treeview.item(item, "values")[0] == self.selected_item:
                    self.treeview.item(item, values=new_data)
                    break
            
            # Update Excel file
            path = "products.xlsx"
            try:
                workbook = openpyxl.load_workbook(path)
                sheet = workbook.active

                for row in sheet.iter_rows(values_only=False):
                    if row[0].value == name: 
                        row[1].value = category
                        row[2].value = product_type
                        row[3].value = quantity
                        row[4].value = price
                        row[5].value = current_date
                        row[6].value = current_time
                        break

                workbook.save(path)
                messagebox.showinfo("Success", "Product updated successfully!")
                self.reset()
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while updating the product: {e}")
  
# DELETE FUN
    
    def delete_product(self):
        if not self.selected_item:
            messagebox.showwarning("Select Item", "Please select an item to delete.")
            return
        
        # Show confirmation dialog
        response = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this product?")
        
        if response:
            # Remove from Treeview
            for item in self.treeview.get_children():
                if self.treeview.item(item, "values")[0] == self.selected_item:
                    self.treeview.delete(item)
                    break

            # Remove from Excel file
            path = "products.xlsx"
            try:
                workbook = openpyxl.load_workbook(path)
                sheet = workbook.active

                for i, row in enumerate(sheet.iter_rows(values_only=False), start=1):
                    if row[0].value == self.selected_item:
                        sheet.delete_rows(i, 1)
                        break

                workbook.save(path)
                messagebox.showinfo("Success", "Product deleted successfully!")
                self.reset()
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while deleting the product: {e}")

# INSRTION FUNCTIONALTY

    def insert_product(self):
        validated_data = self.validate_inputs()
        if not validated_data:
            return
        
        name, category, product_type, quantity, price = validated_data
        current_date = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
        
        unique_id = str(uuid.uuid4())
        row_values = [name, category, product_type, quantity, price, current_date, current_time]
        
        # Insert in a separate thread to keep UI responsive
        threading.Thread(target=self._insert_product, args=(row_values,)).start()
    
    def _insert_product(self, row_values):
        path = "products.xlsx"
        try:
            if not os.path.exists(path):
                self.create_excel_file(path)
            
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active

            # Check for duplicate product name
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0].lower() == row_values[0].lower():  # Assuming Name is unique
                    messagebox.showwarning("Product Exists", f"The product '{row_values[0]}' already exists.")
                    return

            # Append the new product
            sheet.append(row_values)
            workbook.save(path)
            
            # Insert into Treeview
            self.treeview.insert('', tk.END, values=row_values)
            
            # Update the stored data
            self.all_data.append(row_values)
            
            # Clear the input fields
            self.reset()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while inserting the product: {e}")
    
    def validate_inputs(self):
        try:
            name = self.product_name_entry.get().strip()
            if not name:
                messagebox.showerror("Error", "Please enter a valid Product Name!")
                return None

            category = self.category_combobox.get()
            if category == "Flash USB" or category == "Memory Card":  # Specify the category you want to check
                product_type = self.memoryType_combobox.get()
            else:
                product_type = self.type_combobox.get()
            quantity = int(self.quantity_spinbox.get())
            price = float(self.price_spinbox.get())
            
            return name, category, product_type, quantity, price
        except ValueError:
            messagebox.showerror("Error", "Ensure all fields are entered correctly!")
            return None

# INSRTION FUNCTIONALTY 
    
    def search_data(self, event=None):
        search_term = self.search_entry.get().lower()
        self.treeview.delete(*self.treeview.get_children())

        for value_tuple in self.all_data[1:]:
            if any(search_term in str(cell).lower() for cell in value_tuple):
                self.treeview.insert('', tk.END, values=value_tuple)
                
    def reset(self):
        if self.selected_item:
            for item in self.treeview.get_children():
                if self.treeview.item(item, "values")[0] == self.selected_item:
                    self.treeview.selection_remove(item)
                    break
        self.selected_item = None
        self.insert_button.config(state="normal")
        self.type_combobox.config(state="disabled")
        self.memoryType_combobox.config(state="disabled")
        self.update_button.config(state="disabled")
        self.delete_button.config(state="disabled")
        self.canceled_button.config(state="disabled")
        self.Sales_button.config(state="disabled")
        
        # Reset input fields
        self.product_name_entry.config(state="normal")
        self.product_name_entry.delete(0, "end")
        self.category_combobox.set(self.combo_list_ProductCategory[0])
        self.type_combobox.set(self.combo_list_ProductType[0])
        self.quantity_spinbox.delete(0, "end")
        self.price_spinbox.delete(0, "end")

#FUN PART ------------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCrafterApp(root)
    root.mainloop()