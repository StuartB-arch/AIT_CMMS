"""
AIT CMMS - MRO Stock Management Module
Add this to your existing AIT_CMMS_REV3.py file
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import os
from PIL import Image, ImageTk
import shutil
import csv

class MROStockManager:
    """MRO (Maintenance, Repair, Operations) Stock Management"""
    
    def ensure_image_directory(self):
        """Create directory for storing part images if it doesn't exist"""
        image_dir = "mro_images"
        if not os.path.exists(image_dir):
            os.makedirs(image_dir)
        return image_dir
    
    
    
    
    
    
    
    
    
    def clear_all_inventory(self):
        """Clear ALL MRO stock inventory - add this inside MROStockManager class"""
        from tkinter import messagebox
    
        # Get the main app reference (passed during __init__)
        # Check how your __init__ stores it - common patterns:
        main_app = self.parent_app
        if hasattr(self, 'parent'):
            main_app = self.parent
        elif hasattr(self, 'app'):
            main_app = self.app
        elif hasattr(self, 'main_app'):
            main_app = self.main_app
        elif hasattr(self, 'cmms'):
            main_app = self.cmms
    
        if not main_app:
            messagebox.showerror("Error", "Cannot access main application")
            return
    
        # Get count
        cursor = main_app.conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM mro_inventory')
        total_count = cursor.fetchone()[0]
    
        if total_count == 0:
            messagebox.showinfo("No Items", "There are no MRO inventory items to clear.")
            return
    
        # Confirmation
        result = messagebox.askyesno(
            "âš ï¸ Confirm Clear All MRO Inventory",
            f"Are you sure you want to DELETE ALL {total_count} MRO inventory items?\n\n"
            "âš ï¸ WARNING: This action cannot be undone!\n"
            "âš ï¸ ALL stock records will be permanently deleted!\n\n"
            "Are you ABSOLUTELY SURE?",
            icon='warning'
        )
    
        if not result:
            return
    
        # Double confirmation
        double_check = messagebox.askyesno(
            "âš ï¸ Final Confirmation",
            f"FINAL WARNING!\n\n"
            f"You are about to permanently delete {total_count} inventory items.\n\n"
            "This cannot be reversed!\n\n"
            "Click YES to proceed with deletion.",
            icon='warning'
        )
    
        if not double_check:
            messagebox.showinfo("Cancelled", "Clear operation cancelled.")
            return
    
        try:
            # Delete all
            cursor.execute('DELETE FROM mro_inventory')
            main_app.conn.commit()
            
            # Refresh display
            if hasattr(self, 'load_mro_inventory'):
                self.load_mro_inventory()
        
            # Update status
            if hasattr(main_app, 'update_status'):
                main_app.update_status(f"âœ… Cleared {total_count} MRO items")
        
            messagebox.showinfo(
                "Success", 
                f"All {total_count} MRO inventory items deleted."
            )
        
        except Exception as e:
            main_app.conn.rollback()
            messagebox.showerror("Error", f"Failed to clear: {str(e)}")
    
    
    
    
    
    
    
    def __init__(self, parent_app):
        self.parent_app = parent_app
        self.conn = parent_app.conn
        self.root = parent_app.root
        self.init_mro_database()
        self.ensure_image_directory() 
        
    def init_mro_database(self):
        """Initialize MRO inventory table"""
        cursor = self.conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS mro_inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                part_number TEXT UNIQUE NOT NULL,
                model_number TEXT,
                equipment TEXT,
                engineering_system TEXT,
                unit_of_measure TEXT,
                quantity_in_stock REAL DEFAULT 0,
                unit_price REAL DEFAULT 0,
                minimum_stock REAL DEFAULT 0,
                supplier TEXT,
                location TEXT,
                rack TEXT,
                row TEXT,
                bin TEXT,
                picture_1_path TEXT,
                picture_2_path TEXT,
                notes TEXT,
                last_updated TEXT DEFAULT CURRENT_TIMESTAMP,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'Active'
            )
        ''')
        
        # Create index for faster searches
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_mro_part_number 
            ON mro_inventory(part_number)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_mro_name 
            ON mro_inventory(name)
        ''')
        
        # Stock transactions table for tracking stock movements
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS mro_stock_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_number TEXT NOT NULL,
                transaction_type TEXT NOT NULL,
                quantity REAL NOT NULL,
                transaction_date TEXT DEFAULT CURRENT_TIMESTAMP,
                technician_name TEXT,
                work_order TEXT,
                notes TEXT,
                FOREIGN KEY (part_number) REFERENCES mro_inventory (part_number)
            )
        ''')
        
        self.conn.commit()
        print("MRO inventory database initialized")
    
    def create_mro_tab(self, notebook):
        """Create MRO Stock Management tab"""
        mro_frame = ttk.Frame(notebook)
        notebook.add(mro_frame, text='MRO Stock')
        
        # Top controls frame
        controls_frame = ttk.LabelFrame(mro_frame, text="MRO Stock Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        # Buttons row 1
        btn_frame1 = ttk.Frame(controls_frame)
        btn_frame1.pack(fill='x', pady=5)
        
        ttk.Button(btn_frame1, text="+ Add New Part", 
                  command=self.add_part_dialog, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame1, text="Edit Selected Part", 
                  command=self.edit_selected_part, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame1, text="Delete Selected Part", 
                  command=self.delete_selected_part, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame1, text="View Full Details", 
                  command=self.view_part_details, width=20).pack(side='left', padx=5)
        
        # Buttons row 2
        btn_frame2 = ttk.Frame(controls_frame)
        btn_frame2.pack(fill='x', pady=5)
        
        ttk.Button(btn_frame2, text="Import from File", 
                  command=self.import_from_file, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame2, text="Export to CSV", 
                  command=self.export_to_csv, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame2, text="Stock Report", 
                  command=self.generate_stock_report, width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame2, text="! Low Stock Alert", 
                  command=self.show_low_stock, width=20).pack(side='left', padx=5)
        
        ttk.Button(controls_frame, text="CLEAR ALL", 
                  command=lambda: self.clear_all_inventory()).pack(side='right', padx=5)
        # Search and filter frame
        search_frame = ttk.LabelFrame(mro_frame, text="Search & Filter", padding=10)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        # Search bar
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.mro_search_var = tk.StringVar()
        self.mro_search_var.trace('w', self.filter_mro_list)
        ttk.Entry(search_frame, textvariable=self.mro_search_var, 
                 width=40).pack(side='left', padx=5)
        
        # Filter by category
        ttk.Label(search_frame, text="System:").pack(side='left', padx=5)
        self.mro_system_filter = tk.StringVar(value='All')
        system_combo = ttk.Combobox(search_frame, textvariable=self.mro_system_filter,
                                    values=['All', 'Mechanical', 'Electrical', 'Pneumatic', 'Hydraulic'],
                                    width=15, state='readonly')
        system_combo.pack(side='left', padx=5)
        system_combo.bind('<<ComboboxSelected>>', self.filter_mro_list)
        
        # Status filter
        ttk.Label(search_frame, text="Status:").pack(side='left', padx=5)
        self.mro_status_filter = tk.StringVar(value='All')
        status_combo = ttk.Combobox(search_frame, textvariable=self.mro_status_filter,
                                    values=['All', 'Active', 'Inactive', 'Low Stock'],
                                    width=15, state='readonly')
        status_combo.pack(side='left', padx=5)
        status_combo.bind('<<ComboboxSelected>>', self.filter_mro_list)
        
        ttk.Button(search_frame, text="Refresh", 
                  command=self.refresh_mro_list).pack(side='left', padx=5)
        
        # Inventory list
        list_frame = ttk.LabelFrame(mro_frame, text="MRO Inventory", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create treeview
        columns = ('Part Number', 'Name', 'Model', 'Equipment', 'System', 'Qty', 
                  'Min Stock', 'Unit', 'Price', 'Location', 'Status')
        self.mro_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        column_widths = {
            'Part Number': 120,
            'Name': 200,
            'Model': 100,
            'Equipment': 120,
            'System': 100,
            'Qty': 70,
            'Min Stock': 80,
            'Unit': 60,
            'Price': 80,
            'Location': 100,
            'Status': 80
        }
        
        for col in columns:
            self.mro_tree.heading(col, text=col, command=lambda c=col: self.sort_mro_column(c))
            self.mro_tree.column(col, width=column_widths[col], anchor='center')
        
        # Scrollbars
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=self.mro_tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient='horizontal', command=self.mro_tree.xview)
        self.mro_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Pack elements
        self.mro_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Double-click to view details
        self.mro_tree.bind('<Double-1>', lambda e: self.view_part_details())
        
        # Statistics frame
        stats_frame = ttk.LabelFrame(mro_frame, text="Inventory Statistics", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=5)
        
        self.mro_stats_label = ttk.Label(stats_frame, text="Loading...", 
                                         font=('Arial', 10))
        self.mro_stats_label.pack()
        
        # Load initial data
        self.refresh_mro_list()
        
        return mro_frame
    
    def add_part_dialog(self):
        """Dialog to add new part"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New MRO Part")
        dialog.geometry("800x900")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Create scrollable frame
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
    
        # âœ… FIX THIS LINE - it was incomplete!
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Form fields
        fields = {}
        row = 0
        
        # Basic Information
        ttk.Label(scrollable_frame, text="BASIC INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                              sticky='w', padx=10, pady=10)
        row += 1
    
        field_configs = [
            ('Name*', 'name'),
            ('Part Number*', 'part_number'),
            ('Model Number', 'model_number'),
            ('Equipment', 'equipment'),
        ]
    
        for label, field_name in field_configs:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Stock Information
        ttk.Label(scrollable_frame, text="STOCK INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        stock_fields = [
            ('Engineering System*', 'engineering_system'),
            ('Unit of Measure*', 'unit_of_measure'),
            ('Quantity in Stock*', 'quantity_in_stock'),
            ('Unit Price', 'unit_price'),
            ('Minimum Stock*', 'minimum_stock'),
            ('Supplier', 'supplier'),
        ]
    
        for label, field_name in stock_fields:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Location Information
        ttk.Label(scrollable_frame, text="LOCATION INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        location_fields = [
            ('Location*', 'location'),
            ('Rack', 'rack'),
            ('Row', 'row'),
            ('Bin', 'bin'),
        ]
    
        for label, field_name in location_fields:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Pictures
        ttk.Label(scrollable_frame, text="PICTURES", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        fields['picture_1'] = tk.StringVar()
        fields['picture_2'] = tk.StringVar()
        
        ttk.Label(scrollable_frame, text="Picture 1:").grid(row=row, column=0, 
                                                            sticky='w', padx=10, pady=5)
        pic1_frame = ttk.Frame(scrollable_frame)
        pic1_frame.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        ttk.Entry(pic1_frame, textvariable=fields['picture_1'], width=35).pack(side='left')
        ttk.Button(pic1_frame, text="Browse", 
                command=lambda: self.browse_image(fields['picture_1'])).pack(side='left', padx=5)
        row += 1
    
        ttk.Label(scrollable_frame, text="Picture 2:").grid(row=row, column=0, 
                                                            sticky='w', padx=10, pady=5)
        pic2_frame = ttk.Frame(scrollable_frame)
        pic2_frame.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        ttk.Entry(pic2_frame, textvariable=fields['picture_2'], width=35).pack(side='left')
        ttk.Button(pic2_frame, text="Browse", 
                command=lambda: self.browse_image(fields['picture_2'])).pack(side='left', padx=5)
        row += 1
    
        # Notes
        ttk.Label(scrollable_frame, text="Notes:").grid(row=row, column=0, 
                                                        sticky='nw', padx=10, pady=5)
        fields['notes'] = tk.Text(scrollable_frame, width=50, height=5)
        fields['notes'].grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Buttons
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
    
        def save_part():
            try:
                # Validate required fields
                required = ['name', 'part_number', 'engineering_system', 
                        'unit_of_measure', 'quantity_in_stock', 'minimum_stock', 'location']
            
                for field in required:
                    if field in ['notes', 'picture_1', 'picture_2']:
                        continue
                    value = fields[field].get() if hasattr(fields[field], 'get') else ''
                    if not value:
                        messagebox.showerror("Error", f"Please fill in: {field.replace('_', ' ').title()}")
                        return
            
                # Validate image paths if provided
                pic1_path = fields['picture_1'].get()
                pic2_path = fields['picture_2'].get()
        
                if pic1_path and not os.path.exists(pic1_path):
                    messagebox.showwarning("Warning", f"Picture 1 path not found: {pic1_path}")
                    # You can choose to continue or return here
            
                if pic2_path and not os.path.exists(pic2_path):
                    messagebox.showwarning("Warning", f"Picture 2 path not found: {pic2_path}")
            
            
            
            
            
            
                # Insert into database
                cursor = self.conn.cursor()
            
                notes_text = fields['notes'].get('1.0', 'end-1c') if 'notes' in fields else ''
            
                cursor.execute('''
                    INSERT INTO mro_inventory (
                        name, part_number, model_number, equipment, engineering_system,
                        unit_of_measure, quantity_in_stock, unit_price, minimum_stock,
                        supplier, location, rack, row, bin, picture_1_path, picture_2_path, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    fields['name'].get(),
                    fields['part_number'].get(),
                    fields['model_number'].get(),
                    fields['equipment'].get(),
                    fields['engineering_system'].get(),
                    fields['unit_of_measure'].get(),
                    float(fields['quantity_in_stock'].get() or 0),
                    float(fields['unit_price'].get() or 0),
                    float(fields['minimum_stock'].get() or 0),
                    fields['supplier'].get(),
                    fields['location'].get(),
                    fields['rack'].get(),
                    fields['row'].get(),
                    fields['bin'].get(),
                    fields['picture_1'].get(),
                    fields['picture_2'].get(),
                    notes_text
                ))
            
                self.conn.commit()
                messagebox.showinfo("Success", "Part added successfully!")
                dialog.destroy()
                self.refresh_mro_list()
            
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Part number already exists!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add part: {str(e)}")
    
        ttk.Button(btn_frame, text="Save Part", command=save_part, width=20).pack(side='left', padx=10)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=20).pack(side='left', padx=10)
        
        # âœ… CRITICAL: Pack canvas and scrollbar - THIS MUST BE AT THE END!
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    
    def browse_image(self, var):
        """Browse for image file and copy to application directory"""
        file_path = filedialog.askopenfilename(
            title="Select Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")]
        )
        if file_path:
            try:
                # Create images directory
                image_dir = self.ensure_image_directory()
                
                # Generate unique filename
                original_name = os.path.basename(file_path)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = f"{timestamp}_{original_name}"
                dest_path = os.path.join(image_dir, new_filename)
                
                # Copy image to application directory
                shutil.copy2(file_path, dest_path)
                
                # Store the relative path in the database
                var.set(dest_path)
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy image: {str(e)}")
                var.set(file_path)  # Fallback to original path
    
    
    
    def edit_selected_part(self):
        """Edit selected part - FIXED VERSION with proper part lookup"""
        selected = self.mro_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a part to edit")
            return

        item = self.mro_tree.item(selected[0])
        part_number = item['values'][0]  # Part Number is the first column

        # Debug: Print what we're trying to find
        print(f"Looking for part number: '{part_number}' (type: {type(part_number)})")

        # Get full part data - use direct string comparison first
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM mro_inventory WHERE part_number = ?', (str(part_number),))
        part_data = cursor.fetchone()

        # If not found, try alternative lookups
        if not part_data:
            # If still not found, try to find what's actually in the treeview
            print("Part not found in database. Checking treeview data...")

            # Get all items in treeview to debug
            all_items = self.mro_tree.get_children()
            for item_id in all_items:
                item_data = self.mro_tree.item(item_id)
                print(f"Treeview item: {item_data['values'][0]} - {item_data['values'][1]}")

            messagebox.showerror("Error",
                            f"Part not found in database: '{part_number}'\n"
                            f"Please refresh the list and try again.")
            return

        # Create edit dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Part: {part_data[2]}")  # part_number is at index 2
        dialog.geometry("800x900")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Create scrollable frame
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
    
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
    
        # Parse part_data
        columns = ['id', 'name', 'part_number', 'model_number', 'equipment', 'engineering_system',
                  'unit_of_measure', 'quantity_in_stock', 'unit_price', 'minimum_stock',
                  'supplier', 'location', 'rack', 'row', 'bin', 'picture_1_path', 
                  'picture_2_path', 'notes', 'last_updated', 'created_date', 'status']
    
        part_dict = dict(zip(columns, part_data))
    
        # Form fields
        fields = {}
        row = 0
    
        # Basic Information
        ttk.Label(scrollable_frame, text="BASIC INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        field_configs = [
            ('Name*', 'name'),
            ('Part Number*', 'part_number'),
            ('Model Number', 'model_number'),
            ('Equipment', 'equipment'),
        ]
    
        for label, field_name in field_configs:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].insert(0, str(part_dict.get(field_name) or ''))
            if field_name == 'part_number':
                fields[field_name].config(state='readonly')  # Don't allow changing part number
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Engineering System
        ttk.Label(scrollable_frame, text="Engineering System*").grid(row=row, column=0, 
                                                                    sticky='w', padx=10, pady=5)
        fields['engineering_system'] = ttk.Combobox(scrollable_frame,
                                                     values=['Mechanical', 'Electrical', 'Pneumatic', 'Hydraulic'],
                                                     width=47, state='readonly')
        fields['engineering_system'].set(part_dict.get('engineering_system') or '')
        fields['engineering_system'].grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Stock Information
        ttk.Label(scrollable_frame, text="STOCK INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        stock_fields = [
            ('Unit of Measure*', 'unit_of_measure'),
            ('Quantity in Stock*', 'quantity_in_stock'),
            ('Unit Price ($)', 'unit_price'),
            ('Minimum Stock*', 'minimum_stock'),
            ('Supplier', 'supplier'),
        ]
    
        for label, field_name in stock_fields:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].insert(0, str(part_dict.get(field_name) or ''))
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Location Information
        ttk.Label(scrollable_frame, text="LOCATION INFORMATION", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        location_fields = [
            ('Location*', 'location'),
            ('Rack', 'rack'),
            ('Row', 'row'),
            ('Bin', 'bin'),
        ]
    
        for label, field_name in location_fields:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
            fields[field_name] = ttk.Entry(scrollable_frame, width=50)
            fields[field_name].insert(0, str(part_dict.get(field_name) or ''))
            fields[field_name].grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1
    
        # Status
        ttk.Label(scrollable_frame, text="Status*").grid(row=row, column=0, 
                                                        sticky='w', padx=10, pady=5)
        fields['status'] = ttk.Combobox(scrollable_frame,
                                       values=['Active', 'Inactive'],
                                       width=47, state='readonly')
        fields['status'].set(part_dict.get('status') or 'Active')
        fields['status'].grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Pictures
        ttk.Label(scrollable_frame, text="PICTURES", 
                font=('Arial', 11, 'bold')).grid(row=row, column=0, columnspan=2, 
                                                sticky='w', padx=10, pady=10)
        row += 1
    
        fields['picture_1'] = tk.StringVar(value=part_dict.get('picture_1_path') or '')
        fields['picture_2'] = tk.StringVar(value=part_dict.get('picture_2_path') or '')
        
        ttk.Label(scrollable_frame, text="Picture 1:").grid(row=row, column=0, 
                                                            sticky='w', padx=10, pady=5)
        pic1_frame = ttk.Frame(scrollable_frame)
        pic1_frame.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        ttk.Entry(pic1_frame, textvariable=fields['picture_1'], width=35).pack(side='left')
        ttk.Button(pic1_frame, text="Browse", 
                command=lambda: self.browse_image(fields['picture_1'])).pack(side='left', padx=5)
        row += 1
    
        ttk.Label(scrollable_frame, text="Picture 2:").grid(row=row, column=0, 
                                                            sticky='w', padx=10, pady=5)
        pic2_frame = ttk.Frame(scrollable_frame)
        pic2_frame.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        ttk.Entry(pic2_frame, textvariable=fields['picture_2'], width=35).pack(side='left')
        ttk.Button(pic2_frame, text="Browse", 
                command=lambda: self.browse_image(fields['picture_2'])).pack(side='left', padx=5)
        row += 1
    
        # Notes
        ttk.Label(scrollable_frame, text="Notes:").grid(row=row, column=0, 
                                                        sticky='nw', padx=10, pady=5)
        fields['notes'] = tk.Text(scrollable_frame, width=50, height=5)
        fields['notes'].insert('1.0', part_dict.get('notes') or '')
        fields['notes'].grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Buttons
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
    
        def update_part():
            try:
                # Validate image paths
                pic1_path = fields['picture_1'].get()
                pic2_path = fields['picture_2'].get()
            
                if pic1_path and not os.path.exists(pic1_path):
                    messagebox.showwarning("Warning", f"Picture 1 path not found: {pic1_path}")
            
                if pic2_path and not os.path.exists(pic2_path):
                    messagebox.showwarning("Warning", f"Picture 2 path not found: {pic2_path}")
            
                cursor = self.conn.cursor()
            
                notes_text = fields['notes'].get('1.0', 'end-1c')
            
                cursor.execute('''
                    UPDATE mro_inventory SET
                        name = ?, model_number = ?, equipment = ?, engineering_system = ?,
                        unit_of_measure = ?, quantity_in_stock = ?, unit_price = ?, 
                        minimum_stock = ?, supplier = ?, location = ?, rack = ?, 
                        row = ?, bin = ?, picture_1_path = ?, picture_2_path = ?, 
                        notes = ?, status = ?, last_updated = ?
                    WHERE part_number = ?
                ''', (
                    fields['name'].get(),
                    fields['model_number'].get(),
                    fields['equipment'].get(),
                    fields['engineering_system'].get(),
                    fields['unit_of_measure'].get(),
                    float(fields['quantity_in_stock'].get() or 0),
                    float(fields['unit_price'].get() or 0),
                    float(fields['minimum_stock'].get() or 0),
                    fields['supplier'].get(),
                    fields['location'].get(),
                    fields['rack'].get(),
                    fields['row'].get(),
                    fields['bin'].get(),
                    fields['picture_1'].get(),
                    fields['picture_2'].get(),
                    notes_text,
                    fields['status'].get(),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    part_dict['part_number']  # Use the original part number from the database
                ))
            
                self.conn.commit()
                messagebox.showinfo("Success", "Part updated successfully!")
                dialog.destroy()
                self.refresh_mro_list()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update part: {str(e)}")
    
        ttk.Button(btn_frame, text="Update Part", command=update_part, width=20).pack(side='left', padx=10)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=20).pack(side='left', padx=10)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        
    
    def delete_selected_part(self):
        """Delete selected part"""
        selected = self.mro_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a part to delete")
            return

        item = self.mro_tree.item(selected[0])
        part_number = str(item['values'][0])  # Convert to string
        part_name = item['values'][1]

        # Get part data to find image paths
        cursor = self.conn.cursor()
        cursor.execute('SELECT picture_1_path, picture_2_path FROM mro_inventory WHERE part_number = ?', (part_number,))
        part_data = cursor.fetchone()

        result = messagebox.askyesno("Confirm Delete",
                                    f"Are you sure you want to delete:\n\n"
                                    f"Part Number: {part_number}\n"
                                    f"Name: {part_name}\n\n"
                                    f"This action cannot be undone!")

        if result:
            try:
                # Delete associated images
                if part_data:
                    pic1_path, pic2_path = part_data
                    for pic_path in [pic1_path, pic2_path]:
                        if pic_path and os.path.exists(pic_path) and pic_path.startswith("mro_images"):
                            try:
                                os.remove(pic_path)
                            except:
                                pass  # Ignore errors when deleting images

                # Delete from database
                cursor.execute('DELETE FROM mro_inventory WHERE part_number = ?', (part_number,))
                self.conn.commit()
                messagebox.showinfo("Success", "Part deleted successfully!")
                self.refresh_mro_list()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete part: {str(e)}")
    
    def view_part_details(self):
        """View full details of selected part"""
        selected = self.mro_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a part to view")
            return

        item = self.mro_tree.item(selected[0])
        part_number = str(item['values'][0])  # Convert to string immediately

        # Get full part data
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM mro_inventory WHERE part_number = ?', (part_number,))
        part_data = cursor.fetchone()

        # If not found, try with stripped whitespace
        if not part_data:
            clean_part_number = part_number.strip()
            cursor.execute('SELECT * FROM mro_inventory WHERE part_number = ?', (clean_part_number,))
            part_data = cursor.fetchone()

        if not part_data:
            messagebox.showerror("Error", f"Part not found: {part_number}")
            return
    
        
        # Create details dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Part Details: {part_number}")
        dialog.geometry("900x800")
        dialog.transient(self.root)
        
        # Create scrollable frame
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Parse part_data
        columns = ['id', 'name', 'part_number', 'model_number', 'equipment', 'engineering_system',
                  'unit_of_measure', 'quantity_in_stock', 'unit_price', 'minimum_stock',
                  'supplier', 'location', 'rack', 'row', 'bin', 'picture_1_path', 
                  'picture_2_path', 'notes', 'last_updated', 'created_date', 'status']
        
        part_dict = dict(zip(columns, part_data))
        
        # Title
        title_frame = ttk.Frame(scrollable_frame)
        title_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(title_frame, text=f"{part_dict['name']}", 
                 font=('Arial', 16, 'bold')).pack()
        ttk.Label(title_frame, text=f"Part Number: {part_dict['part_number']}", 
                 font=('Arial', 12)).pack()
        
        # Basic Info
        basic_frame = ttk.LabelFrame(scrollable_frame, text="Basic Information", padding=15)
        basic_frame.pack(fill='x', padx=10, pady=5)
        
        basic_info = [
            ('Name:', part_dict['name']),
            ('Part Number:', part_dict['part_number']),
            ('Model Number:', part_dict['model_number']),
            ('Equipment:', part_dict['equipment']),
            ('Engineering System:', part_dict['engineering_system']),
            ('Status:', part_dict['status']),
        ]
        
        for label, value in basic_info:
            row_frame = ttk.Frame(basic_frame)
            row_frame.pack(fill='x', pady=2)
            ttk.Label(row_frame, text=label, font=('Arial', 10, 'bold'), 
                     width=20).pack(side='left')
            ttk.Label(row_frame, text=str(value or 'N/A'), 
                     font=('Arial', 10)).pack(side='left')
        
        # Stock Info
        stock_frame = ttk.LabelFrame(scrollable_frame, text="Stock Information", padding=15)
        stock_frame.pack(fill='x', padx=10, pady=5)
        
        qty = float(part_dict['quantity_in_stock'] or 0)
        min_stock = float(part_dict['minimum_stock'] or 0)
        stock_status = "âœ… OK" if qty >= min_stock else "âš ï¸ LOW STOCK"
        
        stock_info = [
            ('Quantity in Stock:', f"{qty} {part_dict['unit_of_measure']}"),
            ('Minimum Stock:', f"{min_stock} {part_dict['unit_of_measure']}"),
            ('Stock Status:', stock_status),
            ('Unit Price:', f"${float(part_dict['unit_price'] or 0):.2f}"),
            ('Total Value:', f"${qty * float(part_dict['unit_price'] or 0):.2f}"),
            ('Supplier:', part_dict['supplier']),
        ]
        
        for label, value in stock_info:
            row_frame = ttk.Frame(stock_frame)
            row_frame.pack(fill='x', pady=2)
            ttk.Label(row_frame, text=label, font=('Arial', 10, 'bold'), 
                     width=20).pack(side='left')
            label_widget = ttk.Label(row_frame, text=str(value or 'N/A'), 
                                    font=('Arial', 10))
            if 'LOW STOCK' in str(value):
                label_widget.config(foreground='red')
            label_widget.pack(side='left')
        
        # Location Info
        loc_frame = ttk.LabelFrame(scrollable_frame, text="Location Information", padding=15)
        loc_frame.pack(fill='x', padx=10, pady=5)
        
        loc_info = [
            ('Location:', part_dict['location']),
            ('Rack:', part_dict['rack']),
            ('Row:', part_dict['row']),
            ('Bin:', part_dict['bin']),
        ]
        
        for label, value in loc_info:
            row_frame = ttk.Frame(loc_frame)
            row_frame.pack(fill='x', pady=2)
            ttk.Label(row_frame, text=label, font=('Arial', 10, 'bold'), 
                     width=20).pack(side='left')
            ttk.Label(row_frame, text=str(value or 'N/A'), 
                     font=('Arial', 10)).pack(side='left')
        
        # Notes
        if part_dict['notes']:
            notes_frame = ttk.LabelFrame(scrollable_frame, text="Notes", padding=15)
            notes_frame.pack(fill='x', padx=10, pady=5)
            notes_text = tk.Text(notes_frame, height=4, wrap='word')
            notes_text.insert('1.0', part_dict['notes'])
            notes_text.config(state='disabled')
            notes_text.pack(fill='x')
        
        # Pictures
        pics_frame = ttk.LabelFrame(scrollable_frame, text="Pictures", padding=15)
        pics_frame.pack(fill='both', expand=True, padx=10, pady=5)

        pic_container = ttk.Frame(pics_frame)
        pic_container.pack(fill='both', expand=True)

        for idx, pic_path in enumerate([part_dict['picture_1_path'], part_dict['picture_2_path']], 1):
            if pic_path and os.path.exists(pic_path):
                try:
                    pic_frame = ttk.LabelFrame(pic_container, text=f"Picture {idx}", padding=10)
                    pic_frame.pack(side='left', padx=10, pady=5, fill='both', expand=True)
                    
                    img = Image.open(pic_path)
                    # Resize image to fit in dialog
                    img.thumbnail((250, 250), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    
                    pic_label = ttk.Label(pic_frame, image=photo)
                    pic_label.image = photo  # Keep reference to prevent garbage collection
                    pic_label.pack()
                    
                    ttk.Label(pic_frame, text=os.path.basename(pic_path), 
                            font=('Arial', 8)).pack()
                     
                except Exception as e:
                    error_frame = ttk.Frame(pic_container)
                    error_frame.pack(side='left', padx=10, pady=5)
                    ttk.Label(error_frame, text=f"Picture {idx}: Error loading\n{str(e)}", 
                            foreground='red', font=('Arial', 8)).pack()
            else:
                no_image_frame = ttk.Frame(pic_container)
                no_image_frame.pack(side='left', padx=10, pady=5)
                ttk.Label(no_image_frame, text=f"Picture {idx}: No image", 
                        foreground='gray').pack()
        
        # Dates
        dates_frame = ttk.LabelFrame(scrollable_frame, text="Record Information", padding=15)
        dates_frame.pack(fill='x', padx=10, pady=5)
        
        date_info = [
            ('Created:', part_dict['created_date']),
            ('Last Updated:', part_dict['last_updated']),
        ]
        
        for label, value in date_info:
            row_frame = ttk.Frame(dates_frame)
            row_frame.pack(fill='x', pady=2)
            ttk.Label(row_frame, text=label, font=('Arial', 10, 'bold'), 
                     width=20).pack(side='left')
            ttk.Label(row_frame, text=str(value or 'N/A'), 
                     font=('Arial', 10)).pack(side='left')
        
        # Buttons
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Edit Part", 
                  command=lambda: [dialog.destroy(), self.edit_selected_part()], 
                  width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Stock Transaction", 
                  command=lambda: [dialog.destroy(), self.stock_transaction_dialog(part_number)], 
                  width=20).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Close", 
                  command=dialog.destroy, width=20).pack(side='left', padx=5)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def stock_transaction_dialog(self, part_number):
        """Dialog for stock transactions (add/remove stock)"""
        part_number = str(part_number)  # Ensure string type

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Stock Transaction: {part_number}")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()

        # Get current stock
        cursor = self.conn.cursor()
        cursor.execute('SELECT quantity_in_stock, unit_of_measure, name FROM mro_inventory WHERE part_number = ?',
                      (part_number,))
        result = cursor.fetchone()
        current_stock = result[0] if result else 0
        unit = result[1] if result else ''
        part_name = result[2] if result else ''
        
        ttk.Label(dialog, text=f"Part: {part_name}", 
                 font=('Arial', 12, 'bold')).pack(pady=10)
        ttk.Label(dialog, text=f"Current Stock: {current_stock} {unit}", 
                 font=('Arial', 11)).pack(pady=5)
        
        # Transaction type
        ttk.Label(dialog, text="Transaction Type:").pack(pady=5)
        trans_type = tk.StringVar(value='Add')
        ttk.Radiobutton(dialog, text="âž• Add Stock", variable=trans_type, 
                       value='Add').pack()
        ttk.Radiobutton(dialog, text="âž– Remove Stock", variable=trans_type, 
                       value='Remove').pack()
        
        # Quantity
        ttk.Label(dialog, text="Quantity:").pack(pady=5)
        qty_entry = ttk.Entry(dialog, width=20)
        qty_entry.pack(pady=5)
        
        # Work order
        ttk.Label(dialog, text="Work Order (Optional):").pack(pady=5)
        wo_entry = ttk.Entry(dialog, width=30)
        wo_entry.pack(pady=5)
        
        # Notes
        ttk.Label(dialog, text="Notes:").pack(pady=5)
        notes_text = tk.Text(dialog, height=4, width=50)
        notes_text.pack(pady=5)
        
        def process_transaction():
            try:
                qty = float(qty_entry.get())
                trans_type_val = trans_type.get()
                
                if trans_type_val == 'Remove':
                    qty = -qty
                
                new_stock = current_stock + qty
                
                if new_stock < 0:
                    messagebox.showerror("Error", "Cannot remove more stock than available!")
                    return
                
                # Update stock
                cursor = self.conn.cursor()
                cursor.execute('''
                    UPDATE mro_inventory 
                    SET quantity_in_stock = ?, last_updated = ?
                    WHERE part_number = ?
                ''', (new_stock, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), part_number))
                
                # Log transaction
                cursor.execute('''
                    INSERT INTO mro_stock_transactions 
                    (part_number, transaction_type, quantity, technician_name, 
                     work_order, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    part_number,
                    trans_type_val,
                    abs(qty),
                    self.parent_app.current_user if hasattr(self.parent_app, 'current_user') else 'System',
                    wo_entry.get(),
                    notes_text.get('1.0', 'end-1c')
                ))
                
                self.conn.commit()
                messagebox.showinfo("Success", 
                                  f"Stock updated!\n"
                                  f"Previous: {current_stock} {unit}\n"
                                  f"Change: {qty:+.1f} {unit}\n"
                                  f"New Stock: {new_stock} {unit}")
                dialog.destroy()
                self.refresh_mro_list()
                
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid quantity")
            except Exception as e:
                messagebox.showerror("Error", f"Transaction failed: {str(e)}")
        
        ttk.Button(dialog, text="Process Transaction", 
                  command=process_transaction, width=25).pack(pady=10)
        ttk.Button(dialog, text="Cancel", 
                  command=dialog.destroy, width=25).pack(pady=5)
    
    def import_from_file(self):
        """Import parts from inventory.txt or CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select Inventory File",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            imported_count = 0
            skipped_count = 0
            
            with open(file_path, 'r', encoding='utf-8') as f:
                if file_path.endswith('.csv'):
                    reader = csv.DictReader(f)
                    for row in reader:
                        try:
                            self.import_part_from_dict(row)
                            imported_count += 1
                        except:
                            skipped_count += 1
                else:
                    # Parse text file format
                    content = f.read()
                    # You can customize this based on your inventory.txt format
                    messagebox.showinfo("Info", 
                                      "Please use CSV format for bulk import.\n\n"
                                      "Required columns:\n"
                                      "Name, Part Number, Model Number, Equipment, "
                                      "Engineering System, Unit of Measure, Quantity in Stock, "
                                      "Unit Price, Minimum Stock, Supplier, Location, Rack, Row, Bin")
                    return
            
            self.conn.commit()
            messagebox.showinfo("Import Complete", 
                              f"Successfully imported: {imported_count} parts\n"
                              f"Skipped (duplicates/errors): {skipped_count} parts")
            self.refresh_mro_list()
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import file:\n{str(e)}")
    
    def import_part_from_dict(self, data):
        """Import a single part from dictionary"""
        cursor = self.conn.cursor()
        
        cursor.execute('''
            INSERT OR IGNORE INTO mro_inventory (
                name, part_number, model_number, equipment, engineering_system,
                unit_of_measure, quantity_in_stock, unit_price, minimum_stock,
                supplier, location, rack, row, bin
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data.get('Name', ''),
            data.get('Part Number', ''),
            data.get('Model Number', ''),
            data.get('Equipment', ''),
            data.get('Engineering System', ''),
            data.get('Unit of Measure', ''),
            float(data.get('Quantity in Stock', 0) or 0),
            float(data.get('Unit Price', 0) or 0),
            float(data.get('Minimum Stock', 0) or 0),
            data.get('Supplier', ''),
            data.get('Location', ''),
            data.get('Rack', ''),
            data.get('Row', ''),
            data.get('Bin', '')
        ))
    
    def export_to_csv(self):
        """Export inventory to CSV"""
        file_path = filedialog.asksaveasfilename(
            title="Export Inventory",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT * FROM mro_inventory ORDER BY part_number')
            rows = cursor.fetchall()
            
            columns = ['ID', 'Name', 'Part Number', 'Model Number', 'Equipment', 
                      'Engineering System', 'Unit of Measure', 'Quantity in Stock', 
                      'Unit Price', 'Minimum Stock', 'Supplier', 'Location', 'Rack', 
                      'Row', 'Bin', 'Picture 1', 'Picture 2', 'Notes', 
                      'Last Updated', 'Created Date', 'Status']
            
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(columns)
                writer.writerows(rows)
            
            messagebox.showinfo("Success", f"Inventory exported to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")
    
    def generate_stock_report(self):
        """Generate comprehensive stock report"""
        report_dialog = tk.Toplevel(self.root)
        report_dialog.title("Stock Report")
        report_dialog.geometry("900x700")
        report_dialog.transient(self.root)
        
        # Report text
        report_frame = ttk.Frame(report_dialog)
        report_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        report_text = tk.Text(report_frame, wrap='word', font=('Courier', 10))
        report_scrollbar = ttk.Scrollbar(report_frame, command=report_text.yview)
        report_text.configure(yscrollcommand=report_scrollbar.set)
        
        report_text.pack(side='left', fill='both', expand=True)
        report_scrollbar.pack(side='right', fill='y')
        
        # Generate report
        cursor = self.conn.cursor()
        
        report = []
        report.append("=" * 80)
        report.append("MRO INVENTORY STOCK REPORT")
        report.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("=" * 80)
        report.append("")
        
        # Summary statistics
        cursor.execute('SELECT COUNT(*) FROM mro_inventory WHERE status = "Active"')
        total_parts = cursor.fetchone()[0]
        
        cursor.execute('SELECT SUM(quantity_in_stock * unit_price) FROM mro_inventory WHERE status = "Active"')
        total_value = cursor.fetchone()[0] or 0
        
        cursor.execute('''
            SELECT COUNT(*) FROM mro_inventory 
            WHERE quantity_in_stock < minimum_stock AND status = "Active"
        ''')
        low_stock_count = cursor.fetchone()[0]
        
        report.append("SUMMARY")
        report.append("-" * 80)
        report.append(f"Total Active Parts: {total_parts}")
        report.append(f"Total Inventory Value: ${total_value:,.2f}")
        report.append(f"Low Stock Items: {low_stock_count}")
        report.append("")
        
        # Low stock items
        if low_stock_count > 0:
            report.append("LOW STOCK ALERTS")
            report.append("-" * 80)
            cursor.execute('''
                SELECT part_number, name, quantity_in_stock, minimum_stock, 
                       unit_of_measure, location
                FROM mro_inventory 
                WHERE quantity_in_stock < minimum_stock AND status = "Active"
                ORDER BY (minimum_stock - quantity_in_stock) DESC
            ''')
            
            for row in cursor.fetchall():
                part_no, name, qty, min_qty, unit, loc = row
                deficit = min_qty - qty
                report.append(f"  Part: {part_no} - {name}")
                report.append(f"  Current: {qty} {unit} | Minimum: {min_qty} {unit} | Deficit: {deficit} {unit}")
                report.append(f"  Location: {loc}")
                report.append("")
        
        # Inventory by system
        report.append("INVENTORY BY ENGINEERING SYSTEM")
        report.append("-" * 80)
        cursor.execute('''
            SELECT engineering_system, COUNT(*), SUM(quantity_in_stock * unit_price)
            FROM mro_inventory 
            WHERE status = "Active"
            GROUP BY engineering_system
            ORDER BY engineering_system
        ''')
        
        for row in cursor.fetchall():
            system, count, value = row
            report.append(f"  {system or 'Unknown'}: {count} parts, ${value or 0:,.2f} value")
        
        report.append("")
        report.append("=" * 80)
        report.append("END OF REPORT")
        report.append("=" * 80)
        
        report_text.insert('1.0', '\n'.join(report))
        report_text.config(state='disabled')
        
        # Export button
        def export_report():
            file_path = filedialog.asksaveasfilename(
                title="Export Stock Report",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(report))
                messagebox.showinfo("Success", f"Report exported to:\n{file_path}")
        
        ttk.Button(report_dialog, text="ðŸ“¤ Export Report", 
                  command=export_report).pack(pady=10)
    
    def show_low_stock(self):
        """Show low stock alert dialog"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT part_number, name, quantity_in_stock, minimum_stock, 
                   unit_of_measure, location, supplier
            FROM mro_inventory 
            WHERE quantity_in_stock < minimum_stock AND status = "Active"
            ORDER BY (minimum_stock - quantity_in_stock) DESC
        ''')
        
        low_stock_items = cursor.fetchall()
        
        if not low_stock_items:
            messagebox.showinfo("Stock Status", "âœ… All items are adequately stocked!")
            return
        
        # Create alert dialog
        alert_dialog = tk.Toplevel(self.root)
        alert_dialog.title(f"âš ï¸ Low Stock Alert ({len(low_stock_items)} items)")
        alert_dialog.geometry("1000x600")
        alert_dialog.transient(self.root)
        
        ttk.Label(alert_dialog, 
                 text=f"âš ï¸ {len(low_stock_items)} items are below minimum stock level",
                 font=('Arial', 12, 'bold'), foreground='red').pack(pady=10)
        
        # Create treeview
        tree_frame = ttk.Frame(alert_dialog)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        columns = ('Part Number', 'Name', 'Current', 'Minimum', 'Deficit', 
                  'Unit', 'Location', 'Supplier')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)
        
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        
        # Populate tree
        for item in low_stock_items:
            part_no, name, current, minimum, unit, location, supplier = item
            deficit = minimum - current
            tree.insert('', 'end', values=(
                part_no, name, f"{current:.1f}", f"{minimum:.1f}", 
                f"{deficit:.1f}", unit, location or 'N/A', supplier or 'N/A'
            ))
        
        ttk.Button(alert_dialog, text="Close", 
                  command=alert_dialog.destroy).pack(pady=10)
    
    def refresh_mro_list(self):
        """Refresh MRO inventory list"""
        self.filter_mro_list()
        self.update_mro_statistics()
    
    def filter_mro_list(self, *args):
        """Filter MRO list based on search and filters - FIXED part number display"""
        search_term = self.mro_search_var.get().lower()
        system_filter = self.mro_system_filter.get()
        status_filter = self.mro_status_filter.get()
    
        # Clear existing items
        for item in self.mro_tree.get_children():
            self.mro_tree.delete(item)
    
        # Build query with case-insensitive comparisons
        query = 'SELECT * FROM mro_inventory WHERE 1=1'
        params = []
    
        # System filter - CASE INSENSITIVE
        if system_filter != 'All':
            query += ' AND LOWER(IFNULL(engineering_system, "")) = LOWER(?)'
            params.append(system_filter)
    
        # Status filter
        if status_filter == 'Low Stock':
            query += ' AND quantity_in_stock < minimum_stock'
        elif status_filter != 'All':
            # CASE INSENSITIVE status check
            query += ' AND LOWER(IFNULL(status, "")) = LOWER(?)'
            params.append(status_filter)
    
        # Search filter
        if search_term:
            query += ''' AND (
                LOWER(name) LIKE ? OR 
                LOWER(part_number) LIKE ? OR 
                LOWER(model_number) LIKE ? OR 
                LOWER(equipment) LIKE ? OR 
                LOWER(location) LIKE ?
            )'''
            search_param = f'%{search_term}%'
            params.extend([search_param] * 5)
    
        query += ' ORDER BY part_number'
    
        try:
            cursor = self.conn.cursor()
            cursor.execute(query, params)
        
            for row in cursor.fetchall():
                # Determine status color
                qty = float(row[7])
                min_stock = float(row[9])
                status = '⚠️ LOW' if qty < min_stock else (row[20] or 'Active')
            
                # CRITICAL FIX: Use the exact part_number from database (row[2])
                # Don't convert to string or modify it in any way
                part_number_display = row[2]  # This is the exact part_number from DB
            
                self.mro_tree.insert('', 'end', values=(
                    part_number_display,   # Part Number - use exact value from DB
                    row[1],   # Name
                    row[3],   # Model
                    row[4],   # Equipment
                    row[5],   # System
                    f"{qty:.1f}",  # Qty
                    f"{min_stock:.1f}",  # Min Stock
                    row[6],   # Unit
                    f"${float(row[8]):.2f}",  # Price
                    row[11],  # Location
                    status    # Status
                ), tags=('low_stock',) if qty < min_stock else ())
        
            # Color low stock items
            self.mro_tree.tag_configure('low_stock', background='#ffcccc')
        
        except Exception as e:
            print(f"Filter error: {e}")
            import traceback
            traceback.print_exc()
    

    def update_mro_statistics(self):
        """Update inventory statistics"""
        cursor = self.conn.cursor()
        
        cursor.execute('SELECT COUNT(*) FROM mro_inventory WHERE status = "Active"')
        total = cursor.fetchone()[0]
        
        cursor.execute('SELECT SUM(quantity_in_stock * unit_price) FROM mro_inventory WHERE status = "Active"')
        value = cursor.fetchone()[0] or 0
        
        cursor.execute('''
            SELECT COUNT(*) FROM mro_inventory 
            WHERE quantity_in_stock < minimum_stock AND status = "Active"
        ''')
        low_stock = cursor.fetchone()[0]
        
        stats_text = (f"Total Parts: {total} | "
                     f"Total Value: ${value:,.2f} | "
                     f"Low Stock Items: {low_stock}")
        
        self.mro_stats_label.config(text=stats_text)
    
    def sort_mro_column(self, col):
        """Sort MRO treeview by column"""
        # Implement sorting logic here
        pass


# ============================================================================
# INTEGRATION INSTRUCTIONS
# ============================================================================
"""
To integrate this MRO Stock Management into your existing CMMS application:

1. Add this import at the top of your AIT_CMMS_REV3.py file:
   from mro_stock_module import MROStockManager

2. In your AIT_CMMS class __init__ method, add:
   self.mro_manager = MROStockManager(self)

3. In your create_all_manager_tabs() or create_gui() method, add:
   self.mro_manager.create_mro_tab(self.notebook)

4. The MRO Stock system will automatically use your existing SQLite database.

Example integration code:

    def create_all_manager_tabs(self):
        # ... your existing tabs ...
        
        # Add MRO Stock tab
        self.mro_manager.create_mro_tab(self.notebook)
"""