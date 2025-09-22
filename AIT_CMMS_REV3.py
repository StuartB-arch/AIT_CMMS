#!/usr/bin/env python3
"""
AIT Complete CMMS - Computerized Maintenance Management System
Fully functional CMMS with automatic PM scheduling, technician assignment, and comprehensive reporting
"""

import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import json
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import calendar
import random
import math
import re
from pathlib import Path




class DateStandardizer:
    """Utility class to standardize all dates in the CMMS database to YYYY-MM-DD format"""
    
    def __init__(self, conn):
        self.conn = conn
        self.date_patterns = [
            r'^\d{1,2}/\d{1,2}/\d{2}$',      # MM/DD/YY or M/D/YY
            r'^\d{1,2}/\d{1,2}/\d{4}$',     # MM/DD/YYYY or M/D/YYYY
            r'^\d{1,2}-\d{1,2}-\d{2}$',      # MM-DD-YY or M-D-YY
            r'^\d{1,2}-\d{1,2}-\d{4}$',     # MM-DD-YYYY or M-D-YYYY
            r'^\d{4}-\d{1,2}-\d{1,2}$'      # YYYY-MM-DD (already correct)
        ]
        
        self.date_formats = [
            '%m/%d/%y', '%#m/%#d/%y', '%-m/%-d/%y',  # Handle leading zeros
            '%m/%d/%Y', '%#m/%#d/%Y', '%-m/%-d/%Y',
            '%m-%d-%y', '%#m-%#d-%y', '%-m/%-d/%y',
            '%m-%d-%Y', '%#m-%#d-%Y', '%-m/%-d/%Y',
            '%Y-%m-%d'  # Target format
        ]
    
    def parse_date_flexible(self, date_str):
        """Parse date string using multiple formats and return standardized YYYY-MM-DD"""
        if not date_str or date_str.strip() == '':
            return None
            
        date_str = str(date_str).strip()
        
        # Already in correct format
        if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
            try:
                # Validate it's a real date
                datetime.strptime(date_str, '%Y-%m-%d')
                return date_str
            except ValueError:
                pass
        
        # Try all possible formats
        for date_format in self.date_formats:
            try:
                parsed_date = datetime.strptime(date_str, date_format)
                
                # Handle 2-digit years (assume 20xx if < 50, 19xx if >= 50)
                if parsed_date.year < 1950:
                    if parsed_date.year < 50:
                        parsed_date = parsed_date.replace(year=parsed_date.year + 2000)
                    else:
                        parsed_date = parsed_date.replace(year=parsed_date.year + 1900)
                
                return parsed_date.strftime('%Y-%m-%d')
                
            except ValueError:
                continue
        
        # If no format worked, return None
        print(f"Could not parse date: '{date_str}'")
        return None
    
    def standardize_all_dates(self):
        """Standardize all dates in the database to YYYY-MM-DD format"""
        cursor = self.conn.cursor()
        total_updated = 0
        errors = []
        
        # Tables and their date columns to standardize
        tables_to_update = {
            'equipment': [
                'last_monthly_pm', 'last_six_month_pm', 'last_annual_pm',
                'next_monthly_pm', 'next_six_month_pm', 'next_annual_pm'
            ],
            'pm_completions': [
                'completion_date', 'pm_due_date', 'next_annual_pm_date'
            ],
            'weekly_pm_schedules': [
                'week_start_date', 'scheduled_date', 'completion_date'
            ],
            'corrective_maintenance': [
                'created_date', 'completion_date'
            ],
            'cannot_find_assets': [
                'report_date'
            ],
            'run_to_failure_assets': [
                'completion_date'
            ]
        }
        
        for table, date_columns in tables_to_update.items():
            print(f"Processing table: {table}")
            
            try:
                # Get all rows from table
                cursor.execute(f'SELECT * FROM {table}')
                rows = cursor.fetchall()
                
                # Get column names
                cursor.execute(f'PRAGMA table_info({table})')
                column_info = cursor.fetchall()
                column_names = [col[1] for col in column_info]
                
                for row in rows:
                    row_dict = dict(zip(column_names, row))
                    updates_needed = {}
                    
                    # Check each date column
                    for date_col in date_columns:
                        if date_col in row_dict and row_dict[date_col]:
                            original_date = row_dict[date_col]
                            standardized_date = self.parse_date_flexible(original_date)
                            
                            if standardized_date and standardized_date != original_date:
                                updates_needed[date_col] = standardized_date
                    
                    # Update row if any dates need standardizing
                    if updates_needed:
                        update_parts = []
                        values = []
                        
                        for col, new_value in updates_needed.items():
                            update_parts.append(f'{col} = ?')
                            values.append(new_value)
                        
                        # Identify primary key or unique identifier
                        if table == 'equipment':
                            where_clause = 'bfm_equipment_no = ?'
                            values.append(row_dict['bfm_equipment_no'])
                        elif 'id' in row_dict:
                            where_clause = 'id = ?'
                            values.append(row_dict['id'])
                        else:
                            # Skip if no clear identifier
                            continue
                        
                        update_sql = f"UPDATE {table} SET {', '.join(update_parts)} WHERE {where_clause}"
                        
                        try:
                            cursor.execute(update_sql, values)
                            total_updated += 1
                            print(f"Updated {table} - {updates_needed}")
                        except Exception as e:
                            errors.append(f"Error updating {table}: {str(e)}")
                            
            except Exception as e:
                errors.append(f"Error processing table {table}: {str(e)}")
                continue
        
        # Commit changes
        try:
            self.conn.commit()
            return total_updated, errors
        except Exception as e:
            self.conn.rollback()
            errors.append(f"Error committing changes: {str(e)}")
            return 0, errors


class AITCMMSSystem:
    """Complete AIT CMMS - Computerized Maintenance Management System"""
    
    def check_empty_database_and_offer_restore(self):
        """Check if database is empty and offer to restore from backup"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM equipment')
            equipment_count = cursor.fetchone()[0]
            
            if equipment_count == 0:
                # Database is empty, offer restore
                result = messagebox.askyesno(
                    "Empty Database Detected",
                    "The database appears to be empty.\n\n"
                    "Would you like to restore data from a previous backup?\n\n"
                    "Click 'Yes' to browse available backups\n"
                    "Click 'No' to continue with empty database",
                    icon='question'
                )
                
                if result:
                    self.create_database_restore_dialog()
                    
        except Exception as e:
            print(f"Error checking empty database: {e}")
    
    
    def create_database_restore_dialog(self):
        """Create dialog to restore database from SharePoint backups - FIXED with proper buttons"""
        if not hasattr(self, 'backup_sync_dir') or not self.backup_sync_dir:
            messagebox.showerror("Error", "No backup directory configured. Please restart the application.")
            return
    
        dialog = tk.Toplevel(self.root)
        dialog.title("Restore Database from Backup")
        dialog.geometry("1000x700")  # Made larger
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Instructions
        instructions_frame = ttk.LabelFrame(dialog, text="Database Restore", padding=15)
        instructions_frame.pack(fill='x', padx=10, pady=5)
        
        instructions_text = f"""Select a backup file to restore your database from SharePoint.

    Current backup location: {self.backup_sync_dir}

    WARNING: Restoring a backup will:
    • Close the current database
    • Replace it with the selected backup
    • All unsaved changes will be lost
    • The application will reload with the restored data"""
    
        ttk.Label(instructions_frame, text=instructions_text, font=('Arial', 10)).pack(anchor='w')
    
        # Backup files list
        files_frame = ttk.LabelFrame(dialog, text="Available Backup Files (Last 15)", padding=10)
        files_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        # Create treeview for backup files
        self.backup_files_tree = ttk.Treeview(files_frame,
                                            columns=('Filename', 'Date Created', 'Size', 'Age'),
                                            show='headings')
    
        # Configure columns
        backup_columns = {
            'Filename': ('Backup Filename', 350),
            'Date Created': ('Date Created', 150),
            'Size': ('File Size', 100),
            'Age': ('Age (Days)', 100)
        }
    
        for col, (heading, width) in backup_columns.items():
            self.backup_files_tree.heading(col, text=heading)
            self.backup_files_tree.column(col, width=width)
    
        # Scrollbars
        backup_v_scrollbar = ttk.Scrollbar(files_frame, orient='vertical', command=self.backup_files_tree.yview)
        backup_h_scrollbar = ttk.Scrollbar(files_frame, orient='horizontal', command=self.backup_files_tree.xview)
        self.backup_files_tree.configure(yscrollcommand=backup_v_scrollbar.set, xscrollcommand=backup_h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.backup_files_tree.grid(row=0, column=0, sticky='nsew')
        backup_v_scrollbar.grid(row=0, column=1, sticky='ns')
        backup_h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        files_frame.grid_rowconfigure(0, weight=1)
        files_frame.grid_columnconfigure(0, weight=1)
        
        # Selection info
        selection_frame = ttk.LabelFrame(dialog, text="Selected Backup Info", padding=10)
        selection_frame.pack(fill='x', padx=10, pady=5)
    
        self.backup_info_label = ttk.Label(selection_frame, text="Loading backup files...", 
                                        font=('Arial', 10), foreground='blue')
        self.backup_info_label.pack(anchor='w')
    
        # Bind selection event
        self.backup_files_tree.bind('<<TreeviewSelect>>', self.on_backup_file_select)
    
        # Action buttons - FIXED with proper layout
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=15)
    
        # Left side buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side='left')
    
        ttk.Button(left_buttons, text="Refresh List", 
                command=self.load_backup_files).pack(side='left', padx=5)
        ttk.Button(left_buttons, text="Preview Backup", 
                command=self.preview_selected_backup).pack(side='left', padx=5)
    
        # Right side buttons  
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side='right')
    
        ttk.Button(right_buttons, text="Cancel", 
                command=dialog.destroy).pack(side='right', padx=5)
    
        # Main restore button - prominent in center
        center_buttons = ttk.Frame(button_frame)
        center_buttons.pack(expand=True)
        
        self.restore_button = ttk.Button(center_buttons, text="RESTORE SELECTED BACKUP", 
                                        command=self.restore_selected_backup, 
                                        state='disabled',
                                        width=25)
        self.restore_button.pack(pady=5)
    
        # Load backup files after creating the dialog
        self.root.after(100, self.load_backup_files)  # Load after dialog is fully created

    
    

    def load_backup_files(self):
        """Load available backup files from SharePoint - FIXED to show multiple files"""
        try:
            if not os.path.exists(self.backup_sync_dir):
                if hasattr(self, 'backup_info_label'):
                    self.backup_info_label.config(text="Backup directory not found", foreground='red')
                return
        
            # Clear existing items
            for item in self.backup_files_tree.get_children():
                self.backup_files_tree.delete(item)
        
            # Get all backup files
            backup_files = []
            try:
                all_files = os.listdir(self.backup_sync_dir)
                print(f"DEBUG: Found {len(all_files)} total files in backup directory")
            
                for filename in all_files:
                    if filename.startswith('ait_cmms_backup_') and filename.endswith('.db'):
                        file_path = os.path.join(self.backup_sync_dir, filename)
                        try:
                            # Get file stats
                            stat = os.stat(file_path)
                            file_size = stat.st_size
                            modified_time = datetime.fromtimestamp(stat.st_mtime)
                            age_days = (datetime.now() - modified_time).days
                        
                            backup_files.append({
                                'filename': filename,
                                'filepath': file_path,
                                'size': file_size,
                                'modified': modified_time,
                                'age_days': age_days
                            })
                            print(f"DEBUG: Added backup file: {filename}")
                        except Exception as e:
                            print(f"Error reading backup file {filename}: {e}")
                            continue
            except Exception as e:
                print(f"Error listing backup directory: {e}")
                if hasattr(self, 'backup_info_label'):
                    self.backup_info_label.config(text=f"Error reading backup directory: {str(e)}", foreground='red')
                return
        
            print(f"DEBUG: Total backup files found: {len(backup_files)}")
        
            # Sort by modification time (newest first)
            backup_files.sort(key=lambda x: x['modified'], reverse=True)
            
            # Limit to last 15 backups for better performance
            backup_files = backup_files[:15]
        
            # Add to tree
            for backup in backup_files:
                # Format file size
                size_mb = backup['size'] / (1024 * 1024)
                size_str = f"{size_mb:.1f} MB" if size_mb >= 1 else f"{backup['size']} bytes"
            
                item_id = self.backup_files_tree.insert('', 'end', values=(
                    backup['filename'],
                    backup['modified'].strftime('%Y-%m-%d %H:%M:%S'),
                    size_str,
                    f"{backup['age_days']} days"
                ))
            
                print(f"DEBUG: Inserted item: {backup['filename']}")
        
            # Update info label
            if hasattr(self, 'backup_info_label'):
                if backup_files:
                    self.backup_info_label.config(text=f"Found {len(backup_files)} backup files", foreground='green')
                else:
                    self.backup_info_label.config(text="No backup files found in directory", foreground='orange')
                
        except Exception as e:
            print(f"Error loading backup files: {e}")
            if hasattr(self, 'backup_info_label'):
                self.backup_info_label.config(text=f"Error loading backups: {str(e)}", foreground='red')


    def on_backup_file_select(self, event):
        """Handle backup file selection - ENHANCED"""
        try:
            selected = self.backup_files_tree.selection()
            if selected:
                item = self.backup_files_tree.item(selected[0])
                filename = item['values'][0]
                date_created = item['values'][1]
                file_size = item['values'][2]
                age = item['values'][3]
                
                # Show backup info
                info_text = f"✓ SELECTED: {filename}\n"
                info_text += f"Created: {date_created}\n"
                info_text += f"Size: {file_size}\n"
                info_text += f"Age: {age}\n\n"
                info_text += "Click 'RESTORE SELECTED BACKUP' to proceed"
            
                self.backup_info_label.config(text=info_text, foreground='darkgreen')
            
                # Enable restore button
                self.restore_button.config(state='normal')
                self.restore_button.config(text=f"RESTORE: {filename}")
            else:
                self.backup_info_label.config(text="Select a backup file to see details", foreground='gray')
                self.restore_button.config(state='disabled')
                self.restore_button.config(text="RESTORE SELECTED BACKUP")
        except Exception as e:
            print(f"Error in backup file selection: {e}")
            self.backup_info_label.config(text="Error selecting backup file", foreground='red')


    def preview_selected_backup(self):
        """Preview selected backup file contents"""
        selected = self.backup_files_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a backup file to preview")
            return
    
        try:
            item = self.backup_files_tree.item(selected[0])
            filename = item['values'][0]
            filepath = os.path.join(self.backup_sync_dir, filename)
        
            if not os.path.exists(filepath):
                messagebox.showerror("Error", f"Backup file not found: {filename}")
                return
        
            # Create preview dialog
            preview_dialog = tk.Toplevel(self.root)
            preview_dialog.title(f"Preview Backup: {filename}")
            preview_dialog.geometry("800x600")
            preview_dialog.transient(self.root)
            preview_dialog.grab_set()
        
            # Preview text area
            text_frame = ttk.Frame(preview_dialog)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
            
            preview_text = tk.Text(text_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=preview_text.yview)
            preview_text.configure(yscrollcommand=scrollbar.set)
            
            preview_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
            # Connect to backup database and get preview info
            try:
                backup_conn = sqlite3.connect(filepath)
                backup_cursor = backup_conn.cursor()
                
                preview_info = f"BACKUP DATABASE PREVIEW\n"
                preview_info += f"File: {filename}\n"
                preview_info += f"=" * 80 + "\n\n"
                
                # Get table counts
                tables = [
                    ('equipment', 'Equipment/Assets'),
                    ('pm_completions', 'PM Completions'),
                    ('weekly_pm_schedules', 'Weekly Schedules'),
                    ('corrective_maintenance', 'Corrective Maintenance'),
                    ('cannot_find_assets', 'Cannot Find Assets'),
                    ('run_to_failure_assets', 'Run to Failure Assets'),
                    ('pm_templates', 'PM Templates')
                ]
            
                preview_info += "DATABASE CONTENTS:\n"
                preview_info += "-" * 40 + "\n"
            
                total_records = 0
                for table_name, display_name in tables:
                    try:
                        backup_cursor.execute(f'SELECT COUNT(*) FROM {table_name}')
                        count = backup_cursor.fetchone()[0]
                        total_records += count
                        preview_info += f"{display_name}: {count} records\n"
                    except Exception as e:
                        preview_info += f"{display_name}: Error reading ({str(e)})\n"
            
                preview_info += f"\nTotal Records: {total_records}\n\n"
            
                # Get some sample equipment data
                try:
                    backup_cursor.execute('''
                        SELECT bfm_equipment_no, description, status 
                        FROM equipment 
                        ORDER BY updated_date DESC 
                        LIMIT 10
                    ''')
                    equipment_sample = backup_cursor.fetchall()
                
                    if equipment_sample:
                        preview_info += "RECENT EQUIPMENT (Sample):\n"
                        preview_info += "-" * 40 + "\n"
                        for bfm_no, desc, status in equipment_sample:
                            desc_short = (desc[:30] + '...') if desc and len(desc) > 30 else (desc or 'No description')
                            preview_info += f"{bfm_no}: {desc_short} ({status or 'Active'})\n"
                        preview_info += "\n"
                except:
                    pass
            
                # Get recent PM completions
                try:
                    backup_cursor.execute('''
                        SELECT completion_date, COUNT(*) as count
                        FROM pm_completions 
                        GROUP BY completion_date 
                        ORDER BY completion_date DESC 
                        LIMIT 10
                    ''')
                    pm_dates = backup_cursor.fetchall()
                
                    if pm_dates:
                        preview_info += "RECENT PM ACTIVITY:\n"
                        preview_info += "-" * 40 + "\n"
                        for date, count in pm_dates:
                            preview_info += f"{date}: {count} PM completions\n"
                        preview_info += "\n"
                except:
                    pass
                
                # Database metadata
                try:
                    backup_cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                    all_tables = [row[0] for row in backup_cursor.fetchall()]
                    preview_info += f"DATABASE STRUCTURE:\n"
                    preview_info += "-" * 40 + "\n"
                    preview_info += f"Total Tables: {len(all_tables)}\n"
                    preview_info += f"Tables: {', '.join(all_tables)}\n"
                except:
                    pass
            
                backup_conn.close()
            
                preview_text.insert('1.0', preview_info)
                preview_text.config(state='disabled')
            
            except Exception as e:
                preview_text.insert('1.0', f"Error previewing backup database:\n{str(e)}")
                preview_text.config(state='disabled')
        
            # Close button
            ttk.Button(preview_dialog, text="Close", command=preview_dialog.destroy).pack(pady=10)
        
        except Exception as e:
            messagebox.showerror("Preview Error", f"Failed to preview backup: {str(e)}")

    def restore_selected_backup(self):
        """Restore the selected backup file"""
        selected = self.backup_files_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a backup file to restore")
            return
    
        try:
            item = self.backup_files_tree.item(selected[0])
            filename = item['values'][0]
            date_created = item['values'][1]
            file_size = item['values'][2]
        
            # Get the full file path
            source_filepath = os.path.join(self.backup_sync_dir, filename)
        
            if not os.path.exists(source_filepath):
                messagebox.showerror("Error", f"Backup file not found: {filename}")
                return
        
            # Confirmation dialog with detailed info
            confirm_msg = f"""RESTORE DATABASE FROM BACKUP

    Selected Backup:
    • File: {filename}
    • Created: {date_created}
    • Size: {file_size}

    WARNING: This action will:
    • Close the current database
    • Replace it completely with the backup data
    • All current unsaved changes will be lost
    • The application will reload with the backup data

    This action cannot be undone.

    Are you sure you want to proceed?"""
        
            result = messagebox.askyesno("Confirm Database Restore", confirm_msg, 
                                        icon='warning', default='no')
        
            if not result:
                return
        
            # Create progress dialog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Restoring Database...")
            progress_dialog.geometry("400x150")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            
            ttk.Label(progress_dialog, text="Restoring database from backup...", 
                    font=('Arial', 12)).pack(pady=20)
        
            progress_var = tk.StringVar(value="Preparing restore...")
            progress_label = ttk.Label(progress_dialog, textvariable=progress_var)
            progress_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_dialog, mode='indeterminate')
            progress_bar.pack(pady=10, padx=20, fill='x')
            progress_bar.start()
        
            # Update GUI
            self.root.update()
            
            # Perform the restore
            current_db_path = 'ait_cmms_database.db'
        
            # Step 1: Close current database connection
            progress_var.set("Closing current database...")
            self.root.update()
        
            if hasattr(self, 'conn'):
                try:
                    self.conn.close()
                except:
                    pass
        
            # Step 2: Backup current database (just in case)
            progress_var.set("Backing up current database...")
            self.root.update()
        
            if os.path.exists(current_db_path):
                backup_current_path = f"{current_db_path}.pre_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copy2(current_db_path, backup_current_path)
        
            # Step 3: Copy backup to current location
            progress_var.set("Restoring backup data...")
            self.root.update()
        
            shutil.copy2(source_filepath, current_db_path)
        
            # Step 4: Reconnect to database
            progress_var.set("Reconnecting to database...")
            self.root.update()
        
            self.conn = sqlite3.connect(current_db_path)
        
            # Step 5: Refresh all data displays
            progress_var.set("Refreshing application data...")
            self.root.update()
        
            # Refresh all displays
            self.load_equipment_data()
            self.refresh_equipment_list()
            self.load_recent_completions()
            self.load_corrective_maintenance()
            if hasattr(self, 'load_cannot_find_assets'):
                self.load_cannot_find_assets()
            if hasattr(self, 'load_run_to_failure_assets'):
                self.load_run_to_failure_assets()
            if hasattr(self, 'load_pm_templates'):
                self.load_pm_templates()
        
            # Update statistics
            if hasattr(self, 'update_equipment_statistics'):
                self.update_equipment_statistics()
        
            progress_bar.stop()
            progress_dialog.destroy()
        
            # Close the restore dialog
            if hasattr(self, 'backup_files_tree'):
                # Find and close the restore dialog
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Toplevel) and "Restore Database" in widget.title():
                        widget.destroy()
                        break
        
            # Show success message
            messagebox.showinfo("Restore Complete", 
                               f"Database successfully restored from backup!\n\n"
                               f"Restored from: {filename}\n"
                               f"Created: {date_created}\n"
                               f"The application has been refreshed with the restored data.")
        
            self.update_status(f"Database restored from backup: {filename}")
        
        except Exception as e:
            # Try to reconnect to original database
            try:
                self.conn = sqlite3.connect('ait_cmms_database.db')
            except:
                pass
        
            messagebox.showerror("Restore Error", f"Failed to restore database backup:\n\n{str(e)}")
            print(f"Database restore error: {e}")

    def add_database_restore_button(self):
        """Add database restore button to the equipment tab"""
        try:
            if hasattr(self, 'equipment_frame'):
                 #Find the controls frame and add the button
                for widget in self.equipment_frame.winfo_children():
                    if isinstance(widget, ttk.LabelFrame) and "Equipment Controls" in widget['text']:
                        ttk.Button(widget, text="📁 Restore Database from Backup", 
                                 command=self.create_database_restore_dialog,
                                 width=30).pack(side='left', padx=5)
                        break
        except Exception as e:
            print(f"Error adding restore button: {e}")
    
    
    
    
    
    def add_logo_to_main_window(self):
        """Add AIT logo to the main application window - LEFT SIDE ONLY"""
        try:
            from tkinter import PhotoImage
            from PIL import Image, ImageTk
            
            # Get the directory where the script is located
            script_dir = os.path.dirname(os.path.abspath(__file__))
            img_dir = os.path.join(script_dir, "img")
            logo_path = os.path.join(img_dir, "ait_logo.png")
        
            # Create img directory if it doesn't exist
            if not os.path.exists(img_dir):
                os.makedirs(img_dir)
                print(f"Created img directory: {img_dir}")
        
            # Alternative paths to try
            alternative_paths = [
                os.path.join(script_dir, "ait_logo.png"),  # Same directory as script
                os.path.join(script_dir, "img", "ait_logo.png"),  # img subdirectory
                "ait_logo.png"  # Current working directory
            ]
        
            logo_found = False
            for path in alternative_paths:
                if os.path.exists(path):
                    logo_path = path
                    logo_found = True
                    print(f"Found logo at: {logo_path}")
                    break
        
            if not logo_found:
                print(f"Logo file not found. Tried paths: {alternative_paths}")
                print("Please place your logo file in one of these locations.")
                return
            
            if os.path.exists(logo_path):
                # Open and resize image for tkinter
                pil_image = Image.open(logo_path)
                pil_image = pil_image.resize((200, 60), Image.Resampling.LANCZOS)  # Reasonable size for left corner
                
                # Convert to PhotoImage
                self.logo_image = ImageTk.PhotoImage(pil_image)
                
                # Create logo frame at top left of window
                logo_frame = ttk.Frame(self.root)
                logo_frame.pack(side='top', fill='x', padx=10, pady=5)
            
                # Add logo label (left aligned)
                logo_label = ttk.Label(logo_frame, image=self.logo_image)
                logo_label.pack(side='left')
                
                # Optional: Add a subtle separator line below
                separator = ttk.Separator(self.root, orient='horizontal')
                separator.pack(fill='x', padx=10, pady=2)
            
        except ImportError:
            print("PIL (Pillow) not installed. Install with: pip install Pillow")
        except Exception as e:
            print(f"Error loading logo: {e}")
    
    
    
    
    def sync_database_on_startup(self):
        """Lighter sync check after initialization (already handled in pre-init)"""
        try:
            if not hasattr(self, 'backup_sync_dir'):
                return False
        
            # Just update the status, heavy lifting was done in pre-init sync
            if os.path.exists('ait_cmms_database.db'):
                self.update_status("Database loaded - SharePoint sync completed during startup")
                return True
            else:
                self.update_status("Database initialization completed")
                return False
            
        except Exception as e:
            print(f"Error in startup sync check: {e}")
            return False
    
    
    
    
    
    
    

    def get_sharepoint_backup_path(self):
        """Get path to specific SharePoint PM/CM folder with comprehensive fallbacks"""
        try:
            home_dir = os.path.expanduser("~")
        
            # Primary target: Your specific SharePoint PM/CM path
            primary_sharepoint_path = os.path.join(home_dir, "Advanced Integration Technology", "PM CM - Documents", "General", "Asset Maintenance", "CMMS_Backups")
            
            # Check if the SharePoint PM/CM path exists
            sharepoint_parent = os.path.join(home_dir, "Advanced Integration Technology", "PM CM - Documents", "General", "Asset Maintenance")
            if os.path.exists(sharepoint_parent) and os.path.isdir(sharepoint_parent):
                # Create the CMMS_Backups subfolder in Asset Maintenance
                os.makedirs(primary_sharepoint_path, exist_ok=True)
                
                # Test write permissions
                test_file = os.path.join(primary_sharepoint_path, "test_write.tmp")
                try:
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    print(f"Using SharePoint PM/CM Asset Maintenance folder: {primary_sharepoint_path}")
                    return primary_sharepoint_path
                except Exception as e:
                    print(f"Cannot write to SharePoint PM/CM path: {e}")
        
            # Fallback 1: Try other locations in the same SharePoint site
            fallback_sharepoint_paths = [
                os.path.join(home_dir, "Advanced Integration Technology", "PM CM - Documents", "General", "CMMS_Backups"),
                os.path.join(home_dir, "Advanced Integration Technology", "PM CM - Documents", "CMMS_Backups"),
                os.path.join(home_dir, "Advanced Integration Technology", "CMMS_Backups")
            ]
        
            for fallback_path in fallback_sharepoint_paths:
                parent_dir = os.path.dirname(fallback_path)
                if os.path.exists(parent_dir) and os.path.isdir(parent_dir):
                    try:
                        os.makedirs(fallback_path, exist_ok=True)
                        # Test write permissions
                        test_file = os.path.join(fallback_path, "test_write.tmp")
                        with open(test_file, 'w') as f:
                            f.write("test")
                        os.remove(test_file)
                        print(f"Using SharePoint fallback location: {fallback_path}")
                        return fallback_path
                    except Exception as e:
                        print(f"Cannot write to fallback SharePoint path {fallback_path}: {e}")
                        continue
        
            # Fallback 2: Work OneDrive root with PM/CM structure
            work_onedrive_path = os.path.join(home_dir, "OneDrive - Advanced Integration Technology", "PM-CM", "CMMS_Backups")
            work_onedrive_root = os.path.join(home_dir, "OneDrive - Advanced Integration Technology")
            if os.path.exists(work_onedrive_root) and os.path.isdir(work_onedrive_root):
                try:
                    os.makedirs(work_onedrive_path, exist_ok=True)
                    # Test write permissions
                    test_file = os.path.join(work_onedrive_path, "test_write.tmp")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    print(f"Using work OneDrive with PM/CM structure: {work_onedrive_path}")
                    return work_onedrive_path
                except Exception as e:
                    print(f"Cannot write to work OneDrive PM/CM path: {e}")
        
            # Fallback 3: Work OneDrive root
            work_onedrive_basic = os.path.join(home_dir, "OneDrive - Advanced Integration Technology", "AIT_CMMS_Backups")
            if os.path.exists(work_onedrive_root) and os.path.isdir(work_onedrive_root):
                try:
                    os.makedirs(work_onedrive_basic, exist_ok=True)
                    # Test write permissions
                    test_file = os.path.join(work_onedrive_basic, "test_write.tmp")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    print(f"Using work OneDrive basic location: {work_onedrive_basic}")
                    return work_onedrive_basic
                except Exception as e:
                    print(f"Cannot write to work OneDrive basic path: {e}")
        
            # Fallback 4: Personal OneDrive
            personal_onedrive_path = os.path.join(home_dir, "OneDrive", "AIT_CMMS_Backups")
            personal_onedrive_root = os.path.join(home_dir, "OneDrive")
            if os.path.exists(personal_onedrive_root) and os.path.isdir(personal_onedrive_root):
                try:
                    os.makedirs(personal_onedrive_path, exist_ok=True)
                    # Test write permissions
                    test_file = os.path.join(personal_onedrive_path, "test_write.tmp")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    print(f"Using personal OneDrive: {personal_onedrive_path}")
                    return personal_onedrive_path
                except Exception as e:
                    print(f"Cannot write to personal OneDrive: {e}")
        
            # Fallback 5: Local backup (ultimate fallback)
            local_backup_path = os.path.join(home_dir, "AIT_CMMS_Backups")
            os.makedirs(local_backup_path, exist_ok=True)
            print(f"Using local backup (no OneDrive available): {local_backup_path}")
            return local_backup_path
        
        except Exception as e:
            print(f"Error finding backup path: {e}")
            # Ultimate emergency fallback
            emergency_path = os.path.join(os.path.expanduser("~"), "AIT_CMMS_Backups")
            os.makedirs(emergency_path, exist_ok=True)
            print(f"Using emergency local backup: {emergency_path}")
            return emergency_path



    def schedule_sharepoint_only_backups(self, sync_dir):
        """Schedule automatic backups to SharePoint only"""
        try:
            # Create initial backup immediately (5 seconds after startup)
            self.root.after(5000, lambda: self.sharepoint_only_backup(sync_dir))
        
            # Schedule daily backups - run every 24 hours
            self.root.after(24 * 60 * 60 * 1000, lambda: self.recurring_sharepoint_backup(sync_dir))
        
            print(f"Scheduled SharePoint-only backups to: {sync_dir}")
        
        except Exception as e:
            print(f"Error scheduling SharePoint backups: {e}")

    def recurring_sharepoint_backup(self, sync_dir):
        """Recurring backup function that runs daily"""
        try:
            # Perform the backup
            self.sharepoint_only_backup(sync_dir)
        
            # Reschedule for next day (24 hours later)
            self.root.after(24 * 60 * 60 * 1000, lambda: self.recurring_sharepoint_backup(sync_dir))
        
        except Exception as e:
            print(f"Error in recurring backup: {e}")
            # Try to reschedule anyway
            self.root.after(24 * 60 * 60 * 1000, lambda: self.recurring_sharepoint_backup(sync_dir))







    def setup_existing_database_with_sharepoint_backup(self):
        """Set up existing database with SharePoint-only backups - FIXED VERSION"""
        try:
            # First, check if database file exists
            db_file = 'ait_cmms_database.db'
            if not os.path.exists(db_file):
                messagebox.showerror("Database Error", 
                                "Database file not found: ait_cmms_database.db\n\n"
                                "Please ensure the database file is in the same directory as the application.")
                return None
    
            # Get the SharePoint-synced backup path
            sync_dir = self.get_sharepoint_backup_path()
        
            if not sync_dir:
                print("Could not establish backup directory")
                return None
        
            print(f"Using backup directory: {sync_dir}")
        
            # Store sync_dir for later use (when GUI is created)
            self.backup_sync_dir = sync_dir
            
            
            
            # Schedule automatic backups to SharePoint only
            self.schedule_sharepoint_only_backups(sync_dir)
        
            return sync_dir
        
        except Exception as e:
            error_msg = f"Error setting up database with SharePoint backup: {e}"
            print(error_msg)
            self.update_status(error_msg)
            return None


    def schedule_sharepoint_only_backups(self, sync_dir):
        """Schedule daily automatic backups only"""
        try:
            # Schedule daily backups - run every 24 hours from now
            self.root.after(24 * 60 * 60 * 1000, lambda: self.recurring_sharepoint_backup(sync_dir))
        
            print(f"Scheduled daily SharePoint backups to: {sync_dir}")
        
        except Exception as e:
            print(f"Error scheduling SharePoint backups: {e}")

    def on_closing(self):
        """Handle application closing with final backup"""
        try:
            if hasattr(self, 'backup_sync_dir'):
                print("Creating final backup before closing...")
                self.sharepoint_only_backup(self.backup_sync_dir)
                print("Final backup completed")
        
            # Close database connection
            if hasattr(self, 'conn'):
                self.conn.close()
        
            # Destroy the main window
            self.root.destroy()
        
        except Exception as e:
            print(f"Error during closing: {e}")
            self.root.destroy()




    def create_initial_backup(self, sync_dir, db_file):
        """Create initial backup of existing database"""
        try:
            if not os.path.exists(db_file):
                print(f"Database file {db_file} not found")
                return
            
            # Create timestamped backup filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(sync_dir, f"ait_cmms_backup_{timestamp}.db")
            
            # Create backup of existing database
            shutil.copy2(db_file, backup_file)
            
            print(f"Initial backup created: {backup_file}")
            self.update_status(f"Initial backup created in backup folder")
        
        except Exception as e:
            print(f"Error creating initial backup: {e}")
            self.update_status(f"Backup error: {str(e)}")

    def sharepoint_only_backup(self, sync_dir):
        """Create backup directly in SharePoint folder only - FIXED VERSION"""
        try:
            db_file = 'ait_cmms_database.db'
        
            # Check if database file exists
            if not os.path.exists(db_file):
                print(f"Database file {db_file} not found for backup")
                return
            
            # Close current connection temporarily for clean backup
            if hasattr(self, 'conn'):
                try:
                    self.conn.close()
                except:
                    pass
        
            # Create timestamped backup filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(sync_dir, f"ait_cmms_backup_{timestamp}.db")
            
            # Copy database file directly to backup location
            shutil.copy2(db_file, backup_file)
            
            # Reopen connection
            self.conn = sqlite3.connect(db_file)
            
            # Clean up old backups (keep last 10)
            self.cleanup_old_backups(sync_dir, keep_last=10)
            
            self.update_status(f"Backup created: {os.path.basename(backup_file)}")
            print(f"Backup successful: {backup_file}")
        
        except Exception as e:
            error_msg = f"Backup failed: {e}"
            self.update_status(error_msg)
            print(error_msg)
            # Try to reopen connection if it failed
            try:
                self.conn = sqlite3.connect('ait_cmms_database.db')
            except:
                pass

    def cleanup_old_backups(self, backup_dir, keep_last=10):
        """Keep only the most recent backup files in SharePoint"""
        try:
            # Get all backup files
            backup_files = []
            for f in os.listdir(backup_dir):
                if f.startswith('ait_cmms_backup_') and f.endswith('.db'):
                    full_path = os.path.join(backup_dir, f)
                    backup_files.append((full_path, os.path.getmtime(full_path)))
    
            if len(backup_files) <= keep_last:
                print(f"Only {len(backup_files)} backups found, no cleanup needed")
                return  # No cleanup needed
    
            # Sort by modification time (newest first)
            backup_files.sort(key=lambda x: x[1], reverse=True)
        
            # Delete older backups
            deleted_count = 0
            for old_backup_path, _ in backup_files[keep_last:]:
                try:
                    os.remove(old_backup_path)
                    print(f"Removed old SharePoint backup: {os.path.basename(old_backup_path)}")
                    deleted_count += 1
                except Exception as e:
                    print(f"Error removing old backup {old_backup_path}: {e}")
                
            if deleted_count > 0:
                print(f"Cleaned up {deleted_count} old SharePoint backups, kept {keep_last} most recent")
            
        except Exception as e:
            print(f"Error cleaning up old backups: {e}")
    
    
    def cleanup_local_backups(self):
        """Remove old timestamped local backup files, keep only the single local backup"""
        try:
            import glob
        
            # Find all timestamped local backup files
            pattern = "ait_cmms_database.db.local_backup_*"
            old_backups = glob.glob(pattern)
        
            removed_count = 0
            for backup_file in old_backups:
                try:
                    os.remove(backup_file)
                    print(f"Removed old local backup: {backup_file}")
                    removed_count += 1
                except Exception as e:
                    print(f"Error removing {backup_file}: {e}")
        
            if removed_count > 0:
                print(f"Cleaned up {removed_count} old local backup files")
            
        except Exception as e:
            print(f"Error cleaning up local backups: {e}")
    
    
    
    
    
    
    def test_backup_now(self):
        """Manual backup test for debugging"""
        try:
            sync_dir = self.get_sharepoint_backup_path()
            if sync_dir:
                self.sharepoint_only_backup(sync_dir)
                messagebox.showinfo("Test Complete", f"Backup test completed.\nBackup location: {sync_dir}")
            else:
                messagebox.showerror("Test Failed", "Could not determine backup location")
        except Exception as e:
            messagebox.showerror("Test Error", f"Backup test failed: {str(e)}")
    
    
    

    def init_pm_templates_database(self):
        """Initialize PM templates database tables"""
        cursor = self.conn.cursor()
    
        # PM Templates table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pm_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bfm_equipment_no TEXT,
                template_name TEXT,
                pm_type TEXT,
                checklist_items TEXT,  -- JSON string
                special_instructions TEXT,
                safety_notes TEXT,
                estimated_hours REAL,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
    
        # Default checklist items for fallback
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS default_pm_checklist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pm_type TEXT,
                step_number INTEGER,
                description TEXT,
                is_active BOOLEAN DEFAULT 1
            )
        ''')
    
        # Insert default checklist if empty
        cursor.execute('SELECT COUNT(*) FROM default_pm_checklist')
        if cursor.fetchone()[0] == 0:
            default_items = [
                (1, "Special Equipment Used (List):"),
                (2, "Validate your maintenance with Date / Stamp / Hours"),
                (3, "Refer to drawing when performing maintenance"),
                (4, "Make sure all instruments are properly calibrated"),
                (5, "Make sure tool is properly identified"),
                (6, "Make sure all mobile mechanisms move fluidly"),
                (7, "Visually inspect the welds"),
                (8, "Take note of any anomaly or defect (create a CM if needed)"),
                (9, "Check all screws. Tighten if needed."),
                (10, "Check the pins for wear"),
                (11, "Make sure all tooling is secured to the equipment with cable"),
                (12, "Ensure all tags (BFM and SAP) are applied and securely fastened"),
                (13, "All documentation are picked up from work area"),
                (14, "All parts and tools have been picked up"),
                (15, "Workspace has been cleaned up"),
                (16, "Dry runs have been performed (tests, restarts, etc.)"),
                (17, "Ensure that AIT Sticker is applied")
            ]
        
            for step_num, description in default_items:
                cursor.execute('''
                    INSERT INTO default_pm_checklist (pm_type, step_number, description)
                    VALUES ('All', ?, ?)
                ''', (step_num, description))
    
        self.conn.commit()

    def create_custom_pm_templates_tab(self):
        """Create PM Templates management tab"""
        self.pm_templates_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pm_templates_frame, text="PM Templates")
    
        # Controls
        controls_frame = ttk.LabelFrame(self.pm_templates_frame, text="PM Template Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Button(controls_frame, text="Create Custom Template", 
                command=self.create_custom_pm_template_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Edit Template", 
                command=self.edit_pm_template_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Preview Template", 
                command=self.preview_pm_template).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Delete Template", 
                command=self.delete_pm_template).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export Template to PDF", 
                command=self.export_custom_template_pdf).pack(side='left', padx=5)
    
        # Search frame
        search_frame = ttk.Frame(self.pm_templates_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Label(search_frame, text="Search Templates:").pack(side='left', padx=5)
        self.template_search_var = tk.StringVar()
        self.template_search_var.trace('w', self.filter_template_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.template_search_var, width=30)
        search_entry.pack(side='left', padx=5)
    
        # Templates list
        list_frame = ttk.LabelFrame(self.pm_templates_frame, text="PM Templates", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        self.templates_tree = ttk.Treeview(list_frame,
                                        columns=('BFM No', 'Template Name', 'PM Type', 'Steps', 'Est Hours', 'Updated'),
                                        show='headings')
    
        template_columns = {
            'BFM No': ('BFM Equipment No', 120),
            'Template Name': ('Template Name', 200),
            'PM Type': ('PM Type', 100),
            'Steps': ('# Steps', 80),
            'Est Hours': ('Est Hours', 80),
            'Updated': ('Last Updated', 120)
        }
    
        for col, (heading, width) in template_columns.items():
            self.templates_tree.heading(col, text=heading)
            self.templates_tree.column(col, width=width)
    
        # Scrollbars
        template_v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.templates_tree.yview)
        template_h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.templates_tree.xview)
        self.templates_tree.configure(yscrollcommand=template_v_scrollbar.set, xscrollcommand=template_h_scrollbar.set)
    
        # Pack treeview and scrollbars
        self.templates_tree.grid(row=0, column=0, sticky='nsew')
        template_v_scrollbar.grid(row=0, column=1, sticky='ns')
        template_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
    
        # Load templates
        self.load_pm_templates()

    def create_custom_pm_template_dialog(self):
        """Dialog to create custom PM template for specific equipment"""
        print("DEBUG: Starting create_custom_pm_template_dialog method")  # Add this line
    
        dialog = tk.Toplevel(self.root)
        """Dialog to create custom PM template for specific equipment"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Custom PM Template")
        dialog.geometry("800x750")
        dialog.transient(self.root)
        dialog.grab_set()

        # Equipment selection
        header_frame = ttk.LabelFrame(dialog, text="Template Information", padding=10)
        header_frame.pack(fill='x', padx=10, pady=5)

        # BFM Equipment selection
        ttk.Label(header_frame, text="BFM Equipment Number:").grid(row=0, column=0, sticky='w', pady=5)
        bfm_var = tk.StringVar()
        bfm_combo = ttk.Combobox(header_frame, textvariable=bfm_var, width=25)
        bfm_combo.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        # Populate equipment list
        cursor = self.conn.cursor()
        cursor.execute('SELECT bfm_equipment_no, description FROM equipment ORDER BY bfm_equipment_no')
        equipment_list = cursor.fetchall()
        bfm_combo['values'] = [f"{bfm} - {desc[:30]}..." if len(desc) > 30 else f"{bfm} - {desc}" 
                            for bfm, desc in equipment_list]

        # Template name
        ttk.Label(header_frame, text="Template Name:").grid(row=0, column=2, sticky='w', pady=5, padx=(20,5))
        template_name_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=template_name_var, width=25).grid(row=0, column=3, sticky='w', padx=5, pady=5)

        # PM Type
        ttk.Label(header_frame, text="PM Type:").grid(row=1, column=0, sticky='w', pady=5)
        pm_type_var = tk.StringVar(value='Annual')
        pm_type_combo = ttk.Combobox(header_frame, textvariable=pm_type_var, 
                                    values=['Monthly', 'Six Month', 'Annual'], width=22)
        pm_type_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # Estimated hours
        ttk.Label(header_frame, text="Estimated Hours:").grid(row=1, column=2, sticky='w', pady=5, padx=(20,5))
        est_hours_var = tk.StringVar(value="1.0")
        ttk.Entry(header_frame, textvariable=est_hours_var, width=10).grid(row=1, column=3, sticky='w', padx=5, pady=5)

        # Custom checklist section
        checklist_frame = ttk.LabelFrame(dialog, text="Custom PM Checklist", padding=10)
        checklist_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Checklist controls
        controls_subframe = ttk.Frame(checklist_frame)
        controls_subframe.pack(fill='x', pady=5)
    
        # Checklist listbox with scrollbar
        list_frame = ttk.Frame(checklist_frame)
        list_frame.pack(fill='both', expand=True, pady=5)

        checklist_listbox = tk.Listbox(list_frame, height=15, font=('Arial', 9))
        list_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=checklist_listbox.yview)
        checklist_listbox.configure(yscrollcommand=list_scrollbar.set)

        checklist_listbox.pack(side='left', fill='both', expand=True)
        list_scrollbar.pack(side='right', fill='y')

        # Step editing
        edit_frame = ttk.LabelFrame(checklist_frame, text="Edit Selected Step", padding=5)
        edit_frame.pack(fill='x', pady=5)

        step_text_var = tk.StringVar()
        step_entry = ttk.Entry(edit_frame, textvariable=step_text_var, width=80)
        step_entry.pack(side='left', fill='x', expand=True, padx=5)

        # Special instructions and safety notes
        notes_frame = ttk.LabelFrame(dialog, text="Additional Information", padding=10)
        notes_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(notes_frame, text="Special Instructions:").grid(row=0, column=0, sticky='nw', pady=2)
        special_instructions_text = tk.Text(notes_frame, height=3, width=50)
        special_instructions_text.grid(row=0, column=1, sticky='ew', padx=5, pady=2)

        ttk.Label(notes_frame, text="Safety Notes:").grid(row=1, column=0, sticky='nw', pady=2)
        safety_notes_text = tk.Text(notes_frame, height=3, width=50)
        safety_notes_text.grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        notes_frame.grid_columnconfigure(1, weight=1)

        # DEFINE ALL HELPER FUNCTIONS FIRST
        def add_checklist_step():
            step_text = step_text_var.get().strip()
            if step_text:
                step_num = checklist_listbox.size() + 1
                checklist_listbox.insert('end', f"{step_num}. {step_text}")
                step_text_var.set('')

        def remove_checklist_step():
            selection = checklist_listbox.curselection()
            if selection:
                checklist_listbox.delete(selection[0])
                renumber_steps()

        def renumber_steps():
            items = []
            for i in range(checklist_listbox.size()):
                step_text = checklist_listbox.get(i)
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                items.append(f"{i+1}. {step_content}")
        
            checklist_listbox.delete(0, 'end')
            for item in items:
                checklist_listbox.insert('end', item)

        def update_selected_step():
            selection = checklist_listbox.curselection()
            if selection and step_text_var.get().strip():
                step_num = selection[0] + 1
                new_text = f"{step_num}. {step_text_var.get().strip()}"
                checklist_listbox.delete(selection[0])
                checklist_listbox.insert(selection[0], new_text)

        def move_step_up():
            selection = checklist_listbox.curselection()
            if selection and selection[0] > 0:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx-1, item)
                checklist_listbox.selection_set(idx-1)
                renumber_steps()

        def move_step_down():
            selection = checklist_listbox.curselection()
            if selection and selection[0] < checklist_listbox.size()-1:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx+1, item)
                checklist_listbox.selection_set(idx+1)
                renumber_steps()

        def load_default_template():
            cursor = self.conn.cursor()
            cursor.execute('SELECT description FROM default_pm_checklist ORDER BY step_number')
            default_steps = cursor.fetchall()
        
            checklist_listbox.delete(0, 'end')
            for i, (step,) in enumerate(default_steps, 1):
                checklist_listbox.insert('end', f"{i}. {step}")

        def save_template():
            try:
                # Validate inputs
                if not bfm_var.get():
                    messagebox.showerror("Error", "Please select equipment")
                    return
            
                if not template_name_var.get().strip():
                    messagebox.showerror("Error", "Please enter template name")
                    return
            
                # Extract BFM number from combo selection
                bfm_no = bfm_var.get().split(' - ')[0]
            
                # Get checklist items
                checklist_items = []
                for i in range(checklist_listbox.size()):
                    step_text = checklist_listbox.get(i)
                    step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                    checklist_items.append(step_content)
            
                if not checklist_items:
                    messagebox.showerror("Error", "Please add at least one checklist item")
                    return
            
                # Save to database
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO pm_templates 
                    (bfm_equipment_no, template_name, pm_type, checklist_items, 
                    special_instructions, safety_notes, estimated_hours)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    bfm_no,
                    template_name_var.get().strip(),
                    pm_type_var.get(),
                    json.dumps(checklist_items),
                    special_instructions_text.get('1.0', 'end-1c'),
                    safety_notes_text.get('1.0', 'end-1c'),
                    float(est_hours_var.get() or 1.0)
                ))
            
                self.conn.commit()
                messagebox.showinfo("Success", "Custom PM template created successfully!")
                dialog.destroy()
                self.load_pm_templates()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {str(e)}")

        def on_step_select(event):
            selection = checklist_listbox.curselection()
            if selection:
                step_text = checklist_listbox.get(selection[0])
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                step_text_var.set(step_content)

        # NOW CREATE BUTTONS - AFTER ALL FUNCTIONS ARE DEFINED
        ttk.Button(controls_subframe, text="Add Step", command=add_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Remove Step", command=remove_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Load Default Template", command=load_default_template).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Up", command=move_step_up).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Down", command=move_step_down).pack(side='left', padx=5)

        ttk.Button(edit_frame, text="Update Step", command=update_selected_step).pack(side='right', padx=5)

        # Bind listbox selection
        checklist_listbox.bind('<<ListboxSelect>>', on_step_select)

        # Load default template initially
        load_default_template()

        # Save and Cancel buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Save Template", command=save_template).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)

    def load_pm_templates(self):
        """Load PM templates into the tree"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                    pt.checklist_items, pt.estimated_hours, pt.updated_date
                FROM pm_templates pt
                ORDER BY pt.bfm_equipment_no, pt.template_name
            ''')
        
            # Clear existing items
            for item in self.templates_tree.get_children():
                self.templates_tree.delete(item)
        
            # Add templates
            for template in cursor.fetchall():
                bfm_no, name, pm_type, checklist_json, est_hours, updated = template
            
                # Count checklist items
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    step_count = len(checklist_items)
                except:
                    step_count = 0
            
                self.templates_tree.insert('', 'end', values=(
                    bfm_no, name, pm_type, step_count, f"{est_hours:.1f}h", updated[:10]
                ))
            
        except Exception as e:
            print(f"Error loading PM templates: {e}")

    def filter_template_list(self, *args):
        """Filter template list based on search term"""
        search_term = self.template_search_var.get().lower()
    
        try:
            cursor = self.conn.cursor()
            if search_term:
                cursor.execute('''
                    SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                        pt.checklist_items, pt.estimated_hours, pt.updated_date
                    FROM pm_templates pt
                    WHERE LOWER(pt.bfm_equipment_no) LIKE ? 
                    OR LOWER(pt.template_name) LIKE ?
                    ORDER BY pt.bfm_equipment_no, pt.template_name
                ''', (f'%{search_term}%', f'%{search_term}%'))
            else:
                cursor.execute('''
                    SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                        pt.checklist_items, pt.estimated_hours, pt.updated_date
                    FROM pm_templates pt
                    ORDER BY pt.bfm_equipment_no, pt.template_name
                ''')
        
            # Clear and repopulate
            for item in self.templates_tree.get_children():
                self.templates_tree.delete(item)
        
            for template in cursor.fetchall():
                bfm_no, name, pm_type, checklist_json, est_hours, updated = template
            
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    step_count = len(checklist_items)
                except:
                    step_count = 0
            
                self.templates_tree.insert('', 'end', values=(
                    bfm_no, name, pm_type, step_count, f"{est_hours:.1f}h", updated[:10]
                ))
    
        except Exception as e:
            print(f"Error filtering templates: {e}")

    def edit_pm_template_dialog(self):
        """Edit existing PM template with full functionality"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to edit")
            return

        # Get selected template data
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]

        # Fetch full template data
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT id, bfm_equipment_no, template_name, pm_type, checklist_items, 
                special_instructions, safety_notes, estimated_hours
            FROM pm_templates 
            WHERE bfm_equipment_no = ? AND template_name = ?
        ''', (bfm_no, template_name))

        template_data = cursor.fetchone()
        if not template_data:
            messagebox.showerror("Error", "Template not found")
            return

        # Extract template data
        (template_id, orig_bfm_no, orig_name, orig_pm_type, orig_checklist_json, 
        orig_instructions, orig_safety, orig_hours) = template_data

        # Parse checklist items
        try:
            orig_checklist_items = json.loads(orig_checklist_json) if orig_checklist_json else []
        except:
            orig_checklist_items = []

        # Create edit dialog (similar structure to create dialog)
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit PM Template - {template_name}")
        dialog.geometry("800x750")
        dialog.transient(self.root)
        dialog.grab_set()

        # Template information (pre-populated)
        header_frame = ttk.LabelFrame(dialog, text="Template Information", padding=10)
        header_frame.pack(fill='x', padx=10, pady=5)

        # BFM Equipment (read-only)
        ttk.Label(header_frame, text="BFM Equipment Number:").grid(row=0, column=0, sticky='w', pady=5)
        bfm_var = tk.StringVar(value=orig_bfm_no)
        bfm_label = ttk.Label(header_frame, text=orig_bfm_no, font=('Arial', 10, 'bold'))
        bfm_label.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        # Template name (editable)
        ttk.Label(header_frame, text="Template Name:").grid(row=0, column=2, sticky='w', pady=5, padx=(20,5))
        template_name_var = tk.StringVar(value=orig_name)
        ttk.Entry(header_frame, textvariable=template_name_var, width=25).grid(row=0, column=3, sticky='w', padx=5, pady=5)

        # PM Type (editable)
        ttk.Label(header_frame, text="PM Type:").grid(row=1, column=0, sticky='w', pady=5)
        pm_type_var = tk.StringVar(value=orig_pm_type)
        pm_type_combo = ttk.Combobox(header_frame, textvariable=pm_type_var, 
                                values=['Monthly', 'Six Month', 'Annual'], width=22)
        pm_type_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # Estimated hours (editable)
        ttk.Label(header_frame, text="Estimated Hours:").grid(row=1, column=2, sticky='w', pady=5, padx=(20,5))
        est_hours_var = tk.StringVar(value=str(orig_hours))
        ttk.Entry(header_frame, textvariable=est_hours_var, width=10).grid(row=1, column=3, sticky='w', padx=5, pady=5)

        # Custom checklist section
        checklist_frame = ttk.LabelFrame(dialog, text="Edit PM Checklist", padding=10)
        checklist_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Checklist controls
        controls_subframe = ttk.Frame(checklist_frame)
        controls_subframe.pack(fill='x', pady=5)

        # Checklist listbox with scrollbar
        list_frame = ttk.Frame(checklist_frame)
        list_frame.pack(fill='both', expand=True, pady=5)

        checklist_listbox = tk.Listbox(list_frame, height=15, font=('Arial', 9))
        list_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=checklist_listbox.yview)
        checklist_listbox.configure(yscrollcommand=list_scrollbar.set)

        checklist_listbox.pack(side='left', fill='both', expand=True)
        list_scrollbar.pack(side='right', fill='y')

        # Step editing
        edit_frame = ttk.LabelFrame(checklist_frame, text="Edit Selected Step", padding=5)
        edit_frame.pack(fill='x', pady=5)

        step_text_var = tk.StringVar()
        step_entry = ttk.Entry(edit_frame, textvariable=step_text_var, width=80)
        step_entry.pack(side='left', fill='x', expand=True, padx=5)

        # Special instructions and safety notes (pre-populated)
        notes_frame = ttk.LabelFrame(dialog, text="Additional Information", padding=10)
        notes_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(notes_frame, text="Special Instructions:").grid(row=0, column=0, sticky='nw', pady=2)
        special_instructions_text = tk.Text(notes_frame, height=3, width=50)
        special_instructions_text.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        special_instructions_text.insert('1.0', orig_instructions or '')

        ttk.Label(notes_frame, text="Safety Notes:").grid(row=1, column=0, sticky='nw', pady=2)
        safety_notes_text = tk.Text(notes_frame, height=3, width=50)
        safety_notes_text.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        safety_notes_text.insert('1.0', orig_safety or '')

        notes_frame.grid_columnconfigure(1, weight=1)

        # Helper functions (same as create dialog)
        def add_checklist_step():
            step_text = step_text_var.get().strip()
            if step_text:
                step_num = checklist_listbox.size() + 1
                checklist_listbox.insert('end', f"{step_num}. {step_text}")
                step_text_var.set('')

        def remove_checklist_step():
            selection = checklist_listbox.curselection()
            if selection:
                checklist_listbox.delete(selection[0])
                renumber_steps()

        def renumber_steps():
            items = []
            for i in range(checklist_listbox.size()):
                step_text = checklist_listbox.get(i)
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                items.append(f"{i+1}. {step_content}")
        
            checklist_listbox.delete(0, 'end')
            for item in items:
                checklist_listbox.insert('end', item)

        def update_selected_step():
            selection = checklist_listbox.curselection()
            if selection and step_text_var.get().strip():
                step_num = selection[0] + 1
                new_text = f"{step_num}. {step_text_var.get().strip()}"
                checklist_listbox.delete(selection[0])
                checklist_listbox.insert(selection[0], new_text)

        def move_step_up():
            selection = checklist_listbox.curselection()
            if selection and selection[0] > 0:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx-1, item)
                checklist_listbox.selection_set(idx-1)
                renumber_steps()

        def move_step_down():
            selection = checklist_listbox.curselection()
            if selection and selection[0] < checklist_listbox.size()-1:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx+1, item)
                checklist_listbox.selection_set(idx+1)
                renumber_steps()

        def save_changes():
            try:
                # Validate inputs
                if not template_name_var.get().strip():
                    messagebox.showerror("Error", "Please enter template name")
                    return

                # Get updated checklist items
                checklist_items = []
                for i in range(checklist_listbox.size()):
                    step_text = checklist_listbox.get(i)
                    step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                    checklist_items.append(step_content)

                if not checklist_items:
                    messagebox.showerror("Error", "Please add at least one checklist item")
                    return

                # Update database
                cursor = self.conn.cursor()
                cursor.execute('''
                    UPDATE pm_templates SET
                    template_name = ?,
                    pm_type = ?,
                    checklist_items = ?,
                    special_instructions = ?,
                    safety_notes = ?,
                    estimated_hours = ?,
                    updated_date = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (
                    template_name_var.get().strip(),
                    pm_type_var.get(),
                    json.dumps(checklist_items),
                    special_instructions_text.get('1.0', 'end-1c'),
                    safety_notes_text.get('1.0', 'end-1c'),
                    float(est_hours_var.get() or 1.0),
                    template_id
                ))

                self.conn.commit()
                messagebox.showinfo("Success", "PM template updated successfully!")
                dialog.destroy()
                self.load_pm_templates()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to update template: {str(e)}")

        def on_step_select(event):
            selection = checklist_listbox.curselection()
            if selection:
                step_text = checklist_listbox.get(selection[0])
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                step_text_var.set(step_content)

        # Create buttons
        ttk.Button(controls_subframe, text="Add Step", command=add_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Remove Step", command=remove_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Up", command=move_step_up).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Down", command=move_step_down).pack(side='left', padx=5)

        ttk.Button(edit_frame, text="Update Step", command=update_selected_step).pack(side='right', padx=5)

        # Bind listbox selection
        checklist_listbox.bind('<<ListboxSelect>>', on_step_select)

        # Load existing checklist items
        checklist_listbox.delete(0, 'end')
        for i, item in enumerate(orig_checklist_items, 1):
            checklist_listbox.insert('end', f"{i}. {item}")

        # Save and Cancel buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Save Changes", command=save_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)

    def preview_pm_template(self):
        """Preview selected PM template"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to preview")
            return
    
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
    
        # Get template data
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT pt.*, e.description, e.sap_material_no, e.location
            FROM pm_templates pt
            LEFT JOIN equipment e ON pt.bfm_equipment_no = e.bfm_equipment_no
            WHERE pt.bfm_equipment_no = ? AND pt.template_name = ?
        ''', (bfm_no, template_name))
    
        template_data = cursor.fetchone()
        if not template_data:
            messagebox.showerror("Error", "Template not found")
            return
    
        # Create preview dialog
        preview_dialog = tk.Toplevel(self.root)
        preview_dialog.title(f"PM Template Preview - {bfm_no}")
        preview_dialog.geometry("700x600")
        preview_dialog.transient(self.root)
        preview_dialog.grab_set()
    
        # Template info
        info_frame = ttk.LabelFrame(preview_dialog, text="Template Information", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
    
        info_text = f"Equipment: {bfm_no} - {template_data[9] or 'N/A'}\n"
        info_text += f"Template: {template_data[2]}\n"
        info_text += f"PM Type: {template_data[3]}\n"
        info_text += f"Estimated Hours: {template_data[7]:.1f}h"
    
        ttk.Label(info_frame, text=info_text, font=('Arial', 10)).pack(anchor='w')
    
        # Checklist preview
        checklist_frame = ttk.LabelFrame(preview_dialog, text="PM Checklist", padding=10)
        checklist_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        checklist_text = tk.Text(checklist_frame, wrap='word', font=('Arial', 10))
        scrollbar = ttk.Scrollbar(checklist_frame, orient='vertical', command=checklist_text.yview)
        checklist_text.configure(yscrollcommand=scrollbar.set)
        
        # Format checklist content
        try:
            checklist_items = json.loads(template_data[4]) if template_data[4] else []
            content = "PM CHECKLIST:\n" + "="*50 + "\n\n"
        
            for i, item in enumerate(checklist_items, 1):
                content += f"{i:2d}. {item}\n"
        
            if template_data[5]:  # Special instructions
                content += f"\n\nSPECIAL INSTRUCTIONS:\n{template_data[5]}\n"
        
            if template_data[6]:  # Safety notes
                content += f"\n\nSAFETY NOTES:\n{template_data[6]}\n"
        
            checklist_text.insert('1.0', content)
            checklist_text.config(state='disabled')
        
        except Exception as e:
            checklist_text.insert('1.0', f"Error loading template: {str(e)}")
            checklist_text.config(state='disabled')
    
        checklist_text.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
        # Buttons
        button_frame = ttk.Frame(preview_dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)
    
        ttk.Button(button_frame, text="Close", command=preview_dialog.destroy).pack(side='right', padx=5)

    def delete_pm_template(self):
        """Delete selected PM template with enhanced confirmation"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to delete")
            return

        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
        pm_type = item['values'][2]
        steps_count = item['values'][3]

        # Enhanced confirmation dialog
        result = messagebox.askyesno("Confirm Delete", 
                                f"Delete PM template '{template_name}'?\n\n"
                                f"Equipment: {bfm_no}\n"
                                f"PM Type: {pm_type}\n"
                                f"Steps: {steps_count}\n\n"
                                f"This action cannot be undone.\n"
                                f"Any equipment using this template will revert to default PM procedures.")

        if result:
            try:
                cursor = self.conn.cursor()
                cursor.execute('''
                    DELETE FROM pm_templates 
                    WHERE bfm_equipment_no = ? AND template_name = ?
                ''', (bfm_no, template_name))

                self.conn.commit()
                messagebox.showinfo("Success", f"Template '{template_name}' deleted successfully!")
                self.load_pm_templates()
                self.update_status(f"Deleted PM template: {template_name}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete template: {str(e)}")

    def export_custom_template_pdf(self):
        """Export custom template as PDF form"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to export")
            return
    
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
    
        # Get template and equipment data
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT pt.*, e.sap_material_no, e.description, e.tool_id_drawing_no, e.location
            FROM pm_templates pt
            LEFT JOIN equipment e ON pt.bfm_equipment_no = e.bfm_equipment_no
            WHERE pt.bfm_equipment_no = ? AND pt.template_name = ?
        ''', (bfm_no, template_name))
    
        template_data = cursor.fetchone()
        if not template_data:
            messagebox.showerror("Error", "Template not found")
            return
    
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Custom_PM_Template_{bfm_no}_{template_name.replace(' ', '_')}_{timestamp}.pdf"
        
            # Create custom PDF using the template data
            self.create_custom_pm_template_pdf(filename, template_data)
        
            messagebox.showinfo("Success", f"Custom PM template exported to: {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export template: {str(e)}")

    def create_custom_pm_template_pdf(self, filename, template_data):
        """Create PDF with custom PM template"""
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib import colors
        
            doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=36, leftMargin=36,
                                topMargin=36, bottomMargin=36)
        
            styles = getSampleStyleSheet()
            story = []
        
            # Extract template data
            (template_id, bfm_no, template_name, pm_type, checklist_json,
            special_instructions, safety_notes, estimated_hours, created_date, updated_date,
            sap_no, description, tool_id, location) = template_data
        
            # Custom styles
            cell_style = ParagraphStyle(
                'CellStyle',
                parent=styles['Normal'],
                fontSize=8,
                leading=10,
                wordWrap='LTR'
            )
        
            header_cell_style = ParagraphStyle(
                'HeaderCellStyle',
                parent=styles['Normal'],
                fontSize=9,
                fontName='Helvetica-Bold',
                leading=11,
                wordWrap='LTR'
            )
        
            company_style = ParagraphStyle(
                'CompanyStyle',
                parent=styles['Heading1'],
                fontSize=14,
                fontName='Helvetica-Bold',
                alignment=1,
                textColor=colors.darkblue
            )
        
            # Header
            story.append(Paragraph("AIT - BUILDING THE FUTURE OF AEROSPACE", company_style))
            story.append(Spacer(1, 15))
        
            # Equipment information table
            equipment_data = [
                [
                    Paragraph('(SAP) Material Number:', header_cell_style), 
                    Paragraph(str(sap_no or ''), cell_style), 
                    Paragraph('Tool ID / Drawing Number:', header_cell_style), 
                    Paragraph(str(tool_id or ''), cell_style)
                ],
                [
                    Paragraph('(BFM) Equipment Number:', header_cell_style), 
                    Paragraph(str(bfm_no), cell_style), 
                    Paragraph('Description of Equipment:', header_cell_style), 
                    Paragraph(str(description or ''), cell_style)
                ],
                [
                    Paragraph('Custom Template:', header_cell_style), 
                    Paragraph(str(template_name), cell_style), 
                    Paragraph('Location of Equipment:', header_cell_style), 
                    Paragraph(str(location or ''), cell_style)
                ],
                [
                    Paragraph('Maintenance Technician:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('PM Cycle:', header_cell_style), 
                    Paragraph(str(pm_type), cell_style)
                ],
                [
                    Paragraph('Estimated Hours:', header_cell_style), 
                    Paragraph(f'{estimated_hours:.1f}h', cell_style), 
                    Paragraph('Date of Current PM:', header_cell_style), 
                    Paragraph('', cell_style)
                ]
            ]
        
            if safety_notes:
                equipment_data.append([
                    Paragraph(f'SAFETY: {safety_notes}', cell_style), 
                    '', '', ''
                ])
        
            equipment_data.append([
                Paragraph(f'Printed: {datetime.now().strftime("%m/%d/%Y")}', cell_style), 
                '', '', ''
            ])
        
            equipment_table = Table(equipment_data, colWidths=[1.8*inch, 1.7*inch, 1.8*inch, 1.7*inch])
            equipment_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('SPAN', (0, -2), (-1, -2)),  # Safety spans all columns
                ('SPAN', (0, -1), (-1, -1)),  # Printed date spans all columns
            ]))
        
            story.append(equipment_table)
            story.append(Spacer(1, 15))
        
            # Custom checklist table
            checklist_data = [
                [
                    Paragraph('', header_cell_style), 
                    Paragraph('CUSTOM PM CHECKLIST:', header_cell_style), 
                    Paragraph('', header_cell_style), 
                    Paragraph('Completed', header_cell_style), 
                    Paragraph('Labor Time', header_cell_style)
                ]
            ]
        
            # Add custom checklist items
            try:
                checklist_items = json.loads(checklist_json) if checklist_json else []
            except:
                checklist_items = []
        
            if not checklist_items:
                checklist_items = ["No custom checklist defined - using default steps"]
        
            for idx, item in enumerate(checklist_items, 1):
                checklist_data.append([
                    Paragraph(str(idx), cell_style), 
                    Paragraph(item, cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Yes', cell_style), 
                    Paragraph('hours    minutes', cell_style)
                ])
        
            checklist_table = Table(checklist_data, colWidths=[0.3*inch, 4.2*inch, 0.4*inch, 0.7*inch, 1.4*inch])
            checklist_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]))
        
            story.append(checklist_table)
            story.append(Spacer(1, 15))
        
            # Special instructions section
            if special_instructions and special_instructions.strip():
                instructions_data = [
                    [Paragraph('SPECIAL INSTRUCTIONS:', header_cell_style)],
                    [Paragraph(special_instructions, cell_style)]
                ]
            
                instructions_table = Table(instructions_data, colWidths=[7*inch])
                instructions_table.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ]))
            
                story.append(instructions_table)
                story.append(Spacer(1, 15))
        
            # Completion section
            completion_data = [
                [
                    Paragraph('Notes from Technician:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Next Annual PM Date:', header_cell_style)
                ],
                [
                    Paragraph('', cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('', cell_style)
                ],
                [
                    Paragraph('All Data Entered Into System:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Total Time', header_cell_style)
                ],
                [
                    Paragraph('Document Name', header_cell_style), 
                    Paragraph('Revision', header_cell_style), 
                    Paragraph('', cell_style)
                ],
                [
                    Paragraph(f'Custom_PM_Template_{template_name}', cell_style), 
                    Paragraph('A1', cell_style), 
                    Paragraph('', cell_style)
                ]
            ]
        
            completion_table = Table(completion_data, colWidths=[2.8*inch, 2.2*inch, 2*inch])
            completion_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
        
            story.append(completion_table)
        
            # Build PDF
            doc.build(story)
        
        except Exception as e:
            print(f"Error creating custom PM template PDF: {e}")
            raise

    # Additional methods to integrate with existing PM completion system

    def get_pm_template_for_equipment(self, bfm_no, pm_type):
        """Get custom PM template for specific equipment and PM type"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT checklist_items, special_instructions, safety_notes, estimated_hours
                FROM pm_templates 
                WHERE bfm_equipment_no = ? AND pm_type = ?
                ORDER BY updated_date DESC LIMIT 1
            ''', (bfm_no, pm_type))
        
            result = cursor.fetchone()
            if result:
                checklist_json, special_instructions, safety_notes, estimated_hours = result
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    return {
                        'checklist_items': checklist_items,
                        'special_instructions': special_instructions,
                        'safety_notes': safety_notes,
                        'estimated_hours': estimated_hours
                    }
                except:
                    return None
            return None
        
        except Exception as e:
            print(f"Error getting PM template: {e}")
            return None

    def update_pm_completion_form_with_template(self):
        """Update PM completion form when equipment is selected"""
        bfm_no = self.completion_bfm_var.get().strip()
        pm_type = self.pm_type_var.get()
    
        if bfm_no and pm_type:
            template = self.get_pm_template_for_equipment(bfm_no, pm_type)
            if template:
                # Update estimated hours
                self.labor_hours_var.set(str(int(template['estimated_hours'])))
                self.labor_minutes_var.set(str(int((template['estimated_hours'] % 1) * 60)))
            
                # Show template info
                self.update_status(f"Custom template found for {bfm_no} - {pm_type} PM")
            else:
                self.update_status(f"No custom template found for {bfm_no} - {pm_type} PM, using default")

    def create_equipment_pm_lookup_with_templates(self):
        """Enhanced equipment lookup that shows custom templates"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Equipment PM Schedule & Templates")
        dialog.geometry("900x700")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Equipment search
        search_frame = ttk.LabelFrame(dialog, text="Equipment Search", padding=15)
        search_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Label(search_frame, text="BFM Equipment Number:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
    
        bfm_var = tk.StringVar()
        bfm_entry = ttk.Entry(search_frame, textvariable=bfm_var, width=20, font=('Arial', 11))
        bfm_entry.grid(row=0, column=1, padx=10, pady=5)
    
        search_btn = ttk.Button(search_frame, text="Look Up Equipment", 
                            command=lambda: self.lookup_equipment_with_templates(bfm_var.get().strip(), results_frame))
        search_btn.grid(row=0, column=2, padx=10, pady=5)
    
        # Results frame
        results_frame = ttk.LabelFrame(dialog, text="Equipment Information & Templates", padding=10)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        bfm_entry.focus_set()
        bfm_entry.bind('<Return>', lambda e: search_btn.invoke())

    def lookup_equipment_with_templates(self, bfm_no, parent_frame):
        """Lookup equipment with custom template information"""
        if not bfm_no:
            messagebox.showwarning("Warning", "Please enter a BFM Equipment Number")
            return
    
        try:
            cursor = self.conn.cursor()
        
            # Clear previous results
            for widget in parent_frame.winfo_children():
                widget.destroy()
        
            # Get equipment info
            cursor.execute('''
                SELECT sap_material_no, description, location, status
                FROM equipment 
                WHERE bfm_equipment_no = ?
            ''', (bfm_no,))
        
            equipment_data = cursor.fetchone()
            if not equipment_data:
                error_label = ttk.Label(parent_frame, 
                                    text=f"Equipment '{bfm_no}' not found in database",
                                    font=('Arial', 12, 'bold'), foreground='red')
                error_label.pack(pady=20)
                return
        
            # Equipment header
            header_text = f"Equipment: {bfm_no}\n"
            header_text += f"Description: {equipment_data[1] or 'N/A'}\n"
            header_text += f"Location: {equipment_data[2] or 'N/A'}\n"
            header_text += f"Status: {equipment_data[3] or 'Active'}"
        
            header_label = ttk.Label(parent_frame, text=header_text, font=('Arial', 10))
            header_label.pack(pady=10)
        
            # Get custom templates
            cursor.execute('''
                SELECT template_name, pm_type, checklist_items, estimated_hours, updated_date
                FROM pm_templates 
                WHERE bfm_equipment_no = ?
                ORDER BY pm_type, template_name
            ''', (bfm_no,))
        
            templates = cursor.fetchall()
        
            if templates:
                templates_frame = ttk.LabelFrame(parent_frame, text="Custom PM Templates", padding=10)
                templates_frame.pack(fill='x', pady=10)
            
                for template in templates:
                    name, pm_type, checklist_json, est_hours, updated = template
                    try:
                        checklist_items = json.loads(checklist_json) if checklist_json else []
                        step_count = len(checklist_items)
                    except:
                        step_count = 0
                
                    template_text = f"• {name} ({pm_type} PM) - {step_count} steps, {est_hours:.1f}h estimated"
                    ttk.Label(templates_frame, text=template_text, font=('Arial', 9)).pack(anchor='w')
            else:
                no_templates_label = ttk.Label(parent_frame, 
                                            text="No custom PM templates found for this equipment",
                                            font=('Arial', 10), foreground='orange')
                no_templates_label.pack(pady=10)
        
            # Regular PM schedule info (existing functionality)
            self.lookup_equipment_pm_schedule(bfm_no, parent_frame)
        
        except Exception as e:
            error_label = ttk.Label(parent_frame, 
                                text=f"Error looking up equipment: {str(e)}", 
                                font=('Arial', 10), foreground='red')
            error_label.pack(pady=20)
    
    
    
    def update_existing_annual_pm_dates(self):
        """One-time update to spread out existing annual PM dates"""
        try:
            cursor = self.conn.cursor()
        
            # Get all PM completions with the same annual date (like 2026-08-18)
            cursor.execute('''
                SELECT id, bfm_equipment_no, next_annual_pm_date 
                FROM pm_completions 
                WHERE next_annual_pm_date = "2026-08-18"
            ''')
        
            records = cursor.fetchall()
            updated_count = 0
        
            for record_id, bfm_no, current_date in records:
                try:
                    # Apply same offset logic as the new code
                    numeric_part = re.findall(r'\d+', bfm_no)
                    if numeric_part:
                        last_digits = int(numeric_part[-1]) % 61  # 0-60
                        offset_days = last_digits - 30  # -30 to +30 days
                    else:
                        offset_days = (hash(bfm_no) % 61) - 30  # -30 to +30 days
                
                    # Calculate new date
                    base_date = datetime.strptime(current_date, '%Y-%m-%d')
                    new_date = (base_date + timedelta(days=offset_days)).strftime('%Y-%m-%d')
                
                    # Update the record
                    cursor.execute('''
                        UPDATE pm_completions 
                        SET next_annual_pm_date = ? 
                        WHERE id = ?
                    ''', (new_date, record_id))
                
                    updated_count += 1
                
                except Exception as e:
                    print(f"Error updating record {record_id}: {e}")
                    continue
        
            # Also update the equipment table next_annual_pm dates
            cursor.execute('''
                SELECT bfm_equipment_no, next_annual_pm 
                FROM equipment 
                WHERE next_annual_pm = "2026-08-18"
            ''')
        
            equipment_records = cursor.fetchall()
        
            for bfm_no, current_date in equipment_records:
                try:
                    numeric_part = re.findall(r'\d+', bfm_no)
                    if numeric_part:
                        last_digits = int(numeric_part[-1]) % 61
                        offset_days = last_digits - 30
                    else:
                        offset_days = (hash(bfm_no) % 61) - 30
                
                    base_date = datetime.strptime(current_date, '%Y-%m-%d')
                    new_date = (base_date + timedelta(days=offset_days)).strftime('%Y-%m-%d')
                
                    cursor.execute('''
                        UPDATE equipment 
                        SET next_annual_pm = ? 
                        WHERE bfm_equipment_no = ?
                    ''', (new_date, bfm_no))
                
                    updated_count += 1
                
                except Exception as e:
                    print(f"Error updating equipment {bfm_no}: {e}")
                    continue
        
            self.conn.commit()
            messagebox.showinfo("Success", f"Updated {updated_count} records with spread dates!")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update database: {str(e)}")
    
    def set_monthly_annual_pms_only(self):
        """Set all equipment to Monthly and Annual PMs only (disable Six Month PMs)"""
        result = messagebox.askyesno(
            "Confirm PM Update", 
            "This will set ALL equipment to:\n"
            "• Monthly PM: ENABLED\n"
            "• Six Month PM: DISABLED\n" 
            "• Annual PM: ENABLED\n\n"
            "Continue?"
        )
    
        if result:
            try:
                cursor = self.conn.cursor()
            
                # Get count before update
                cursor.execute('SELECT COUNT(*) FROM equipment')
                total_count = cursor.fetchone()[0]
            
                # Update all equipment
                cursor.execute('''
                    UPDATE equipment 
                    SET monthly_pm = 1, 
                        six_month_pm = 0, 
                        annual_pm = 1,
                        updated_date = CURRENT_TIMESTAMP
                ''')
            
                updated_count = cursor.rowcount
                self.conn.commit()
            
                messagebox.showinfo(
                    "Success", 
                    f"Updated {updated_count} equipment records!\n\n"
                    f"All equipment now set to:\n"
                    f"• Monthly PM: Enabled\n"
                    f"• Six Month PM: Disabled\n"
                    f"• Annual PM: Enabled"
                )
            
                # Refresh the equipment list display
                self.refresh_equipment_list()
                self.update_status(f"Updated {updated_count} equipment PM settings")
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update PM settings: {str(e)}")
    
    def standardize_all_database_dates(self):
        """Standardize all dates in the database to YYYY-MM-DD format"""
        
        # Confirmation dialog
        result = messagebox.askyesno(
            "Confirm Date Standardization",
            "This will standardize ALL dates in the database to YYYY-MM-DD format.\n\n"
            "Tables affected:\n"
            "• Equipment (PM dates)\n"
            "• PM Completions\n"
            "• Weekly Schedules\n"
            "• Corrective Maintenance\n"
            "• Cannot Find Assets\n"
            "• Run to Failure Assets\n\n"
            "This action cannot be undone. Continue?",
            icon='warning'
        )
        
        if not result:
            return
        
        try:
            # Create progress dialog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Standardizing Dates...")
            progress_dialog.geometry("400x150")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            
            ttk.Label(progress_dialog, text="Standardizing dates in database...", 
                     font=('Arial', 12)).pack(pady=20)
            
            progress_var = tk.StringVar(value="Initializing...")
            progress_label = ttk.Label(progress_dialog, textvariable=progress_var)
            progress_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_dialog, mode='indeterminate')
            progress_bar.pack(pady=10, padx=20, fill='x')
            progress_bar.start()
            
            # Update GUI
            self.root.update()
            
            # Perform standardization
            progress_var.set("Processing database...")
            self.root.update()
            
            standardizer = DateStandardizer(self.conn)
            total_updated, errors = standardizer.standardize_all_dates()
            
            progress_bar.stop()
            progress_dialog.destroy()
            
            # Show results
            if errors:
                error_msg = f"Date standardization completed with {len(errors)} errors:\n\n"
                error_msg += "\n".join(errors[:10])  # Show first 10 errors
                if len(errors) > 10:
                    error_msg += f"\n... and {len(errors) - 10} more errors"
                
                messagebox.showwarning("Standardization Complete (With Errors)", 
                                     f"Updated {total_updated} records.\n\n{error_msg}")
            else:
                messagebox.showinfo("Success", 
                                  f"Date standardization completed successfully!\n\n"
                                  f"Updated {total_updated} date records to YYYY-MM-DD format.\n\n"
                                  f"All dates are now standardized.")
            
            # Refresh displays
            self.refresh_equipment_list()
            self.load_recent_completions()
            if hasattr(self, 'load_cannot_find_assets'):
                self.load_cannot_find_assets()
            if hasattr(self, 'load_run_to_failure_assets'):
                self.load_run_to_failure_assets()
            
            self.update_status(f"Date standardization complete: {total_updated} records updated")
            
        except Exception as e:
            if 'progress_dialog' in locals():
                progress_dialog.destroy()
            messagebox.showerror("Error", f"Failed to standardize dates: {str(e)}")

    def format_date_display(self, date_str):
        """Format date for consistent display"""
        if not date_str:
            return ''
        try:
            # Parse using flexible method
            standardizer = DateStandardizer(self.conn)
            standardized = standardizer.parse_date_flexible(date_str)
            return standardized if standardized else date_str
        except:
            return date_str

    def get_current_date_standard(self):
        """Get current date in standard format"""
        return datetime.now().strftime('%Y-%m-%d')
    
    def __init__(self, root):
        self.root = root
        self.root.title("AIT Complete CMMS - Computerized Maintenance Management System")
        self.root.geometry("1800x1000")
        try:
            self.root.state('zoomed')  # Maximize window on Windows
        except:
            pass  # Skip if not on Windows
    
        # ===== ROLE-BASED ACCESS CONTROL =====
        self.current_user_role = None  # Will be set by login
        self.user_name = None
    
        # Team members as specified - MUST be defined before login dialog
        self.technicians = [
            "Mark Michaels", "Jerone Bosarge", "Jon Hymel", "Nick Whisenant", 
            "James Dunnam", "Wayne Dunnam", "Nate Williams", "Rey Marikit", "Ronald Houghs",
        ]
    
        # Show login dialog after technicians are defined
        if not self.show_login_dialog():
            self.root.destroy()
            return
    
       

        # ===== CRITICAL FIX: SET UP SHAREPOINT BACKUP FIRST =====
        self.backup_sync_dir = self.get_sharepoint_backup_path()
    
        # ===== SYNC DATABASE BEFORE INITIALIZING =====
        # This will download the latest backup if available BEFORE creating local database
        database_synced = self.sync_database_before_init()
    
        # ===== NOW Initialize database (will use synced version if available) =====
        self.init_database()
        self.init_pm_templates_database()
    
        # Set up ongoing backup system
        if self.backup_sync_dir:
            self.schedule_sharepoint_only_backups(self.backup_sync_dir)
    
    
        # ===== CRITICAL FIX: SET UP SHAREPOINT BACKUP FIRST =====
        self.backup_sync_dir = self.get_sharepoint_backup_path()

        # ===== SYNC DATABASE BEFORE INITIALIZING =====
        database_synced = self.sync_database_before_init()

        # ===== NOW Initialize database =====
        self.init_database()
        self.init_pm_templates_database()
        # Clean up old local backups (keep only one)
        self.cleanup_local_backups()


        # Set up ongoing backup system
        if hasattr(self, 'backup_sync_dir') and self.backup_sync_dir:
            self.schedule_sharepoint_only_backups(self.backup_sync_dir)

        # Light sync check (heavy lifting already done)
        self.sync_database_on_startup()
        
    
        # Add logo header
        self.add_logo_to_main_window()
    
        # PM Frequencies and cycles
        self.pm_frequencies = {
            'Monthly': 30,
            'Six Month': 180,
            'Annual': 365,
            'Run Till Failure': 0
        }
    
        # Weekly PM target
        self.weekly_pm_target = 110
    
        # Initialize data storage
        self.equipment_data = []
        self.current_week_start = self.get_week_start(datetime.now())
    
        # Create GUI based on user role
        self.create_gui()
    
        # Load initial data
        self.load_equipment_data()
        self.check_empty_database_and_offer_restore()
    
        # Add this near the end of __init__, after self.load_equipment_data()
        if self.current_user_role == 'Manager':
            self.add_database_restore_button()  
        if self.current_user_role == 'Manager':
            self.update_equipment_statistics()
    
        print(f"✅ AIT Complete CMMS System initialized successfully for {self.user_name} ({self.current_user_role})")

       
    
        # Add this at the very end of __init__:
    
        # Set up window close handler
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
        print(f"AIT Complete CMMS System initialized successfully for {self.user_name} ({self.current_user_role})")


    
    def sync_database_before_init(self):
        """Download and sync with latest database backup BEFORE initializing local database"""
        try:
            if not hasattr(self, 'backup_sync_dir') or not self.backup_sync_dir:
                print("No backup directory configured, skipping pre-init sync")
                return False
        
            backup_dir = self.backup_sync_dir
            local_db = 'ait_cmms_database.db'
    
            print("Checking for latest database backup before initialization...")
    
            # Get all backup files from SharePoint folder
            if not os.path.exists(backup_dir):
                print("SharePoint backup folder not found, skipping sync")
                return False
        
            backup_files = []
            for f in os.listdir(backup_dir):
                if f.startswith('ait_cmms_backup_') and f.endswith('.db'):
                    full_path = os.path.join(backup_dir, f)
                    backup_files.append((full_path, os.path.getmtime(full_path)))
    
            if not backup_files:
                print("No backup files found in SharePoint, will start with empty database")
                return False
        
            # Find the most recent backup
            backup_files.sort(key=lambda x: x[1], reverse=True)
            latest_backup_path, latest_backup_time = backup_files[0]
    
            # Check if local database exists and compare
            if os.path.exists(local_db):
                local_db_time = os.path.getmtime(local_db)
                local_db_size = os.path.getsize(local_db)
            
                # If SharePoint backup is newer OR if local database is empty, replace it
                if latest_backup_time > local_db_time or local_db_size < 10000:  # 10KB threshold for "empty"
                    print(f"SharePoint backup is newer or local database is empty, syncing...")
                    print(f"Latest backup: {os.path.basename(latest_backup_path)}")
                
                    # MODIFIED: Only create ONE local backup, overwrite if exists
                    if local_db_size > 10000:  # Only backup if local has meaningful content
                        local_backup_name = "ait_cmms_database_local_backup.db"
                        if os.path.exists(local_backup_name):
                            print(f"Overwriting existing local backup: {local_backup_name}")
                        else:
                            print(f"Creating local backup: {local_backup_name}")
                        shutil.copy2(local_db, local_backup_name)
                
                    # Copy SharePoint backup to local database
                    shutil.copy2(latest_backup_path, local_db)
                
                    print(f"Database synced successfully from SharePoint")
                    return True
                else:
                    print("Local database is current and has content, no sync needed")
                    return False
            else:
                # No local database exists, copy the latest backup
                print(f"No local database found, copying latest backup: {os.path.basename(latest_backup_path)}")
                shutil.copy2(latest_backup_path, local_db)
                print(f"Database initialized from SharePoint backup")
                return True
        
        except Exception as e:
            print(f"Error in pre-init database sync: {e}")
            return False
    
    
    
    def show_login_dialog(self):
        """Show login dialog to determine user role with password protection for manager"""
        # Ensure we start fresh
        login_successful = False
    
        def create_login_dialog():
            nonlocal login_successful
            
            login_dialog = tk.Toplevel(self.root)
            login_dialog.title("AIT CMMS - User Login")
            login_dialog.geometry("450x350")
            login_dialog.transient(self.root)
            login_dialog.grab_set()

            # Center the dialog
            login_dialog.update_idletasks()
            x = (login_dialog.winfo_screenwidth() // 2) - (login_dialog.winfo_width() // 2)
            y = (login_dialog.winfo_screenheight() // 2) - (login_dialog.winfo_height() // 2)
            login_dialog.geometry(f"+{x}+{y}")

            # Prevent closing the dialog with X button
            login_dialog.protocol("WM_DELETE_WINDOW", lambda: None)

            # Header
            header_frame = ttk.Frame(login_dialog)
            header_frame.pack(fill='x', padx=20, pady=20)

            ttk.Label(header_frame, text="AIT CMMS LOGIN", 
                    font=('Arial', 16, 'bold')).pack()
            ttk.Label(header_frame, text="Select your role to continue", 
                    font=('Arial', 10)).pack(pady=5)

            # User selection
            user_frame = ttk.LabelFrame(login_dialog, text="Select User", padding=15)
            user_frame.pack(fill='x', padx=20, pady=10)

            selected_user = tk.StringVar()

            # Manager option with password requirement
            manager_frame = ttk.Frame(user_frame)
            manager_frame.pack(fill='x', pady=5)

            ttk.Radiobutton(manager_frame, text="Manager (Full Access)", 
                        variable=selected_user, value="Manager").pack(side='left')
            ttk.Label(manager_frame, text="- Access to all CMMS functions (Password Required)", 
                    font=('Arial', 9), foreground='blue').pack(side='left', padx=10)

            # Password field for manager (initially hidden)
            password_frame = ttk.Frame(user_frame)
            password_frame.pack(fill='x', pady=10)

            ttk.Label(password_frame, text="Manager Password:", font=('Arial', 10, 'bold')).pack(anchor='w')
            password_var = tk.StringVar()
            password_entry = ttk.Entry(password_frame, textvariable=password_var, show="*", width=20)
            password_entry.pack(anchor='w', pady=2)

            # Initially hide password field
            password_frame.pack_forget()

            # Separator
            ttk.Separator(user_frame, orient='horizontal').pack(fill='x', pady=10)

            # Technician options
            ttk.Label(user_frame, text="Technicians (CM Access Only):", 
                    font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0,5))

            # Create technician radio buttons in two columns
            tech_frame = ttk.Frame(user_frame)
            tech_frame.pack(fill='x')

            left_column = ttk.Frame(tech_frame)
            left_column.pack(side='left', fill='both', expand=True)

            right_column = ttk.Frame(tech_frame)
            right_column.pack(side='right', fill='both', expand=True)

            for i, tech in enumerate(self.technicians):
                column = left_column if i < len(self.technicians)//2 else right_column
                ttk.Radiobutton(column, text=tech, 
                            variable=selected_user, value=tech).pack(anchor='w', pady=1)

            def on_user_selection_change(*args):
                """Show/hide password field based on selection"""
                if selected_user.get() == "Manager":
                    password_frame.pack(fill='x', pady=10, after=manager_frame)
                    password_entry.focus_set()
                else:
                    password_frame.pack_forget()

            # Bind to user selection changes
            selected_user.trace('w', on_user_selection_change)

            # Track if login is in progress to prevent double-execution
            login_in_progress = False

            def do_login():
                nonlocal login_successful, login_in_progress
            
                if login_in_progress:
                    return
            
                login_in_progress = True
            
                try:
                    user = selected_user.get()

                    if not user:
                        messagebox.showerror("Error", "Please select a user")
                        return

                    # Handle manager login with password
                    if user == "Manager":
                        entered_password = password_var.get()
                        correct_password = "AIT2584"

                        if not entered_password:
                            messagebox.showerror("Error", "Please enter the manager password")
                            password_entry.focus_set()
                            return

                        if entered_password != correct_password:
                            messagebox.showerror("Access Denied", 
                                            "Incorrect manager password.\n\n"
                                            "Access to manager functions is restricted.")
                            password_var.set("")
                            password_entry.focus_set()
                            return

                        # Password correct
                        self.current_user_role = "Manager"
                        self.user_name = "Manager"

                    else:
                        # Technician login (no password required)
                        self.current_user_role = "Technician"
                        self.user_name = user

                    login_successful = True
                    dialog.quit()
                
                finally:
                    login_in_progress = False

            def cancel_login():
                nonlocal login_successful
                login_successful = False
                dialog.quit()

            # Buttons
            button_frame = ttk.Frame(login_dialog)
            button_frame.pack(side='bottom', fill='x', padx=20, pady=20)

            login_button = ttk.Button(button_frame, text="Login", command=do_login)
            login_button.pack(side='left', padx=5)
            ttk.Button(button_frame, text="Exit", command=cancel_login).pack(side='right', padx=5)

            # Simplified event binding
            def on_enter_key(event):
                if not login_in_progress:
                    do_login()

            # Only bind to password entry
            password_entry.bind('<Return>', on_enter_key)

            return login_dialog

        # Create and run the dialog
        dialog = create_login_dialog()
        dialog.mainloop()
        dialog.destroy()
    
        return login_successful

    

    
    def get_week_start(self, date):
        """Get the start of the week (Monday) for a given date"""
        days_since_monday = date.weekday()
        return date - timedelta(days=days_since_monday)
    
    def init_pm_templates_database(self):
        """Initialize PM templates database tables"""
        cursor = self.conn.cursor()
    
        # PM Templates table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pm_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bfm_equipment_no TEXT,
                template_name TEXT,
                pm_type TEXT,
                checklist_items TEXT,  -- JSON string
                special_instructions TEXT,
                safety_notes TEXT,
                estimated_hours REAL,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
    
        # Default checklist items for fallback
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS default_pm_checklist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pm_type TEXT,
                step_number INTEGER,
                description TEXT,
                is_active BOOLEAN DEFAULT 1
            )
        ''')
    
        # Insert default checklist if empty
        cursor.execute('SELECT COUNT(*) FROM default_pm_checklist')
        if cursor.fetchone()[0] == 0:
            default_items = [
                (1, "Special Equipment Used (List):"),
                (2, "Validate your maintenance with Date / Stamp / Hours"),
                (3, "Refer to drawing when performing maintenance"),
                (4, "Make sure all instruments are properly calibrated"),
                (5, "Make sure tool is properly identified"),
                (6, "Make sure all mobile mechanisms move fluidly"),
                (7, "Visually inspect the welds"),
                (8, "Take note of any anomaly or defect (create a CM if needed)"),
                (9, "Check all screws. Tighten if needed."),
                (10, "Check the pins for wear"),
                (11, "Make sure all tooling is secured to the equipment with cable"),
                (12, "Ensure all tags (BFM and SAP) are applied and securely fastened"),
                (13, "All documentation are picked up from work area"),
                (14, "All parts and tools have been picked up"),
                (15, "Workspace has been cleaned up"),
                (16, "Dry runs have been performed (tests, restarts, etc.)"),
                (17, "Ensure that AIT Sticker is applied")
            ]
        
            for step_num, description in default_items:
                cursor.execute('''
                    INSERT INTO default_pm_checklist (pm_type, step_number, description)
                    VALUES ('All', ?, ?)
                ''', (step_num, description))
    
        self.conn.commit()

    def create_custom_pm_templates_tab(self):
        """Create PM Templates management tab"""
        self.pm_templates_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pm_templates_frame, text="PM Templates")
    
        # Controls
        controls_frame = ttk.LabelFrame(self.pm_templates_frame, text="PM Template Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Button(controls_frame, text="Create Custom Template", 
                command=self.create_custom_pm_template_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Edit Template", 
                command=self.edit_pm_template_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Preview Template", 
                command=self.preview_pm_template).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Delete Template", 
                command=self.delete_pm_template).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export Template to PDF", 
                command=self.export_custom_template_pdf).pack(side='left', padx=5)
    
        # Search frame
        search_frame = ttk.Frame(self.pm_templates_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Label(search_frame, text="Search Templates:").pack(side='left', padx=5)
        self.template_search_var = tk.StringVar()
        self.template_search_var.trace('w', self.filter_template_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.template_search_var, width=30)
        search_entry.pack(side='left', padx=5)
    
        # Templates list
        list_frame = ttk.LabelFrame(self.pm_templates_frame, text="PM Templates", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        self.templates_tree = ttk.Treeview(list_frame,
                                        columns=('BFM No', 'Template Name', 'PM Type', 'Steps', 'Est Hours', 'Updated'),
                                        show='headings')
    
        template_columns = {
            'BFM No': ('BFM Equipment No', 120),
            'Template Name': ('Template Name', 200),
            'PM Type': ('PM Type', 100),
            'Steps': ('# Steps', 80),
            'Est Hours': ('Est Hours', 80),
            'Updated': ('Last Updated', 120)
        }
    
        for col, (heading, width) in template_columns.items():
            self.templates_tree.heading(col, text=heading)
            self.templates_tree.column(col, width=width)
    
        # Scrollbars
        template_v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.templates_tree.yview)
        template_h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.templates_tree.xview)
        self.templates_tree.configure(yscrollcommand=template_v_scrollbar.set, xscrollcommand=template_h_scrollbar.set)
    
        # Pack treeview and scrollbars
        self.templates_tree.grid(row=0, column=0, sticky='nsew')
        template_v_scrollbar.grid(row=0, column=1, sticky='ns')
        template_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
    
        # Load templates
        self.load_pm_templates()

    def create_custom_pm_template_dialog(self):
        """Dialog to create custom PM template for specific equipment"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Custom PM Template")
        dialog.geometry("800x750")
        dialog.transient(self.root)
        dialog.grab_set()

        # Equipment selection
        header_frame = ttk.LabelFrame(dialog, text="Template Information", padding=10)
        header_frame.pack(fill='x', padx=10, pady=5)

        # BFM Equipment selection
        ttk.Label(header_frame, text="BFM Equipment Number:").grid(row=0, column=0, sticky='w', pady=5)
        bfm_var = tk.StringVar()
        bfm_combo = ttk.Combobox(header_frame, textvariable=bfm_var, width=25)
        bfm_combo.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        # Populate equipment list
        cursor = self.conn.cursor()
        cursor.execute('SELECT bfm_equipment_no, description FROM equipment ORDER BY bfm_equipment_no')
        equipment_list = cursor.fetchall()
        bfm_combo['values'] = [f"{bfm} - {desc[:30]}..." if len(desc) > 30 else f"{bfm} - {desc}" 
                            for bfm, desc in equipment_list]

        # Template name
        ttk.Label(header_frame, text="Template Name:").grid(row=0, column=2, sticky='w', pady=5, padx=(20,5))
        template_name_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=template_name_var, width=25).grid(row=0, column=3, sticky='w', padx=5, pady=5)

        # PM Type
        ttk.Label(header_frame, text="PM Type:").grid(row=1, column=0, sticky='w', pady=5)
        pm_type_var = tk.StringVar(value='Annual')
        pm_type_combo = ttk.Combobox(header_frame, textvariable=pm_type_var, 
                                    values=['Monthly', 'Six Month', 'Annual'], width=22)
        pm_type_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # Estimated hours
        ttk.Label(header_frame, text="Estimated Hours:").grid(row=1, column=2, sticky='w', pady=5, padx=(20,5))
        est_hours_var = tk.StringVar(value="1.0")
        ttk.Entry(header_frame, textvariable=est_hours_var, width=10).grid(row=1, column=3, sticky='w', padx=5, pady=5)

        # Custom checklist section
        checklist_frame = ttk.LabelFrame(dialog, text="Custom PM Checklist", padding=10)
        checklist_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Checklist controls
        controls_subframe = ttk.Frame(checklist_frame)
        controls_subframe.pack(fill='x', pady=5)

        # Checklist listbox with scrollbar
        list_frame = ttk.Frame(checklist_frame)
        list_frame.pack(fill='both', expand=True, pady=5)

        checklist_listbox = tk.Listbox(list_frame, height=15, font=('Arial', 9))
        list_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=checklist_listbox.yview)
        checklist_listbox.configure(yscrollcommand=list_scrollbar.set)

        checklist_listbox.pack(side='left', fill='both', expand=True)
        list_scrollbar.pack(side='right', fill='y')

        # Step editing
        edit_frame = ttk.LabelFrame(checklist_frame, text="Edit Selected Step", padding=5)
        edit_frame.pack(fill='x', pady=5)

        step_text_var = tk.StringVar()
        step_entry = ttk.Entry(edit_frame, textvariable=step_text_var, width=80)
        step_entry.pack(side='left', fill='x', expand=True, padx=5)

        # Special instructions and safety notes
        notes_frame = ttk.LabelFrame(dialog, text="Additional Information", padding=10)
        notes_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(notes_frame, text="Special Instructions:").grid(row=0, column=0, sticky='nw', pady=2)
        special_instructions_text = tk.Text(notes_frame, height=3, width=50)
        special_instructions_text.grid(row=0, column=1, sticky='ew', padx=5, pady=2)

        ttk.Label(notes_frame, text="Safety Notes:").grid(row=1, column=0, sticky='nw', pady=2)
        safety_notes_text = tk.Text(notes_frame, height=3, width=50)
        safety_notes_text.grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        notes_frame.grid_columnconfigure(1, weight=1)

        # DEFINE ALL HELPER FUNCTIONS FIRST
        def add_checklist_step():
            step_text = step_text_var.get().strip()
            if step_text:
                step_num = checklist_listbox.size() + 1
                checklist_listbox.insert('end', f"{step_num}. {step_text}")
                step_text_var.set('')

        def remove_checklist_step():
            selection = checklist_listbox.curselection()
            if selection:
                checklist_listbox.delete(selection[0])
                renumber_steps()

        def renumber_steps():
            items = []
            for i in range(checklist_listbox.size()):
                step_text = checklist_listbox.get(i)
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                items.append(f"{i+1}. {step_content}")
        
            checklist_listbox.delete(0, 'end')
            for item in items:
                checklist_listbox.insert('end', item)

        def update_selected_step():
            selection = checklist_listbox.curselection()
            if selection and step_text_var.get().strip():
                step_num = selection[0] + 1
                new_text = f"{step_num}. {step_text_var.get().strip()}"
                checklist_listbox.delete(selection[0])
                checklist_listbox.insert(selection[0], new_text)

        def move_step_up():
            selection = checklist_listbox.curselection()
            if selection and selection[0] > 0:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx-1, item)
                checklist_listbox.selection_set(idx-1)
                renumber_steps()

        def move_step_down():
            selection = checklist_listbox.curselection()
            if selection and selection[0] < checklist_listbox.size()-1:
                idx = selection[0]
                item = checklist_listbox.get(idx)
                checklist_listbox.delete(idx)
                checklist_listbox.insert(idx+1, item)
                checklist_listbox.selection_set(idx+1)
                renumber_steps()

        def load_default_template():
            cursor = self.conn.cursor()
            cursor.execute('SELECT description FROM default_pm_checklist ORDER BY step_number')
            default_steps = cursor.fetchall()
        
            checklist_listbox.delete(0, 'end')
            for i, (step,) in enumerate(default_steps, 1):
                checklist_listbox.insert('end', f"{i}. {step}")

        def save_template():
            try:
                # Validate inputs
                if not bfm_var.get():
                    messagebox.showerror("Error", "Please select equipment")
                    return
            
                if not template_name_var.get().strip():
                    messagebox.showerror("Error", "Please enter template name")
                    return
            
                # Extract BFM number from combo selection
                bfm_no = bfm_var.get().split(' - ')[0]
            
                # Get checklist items
                checklist_items = []
                for i in range(checklist_listbox.size()):
                    step_text = checklist_listbox.get(i)
                    step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                    checklist_items.append(step_content)
            
                if not checklist_items:
                    messagebox.showerror("Error", "Please add at least one checklist item")
                    return
            
                # Save to database
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO pm_templates 
                    (bfm_equipment_no, template_name, pm_type, checklist_items, 
                    special_instructions, safety_notes, estimated_hours)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    bfm_no,
                    template_name_var.get().strip(),
                    pm_type_var.get(),
                    json.dumps(checklist_items),
                    special_instructions_text.get('1.0', 'end-1c'),
                    safety_notes_text.get('1.0', 'end-1c'),
                    float(est_hours_var.get() or 1.0)
                ))
            
                self.conn.commit()
                messagebox.showinfo("Success", "Custom PM template created successfully!")
                dialog.destroy()
                self.load_pm_templates()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {str(e)}")

        def on_step_select(event):
            selection = checklist_listbox.curselection()
            if selection:
                step_text = checklist_listbox.get(selection[0])
                step_content = '. '.join(step_text.split('. ')[1:]) if '. ' in step_text else step_text
                step_text_var.set(step_content)

        # NOW CREATE ALL BUTTONS - AFTER ALL FUNCTIONS ARE DEFINED
        ttk.Button(controls_subframe, text="Add Step", command=add_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Remove Step", command=remove_checklist_step).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Load Default Template", command=load_default_template).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Up", command=move_step_up).pack(side='left', padx=5)
        ttk.Button(controls_subframe, text="Move Down", command=move_step_down).pack(side='left', padx=5)

        ttk.Button(edit_frame, text="Update Step", command=update_selected_step).pack(side='right', padx=5)

        # Bind listbox selection
        checklist_listbox.bind('<<ListboxSelect>>', on_step_select)

        # Load default template initially
        load_default_template()

        # Save and Cancel buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Save Template", command=save_template).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)

    def load_pm_templates(self):
        """Load PM templates into the tree"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                    pt.checklist_items, pt.estimated_hours, pt.updated_date
                FROM pm_templates pt
                ORDER BY pt.bfm_equipment_no, pt.template_name
            ''')
        
            # Clear existing items
            for item in self.templates_tree.get_children():
                self.templates_tree.delete(item)
        
            # Add templates
            for template in cursor.fetchall():
                bfm_no, name, pm_type, checklist_json, est_hours, updated = template
            
                # Count checklist items
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    step_count = len(checklist_items)
                except:
                    step_count = 0
            
                self.templates_tree.insert('', 'end', values=(
                    bfm_no, name, pm_type, step_count, f"{est_hours:.1f}h", updated[:10]
                ))
            
        except Exception as e:
            print(f"Error loading PM templates: {e}")

    def filter_template_list(self, *args):
        """Filter template list based on search term"""
        search_term = self.template_search_var.get().lower()
    
        try:
            cursor = self.conn.cursor()
            if search_term:
                cursor.execute('''
                    SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                        pt.checklist_items, pt.estimated_hours, pt.updated_date
                    FROM pm_templates pt
                    WHERE LOWER(pt.bfm_equipment_no) LIKE ? 
                    OR LOWER(pt.template_name) LIKE ?
                    ORDER BY pt.bfm_equipment_no, pt.template_name
                ''', (f'%{search_term}%', f'%{search_term}%'))
            else:
                cursor.execute('''
                    SELECT pt.bfm_equipment_no, pt.template_name, pt.pm_type, 
                        pt.checklist_items, pt.estimated_hours, pt.updated_date
                    FROM pm_templates pt
                    ORDER BY pt.bfm_equipment_no, pt.template_name
                ''')
        
            # Clear and repopulate
            for item in self.templates_tree.get_children():
                self.templates_tree.delete(item)
        
            for template in cursor.fetchall():
                bfm_no, name, pm_type, checklist_json, est_hours, updated = template
            
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    step_count = len(checklist_items)
                except:
                    step_count = 0
            
                self.templates_tree.insert('', 'end', values=(
                    bfm_no, name, pm_type, step_count, f"{est_hours:.1f}h", updated[:10]
                ))
    
        except Exception as e:
            print(f"Error filtering templates: {e}")

    

    def preview_pm_template(self):
        """Preview selected PM template"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to preview")
            return
    
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
    
        # Get template data
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT pt.*, e.description, e.sap_material_no, e.location
            FROM pm_templates pt
            LEFT JOIN equipment e ON pt.bfm_equipment_no = e.bfm_equipment_no
            WHERE pt.bfm_equipment_no = ? AND pt.template_name = ?
        ''', (bfm_no, template_name))
    
        template_data = cursor.fetchone()
        if not template_data:
            messagebox.showerror("Error", "Template not found")
            return
    
        # Create preview dialog
        preview_dialog = tk.Toplevel(self.root)
        preview_dialog.title(f"PM Template Preview - {bfm_no}")
        preview_dialog.geometry("700x600")
        preview_dialog.transient(self.root)
        preview_dialog.grab_set()
    
        # Template info
        info_frame = ttk.LabelFrame(preview_dialog, text="Template Information", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
    
        info_text = f"Equipment: {bfm_no} - {template_data[9] or 'N/A'}\n"
        info_text += f"Template: {template_data[2]}\n"
        info_text += f"PM Type: {template_data[3]}\n"
        info_text += f"Estimated Hours: {template_data[7]:.1f}h"
    
        ttk.Label(info_frame, text=info_text, font=('Arial', 10)).pack(anchor='w')
    
        # Checklist preview
        checklist_frame = ttk.LabelFrame(preview_dialog, text="PM Checklist", padding=10)
        checklist_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        checklist_text = tk.Text(checklist_frame, wrap='word', font=('Arial', 10))
        scrollbar = ttk.Scrollbar(checklist_frame, orient='vertical', command=checklist_text.yview)
        checklist_text.configure(yscrollcommand=scrollbar.set)
    
        # Format checklist content
        try:
            checklist_items = json.loads(template_data[4]) if template_data[4] else []
            content = "PM CHECKLIST:\n" + "="*50 + "\n\n"
        
            for i, item in enumerate(checklist_items, 1):
                content += f"{i:2d}. {item}\n"
        
            if template_data[5]:  # Special instructions
                content += f"\n\nSPECIAL INSTRUCTIONS:\n{template_data[5]}\n"
        
            if template_data[6]:  # Safety notes
                content += f"\n\nSAFETY NOTES:\n{template_data[6]}\n"
        
            checklist_text.insert('1.0', content)
            checklist_text.config(state='disabled')
        
        except Exception as e:
            checklist_text.insert('1.0', f"Error loading template: {str(e)}")
            checklist_text.config(state='disabled')
    
        checklist_text.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
        # Buttons
        button_frame = ttk.Frame(preview_dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)
    
        ttk.Button(button_frame, text="Close", command=preview_dialog.destroy).pack(side='right', padx=5)

    def delete_pm_template(self):
        """Delete selected PM template"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to delete")
            return
    
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
    
        result = messagebox.askyesno("Confirm Delete", 
                                f"Delete PM template '{template_name}' for {bfm_no}?\n\n"
                                f"This action cannot be undone.")
    
        if result:
            try:
                cursor = self.conn.cursor()
                cursor.execute('''
                    DELETE FROM pm_templates 
                    WHERE bfm_equipment_no = ? AND template_name = ?
                ''', (bfm_no, template_name))
            
                self.conn.commit()
                messagebox.showinfo("Success", "Template deleted successfully!")
                self.load_pm_templates()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete template: {str(e)}")

    def export_custom_template_pdf(self):
        """Export custom template as PDF form"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a template to export")
            return
    
        item = self.templates_tree.item(selected[0])
        bfm_no = item['values'][0]
        template_name = item['values'][1]
    
        # Get template and equipment data
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT pt.*, e.sap_material_no, e.description, e.tool_id_drawing_no, e.location
            FROM pm_templates pt
            LEFT JOIN equipment e ON pt.bfm_equipment_no = e.bfm_equipment_no
            WHERE pt.bfm_equipment_no = ? AND pt.template_name = ?
        ''', (bfm_no, template_name))
    
        template_data = cursor.fetchone()
        if not template_data:
            messagebox.showerror("Error", "Template not found")
            return
    
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Custom_PM_Template_{bfm_no}_{template_name.replace(' ', '_')}_{timestamp}.pdf"
        
            # Create custom PDF using the template data
            self.create_custom_pm_template_pdf(filename, template_data)
        
            messagebox.showinfo("Success", f"Custom PM template exported to: {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export template: {str(e)}")

    def create_custom_pm_template_pdf(self, filename, template_data):
        """Create PDF with custom PM template"""
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib import colors
        
            doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=36, leftMargin=36,
                                topMargin=36, bottomMargin=36)
        
            styles = getSampleStyleSheet()
            story = []
        
            # Extract template data
            (template_id, bfm_no, template_name, pm_type, checklist_json,
            special_instructions, safety_notes, estimated_hours, created_date, updated_date,
            sap_no, description, tool_id, location) = template_data
        
            # Custom styles
            cell_style = ParagraphStyle(
                'CellStyle',
                parent=styles['Normal'],
                fontSize=8,
                leading=10,
                wordWrap='LTR'
            )
        
            header_cell_style = ParagraphStyle(
                'HeaderCellStyle',
                parent=styles['Normal'],
                fontSize=9,
                fontName='Helvetica-Bold',
                leading=11,
                wordWrap='LTR'
            )
        
            company_style = ParagraphStyle(
                'CompanyStyle',
                parent=styles['Heading1'],
                fontSize=14,
                fontName='Helvetica-Bold',
                alignment=1,
                textColor=colors.darkblue
            )
        
            # Header
            story.append(Paragraph("AIT - BUILDING THE FUTURE OF AEROSPACE", company_style))
            story.append(Spacer(1, 15))
        
            # Equipment information table
            equipment_data = [
                [
                    Paragraph('(SAP) Material Number:', header_cell_style), 
                    Paragraph(str(sap_no or ''), cell_style), 
                    Paragraph('Tool ID / Drawing Number:', header_cell_style), 
                    Paragraph(str(tool_id or ''), cell_style)
                ],
                [
                    Paragraph('(BFM) Equipment Number:', header_cell_style), 
                    Paragraph(str(bfm_no), cell_style), 
                    Paragraph('Description of Equipment:', header_cell_style), 
                    Paragraph(str(description or ''), cell_style)
                ],
                [
                    Paragraph('Custom Template:', header_cell_style), 
                    Paragraph(str(template_name), cell_style), 
                    Paragraph('Location of Equipment:', header_cell_style), 
                    Paragraph(str(location or ''), cell_style)
                ],
                [
                    Paragraph('Maintenance Technician:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('PM Cycle:', header_cell_style), 
                    Paragraph(str(pm_type), cell_style)
                ],
                [
                    Paragraph('Estimated Hours:', header_cell_style), 
                    Paragraph(f'{estimated_hours:.1f}h', cell_style), 
                    Paragraph('Date of Current PM:', header_cell_style), 
                    Paragraph('', cell_style)
                ]
            ]
        
            if safety_notes:
                equipment_data.append([
                    Paragraph(f'SAFETY: {safety_notes}', cell_style), 
                    '', '', ''
                ])
        
            equipment_data.append([
                Paragraph(f'Printed: {datetime.now().strftime("%m/%d/%Y")}', cell_style), 
                '', '', ''
            ])
        
            equipment_table = Table(equipment_data, colWidths=[1.8*inch, 1.7*inch, 1.8*inch, 1.7*inch])
            equipment_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('SPAN', (0, -2), (-1, -2)),  # Safety spans all columns
                ('SPAN', (0, -1), (-1, -1)),  # Printed date spans all columns
            ]))
        
            story.append(equipment_table)
            story.append(Spacer(1, 15))
        
            # Custom checklist table
            checklist_data = [
                [
                    Paragraph('', header_cell_style), 
                    Paragraph('CUSTOM PM CHECKLIST:', header_cell_style), 
                    Paragraph('', header_cell_style), 
                    Paragraph('Completed', header_cell_style), 
                    Paragraph('Labor Time', header_cell_style)
                ]
            ]
        
            # Add custom checklist items
            try:
                checklist_items = json.loads(checklist_json) if checklist_json else []
            except:
                checklist_items = []
        
            if not checklist_items:
                checklist_items = ["No custom checklist defined - using default steps"]
        
            for idx, item in enumerate(checklist_items, 1):
                checklist_data.append([
                    Paragraph(str(idx), cell_style), 
                    Paragraph(item, cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Yes', cell_style), 
                    Paragraph('hours    minutes', cell_style)
                ])
        
            checklist_table = Table(checklist_data, colWidths=[0.3*inch, 4.2*inch, 0.4*inch, 0.7*inch, 1.4*inch])
            checklist_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]))
        
            story.append(checklist_table)
            story.append(Spacer(1, 15))
        
            # Special instructions section
            if special_instructions and special_instructions.strip():
                instructions_data = [
                    [Paragraph('SPECIAL INSTRUCTIONS:', header_cell_style)],
                    [Paragraph(special_instructions, cell_style)]
                ]
            
                instructions_table = Table(instructions_data, colWidths=[7*inch])
                instructions_table.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ]))
            
                story.append(instructions_table)
                story.append(Spacer(1, 15))
        
            # Completion section
            completion_data = [
                [
                    Paragraph('Notes from Technician:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Next Annual PM Date:', header_cell_style)
                ],
                [
                    Paragraph('', cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('', cell_style)
                ],
                [
                    Paragraph('All Data Entered Into System:', header_cell_style), 
                    Paragraph('', cell_style), 
                    Paragraph('Total Time', header_cell_style)
                ],
                [
                    Paragraph('Document Name', header_cell_style), 
                    Paragraph('Revision', header_cell_style), 
                    Paragraph('', cell_style)
                ],
                [
                    Paragraph(f'Custom_PM_Template_{template_name}', cell_style), 
                    Paragraph('A1', cell_style), 
                    Paragraph('', cell_style)
                ]
            ]
        
            completion_table = Table(completion_data, colWidths=[2.8*inch, 2.2*inch, 2*inch])
            completion_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
        
            story.append(completion_table)
        
            # Build PDF
            doc.build(story)
        
        except Exception as e:
            print(f"Error creating custom PM template PDF: {e}")
            raise

    # Additional methods to integrate with existing PM completion system

    def get_pm_template_for_equipment(self, bfm_no, pm_type):
        """Get custom PM template for specific equipment and PM type"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT checklist_items, special_instructions, safety_notes, estimated_hours
                FROM pm_templates 
                WHERE bfm_equipment_no = ? AND pm_type = ?
                ORDER BY updated_date DESC LIMIT 1
            ''', (bfm_no, pm_type))
        
            result = cursor.fetchone()
            if result:
                checklist_json, special_instructions, safety_notes, estimated_hours = result
                try:
                    checklist_items = json.loads(checklist_json) if checklist_json else []
                    return {
                        'checklist_items': checklist_items,
                        'special_instructions': special_instructions,
                        'safety_notes': safety_notes,
                        'estimated_hours': estimated_hours
                    }
                except:
                    return None
            return None
        
        except Exception as e:
            print(f"Error getting PM template: {e}")
            return None

    def update_pm_completion_form_with_template(self):
        """Update PM completion form when equipment is selected"""
        bfm_no = self.completion_bfm_var.get().strip()
        pm_type = self.pm_type_var.get()
    
        if bfm_no and pm_type:
            template = self.get_pm_template_for_equipment(bfm_no, pm_type)
            if template:
                # Update estimated hours
                self.labor_hours_var.set(str(int(template['estimated_hours'])))
                self.labor_minutes_var.set(str(int((template['estimated_hours'] % 1) * 60)))
            
                # Show template info
                self.update_status(f"Custom template found for {bfm_no} - {pm_type} PM")
            else:
                self.update_status(f"No custom template found for {bfm_no} - {pm_type} PM, using default")

    def create_equipment_pm_lookup_with_templates(self):
        """Enhanced equipment lookup that shows custom templates"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Equipment PM Schedule & Templates")
        dialog.geometry("900x700")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Equipment search
        search_frame = ttk.LabelFrame(dialog, text="Equipment Search", padding=15)
        search_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Label(search_frame, text="BFM Equipment Number:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
    
        bfm_var = tk.StringVar()
        bfm_entry = ttk.Entry(search_frame, textvariable=bfm_var, width=20, font=('Arial', 11))
        bfm_entry.grid(row=0, column=1, padx=10, pady=5)
    
        search_btn = ttk.Button(search_frame, text="Look Up Equipment", 
                            command=lambda: self.lookup_equipment_with_templates(bfm_var.get().strip(), results_frame))
        search_btn.grid(row=0, column=2, padx=10, pady=5)
    
        # Results frame
        results_frame = ttk.LabelFrame(dialog, text="Equipment Information & Templates", padding=10)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        bfm_entry.focus_set()
        bfm_entry.bind('<Return>', lambda e: search_btn.invoke())

    def lookup_equipment_with_templates(self, bfm_no, parent_frame):
        """Lookup equipment with custom template information"""
        if not bfm_no:
            messagebox.showwarning("Warning", "Please enter a BFM Equipment Number")
            return
    
        try:
            cursor = self.conn.cursor()
        
            # Clear previous results
            for widget in parent_frame.winfo_children():
                widget.destroy()
        
            # Get equipment info
            cursor.execute('''
                SELECT sap_material_no, description, location, status
                FROM equipment 
                WHERE bfm_equipment_no = ?
            ''', (bfm_no,))
        
            equipment_data = cursor.fetchone()
            if not equipment_data:
                error_label = ttk.Label(parent_frame, 
                                    text=f"Equipment '{bfm_no}' not found in database",
                                    font=('Arial', 12, 'bold'), foreground='red')
                error_label.pack(pady=20)
                return
        
            # Equipment header
            header_text = f"Equipment: {bfm_no}\n"
            header_text += f"Description: {equipment_data[1] or 'N/A'}\n"
            header_text += f"Location: {equipment_data[2] or 'N/A'}\n"
            header_text += f"Status: {equipment_data[3] or 'Active'}"
        
            header_label = ttk.Label(parent_frame, text=header_text, font=('Arial', 10))
            header_label.pack(pady=10)
        
            # Get custom templates
            cursor.execute('''
                SELECT template_name, pm_type, checklist_items, estimated_hours, updated_date
                FROM pm_templates 
                WHERE bfm_equipment_no = ?
                ORDER BY pm_type, template_name
            ''', (bfm_no,))
        
            templates = cursor.fetchall()
        
            if templates:
                templates_frame = ttk.LabelFrame(parent_frame, text="Custom PM Templates", padding=10)
                templates_frame.pack(fill='x', pady=10)
            
                for template in templates:
                    name, pm_type, checklist_json, est_hours, updated = template
                    try:
                        checklist_items = json.loads(checklist_json) if checklist_json else []
                        step_count = len(checklist_items)
                    except:
                        step_count = 0
                
                    template_text = f"• {name} ({pm_type} PM) - {step_count} steps, {est_hours:.1f}h estimated"
                    ttk.Label(templates_frame, text=template_text, font=('Arial', 9)).pack(anchor='w')
            else:
                no_templates_label = ttk.Label(parent_frame, 
                                            text="No custom PM templates found for this equipment",
                                            font=('Arial', 10), foreground='orange')
                no_templates_label.pack(pady=10)
        
            # Regular PM schedule info (existing functionality)
            self.lookup_equipment_pm_schedule(bfm_no, parent_frame)
        
        except Exception as e:
            error_label = ttk.Label(parent_frame, 
                                text=f"Error looking up equipment: {str(e)}", 
                                font=('Arial', 10), foreground='red')
            error_label.pack(pady=20) 
    
    
    
    def init_database(self):
        """Initialize comprehensive CMMS database"""
        self.conn = sqlite3.connect('ait_cmms_database.db')
        cursor = self.conn.cursor()
        
        # Equipment/Assets table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sap_material_no TEXT,
                bfm_equipment_no TEXT UNIQUE,
                description TEXT,
                tool_id_drawing_no TEXT,
                location TEXT,
                master_lin TEXT,
                monthly_pm BOOLEAN DEFAULT 0,
                six_month_pm BOOLEAN DEFAULT 0,
                annual_pm BOOLEAN DEFAULT 0,
                last_monthly_pm TEXT,
                last_six_month_pm TEXT,
                last_annual_pm TEXT,
                next_monthly_pm TEXT,
                next_six_month_pm TEXT,
                next_annual_pm TEXT,
                status TEXT DEFAULT 'Active',
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_date TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # PM Completions table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pm_completions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bfm_equipment_no TEXT,
                pm_type TEXT,
                technician_name TEXT,
                completion_date TEXT,
                labor_hours REAL,
                labor_minutes REAL,
                pm_due_date TEXT,
                special_equipment TEXT,
                notes TEXT,
                next_annual_pm_date TEXT,
                document_name TEXT DEFAULT 'Preventive_Maintenance_Form',
                document_revision TEXT DEFAULT 'A2',
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
        
        # Weekly PM Schedules
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS weekly_pm_schedules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_start_date TEXT,
                bfm_equipment_no TEXT,
                pm_type TEXT,
                assigned_technician TEXT,
                status TEXT DEFAULT 'Scheduled',
                scheduled_date TEXT,
                completion_date TEXT,
                labor_hours REAL,
                notes TEXT,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
        
        # Corrective Maintenance (CM) table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS corrective_maintenance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cm_number TEXT UNIQUE,
                bfm_equipment_no TEXT,
                description TEXT,
                priority TEXT,
                assigned_technician TEXT,
                status TEXT DEFAULT 'Open',
                created_date TEXT,
                completion_date TEXT,
                labor_hours REAL,
                notes TEXT,
                root_cause TEXT,
                corrective_action TEXT,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
        
        # Weekly Reports table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS weekly_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_start_date TEXT,
                total_scheduled INTEGER,
                total_completed INTEGER,
                completion_rate REAL,
                technician_performance TEXT,
                report_data TEXT,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cannot_find_assets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bfm_equipment_no TEXT,
                description TEXT,
                location TEXT,
                technician_name TEXT,
                report_date TEXT,
                notes TEXT,
                status TEXT DEFAULT 'Missing',
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS run_to_failure_assets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bfm_equipment_no TEXT,
                description TEXT,
                location TEXT,
                technician_name TEXT,
                completion_date TEXT,
                labor_hours REAL,
                notes TEXT,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (bfm_equipment_no) REFERENCES equipment (bfm_equipment_no)
            )
        ''')
        
        self.conn.commit()
    
    def create_gui(self):
        """Create the main GUI interface based on user role"""
        # Create style
        style = ttk.Style()
        style.theme_use('clam')
    
        # Main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
    
        # Create tabs based on role
        if self.current_user_role == 'Manager':
            # Manager gets all tabs
            self.create_all_manager_tabs()
        else:
            # Technicians only get CM tab
            self.create_technician_tabs()
    
        # Status bar with user info
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side='bottom', fill='x')
    
        self.status_bar = ttk.Label(status_frame, text=f"AIT CMMS Ready - Logged in as: {self.user_name} ({self.current_user_role})", 
                                    relief='sunken')
        self.status_bar.pack(side='left', fill='x', expand=True)
    
        # Role switching button (only for development/testing)
        if self.current_user_role == 'Manager':
            ttk.Button(status_frame, text="Switch to Technician View", 
                    command=self.switch_to_technician_view).pack(side='right', padx=5)

    def create_all_manager_tabs(self):
        """Create all tabs for manager access"""
        self.create_equipment_tab()
        self.create_pm_scheduling_tab()
        self.create_pm_completion_tab()
        self.create_cm_management_tab()
        self.create_analytics_dashboard_tab()
        self.create_cannot_find_tab()
        self.create_run_to_failure_tab()
        self.create_pm_history_search_tab()
        self.create_custom_pm_templates_tab()

    def create_technician_tabs(self):
        """Create limited tabs for technician access"""
        # Only create CM Management tab for technicians
        self.create_cm_management_tab()
    
        # Add a simple info tab explaining their access
        self.create_technician_info_tab()

    def create_technician_info_tab(self):
        """Create an info tab for technicians"""
        info_frame = ttk.Frame(self.notebook)
        self.notebook.add(info_frame, text="System Info")
    
        # Welcome message
        welcome_frame = ttk.LabelFrame(info_frame, text="Welcome to AIT CMMS", padding=20)
        welcome_frame.pack(fill='both', expand=True, padx=20, pady=20)
    
        welcome_text = f"""
    Welcome, {self.user_name}!

    You are logged in as a Technician with access to:
    • Complete Corrective Maintenance (CM) System
    - View ALL team CMs (everyone's entries)
    - Create new CMs
    - Edit existing CMs  
    - Complete CMs
    - View CM history and status

    Team Collaboration:
    • You can see CMs created by all technicians
    • View work assigned to other team members
    • Complete CMs assigned to you or help with others
    • Full visibility of maintenance activities

    For additional system access or questions, please contact your manager.

    System Information:
    • User: {self.user_name}
    • Role: {self.current_user_role}
    • Login Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

    Quick Tips:
    • Use the CM tab to view all corrective maintenance
    • Create new CMs when you discover issues
    • Enter accurate dates when creating CMs
    • Provide detailed descriptions for better tracking
    • Update CM status when work is completed
    • Coordinate with team members through CM system
    """
    
        info_label = ttk.Label(welcome_frame, text=welcome_text, 
                            font=('Arial', 11), justify='left')
        info_label.pack(anchor='w')
    
        # Quick access buttons
        buttons_frame = ttk.Frame(welcome_frame)
        buttons_frame.pack(fill='x', pady=20)
    
        ttk.Button(buttons_frame, text="Create New CM", 
                command=self.create_cm_dialog).pack(side='left', padx=10)
        ttk.Button(buttons_frame, text="View My Assigned CMs", 
                command=self.show_my_cms).pack(side='left', padx=10)
        ttk.Button(buttons_frame, text="Refresh All CMs", 
                command=self.load_corrective_maintenance).pack(side='left', padx=10)

    def show_my_cms(self):
        """Show CMs assigned to current technician"""
        if self.current_user_role != 'Technician':
            return
        
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT cm_number, bfm_equipment_no, description, priority, status, created_date
                FROM corrective_maintenance 
                WHERE assigned_technician = ?
                ORDER BY created_date DESC
            ''', (self.user_name,))
        
            my_cms = cursor.fetchall()
        
            # Create dialog to show results
            dialog = tk.Toplevel(self.root)
            dialog.title(f"My CMs - {self.user_name}")
            dialog.geometry("800x400")
            dialog.transient(self.root)
            dialog.grab_set()
        
            if my_cms:
                # Create tree to display CMs
                tree = ttk.Treeview(dialog, columns=('CM#', 'Equipment', 'Description', 'Priority', 'Status', 'Date'), 
                                show='headings')
            
                for col in ('CM#', 'Equipment', 'Description', 'Priority', 'Status', 'Date'):
                    tree.heading(col, text=col)
                    tree.column(col, width=120)
            
                for cm in my_cms:
                    cm_number, bfm_no, description, priority, status, created_date = cm
                    display_desc = (description[:30] + '...') if len(description) > 30 else description
                    tree.insert('', 'end', values=(cm_number, bfm_no, display_desc, priority, status, created_date))
            
                tree.pack(fill='both', expand=True, padx=10, pady=10)
            else:
                ttk.Label(dialog, text=f"No CMs assigned to {self.user_name}", 
                        font=('Arial', 12)).pack(pady=50)
        
            ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load your CMs: {str(e)}")

    def switch_to_technician_view(self):
        """Switch to technician view for testing (Manager only)"""
        if self.current_user_role != 'Manager':
            return
        
        result = messagebox.askyesno("Switch View", 
                                    "Switch to Technician view?\n\n"
                                    "This will hide all manager functions and only show CM access.\n"
                                    "You'll need to restart the application to get back to Manager view.")
    
        if result:
            # Temporarily switch role
            self.current_user_role = 'Technician'
            self.user_name = 'Test Technician'
        
            # Recreate GUI
            for widget in self.notebook.winfo_children():
                widget.destroy()
        
            self.create_technician_tabs()
            self.status_bar.config(text=f"AIT CMMS - Logged in as: {self.user_name} ({self.current_user_role})")

    def restrict_access(self, function_name):
        """Decorator to restrict access to manager-only functions"""
        def decorator(func):
            def wrapper(*args, **kwargs):
                if self.current_user_role != 'Manager':
                    messagebox.showerror("Access Denied", 
                                    f"Access to {function_name} is restricted to Managers only.\n\n"
                                    f"Current user: {self.user_name} ({self.current_user_role})")
                    return
                return func(*args, **kwargs)
            return wrapper
        return decorator
   
    
    def standardize_all_database_dates(self):
        """Standardize all dates in the database to YYYY-MM-DD format"""
        
        # Confirmation dialog
        result = messagebox.askyesno(
            "Confirm Date Standardization",
            "This will standardize ALL dates in the database to YYYY-MM-DD format.\n\n"
            "Tables affected:\n"
            "• Equipment (PM dates)\n"
            "• PM Completions\n"
            "• Weekly Schedules\n"
            "• Corrective Maintenance\n"
            "• Cannot Find Assets\n"
            "• Run to Failure Assets\n\n"
            "This action cannot be undone. Continue?",
            icon='warning'
        )
        
        if not result:
            return
        
        try:
            # Create progress dialog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Standardizing Dates...")
            progress_dialog.geometry("400x150")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            
            ttk.Label(progress_dialog, text="Standardizing dates in database...", 
                     font=('Arial', 12)).pack(pady=20)
            
            progress_var = tk.StringVar(value="Initializing...")
            progress_label = ttk.Label(progress_dialog, textvariable=progress_var)
            progress_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_dialog, mode='indeterminate')
            progress_bar.pack(pady=10, padx=20, fill='x')
            progress_bar.start()
            
            # Update GUI
            self.root.update()
            
            # Perform standardization
            progress_var.set("Processing database...")
            self.root.update()
            
            standardizer = DateStandardizer(self.conn)
            total_updated, errors = standardizer.standardize_all_dates()
            
            progress_bar.stop()
            progress_dialog.destroy()
            
            # Show results
            if errors:
                error_msg = f"Date standardization completed with {len(errors)} errors:\n\n"
                error_msg += "\n".join(errors[:10])  # Show first 10 errors
                if len(errors) > 10:
                    error_msg += f"\n... and {len(errors) - 10} more errors"
                
                messagebox.showwarning("Standardization Complete (With Errors)", 
                                     f"Updated {total_updated} records.\n\n{error_msg}")
            else:
                messagebox.showinfo("Success", 
                                  f"Date standardization completed successfully!\n\n"
                                  f"Updated {total_updated} date records to YYYY-MM-DD format.\n\n"
                                  f"All dates are now standardized.")
            
            # Refresh displays
            self.refresh_equipment_list()
            self.load_recent_completions()
            if hasattr(self, 'load_cannot_find_assets'):
                self.load_cannot_find_assets()
            if hasattr(self, 'load_run_to_failure_assets'):
                self.load_run_to_failure_assets()
            
            self.update_status(f"Date standardization complete: {total_updated} records updated")
            
        except Exception as e:
            if 'progress_dialog' in locals():
                progress_dialog.destroy()
            messagebox.showerror("Error", f"Failed to standardize dates: {str(e)}")
    
    def add_date_standardization_button(self):
        """Add date standardization button to equipment tab"""
        # Find the controls frame in equipment tab
        for widget in self.equipment_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and "Equipment Controls" in widget['text']:
                ttk.Button(widget, text="🔄 Standardize All Dates (YYYY-MM-DD)", 
                          command=self.standardize_all_database_dates,
                          width=30).pack(side='left', padx=5)
                break
    
    

    def create_equipment_tab(self):
        """Equipment management and data import tab"""
        self.equipment_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.equipment_frame, text="Equipment Management")
        
        # Controls frame
        controls_frame = ttk.LabelFrame(self.equipment_frame, text="Equipment Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        # Add statistics frame after controls_frame
        stats_frame = ttk.LabelFrame(self.equipment_frame, text="Equipment Statistics", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=5)

        # Statistics labels
        self.stats_total_label = ttk.Label(stats_frame, text="Total Assets: 0", font=('Arial', 10, 'bold'))
        self.stats_total_label.pack(side='left', padx=20)

        self.stats_cf_label = ttk.Label(stats_frame, text="Cannot Find: 0", font=('Arial', 10, 'bold'), foreground='red')
        self.stats_cf_label.pack(side='left', padx=20)

        self.stats_rtf_label = ttk.Label(stats_frame, text="Run to Failure: 0", font=('Arial', 10, 'bold'), foreground='orange')
        self.stats_rtf_label.pack(side='left', padx=20)

        self.stats_active_label = ttk.Label(stats_frame, text="Active Assets: 0", font=('Arial', 10, 'bold'), foreground='green')
        self.stats_active_label.pack(side='left', padx=20)

        # Refresh stats button
        ttk.Button(stats_frame, text="Refresh Stats", 
                command=self.update_equipment_statistics).pack(side='right', padx=5)
        
        
        ttk.Button(controls_frame, text="Import Equipment CSV", 
                  command=self.import_equipment_csv).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Add Equipment", 
                  command=self.add_equipment_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Edit Equipment", 
                  command=self.edit_equipment_dialog).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Refresh List", 
                  command=self.refresh_equipment_list).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export Equipment", 
                  command=self.export_equipment_list).pack(side='left', padx=5)
        
        
        
        # Search frame
        search_frame = ttk.Frame(self.equipment_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(search_frame, text="Search Equipment:").pack(side='left', padx=5)
        self.equipment_search_var = tk.StringVar()
        self.equipment_search_var.trace('w', self.filter_equipment_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.equipment_search_var, width=30)
        search_entry.pack(side='left', padx=5)
        
        # Equipment list
        list_frame = ttk.Frame(self.equipment_frame)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Treeview with scrollbars
        self.equipment_tree = ttk.Treeview(list_frame, 
                                         columns=('SAP', 'BFM', 'Description', 'Location', 'LIN', 'Monthly', 'Six Month', 'Annual', 'Status'),
                                         show='headings', height=20)
        
        # Configure columns
        columns_config = {
            'SAP': ('SAP Material No.', 120),
            'BFM': ('BFM Equipment No.', 130),
            'Description': ('Description', 300),
            'Location': ('Location', 100),
            'LIN': ('Master LIN', 80),
            'Monthly': ('Monthly PM', 80),
            'Six Month': ('6-Month PM', 80),
            'Annual': ('Annual PM', 80),
            'Status': ('Status', 80)
        }
        
        for col, (heading, width) in columns_config.items():
            self.equipment_tree.heading(col, text=heading)
            self.equipment_tree.column(col, width=width)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.equipment_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.equipment_tree.xview)
        self.equipment_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.equipment_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
    
    
    
    
    
    
    
    def populate_week_selector(self):
        """Populate dropdown with weeks that have schedules"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT DISTINCT week_start_date 
                FROM weekly_pm_schedules 
                ORDER BY week_start_date DESC
            ''')
            available_weeks = [row[0] for row in cursor.fetchall()]
        
            # Always include current week as an option
            current_week = self.current_week_start.strftime('%Y-%m-%d')
            if current_week not in available_weeks:
                available_weeks.append(current_week)
                available_weeks.sort(reverse=True)
            
            # Update combobox values
            self.week_combo['values'] = available_weeks
        
            # Set to most recent week with data, or current week if no data
            if available_weeks:
                self.week_start_var.set(available_weeks[0])
            
        except Exception as e:
            print(f"Error populating week selector: {e}")

    def load_latest_weekly_schedule(self):
        """Load the most recent weekly schedule on startup"""
        try:
            cursor = self.conn.cursor()
        
            # Find the most recent week with scheduled PMs
            cursor.execute('''
                SELECT week_start_date 
                FROM weekly_pm_schedules 
                ORDER BY week_start_date DESC 
                LIMIT 1
            ''')
        
            latest_week = cursor.fetchone()
        
            if latest_week:
                self.week_start_var.set(latest_week[0])
                self.refresh_technician_schedules()
                self.update_status(f"Loaded latest weekly schedule: {latest_week[0]}")
            else:
                self.update_status("No weekly schedules found")
            
        except Exception as e:
            print(f"Error loading latest weekly schedule: {e}")
    
    
    
    def update_equipment_statistics(self):
        """Update equipment statistics display"""
        try:
            cursor = self.conn.cursor()
        
            # Get total equipment count
            cursor.execute('SELECT COUNT(*) FROM equipment')
            total_assets = cursor.fetchone()[0]
        
            # Get active equipment count
            cursor.execute("SELECT COUNT(*) FROM equipment WHERE status = 'Active' OR status IS NULL")
            active_assets = cursor.fetchone()[0]
        
            # Get Cannot Find count (current missing assets)
            cursor.execute("SELECT COUNT(DISTINCT bfm_equipment_no) FROM cannot_find_assets WHERE status = 'Missing'")
            cannot_find_count = cursor.fetchone()[0]
        
            # Get Run to Failure count
            cursor.execute("SELECT COUNT(*) FROM equipment WHERE status = 'Run to Failure'")
            rtf_count = cursor.fetchone()[0]
        
            # Also check run_to_failure_assets table for additional count
            cursor.execute("SELECT COUNT(DISTINCT bfm_equipment_no) FROM run_to_failure_assets")
            rtf_assets_count = cursor.fetchone()[0]
        
            # Use the higher count for RTF
            rtf_total = max(rtf_count, rtf_assets_count)
        
            # Update labels
            self.stats_total_label.config(text=f"Total Assets: {total_assets}")
            self.stats_active_label.config(text=f"Active Assets: {active_assets}")
            self.stats_cf_label.config(text=f"Cannot Find: {cannot_find_count}")
            self.stats_rtf_label.config(text=f"Run to Failure: {rtf_total}")
        
            # Update status bar
            self.update_status(f"Equipment stats updated - Total: {total_assets}, Active: {active_assets}, CF: {cannot_find_count}, RTF: {rtf_total}")
        
        except Exception as e:
            print(f"Error updating equipment statistics: {e}")
            messagebox.showerror("Error", f"Failed to update equipment statistics: {str(e)}")
    
    def create_pm_scheduling_tab(self):
        """PM Scheduling and assignment tab"""
        self.pm_schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pm_schedule_frame, text="PM Scheduling")
        
        # Controls
        controls_frame = ttk.LabelFrame(self.pm_schedule_frame, text="PM Scheduling Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        
       
        # Week selection with dropdown of available weeks
        ttk.Label(controls_frame, text="Week Starting:").grid(row=0, column=0, sticky='w', padx=5)
        self.week_start_var = tk.StringVar(value=self.current_week_start.strftime('%Y-%m-%d'))

        # Create combobox instead of entry
        self.week_combo = ttk.Combobox(controls_frame, textvariable=self.week_start_var, width=12)
        self.week_combo.grid(row=0, column=1, padx=5)

        # Bind selection change to refresh display
        self.week_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_technician_schedules())

        # Populate with available weeks
        self.populate_week_selector()
        
        ttk.Button(controls_frame, text="Generate Weekly Schedule", 
                  command=self.generate_weekly_assignments).grid(row=0, column=2, padx=5)
        ttk.Button(controls_frame, text="Print PM Forms", 
                  command=self.print_weekly_pm_forms).grid(row=0, column=3, padx=5)
        ttk.Button(controls_frame, text="Export Schedule", 
                  command=self.export_weekly_schedule).grid(row=0, column=4, padx=5)
        # Add this line after your existing buttons in the controls_frame section:
        ttk.Button(controls_frame, text="🔍 Validate Before Scheduling", 
                  command=self.validate_weekly_schedule_before_generation).grid(row=0, column=5, padx=5)
        
        # Schedule display
        schedule_frame = ttk.LabelFrame(self.pm_schedule_frame, text="Weekly PM Schedule", padding=10)
        schedule_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Technician tabs
        self.technician_notebook = ttk.Notebook(schedule_frame)
        self.technician_notebook.pack(fill='both', expand=True)
        
        # Create tabs for each technician
        self.technician_trees = {}
        for tech in self.technicians:
            tech_frame = ttk.Frame(self.technician_notebook)
            self.technician_notebook.add(tech_frame, text=tech)
            
            # Technician's schedule tree
            tech_tree = ttk.Treeview(tech_frame,
                                   columns=('BFM', 'Description', 'PM Type', 'Due Date', 'Status'),
                                   show='headings')
            
            tech_tree.heading('BFM', text='BFM Equipment No.')
            tech_tree.heading('Description', text='Description')
            tech_tree.heading('PM Type', text='PM Type')
            tech_tree.heading('Due Date', text='Due Date')
            tech_tree.heading('Status', text='Status')
            
            for col in ('BFM', 'Description', 'PM Type', 'Due Date', 'Status'):
                tech_tree.column(col, width=150)
            
            tech_tree.pack(fill='both', expand=True, padx=5, pady=5)
            self.technician_trees[tech] = tech_tree
            
            # After creating all the technician trees, add this line:
            self.load_latest_weekly_schedule()
    
    def create_pm_completion_tab(self):
        """PM Completion entry tab"""
        self.pm_completion_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pm_completion_frame, text="PM Completion")
        
        # Completion form
        form_frame = ttk.LabelFrame(self.pm_completion_frame, text="PM Completion Entry", padding=15)
        form_frame.pack(fill='x', padx=10, pady=5)
        
        # Form fields (matching your PM form layout)
        row = 0
        
        # Equipment selection
        ttk.Label(form_frame, text="BFM Equipment Number:").grid(row=row, column=0, sticky='w', pady=5)
        self.completion_bfm_var = tk.StringVar()
        bfm_combo = ttk.Combobox(form_frame, textvariable=self.completion_bfm_var, width=20)
        bfm_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        bfm_combo.bind('<KeyRelease>', self.update_equipment_suggestions)
        self.bfm_combo = bfm_combo
        row += 1
        
        # PM Type
        ttk.Label(form_frame, text="PM Type:").grid(row=row, column=0, sticky='w', pady=5)
        self.pm_type_var = tk.StringVar()
        pm_type_combo = ttk.Combobox(form_frame, textvariable=self.pm_type_var, 
                                   values=['Monthly', 'Six Month', 'Annual', 'CANNOT FIND', 'Run to Failure'], width=20)
        # Bind PM type and equipment changes to template lookup
        pm_type_combo.bind('<<ComboboxSelected>>', lambda e: self.update_pm_completion_form_with_template())
        self.bfm_combo.bind('<KeyRelease>', lambda e: self.update_pm_completion_form_with_template())
        pm_type_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Technician
        ttk.Label(form_frame, text="Maintenance Technician:").grid(row=row, column=0, sticky='w', pady=5)
        self.completion_tech_var = tk.StringVar()
        tech_combo = ttk.Combobox(form_frame, textvariable=self.completion_tech_var, 
                                values=self.technicians, width=20)
        tech_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Labor time
        ttk.Label(form_frame, text="Labor Time:").grid(row=row, column=0, sticky='w', pady=5)
        time_frame = ttk.Frame(form_frame)
        time_frame.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        
        self.labor_hours_var = tk.StringVar(value="0")
        ttk.Entry(time_frame, textvariable=self.labor_hours_var, width=5).pack(side='left')
        ttk.Label(time_frame, text="hours").pack(side='left', padx=5)
        
        self.labor_minutes_var = tk.StringVar(value="0")
        ttk.Entry(time_frame, textvariable=self.labor_minutes_var, width=5).pack(side='left')
        ttk.Label(time_frame, text="minutes").pack(side='left', padx=5)
        row += 1
        
        # PM Due Date
        ttk.Label(form_frame, text="PM Due Date:").grid(row=row, column=0, sticky='w', pady=5)
        self.pm_due_date_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.pm_due_date_var, width=20).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Special Equipment
        ttk.Label(form_frame, text="Special Equipment Used:").grid(row=row, column=0, sticky='w', pady=5)
        self.special_equipment_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.special_equipment_var, width=40).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Notes
        ttk.Label(form_frame, text="Notes from Technician:").grid(row=row, column=0, sticky='nw', pady=5)
        self.notes_text = tk.Text(form_frame, width=50, height=4)
        self.notes_text.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Next Annual PM Date
        ttk.Label(form_frame, text="Next Annual PM Date:").grid(row=row, column=0, sticky='w', pady=5)
        self.next_annual_pm_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.next_annual_pm_var, width=20).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Submit and refresh buttons
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
        
        ttk.Button(buttons_frame, text="Show Equipment PM History", 
                command=lambda: self.show_equipment_pm_history_dialog()).pack(side='left', padx=5)
        
        ttk.Button(buttons_frame, text="Submit PM Completion", 
                command=self.submit_pm_completion).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Refresh List", 
                command=self.load_recent_completions).pack(side='left', padx=5)
        # Add this after the existing buttons in the PM completion tab
        ttk.Button(buttons_frame, text="View Monthly Completions", 
                command=self.view_monthly_completions).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="📅 Check Equipment Schedule", 
                command=self.create_pm_schedule_lookup_dialog).pack(side='left', padx=5)
        
        # Recent completions
        recent_frame = ttk.LabelFrame(self.pm_completion_frame, text="Recent PM Completions", padding=10)
        recent_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.recent_completions_tree = ttk.Treeview(recent_frame,
                                                  columns=('Date', 'BFM', 'PM Type', 'Technician', 'Hours'),
                                                  show='headings')
        
        for col in ('Date', 'BFM', 'PM Type', 'Technician', 'Hours'):
            self.recent_completions_tree.heading(col, text=col)
            self.recent_completions_tree.column(col, width=120)
        
        self.recent_completions_tree.pack(fill='both', expand=True)
        
        # Load recent completions
        self.load_recent_completions()
        
        
    def show_equipment_pm_history_dialog(self):
        """Dialog to look up PM history for any equipment"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Equipment PM History Lookup")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
    
        ttk.Label(dialog, text="Enter BFM Equipment Number:", font=('Arial', 12)).pack(pady=20)
    
        bfm_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=bfm_var, width=20, font=('Arial', 12))
        entry.pack(pady=10)
    
        def lookup_history():
            bfm_no = bfm_var.get().strip()
            if bfm_no:
                dialog.destroy()
                self.show_recent_completions_for_equipment(bfm_no)
            else:
                messagebox.showwarning("Warning", "Please enter a BFM Equipment Number")
    
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
    
        ttk.Button(button_frame, text="Show History", command=lookup_history).pack(side='left', padx=10)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='left', padx=10)
    
        entry.focus_set()
        entry.bind('<Return>', lambda e: lookup_history())
        
        
    def create_pm_schedule_lookup_dialog(self):
        """Create dialog to lookup PM schedule for specific equipment"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Equipment PM Schedule Lookup")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Search section
        search_frame = ttk.LabelFrame(dialog, text="Equipment Search", padding=15)
        search_frame.pack(fill='x', padx=10, pady=5)
    
        # Equipment search
        ttk.Label(search_frame, text="BFM Equipment Number:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
    
        bfm_var = tk.StringVar()
        bfm_entry = ttk.Entry(search_frame, textvariable=bfm_var, width=20, font=('Arial', 11))
        bfm_entry.grid(row=0, column=1, padx=10, pady=5)
    
        # Search button
        search_btn = ttk.Button(search_frame, text="Look Up Schedule", 
                            command=lambda: self.lookup_equipment_pm_schedule(bfm_var.get().strip(), results_frame))
        search_btn.grid(row=0, column=2, padx=10, pady=5)
    
        # Auto-complete functionality
        bfm_entry.bind('<KeyRelease>', lambda e: self.update_equipment_autocomplete(bfm_var, bfm_entry))
    
        # Results display frame
        results_frame = ttk.LabelFrame(dialog, text="PM Schedule Results", padding=10)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        # Instructions
        instructions = ttk.Label(results_frame, 
                            text="Enter a BFM Equipment Number above and click 'Look Up Schedule'\nto see current PM status and next scheduled dates.",
                            font=('Arial', 10), foreground='gray')
        instructions.pack(pady=50)
    
        # Focus on entry field
        bfm_entry.focus_set()
        bfm_entry.bind('<Return>', lambda e: search_btn.invoke())

    def update_equipment_autocomplete(self, bfm_var, entry_widget):
        """Provide autocomplete suggestions for equipment numbers"""
        search_term = bfm_var.get()
        if len(search_term) >= 2:
            try:
                cursor = self.conn.cursor()
                cursor.execute('''
                    SELECT bfm_equipment_no FROM equipment 
                    WHERE LOWER(bfm_equipment_no) LIKE LOWER(?)
                    ORDER BY bfm_equipment_no LIMIT 10
                ''', (f'%{search_term}%',))
            
                suggestions = [row[0] for row in cursor.fetchall()]
            
                # Simple autocomplete - you could enhance this with a dropdown
                if len(suggestions) == 1 and suggestions[0].lower().startswith(search_term.lower()):
                    current_pos = entry_widget.index(tk.INSERT)
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, suggestions[0])
                    entry_widget.icursor(current_pos)
                    entry_widget.select_range(current_pos, tk.END)
                
            except Exception as e:
                print(f"Autocomplete error: {e}")

    def lookup_equipment_pm_schedule(self, bfm_no, parent_frame):
        """Lookup and display PM schedule for specific equipment"""
        if not bfm_no:
            messagebox.showwarning("Warning", "Please enter a BFM Equipment Number")
            return
    
        try:
            cursor = self.conn.cursor()
        
            # Clear previous results
            for widget in parent_frame.winfo_children():
                widget.destroy()
        
            # Get equipment information
            cursor.execute('''
                SELECT sap_material_no, description, location, master_lin, status,
                    monthly_pm, six_month_pm, annual_pm,
                    last_monthly_pm, last_six_month_pm, last_annual_pm,
                    next_monthly_pm, next_six_month_pm, next_annual_pm,
                    updated_date
                FROM equipment 
                WHERE bfm_equipment_no = ?
            ''', (bfm_no,))
        
            equipment_data = cursor.fetchone()
        
            if not equipment_data:
                # Equipment not found
                error_label = ttk.Label(parent_frame, 
                                    text=f"Equipment '{bfm_no}' not found in database",
                                    font=('Arial', 12, 'bold'), foreground='red')
                error_label.pack(pady=20)
                return
        
            # Unpack equipment data
            (sap_no, description, location, master_lin, status,
            monthly_pm, six_month_pm, annual_pm,
            last_monthly, last_six_month, last_annual,
            next_monthly, next_six_month, next_annual,
            updated_date) = equipment_data
        
            # Create scrollable frame for results
            canvas = tk.Canvas(parent_frame)
            scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
        
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
        
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
        
            # Equipment header information
            header_frame = ttk.LabelFrame(scrollable_frame, text="Equipment Information", padding=15)
            header_frame.pack(fill='x', padx=5, pady=5)
        
            # Equipment details in a grid
            ttk.Label(header_frame, text="BFM Equipment No:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=2)
            ttk.Label(header_frame, text=bfm_no, font=('Arial', 10)).grid(row=0, column=1, sticky='w', padx=15, pady=2)
        
            ttk.Label(header_frame, text="SAP Material No:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=2)
            ttk.Label(header_frame, text=sap_no or 'N/A', font=('Arial', 10)).grid(row=0, column=3, sticky='w', padx=15, pady=2)
        
            ttk.Label(header_frame, text="Description:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=2)
            desc_text = (description[:50] + '...') if description and len(description) > 50 else (description or 'N/A')
            ttk.Label(header_frame, text=desc_text, font=('Arial', 10)).grid(row=1, column=1, columnspan=3, sticky='w', padx=15, pady=2)
        
            ttk.Label(header_frame, text="Location:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=2)
            ttk.Label(header_frame, text=location or 'N/A', font=('Arial', 10)).grid(row=2, column=1, sticky='w', padx=15, pady=2)
        
            ttk.Label(header_frame, text="Status:", font=('Arial', 10, 'bold')).grid(row=2, column=2, sticky='w', padx=5, pady=2)
            status_color = 'green' if status == 'Active' else 'red' if status == 'Missing' else 'orange'
            status_label = ttk.Label(header_frame, text=status or 'Active', font=('Arial', 10, 'bold'), foreground=status_color)
            status_label.grid(row=2, column=3, sticky='w', padx=15, pady=2)
        
            # PM Schedule Status
            schedule_frame = ttk.LabelFrame(scrollable_frame, text="PM Schedule Status", padding=15)
            schedule_frame.pack(fill='x', padx=5, pady=5)
        
            # Create PM schedule table
            pm_data = [
                ['PM Type', 'Required', 'Last Completed', 'Next Due', 'Status', 'Days Until Due']
            ]
        
            current_date = datetime.now()
        
            # Monthly PM
            if monthly_pm:
                last_date_str = last_monthly or 'Never'
                next_date_str = next_monthly or 'Not Scheduled'
            
                status_text, days_until = self.calculate_pm_status(last_monthly, next_monthly, 30, current_date)
                pm_data.append(['Monthly', 'Yes', last_date_str, next_date_str, status_text, str(days_until) if days_until is not None else 'N/A'])
            else:
                pm_data.append(['Monthly', 'No', 'N/A', 'N/A', 'Disabled', 'N/A'])
        
            # Six Month PM
            if six_month_pm:
                last_date_str = last_six_month or 'Never'
                next_date_str = next_six_month or 'Not Scheduled'
            
                status_text, days_until = self.calculate_pm_status(last_six_month, next_six_month, 180, current_date)
                pm_data.append(['Six Month', 'Yes', last_date_str, next_date_str, status_text, str(days_until) if days_until is not None else 'N/A'])
            else:
                pm_data.append(['Six Month', 'No', 'N/A', 'N/A', 'Disabled', 'N/A'])
        
            # Annual PM
            if annual_pm:
                last_date_str = last_annual or 'Never'
                next_date_str = next_annual or 'Not Scheduled'
            
                status_text, days_until = self.calculate_pm_status(last_annual, next_annual, 365, current_date)
                pm_data.append(['Annual', 'Yes', last_date_str, next_date_str, status_text, str(days_until) if days_until is not None else 'N/A'])
            else:
                pm_data.append(['Annual', 'No', 'N/A', 'N/A', 'Disabled', 'N/A'])
        
            # Create table display
            for i, row_data in enumerate(pm_data):
                row_frame = ttk.Frame(schedule_frame)
                row_frame.pack(fill='x', pady=1)
            
                for j, cell_data in enumerate(row_data):
                    if i == 0:  # Header row
                        label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10, 'bold'), 
                                        relief='raised', padding=5, width=15)
                    else:  # Data rows
                        # Color code the status column
                        if j == 4:  # Status column
                            if 'Overdue' in cell_data:
                                color = 'red'
                            elif 'Due Soon' in cell_data:
                                color = 'orange'  
                            elif 'Current' in cell_data:
                                color = 'green'
                            else:
                                color = 'gray'
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10, 'bold'), 
                                            foreground=color, padding=3, width=15)
                        else:
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10), 
                                            padding=3, width=15)
                    
                    label.pack(side='left', padx=2)
        
            # Recent PM History
            history_frame = ttk.LabelFrame(scrollable_frame, text="Recent PM History (Last 10)", padding=15)
            history_frame.pack(fill='x', padx=5, pady=5)
        
            cursor.execute('''
                SELECT pm_type, technician_name, completion_date, 
                    (labor_hours + labor_minutes/60.0) as total_hours,
                    SUBSTR(notes, 1, 50) as notes_preview
                FROM pm_completions 
                WHERE bfm_equipment_no = ?
                ORDER BY completion_date DESC LIMIT 10
            ''', (bfm_no,))
        
            recent_completions = cursor.fetchall()
        
            if recent_completions:
                # History table
                history_data = [['Date', 'PM Type', 'Technician', 'Hours', 'Notes']]
            
                for completion in recent_completions:
                    pm_type, technician, comp_date, hours, notes_preview = completion
                    hours_str = f"{hours:.1f}h" if hours else '0h'
                    notes_str = (notes_preview + '...') if notes_preview and len(notes_preview) >= 50 else (notes_preview or '')
                    history_data.append([comp_date, pm_type, technician, hours_str, notes_str])
            
                for i, row_data in enumerate(history_data):
                    row_frame = ttk.Frame(history_frame)
                    row_frame.pack(fill='x', pady=1)
                
                    for j, cell_data in enumerate(row_data):
                        if i == 0:  # Header row
                            width = [10, 10, 15, 8, 25][j]  # Different widths for each column
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10, 'bold'), 
                                        relief='raised', padding=5, width=width)
                        else:
                            width = [10, 10, 15, 8, 25][j]
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 9), 
                                            padding=3, width=width)
                    
                        label.pack(side='left', padx=2)
            else:
                no_history_label = ttk.Label(history_frame, text="No PM completions found for this equipment", 
                                        font=('Arial', 10), foreground='gray')
                no_history_label.pack(pady=10)
        
            # Upcoming schedule (if any)
            upcoming_frame = ttk.LabelFrame(scrollable_frame, text="Upcoming Weekly Schedules", padding=15)
            upcoming_frame.pack(fill='x', padx=5, pady=5)
        
            cursor.execute('''
                SELECT pm_type, assigned_technician, scheduled_date, week_start_date, status
                FROM weekly_pm_schedules 
                WHERE bfm_equipment_no = ? AND scheduled_date >= DATE('now')
                ORDER BY scheduled_date ASC LIMIT 5
            ''', (bfm_no,))
        
            upcoming_schedules = cursor.fetchall()
        
            if upcoming_schedules:
                upcoming_data = [['PM Type', 'Assigned To', 'Scheduled Date', 'Week Start', 'Status']]
            
                for schedule in upcoming_schedules:
                    pm_type, technician, sched_date, week_start, sched_status = schedule
                    upcoming_data.append([pm_type, technician, sched_date, week_start, sched_status])
            
                for i, row_data in enumerate(upcoming_data):
                    row_frame = ttk.Frame(upcoming_frame)
                    row_frame.pack(fill='x', pady=1)
                
                    for j, cell_data in enumerate(row_data):
                        if i == 0:  # Header row
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10, 'bold'), 
                                            relief='raised', padding=5, width=12)
                        else:
                            label = ttk.Label(row_frame, text=cell_data, font=('Arial', 10), 
                                            padding=3, width=12)
                    
                        label.pack(side='left', padx=2)
            else:
                no_upcoming_label = ttk.Label(upcoming_frame, text="No upcoming scheduled PMs found", 
                                            font=('Arial', 10), foreground='gray')
                no_upcoming_label.pack(pady=10)
        
            # Pack the canvas and scrollbar
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        
            # Update scroll region
            scrollable_frame.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        except Exception as e:
            error_label = ttk.Label(parent_frame, 
                                text=f"Error looking up equipment: {str(e)}", 
                                font=('Arial', 10), foreground='red')
            error_label.pack(pady=20)
            print(f"PM Schedule lookup error: {e}")

    def calculate_pm_status(self, last_pm_date, next_pm_date, frequency_days, current_date):
        """Calculate PM status and days until due"""
        try:
            if not last_pm_date and not next_pm_date:
                return "Never Done", None
        
            # Use next_pm_date if available, otherwise calculate from last_pm_date
            if next_pm_date:
                next_due = datetime.strptime(next_pm_date, '%Y-%m-%d')
            elif last_pm_date:
                last_date = datetime.strptime(last_pm_date, '%Y-%m-%d')
                next_due = last_date + timedelta(days=frequency_days)
            else:
                return "Not Scheduled", None
        
            days_until = (next_due - current_date).days
        
            if days_until < 0:
                return f"Overdue ({abs(days_until)} days)", days_until
            elif days_until <= 7:
                return f"Due Soon ({days_until} days)", days_until
            elif days_until <= 30:
                return f"Due in {days_until} days", days_until
            else:
                return f"Current ({days_until} days)", days_until
            
        except ValueError:
            return "Date Error", None
        except Exception as e:
            return "Error", None
        
        
    def view_monthly_completions(self):
        """Open dialog to view completed PMs for a specific month/year"""
        dialog = tk.Toplevel(self.root)
        dialog.title("View Monthly PM Completions")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Month/Year selection frame
        selection_frame = ttk.LabelFrame(dialog, text="Select Month and Year", padding=10)
        selection_frame.pack(fill='x', padx=10, pady=5)
    
        # Month selection
        ttk.Label(selection_frame, text="Month:").grid(row=0, column=0, sticky='w', padx=5)
        month_var = tk.StringVar()
        month_combo = ttk.Combobox(selection_frame, textvariable=month_var, width=12, state='readonly')
        month_combo['values'] = [
            '01 - January', '02 - February', '03 - March', '04 - April',
            '05 - May', '06 - June', '07 - July', '08 - August', 
            '09 - September', '10 - October', '11 - November', '12 - December'
        ]
        month_combo.grid(row=0, column=1, padx=5)
        month_combo.set(f"{datetime.now().month:02d} - {calendar.month_name[datetime.now().month]}")
    
        # Year selection
        ttk.Label(selection_frame, text="Year:").grid(row=0, column=2, sticky='w', padx=5)
        year_var = tk.StringVar(value=str(datetime.now().year))
        year_entry = ttk.Entry(selection_frame, textvariable=year_var, width=8)
        year_entry.grid(row=0, column=3, padx=5)
    
        # Load button
        ttk.Button(selection_frame, text="Load Completions", 
                command=lambda: self.load_monthly_data(month_var, year_var, monthly_tree, summary_text)).grid(row=0, column=4, padx=10)
    
        # Export button
        ttk.Button(selection_frame, text="Export to CSV", 
                command=lambda: self.export_monthly_data(month_var, year_var)).grid(row=0, column=5, padx=5)
    
        # Summary frame
        summary_frame = ttk.LabelFrame(dialog, text="Monthly Summary", padding=10)
        summary_frame.pack(fill='x', padx=10, pady=5)
    
        summary_text = tk.Text(summary_frame, height=6, wrap='word', font=('Courier', 9))
        summary_scrollbar = ttk.Scrollbar(summary_frame, orient='vertical', command=summary_text.yview)
        summary_text.configure(yscrollcommand=summary_scrollbar.set)
    
        summary_text.pack(side='left', fill='both', expand=True)
        summary_scrollbar.pack(side='right', fill='y')
    
        # Completions list frame
        list_frame = ttk.LabelFrame(dialog, text="PM Completions", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        # Treeview for completions
        monthly_tree = ttk.Treeview(list_frame,
                                columns=('Date', 'BFM No', 'Description', 'PM Type', 'Technician', 'Hours', 'Notes'),
                                show='headings')
    
        # Configure columns
        monthly_columns = {
            'Date': ('Completion Date', 100),
            'BFM No': ('BFM Equipment No', 120),
            'Description': ('Equipment Description', 200),
            'PM Type': ('PM Type', 100),
            'Technician': ('Completed By', 120),
            'Hours': ('Labor Hours', 80),
            'Notes': ('Notes Preview', 150)
        }
    
        for col, (heading, width) in monthly_columns.items():
            monthly_tree.heading(col, text=heading)
            monthly_tree.column(col, width=width)
    
        # Scrollbars
        monthly_v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=monthly_tree.yview)
        monthly_h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=monthly_tree.xview)
        monthly_tree.configure(yscrollcommand=monthly_v_scrollbar.set, xscrollcommand=monthly_h_scrollbar.set)
    
        # Pack treeview and scrollbars
        monthly_tree.grid(row=0, column=0, sticky='nsew')
        monthly_v_scrollbar.grid(row=0, column=1, sticky='ns')
        monthly_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
    
        # Load current month by default
        self.load_monthly_data(month_var, year_var, monthly_tree, summary_text)

    
    def load_monthly_data(self, month_var, year_var, tree, summary_text):
        """Load PM completion data for selected month/year with debugging"""
        try:
            # Parse month and year
            month_text = month_var.get()
            month_num = month_text.split(' - ')[0] if month_text else f"{datetime.now().month:02d}"
            year = year_var.get() or str(datetime.now().year)
    
            # Calculate date range for the month
            start_date = f"{year}-{month_num}-01"
    
            # Get last day of month
            year_int = int(year)
            month_int = int(month_num)
            if month_int == 12:
                next_month = 1
                next_year = year_int + 1
            else:
                next_month = month_int + 1
                next_year = year_int
    
            end_date = (datetime(next_year, next_month, 1) - timedelta(days=1)).strftime('%Y-%m-%d')
    
            cursor = self.conn.cursor()
        
            # DEBUG: Print date range
            print(f"DEBUG: Searching date range: {start_date} to {end_date}")
    
            # Get PM completions for the month
            cursor.execute('''
                SELECT 
                    pc.completion_date,
                    pc.bfm_equipment_no,
                    e.description,
                    pc.pm_type,
                    pc.technician_name,
                    (pc.labor_hours + pc.labor_minutes/60.0) as total_hours,
                    pc.notes
                FROM pm_completions pc
                LEFT JOIN equipment e ON pc.bfm_equipment_no = e.bfm_equipment_no
                WHERE pc.completion_date BETWEEN ? AND ?
                ORDER BY pc.completion_date DESC, pc.bfm_equipment_no
            ''', (start_date, end_date))
    
            completions = cursor.fetchall()
            print(f"DEBUG: PM completions found: {len(completions)}")
    
            # Get Cannot Find entries for the month
            cursor.execute('''
                SELECT 
                    cf.report_date,
                    cf.bfm_equipment_no,
                    cf.description,
                    'CANNOT FIND' as pm_type,
                    cf.technician_name,
                    0 as total_hours,
                    cf.notes
                FROM cannot_find_assets cf
                WHERE cf.report_date BETWEEN ? AND ?
                ORDER BY cf.report_date DESC, cf.bfm_equipment_no
            ''', (start_date, end_date))
    
            cannot_finds = cursor.fetchall()
            print(f"DEBUG: Cannot find entries found: {len(cannot_finds)}")
        
            # DEBUG: Check for any dates outside expected range
            cursor.execute('''
                SELECT completion_date, COUNT(*) 
                FROM pm_completions 
                WHERE strftime('%Y-%m', completion_date) = ? 
                GROUP BY completion_date 
                ORDER BY completion_date
            ''', (f"{year}-{month_num}",))
        
            date_counts = cursor.fetchall()
            print(f"DEBUG: All PM completion dates this month: {date_counts}")
        
            cursor.execute('''
                SELECT report_date, COUNT(*) 
                FROM cannot_find_assets 
                WHERE strftime('%Y-%m', report_date) = ? 
                GROUP BY report_date 
                ORDER BY report_date
            ''', (f"{year}-{month_num}",))
        
            cf_date_counts = cursor.fetchall()
            print(f"DEBUG: All cannot find dates this month: {cf_date_counts}")
        
            # Try alternative query to see if we get different results
            cursor.execute('''
                SELECT COUNT(*) FROM pm_completions 
                WHERE strftime('%Y-%m', completion_date) = ?
            ''', (f"{year}-{month_num}",))
            alt_pm_count = cursor.fetchone()[0]
        
            cursor.execute('''
                SELECT COUNT(*) FROM cannot_find_assets 
                WHERE strftime('%Y-%m', report_date) = ?
            ''', (f"{year}-{month_num}",))
            alt_cf_count = cursor.fetchone()[0]
        
            print(f"DEBUG: Alternative count - PM: {alt_pm_count}, CF: {alt_cf_count}, Total: {alt_pm_count + alt_cf_count}")
            
            # Add this debug query after the other debug queries:
            cursor.execute('SELECT COUNT(*) FROM pm_completions')
            total_pm_all = cursor.fetchone()[0]

            cursor.execute('SELECT COUNT(*) FROM cannot_find_assets') 
            total_cf_all = cursor.fetchone()[0]

            print(f"DEBUG: Total records in entire database - PM: {total_pm_all}, CF: {total_cf_all}")

            # Also check what months/years you have data for:
            cursor.execute("SELECT DISTINCT strftime('%Y-%m', completion_date) FROM pm_completions ORDER BY 1")
            pm_months = [row[0] for row in cursor.fetchall()]

            cursor.execute("SELECT DISTINCT strftime('%Y-%m', report_date) FROM cannot_find_assets ORDER BY 1") 
            cf_months = [row[0] for row in cursor.fetchall()]

            print(f"DEBUG: PM completion months available: {pm_months}")
            print(f"DEBUG: Cannot find months available: {cf_months}")
            
            # Add this debug query:
            cursor.execute('SELECT COUNT(*) FROM pm_completions WHERE completion_date IS NULL')
            null_pm_count = cursor.fetchone()[0]

            cursor.execute('SELECT COUNT(*) FROM cannot_find_assets WHERE report_date IS NULL')
            null_cf_count = cursor.fetchone()[0]

            print(f"DEBUG: Records with NULL dates - PM: {null_pm_count}, CF: {null_cf_count}")

            # Also check what those NULL records look like:
            if null_pm_count > 0:
                cursor.execute('SELECT bfm_equipment_no, pm_type, technician_name FROM pm_completions WHERE completion_date IS NULL LIMIT 5')
                null_samples = cursor.fetchall()
                print(f"DEBUG: Sample NULL date PM records: {null_samples}")


            # Add these debug queries to see the most recent entries:
            cursor.execute('''
                SELECT completion_date, COUNT(*) 
                FROM pm_completions 
                WHERE completion_date >= '2025-09-01' 
                GROUP BY completion_date 
                ORDER BY completion_date DESC
            ''')
            recent_entries = cursor.fetchall()
            print(f"DEBUG: All September PM entries by date: {recent_entries}")

            # Check the very last entries added to see if there's recent data entry:
            cursor.execute('''
                SELECT completion_date, bfm_equipment_no, pm_type, technician_name
                FROM pm_completions 
                WHERE completion_date LIKE '2025-09%'
                ORDER BY rowid DESC 
                LIMIT 10
            ''')
            latest_entries = cursor.fetchall()
            print(f"DEBUG: Latest 10 September entries: {latest_entries}")
            
            # Add this debug to find ALL records with single-digit month format:
            cursor.execute('''
                SELECT completion_date, COUNT(*) 
                FROM pm_completions 
                WHERE completion_date LIKE '2025-9-%'
                GROUP BY completion_date 
                ORDER BY completion_date
            ''')
            single_digit_pm = cursor.fetchall()

            cursor.execute('''
                SELECT report_date, COUNT(*) 
                FROM cannot_find_assets 
                WHERE report_date LIKE '2025-9-%'
                GROUP BY report_date 
                ORDER BY report_date
            ''')
            single_digit_cf = cursor.fetchall()

            cursor.execute('SELECT COUNT(*) FROM pm_completions WHERE completion_date LIKE "2025-9-%"')
            total_single_pm = cursor.fetchone()[0]

            cursor.execute('SELECT COUNT(*) FROM cannot_find_assets WHERE report_date LIKE "2025-9-%"')
            total_single_cf = cursor.fetchone()[0]

            print(f"DEBUG: Single-digit month PM records: {single_digit_pm}")
            print(f"DEBUG: Single-digit month CF records: {single_digit_cf}")
            print(f"DEBUG: Total single-digit format - PM: {total_single_pm}, CF: {total_single_cf}")
            print(f"DEBUG: Missing total would be: {224 + total_single_pm + total_single_cf}")


            # Add these debug queries to check for ALL possible date variations:

            # Check for dates with different separators or formats
            cursor.execute('''
                SELECT completion_date, COUNT(*) 
                FROM pm_completions 
                WHERE (completion_date LIKE '%2025%' AND completion_date LIKE '%9%')
                OR (completion_date LIKE '%25-9%')
                OR (completion_date LIKE '%25/9%')
                GROUP BY completion_date 
                ORDER BY completion_date
            ''')
            all_sept_variations = cursor.fetchall()

            # Check the total count using strftime for September (this handles all formats)
            cursor.execute('''
                SELECT completion_date, COUNT(*)
                FROM pm_completions 
                WHERE (strftime('%Y', completion_date) = '2025' AND strftime('%m', completion_date) = '09')
                OR (strftime('%Y', completion_date) = '2025' AND strftime('%m', completion_date) = '9')
                GROUP BY completion_date
                ORDER BY completion_date
            ''')
            strftime_sept = cursor.fetchall()

            cursor.execute('''
                SELECT COUNT(*)
                FROM pm_completions 
                WHERE (strftime('%Y', completion_date) = '2025' AND strftime('%m', completion_date) = '09')
                OR (strftime('%Y', completion_date) = '2025' AND strftime('%m', completion_date) = '9')
            ''')
            total_strftime_pm = cursor.fetchone()[0]

            # Same for cannot_find_assets
            cursor.execute('''
                SELECT COUNT(*)
                FROM cannot_find_assets 
                WHERE (strftime('%Y', report_date) = '2025' AND strftime('%m', report_date) = '09')
                OR (strftime('%Y', report_date) = '2025' AND strftime('%m', report_date) = '9')
            ''')
            total_strftime_cf = cursor.fetchone()[0]

            print(f"DEBUG: All Sept variations found: {all_sept_variations}")
            print(f"DEBUG: Strftime September entries: {strftime_sept}")
            print(f"DEBUG: Total using strftime - PM: {total_strftime_pm}, CF: {total_strftime_cf}, Total: {total_strftime_pm + total_strftime_cf}")

            # Also check if there might be records in other tables
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            all_tables = [row[0] for row in cursor.fetchall()]
            print(f"DEBUG: All tables in database: {all_tables}")


            # Combine both lists
            all_completions = list(completions) + list(cannot_finds)
            all_completions.sort(key=lambda x: x[0], reverse=True)  # Sort by date descending
        
            print(f"DEBUG: Combined total: {len(all_completions)}")

            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)

            # Add completions to tree
            for completion in all_completions:
                date, bfm_no, description, pm_type, technician, hours, notes = completion
                hours_display = f"{hours:.1f}h" if hours else "0.0h"
                notes_preview = (notes[:50] + '...') if notes and len(notes) > 50 else (notes or '')
            
                tree.insert('', 'end', values=(
                    date, bfm_no, description or '', pm_type, technician, hours_display, notes_preview
                ))

            # Generate summary
            month_name = calendar.month_name[month_int]
            summary = f"PM COMPLETIONS SUMMARY - {month_name} {year}\n"
            summary += "=" * 50 + "\n\n"

            # Count by PM type
            pm_type_counts = {}
            total_hours = 0
            technician_counts = {}

            for completion in all_completions:
                pm_type = completion[3]
                technician = completion[4]
                hours = completion[5] or 0
            
                pm_type_counts[pm_type] = pm_type_counts.get(pm_type, 0) + 1
                total_hours += hours
                technician_counts[technician] = technician_counts.get(technician, 0) + 1

            summary += f"Total Completions: {len(all_completions)}\n"
            summary += f"Total Labor Hours: {total_hours:.1f} hours\n\n"

            if pm_type_counts:
                summary += "BY PM TYPE:\n"
                for pm_type, count in sorted(pm_type_counts.items()):
                    summary += f"  {pm_type}: {count}\n"
                summary += "\n"

            if technician_counts:
                summary += "BY TECHNICIAN:\n"
                for tech, count in sorted(technician_counts.items()):
                    avg_hours = sum(c[5] or 0 for c in all_completions if c[4] == tech) / count
                    summary += f"  {tech}: {count} completions, {avg_hours:.1f}h avg\n"

            # Display summary
            summary_text.delete('1.0', 'end')
            summary_text.insert('1.0', summary)

            self.update_status(f"Loaded {len(all_completions)} completions for {month_name} {year}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load monthly data: {str(e)}")
    
    

    def export_monthly_data(self, month_var, year_var):
        """Export monthly completion data to CSV"""
        try:
            # Parse month and year
            month_text = month_var.get()
            month_num = month_text.split(' - ')[0] if month_text else f"{datetime.now().month:02d}"
            year = year_var.get() or str(datetime.now().year)
            month_name = calendar.month_name[int(month_num)]
        
            # Get file path
            filename = filedialog.asksaveasfilename(
            title="Export Monthly PM Completions",
            defaultextension=".csv",
            initialname=f"PM_Completions_{month_name}_{year}.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        
            if filename:
                # Calculate date range
                start_date = f"{year}-{month_num}-01"
                year_int = int(year)
                month_int = int(month_num)
                if month_int == 12:
                    next_month = 1
                    next_year = year_int + 1
                else:
                    next_month = month_int + 1
                    next_year = year_int
                end_date = (datetime(next_year, next_month, 1) - timedelta(days=1)).strftime('%Y-%m-%d')
            
                cursor = self.conn.cursor()
            
                # Get all completion data
                cursor.execute('''
                    SELECT 
                        pc.completion_date,
                        pc.bfm_equipment_no,
                        e.sap_material_no,
                        e.description,
                        e.location,
                        pc.pm_type,
                        pc.technician_name,
                        pc.labor_hours,
                        pc.labor_minutes,
                        (pc.labor_hours + pc.labor_minutes/60.0) as total_hours,
                        pc.special_equipment,
                        pc.notes,
                        pc.pm_due_date,
                        pc.next_annual_pm_date
                    FROM pm_completions pc
                    LEFT JOIN equipment e ON pc.bfm_equipment_no = e.bfm_equipment_no
                    WHERE pc.completion_date BETWEEN ? AND ?
                    UNION ALL
                    SELECT 
                        cf.report_date,
                        cf.bfm_equipment_no,
                        '' as sap_material_no,
                        cf.description,
                        cf.location,
                        'CANNOT FIND' as pm_type,
                        cf.technician_name,
                        0 as labor_hours,
                        0 as labor_minutes,
                        0 as total_hours,
                        '' as special_equipment,
                        cf.notes,
                        '' as pm_due_date,
                        '' as next_annual_pm_date
                    FROM cannot_find_assets cf
                    WHERE cf.report_date BETWEEN ? AND ?
                    ORDER BY completion_date DESC
                ''', (start_date, end_date, start_date, end_date))
            
                data = cursor.fetchall()
            
                # Create DataFrame
                columns = [
                    'Completion Date', 'BFM Equipment No', 'SAP Material No', 
                    'Equipment Description', 'Location', 'PM Type', 'Technician', 
                    'Labor Hours', 'Labor Minutes', 'Total Hours', 'Special Equipment', 
                    'Notes', 'PM Due Date', 'Next Annual PM Date'
                ]
            
                df = pd.DataFrame(data, columns=columns)
                df.to_csv(filename, index=False)
            
                messagebox.showinfo("Success", f"Monthly data exported to: {filename}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export monthly data: {str(e)}")   
        
    
    def create_cannot_find_tab(self):
        """Cannot Find Assets tab"""
        self.cannot_find_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cannot_find_frame, text="Cannot Find Assets")
    
        # Controls
        controls_frame = ttk.LabelFrame(self.cannot_find_frame, text="Cannot Find Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Button(controls_frame, text="Refresh List", 
                command=self.load_cannot_find_assets).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export to PDF", 
                command=self.export_cannot_find_pdf).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Mark as Found", 
                command=self.mark_asset_found).pack(side='left', padx=5)
    
        # Cannot Find list
        list_frame = ttk.LabelFrame(self.cannot_find_frame, text="Missing Assets", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        self.cannot_find_tree = ttk.Treeview(list_frame,
                                        columns=('BFM', 'Description', 'Location', 'Technician', 'Report Date', 'Status'),
                                        show='headings')
    
        columns_config = {
            'BFM': ('BFM Equipment No.', 130),
            'Description': ('Description', 250),
            'Location': ('Location', 120),
            'Technician': ('Reported By', 120),
            'Report Date': ('Report Date', 100),
            'Status': ('Status', 80)
        }
    
        for col, (heading, width) in columns_config.items():
            self.cannot_find_tree.heading(col, text=heading)
            self.cannot_find_tree.column(col, width=width)
    
        # Scrollbars
        cf_v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.cannot_find_tree.yview)
        cf_h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.cannot_find_tree.xview)
        self.cannot_find_tree.configure(yscrollcommand=cf_v_scrollbar.set, xscrollcommand=cf_h_scrollbar.set)
    
        # Pack treeview and scrollbars
        self.cannot_find_tree.grid(row=0, column=0, sticky='nsew')
        cf_v_scrollbar.grid(row=0, column=1, sticky='ns')
        cf_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
    
        # Load initial data
        self.load_cannot_find_assets()
        
        
    def create_run_to_failure_tab(self):
        """Run to Failure Assets tab"""
        self.run_to_failure_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.run_to_failure_frame, text="Run to Failure Assets")
    
        # Controls
        controls_frame = ttk.LabelFrame(self.run_to_failure_frame, text="Run to Failure Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
    
        ttk.Button(controls_frame, text="Refresh List", 
                command=self.load_run_to_failure_assets).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export to PDF", 
                command=self.export_run_to_failure_pdf).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Reactivate Asset", 
                command=self.reactivate_asset).pack(side='left', padx=5)
    
        # Run to Failure list
        list_frame = ttk.LabelFrame(self.run_to_failure_frame, text="Run to Failure Assets", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        self.run_to_failure_tree = ttk.Treeview(list_frame,
                                            columns=('BFM', 'Description', 'Location', 'Technician', 'Completion Date', 'Hours'),
                                            show='headings')
    
        columns_config = {
            'BFM': ('BFM Equipment No.', 130),
            'Description': ('Description', 250),
            'Location': ('Location', 120),
            'Technician': ('Completed By', 120),
            'Completion Date': ('Completion Date', 120),
            'Hours': ('Hours', 80)
        }
    
        for col, (heading, width) in columns_config.items():
            self.run_to_failure_tree.heading(col, text=heading)
            self.run_to_failure_tree.column(col, width=width)
    
        # Scrollbars
        rtf_v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.run_to_failure_tree.yview)
        rtf_h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.run_to_failure_tree.xview)
        self.run_to_failure_tree.configure(yscrollcommand=rtf_v_scrollbar.set, xscrollcommand=rtf_h_scrollbar.set)
    
        # Pack treeview and scrollbars
        self.run_to_failure_tree.grid(row=0, column=0, sticky='nsew')
        rtf_v_scrollbar.grid(row=0, column=1, sticky='ns')
        rtf_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
    
        # Load initial data
        self.load_run_to_failure_assets()


    
    def create_cm_management_tab(self):
        """Enhanced Corrective Maintenance management tab with SharePoint integration and filter"""
        self.cm_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cm_frame, text="Corrective Maintenance")

        # CM controls - Enhanced with SharePoint button
        controls_frame = ttk.LabelFrame(self.cm_frame, text="CM Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)

        # First row of controls
        controls_row1 = ttk.Frame(controls_frame)
        controls_row1.pack(fill='x', pady=(0, 5))
    
        ttk.Button(controls_row1, text="Create New CM", 
                command=self.create_cm_dialog).pack(side='left', padx=5)
        ttk.Button(controls_row1, text="Edit CM", 
                command=self.edit_cm_dialog).pack(side='left', padx=5)
        ttk.Button(controls_row1, text="Complete CM", 
                command=self.complete_cm_dialog).pack(side='left', padx=5)
        ttk.Button(controls_row1, text="Refresh CM List", 
                command=self.load_corrective_maintenance_with_filter).pack(side='left', padx=5)

        # Filter controls
        filter_frame = ttk.Frame(controls_frame)
        filter_frame.pack(fill='x')
        
        ttk.Label(filter_frame, text="Filter by Status:").pack(side='left', padx=(0, 5))
    
        # Create filter dropdown
        self.cm_filter_var = tk.StringVar(value="All")
        self.cm_filter_dropdown = ttk.Combobox(filter_frame, textvariable=self.cm_filter_var, 
                                            values=["All", "Open", "In Progress", "Completed", "On Hold"],
                                            state="readonly", width=15)
        self.cm_filter_dropdown.pack(side='left', padx=5)
        self.cm_filter_dropdown.bind('<<ComboboxSelected>>', self.filter_cm_list)
    
        # Clear filter button
        ttk.Button(filter_frame, text="Clear Filter", 
                command=self.clear_cm_filter).pack(side='left', padx=5)

        # CM list with enhanced columns for SharePoint data
        cm_list_frame = ttk.LabelFrame(self.cm_frame, text="Corrective Maintenance List", padding=10)
        cm_list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Enhanced treeview with additional columns
        self.cm_tree = ttk.Treeview(cm_list_frame,
                                columns=('CM Number', 'BFM', 'Description', 'Priority', 'Assigned', 'Status', 'Created', 'Source'),
                                show='headings')

        cm_columns = {
            'CM Number': 120,
            'BFM': 120,
            'Description': 250,
            'Priority': 80,
            'Assigned': 120,
            'Status': 80,
            'Created': 100,
            'Source': 80  # New column to show if from SharePoint
        }

        for col, width in cm_columns.items():
            self.cm_tree.heading(col, text=col)
            self.cm_tree.column(col, width=width)

        # Scrollbars
        cm_v_scrollbar = ttk.Scrollbar(cm_list_frame, orient='vertical', command=self.cm_tree.yview)
        cm_h_scrollbar = ttk.Scrollbar(cm_list_frame, orient='horizontal', command=self.cm_tree.xview)
        self.cm_tree.configure(yscrollcommand=cm_v_scrollbar.set, xscrollcommand=cm_h_scrollbar.set)

        # Pack treeview and scrollbars
        self.cm_tree.grid(row=0, column=0, sticky='nsew')
        cm_v_scrollbar.grid(row=0, column=1, sticky='ns')
        cm_h_scrollbar.grid(row=1, column=0, sticky='ew')

        cm_list_frame.grid_rowconfigure(0, weight=1)
        cm_list_frame.grid_columnconfigure(0, weight=1)
        
        # Initialize filter data storage
        self.cm_original_data = []
        
        # Load CM data
        self.load_corrective_maintenance_with_filter()

    def load_corrective_maintenance_with_filter(self):
        """Wrapper for your existing load method that adds filter support"""
    
        # Initialize/clear filter data
        self.cm_original_data = []
    
        # Call your existing load method
        self.load_corrective_maintenance()
    
        # After loading, capture data for filtering
        for item in self.cm_tree.get_children():
            item_values = self.cm_tree.item(item, 'values')
            self.cm_original_data.append(item_values)
    
        # Reset filter to show all
        if hasattr(self, 'cm_filter_var'):
            self.cm_filter_var.set("All")
    
    def filter_cm_list(self, event=None):
        """Filter the CM list based on selected status"""
        # Don't filter if no data is loaded yet
        if not hasattr(self, 'cm_original_data') or not self.cm_original_data:
         
            return
        
        selected_filter = self.cm_filter_var.get()
        
    
        # Clear current tree
        for item in self.cm_tree.get_children():
            self.cm_tree.delete(item)
    
        # Filter and display data
        filtered_count = 0
        for item_data in self.cm_original_data:
            # Check if status matches (Status is at index 5)
            if selected_filter == "All" or (len(item_data) > 5 and str(item_data[5]) == selected_filter):
                self.cm_tree.insert('', 'end', values=item_data)
                filtered_count += 1
        
    def clear_cm_filter(self):
        """Clear the filter and show all items"""
        self.cm_filter_var.set("All")
        self.filter_cm_list()
   
    def import_sharepoint_cm_data(self):
        """Import CM data from SharePoint workbook"""
        # Show method selection dialog
        self.show_sharepoint_import_method_dialog()

    def show_sharepoint_import_method_dialog(self):
        """Show dialog to select SharePoint import method"""
        dialog = tk.Toplevel(self.root)
        dialog.title("SharePoint Import Options")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Instructions
        instructions_frame = ttk.LabelFrame(dialog, text="Import Options", padding=15)
        instructions_frame.pack(fill='x', padx=10, pady=5)
    
        instructions = """
        Choose how to import data from SharePoint:
    
        Option 1: Direct SharePoint Integration (Requires authentication)
        Option 2: Manual File Upload (Download Excel file from SharePoint first)
        Option 3: SharePoint REST API (Advanced - requires app registration)
        """
    
        ttk.Label(instructions_frame, text=instructions, font=('Arial', 10)).pack(anchor='w')
    
        # Buttons for different methods
        buttons_frame = ttk.Frame(dialog)
        buttons_frame.pack(fill='x', padx=10, pady=20)
    
        ttk.Button(buttons_frame, text="Option 1: Direct SharePoint", 
                command=lambda: self.try_direct_sharepoint_import(dialog)).pack(pady=5, fill='x')
        ttk.Button(buttons_frame, text="Option 2: Upload Excel File", 
                command=lambda: self.manual_excel_upload(dialog)).pack(pady=5, fill='x')
        ttk.Button(buttons_frame, text="Option 3: REST API Import", 
                command=lambda: self.sharepoint_rest_api_import(dialog)).pack(pady=5, fill='x')
        ttk.Button(buttons_frame, text="Cancel", 
                command=dialog.destroy).pack(pady=5, fill='x')

    def try_direct_sharepoint_import(self, parent_dialog):
        """Attempt direct SharePoint integration"""
        parent_dialog.destroy()
    
        try:
            # This requires additional libraries
            import requests
            from requests_ntlm import HttpNtlmAuth
            import getpass
        
            # Create authentication dialog
            auth_dialog = tk.Toplevel(self.root)
            auth_dialog.title("SharePoint Authentication")
            auth_dialog.geometry("400x300")
            auth_dialog.transient(self.root)
            auth_dialog.grab_set()
        
            # Authentication form
            ttk.Label(auth_dialog, text="SharePoint Credentials:", font=('Arial', 12, 'bold')).pack(pady=10)
            
            # Username
            ttk.Label(auth_dialog, text="Username:").pack(anchor='w', padx=20)
            username_var = tk.StringVar()
            ttk.Entry(auth_dialog, textvariable=username_var, width=40).pack(pady=5, padx=20)
        
            # Password
            ttk.Label(auth_dialog, text="Password:").pack(anchor='w', padx=20)
            password_var = tk.StringVar()
            password_entry = ttk.Entry(auth_dialog, textvariable=password_var, show="*", width=40)
            password_entry.pack(pady=5, padx=20)
        
            # Site URL
            ttk.Label(auth_dialog, text="SharePoint Site URL:").pack(anchor='w', padx=20)
            site_url_var = tk.StringVar(value="https://aitgo.sharepoint.com/sites/PMCM")
            ttk.Entry(auth_dialog, textvariable=site_url_var, width=40).pack(pady=5, padx=20)
        
            def authenticate_and_import():
                try:
                    self.sharepoint_status_label.config(text="Connecting to SharePoint...")
                    self.root.update()
                
                    # This is a simplified example - real implementation would need proper OAuth
                    username = username_var.get()
                    password = password_var.get()
                    site_url = site_url_var.get()
                
                    if not all([username, password, site_url]):
                        messagebox.showerror("Error", "Please fill in all fields")
                        return
                
                    # Attempt connection (this is a basic example)
                    success = self.connect_to_sharepoint_direct(site_url, username, password)
                
                    if success:
                        auth_dialog.destroy()
                        self.sharepoint_status_label.config(text="SharePoint connected successfully")
                    else:
                        messagebox.showerror("Authentication Failed", 
                                        "Could not connect to SharePoint. Try manual upload instead.")
                    
                except Exception as e:
                    messagebox.showerror("Error", f"SharePoint connection failed: {str(e)}")
        
            ttk.Button(auth_dialog, text="Connect", command=authenticate_and_import).pack(pady=20)
            ttk.Button(auth_dialog, text="Cancel", command=auth_dialog.destroy).pack()
        
        except ImportError:
            messagebox.showinfo("Additional Libraries Needed", 
                            "Direct SharePoint integration requires additional libraries.\n"
                            "Please use 'Upload Excel File' option instead.")
            self.manual_excel_upload(None)

    def manual_excel_upload(self, parent_dialog):
        """Manual Excel file upload for SharePoint data"""
        if parent_dialog:
            parent_dialog.destroy()
    
        # Instructions dialog
        instructions = messagebox.showinfo("Manual Upload Instructions", 
                                        "1. Go to the SharePoint link\n"
                                        "2. Open the Asset Maintenance workbook\n"
                                        "3. Go to the 'CMDATA' tab\n"
                                        "4. Download/Save the file as Excel (.xlsx)\n"
                                        "5. Click OK to select the downloaded file")
    
        # File selection
        file_path = filedialog.askopenfilename(
            title="Select Downloaded SharePoint Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("All files", "*.*")]
        )
    
        if file_path:
            try:
                self.process_sharepoint_excel_file(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process Excel file: {str(e)}")

    def process_sharepoint_excel_file(self, file_path):
        """Process the SharePoint Excel file and import CMDATA"""
        try:
            import pandas as pd
        
            self.sharepoint_status_label.config(text="Processing Excel file...")
            self.root.update()
        
            # Read the CMDATA sheet
            try:
                df = pd.read_excel(file_path, sheet_name='CMData')
            except Exception as e:
                # If CMDATA sheet doesn't exist, show available sheets
                xl_file = pd.ExcelFile(file_path)
                available_sheets = xl_file.sheet_names
            
                messagebox.showerror("Sheet Not Found", 
                                f"Could not find 'CMDATA' sheet.\n\n"
                                f"Available sheets: {', '.join(available_sheets)}\n\n"
                                f"Please verify the correct sheet name.")
                return
        
            # Show data preview and column mapping dialog
            self.show_sharepoint_data_preview(df)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
            self.sharepoint_status_label.config(text="Import failed")

    def show_sharepoint_data_preview(self, df):
        """Show preview of SharePoint data and allow column mapping"""
        dialog = tk.Toplevel(self.root)
        dialog.title("SharePoint Data Preview & Mapping")
        dialog.geometry("900x700")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Data preview
        preview_frame = ttk.LabelFrame(dialog, text="Data Preview (First 10 rows)", padding=10)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        # Create treeview for data preview
        preview_columns = list(df.columns)
        preview_tree = ttk.Treeview(preview_frame, columns=preview_columns, show='headings')
    
        # Configure columns
        for col in preview_columns:
            preview_tree.heading(col, text=col)
            preview_tree.column(col, width=100)
    
        # Add data (first 10 rows)
        for index, row in df.head(10).iterrows():
            values = [str(val) if pd.notna(val) else '' for val in row]
            preview_tree.insert('', 'end', values=values)
    
        # Scrollbars
        preview_v_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=preview_tree.yview)
        preview_h_scrollbar = ttk.Scrollbar(preview_frame, orient='horizontal', command=preview_tree.xview)
        preview_tree.configure(yscrollcommand=preview_v_scrollbar.set, xscrollcommand=preview_h_scrollbar.set)
    
        preview_tree.grid(row=0, column=0, sticky='nsew')
        preview_v_scrollbar.grid(row=0, column=1, sticky='ns')
        preview_h_scrollbar.grid(row=1, column=0, sticky='ew')
    
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
    
        # Column mapping
        mapping_frame = ttk.LabelFrame(dialog, text="Map Columns to CM Fields", padding=10)
        mapping_frame.pack(fill='x', padx=10, pady=5)
    
        # Column mappings
        mappings = {}
    
        # CM fields that can be mapped
        cm_fields = [
            ("CM Number/ID", "cm_number"),
            ("Equipment/BFM Number", "bfm_equipment_no"),
            ("Problem Description", "description"),
            ("Priority Level", "priority"),
            ("Assigned Technician", "assigned_technician"),
            ("Status", "status"),
            ("Created Date", "created_date"),
            ("Notes/Comments", "notes")
        ]
    
        # Add "None" option to CSV columns
        column_options = ["(Not in Data)"] + list(df.columns)
    
        row = 0
        for field_name, field_key in cm_fields:
            ttk.Label(mapping_frame, text=field_name + ":").grid(row=row, column=0, sticky='w', pady=2)
        
            mapping_var = tk.StringVar()
            combo = ttk.Combobox(mapping_frame, textvariable=mapping_var, values=column_options, width=30)
            combo.grid(row=row, column=1, padx=10, pady=2)
        
            # Try to auto-match common column names
            for col in df.columns:
                col_lower = col.lower()
                if field_key == 'cm_number' and any(term in col_lower for term in ['cm', 'id', 'number', 'ticket']):
                    mapping_var.set(col)
                    break
                elif field_key == 'bfm_equipment_no' and any(term in col_lower for term in ['bfm', 'equipment', 'asset']):
                    mapping_var.set(col)
                    break
                elif field_key == 'description' and any(term in col_lower for term in ['description', 'problem', 'issue']):
                    mapping_var.set(col)
                    break
                elif field_key == 'priority' and 'priority' in col_lower:
                    mapping_var.set(col)
                    break
                elif field_key == 'assigned_technician' and any(term in col_lower for term in ['technician', 'assigned', 'owner']):
                    mapping_var.set(col)
                    break
                elif field_key == 'status' and 'status' in col_lower:
                    mapping_var.set(col)
                    break
                elif field_key == 'created_date' and any(term in col_lower for term in ['date', 'created', 'opened']):
                    mapping_var.set(col)
                    break
        
            mappings[field_key] = mapping_var
            row += 1
    
        def import_sharepoint_data():
            """Import the mapped SharePoint data"""
            try:
                cursor = self.conn.cursor()
                imported_count = 0
                error_count = 0
            
                self.sharepoint_status_label.config(text="Importing SharePoint data...")
                self.root.update()
            
                for index, row in df.iterrows():
                    try:
                        # Extract mapped data
                        data = {}
                        for field_key, mapping_var in mappings.items():
                            column_name = mapping_var.get()
                            if column_name != "(Not in Data)" and column_name in df.columns:
                                value = row[column_name]
                                if pd.isna(value):
                                    data[field_key] = None
                                else:
                                    # Handle different data types
                                    if field_key == 'created_date':
                                        try:
                                            # Try to parse date
                                            parsed_date = pd.to_datetime(value).strftime('%Y-%m-%d')
                                            data[field_key] = parsed_date
                                        except:
                                            data[field_key] = str(value)
                                    else:
                                        data[field_key] = str(value)
                            else:
                                data[field_key] = None
                    
                        # Generate CM number if not provided
                        if not data.get('cm_number'):
                            data['cm_number'] = f"SP-{datetime.now().strftime('%Y%m%d')}-{index+1:04d}"
                    
                        # Set defaults
                        if not data.get('priority'):
                            data['priority'] = 'Medium'
                        if not data.get('status'):
                            data['status'] = 'Open'
                        if not data.get('created_date'):
                            data['created_date'] = datetime.now().strftime('%Y-%m-%d')
                    
                        # Insert into database with source tracking
                        cursor.execute('''
                            INSERT OR REPLACE INTO corrective_maintenance 
                            (cm_number, bfm_equipment_no, description, priority, assigned_technician, 
                            status, created_date, notes)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            data.get('cm_number'),
                            data.get('bfm_equipment_no'),
                            data.get('description'),
                            data.get('priority'),
                            data.get('assigned_technician'),
                            data.get('status'),
                            data.get('created_date'),
                            f"Imported from SharePoint: {data.get('notes', '')}"
                        ))
                    
                        imported_count += 1
                    
                    except Exception as e:
                        print(f"Error importing row {index}: {e}")
                        error_count += 1
                        continue
            
                self.conn.commit()
                dialog.destroy()
            
                # Show results
                result_msg = f"SharePoint import completed!\n\n"
                result_msg += f"Successfully imported: {imported_count} records\n"
                if error_count > 0:
                    result_msg += f"Skipped (errors): {error_count} records\n"
                result_msg += f"\nTotal processed: {imported_count + error_count} records"
            
                messagebox.showinfo("Import Results", result_msg)
            
                # Refresh CM list
                self.load_corrective_maintenance()
                self.sharepoint_status_label.config(text=f"Imported {imported_count} CMs from SharePoint")
                self.update_status(f"Imported {imported_count} CM records from SharePoint")
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import data: {str(e)}")
                self.sharepoint_status_label.config(text="Import failed")
    
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)
    
        ttk.Button(button_frame, text="Import Data", command=import_sharepoint_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)

    def sharepoint_rest_api_import(self, parent_dialog):
        """SharePoint REST API import (advanced method)"""
        if parent_dialog:
            parent_dialog.destroy()
    
        messagebox.showinfo("REST API Import", 
                        "REST API import requires:\n"
                        "1. App registration in Azure AD\n"
                        "2. SharePoint app permissions\n"
                        "3. Client ID and secret\n\n"
                        "This is an advanced method typically configured by IT.\n"
                        "Please use the manual upload option for now.")

    def connect_to_sharepoint_direct(self, site_url, username, password):
        """Attempt direct SharePoint connection"""
        try:
            # This is a simplified example - real implementation would use proper authentication
            # You would need libraries like Office365-REST-Python-Client or similar
        
            # Placeholder for actual SharePoint connection
            # In reality, you'd need proper OAuth2 authentication
        
            return False  # For now, return False to redirect to manual upload
        
        except Exception as e:
            print(f"SharePoint connection error: {e}")
            return False

    # Enhanced load_corrective_maintenance to show source
    def load_corrective_maintenance(self):
        """Load corrective maintenance data with enhanced source tracking"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT cm_number, bfm_equipment_no, description, priority, 
                    assigned_technician, status, created_date, notes
                FROM corrective_maintenance 
                ORDER BY created_date DESC
            ''')
        
            # Clear existing items
            for item in self.cm_tree.get_children():
                self.cm_tree.delete(item)
        
            # Add CM records
            for cm in cursor.fetchall():
                cm_number, bfm_no, description, priority, assigned, status, created, notes = cm
            
                # Determine source
                source = "SharePoint" if notes and "Imported from SharePoint" in notes else "Manual"
            
                # Truncate description for display
                display_desc = (description[:47] + '...') if description and len(description) > 50 else (description or '')
            
                self.cm_tree.insert('', 'end', values=(
                    cm_number, bfm_no, display_desc, priority, assigned, status, created, source
                ))
            
        except Exception as e:
            print(f"Error loading corrective maintenance: {e}")
        
    

    
    
    
    def create_analytics_dashboard_tab(self):
        """Analytics and dashboard tab"""
        self.analytics_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.analytics_frame, text="Analytics Dashboard")
        
        # Analytics controls
        controls_frame = ttk.LabelFrame(self.analytics_frame, text="Analytics Controls", padding=10)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls_frame, text="Refresh Dashboard", 
                  command=self.refresh_analytics_dashboard).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Equipment Analytics", 
                  command=self.show_equipment_analytics).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="PM Trends", 
                  command=self.show_pm_trends).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="Export Analytics", 
                  command=self.export_analytics).pack(side='left', padx=5)
        
        # Dashboard display
        dashboard_frame = ttk.LabelFrame(self.analytics_frame, text="Analytics Dashboard", padding=10)
        dashboard_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create notebook for different analytics views
        self.analytics_notebook = ttk.Notebook(dashboard_frame)
        self.analytics_notebook.pack(fill='both', expand=True)
        
        # Overview tab
        overview_frame = ttk.Frame(self.analytics_notebook)
        self.analytics_notebook.add(overview_frame, text="Overview")
        
        self.analytics_text = tk.Text(overview_frame, wrap='word', font=('Courier', 10))
        analytics_scrollbar = ttk.Scrollbar(overview_frame, orient='vertical', command=self.analytics_text.yview)
        self.analytics_text.configure(yscrollcommand=analytics_scrollbar.set)
        
        self.analytics_text.pack(side='left', fill='both', expand=True)
        analytics_scrollbar.pack(side='right', fill='y')
        
        # Load initial analytics
        self.refresh_analytics_dashboard()
    
    def update_status(self, message):
        """Update status bar with message"""
        if hasattr(self, 'status_bar'):
            self.status_bar.config(text=f"AIT CMMS - {message}")
            self.root.update_idletasks()
        else:
            print(f"STATUS: {message}")
    
    
    def update_equipment_suggestions(self, event):
        """Update equipment suggestions in completion form"""
        search_term = self.completion_bfm_var.get().lower()
        
        if len(search_term) >= 2:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT bfm_equipment_no FROM equipment 
                WHERE LOWER(bfm_equipment_no) LIKE ? OR LOWER(description) LIKE ?
                ORDER BY bfm_equipment_no LIMIT 10
            ''', (f'%{search_term}%', f'%{search_term}%'))
            
            suggestions = [row[0] for row in cursor.fetchall()]
            self.bfm_combo['values'] = suggestions
    
    
    
    def load_latest_weekly_schedule(self):
        """Load the most recent weekly schedule on startup"""
        try:
            cursor = self.conn.cursor()
        
            # Find the most recent week with scheduled PMs
            cursor.execute('''
                SELECT week_start_date 
                FROM weekly_pm_schedules 
                ORDER BY week_start_date DESC 
                LIMIT 1
            ''')
        
            latest_week = cursor.fetchone()
        
            if latest_week:
                # Set the week start variable to the latest week
                self.week_start_var.set(latest_week[0])
                # Refresh the display with this week's data
                self.refresh_technician_schedules()
                self.update_status(f"Loaded latest weekly schedule: {latest_week[0]}")
            else:
                # No schedules exist, keep current week
                self.update_status("No weekly schedules found")
            
        except Exception as e:
            print(f"Error loading latest weekly schedule: {e}")
    
    
    
    
    
    
    def submit_pm_completion(self):
        """Enhanced PM completion with validation and verification - PREVENTS DUPLICATES"""
        try:
            # Validate required fields
            if not self.completion_bfm_var.get():
                messagebox.showerror("Error", "Please enter BFM Equipment Number")
                return

            if not self.pm_type_var.get():
                messagebox.showerror("Error", "Please select PM Type")
                return

            if not self.completion_tech_var.get():
                messagebox.showerror("Error", "Please select Technician")
                return

            # Get form data
            bfm_no = self.completion_bfm_var.get().strip()
            pm_type = self.pm_type_var.get()
            technician = self.completion_tech_var.get()
            labor_hours = float(self.labor_hours_var.get() or 0)
            labor_minutes = float(self.labor_minutes_var.get() or 0)
            pm_due_date = self.pm_due_date_var.get().strip()
            special_equipment = self.special_equipment_var.get()
            notes = self.notes_text.get('1.0', 'end-1c')
            next_annual_pm = self.next_annual_pm_var.get().strip()

            # Use PM Due Date as completion date if provided, otherwise today's date
            if pm_due_date:
                completion_date = pm_due_date
            else:
                completion_date = datetime.now().strftime('%Y-%m-%d')

            cursor = self.conn.cursor()

            # 🔍 ENHANCED VALIDATION - Check for recent duplicates
            validation_result = self.validate_pm_completion(cursor, bfm_no, pm_type, technician, completion_date)
            if not validation_result['valid']:
                # Show detailed warning dialog
                response = messagebox.askyesno(
                    "⚠️ Potential Duplicate PM Detected", 
                    f"{validation_result['message']}\n\n"
                    f"Details:\n"
                    f"• Equipment: {bfm_no}\n"
                    f"• PM Type: {pm_type}\n"
                    f"• Technician: {technician}\n"
                    f"• Completion Date: {completion_date}\n\n"
                    f"Do you want to proceed anyway?\n\n"
                    f"Click 'No' to review and make changes.",
                    icon='warning'
                )
                if not response:
                    self.update_status("PM submission cancelled - potential duplicate detected")
                    return

            # Auto-calculate next annual PM date if blank
            if not next_annual_pm and pm_type in ['Monthly', 'Six Month', 'Annual']:
                try:
                    completion_dt = datetime.strptime(completion_date, '%Y-%m-%d')
                except ValueError:
                    completion_dt = datetime.now()

                # ONLY set annual PM date when completing an Annual PM
                if pm_type == 'Annual':
                    next_annual_dt = completion_dt + timedelta(days=365)

                    # Add equipment-specific offset to spread annual PMs
                    try:
                        import re
                        numeric_part = re.findall(r'\d+', bfm_no)
                        if numeric_part:
                            last_digits = int(numeric_part[-1]) % 61
                            offset_days = last_digits - 30  # -30 to +30 days
                        else:
                            offset_days = (hash(bfm_no) % 61) - 30
        
                        next_annual_dt = next_annual_dt + timedelta(days=offset_days)
                    except Exception:
                        import random
                        offset_days = random.randint(-21, 21)
                        next_annual_dt = next_annual_dt + timedelta(days=offset_days)

                    next_annual_pm = next_annual_dt.strftime('%Y-%m-%d')
                    self.next_annual_pm_var.set(next_annual_pm)
                # For Monthly and Six Month PMs, DO NOT change the existing Annual PM date
                # This preserves the independent Annual PM schedule

            # Handle different PM types with TRANSACTION SAFETY
            try:
                # Start transaction
                cursor.execute('BEGIN TRANSACTION')

                if pm_type == 'CANNOT FIND':
                    success = self.process_cannot_find_pm(cursor, bfm_no, technician, completion_date, notes)
                
                elif pm_type == 'Run to Failure':
                    success = self.process_run_to_failure_pm(cursor, bfm_no, technician, completion_date, 
                                                        labor_hours + (labor_minutes/60), notes)
                
                else:  # Normal PM (Monthly, Six Month, Annual)
                    success = self.process_normal_pm_completion(cursor, bfm_no, pm_type, technician, 
                                                            completion_date, labor_hours, labor_minutes, 
                                                            pm_due_date, special_equipment, notes, next_annual_pm)

                if success:
                    # Commit transaction
                    cursor.execute('COMMIT')
                
                    # 🔍 VERIFY the completion was saved correctly
                    verification_result = self.verify_pm_completion_saved(cursor, bfm_no, pm_type, technician, completion_date)
                
                    if verification_result['verified']:
                        messagebox.showinfo("✅ Success", 
                                        f"PM completion recorded and verified!\n\n"
                                        f"Equipment: {bfm_no}\n"
                                        f"PM Type: {pm_type}\n"
                                        f"Technician: {technician}\n"
                                        f"Date: {completion_date}\n\n"
                                        f"✅ Database verification passed")
                    
                        # Clear form and refresh displays
                        self.clear_completion_form()
                        self.load_recent_completions()
                        if hasattr(self, 'refresh_technician_schedules'):
                            self.refresh_technician_schedules()
                        self.update_status(f"✅ PM completed and verified: {bfm_no} - {pm_type} by {technician}")
                    else:
                        messagebox.showerror("⚠️ Warning", 
                                        f"PM was saved but verification failed!\n\n"
                                        f"{verification_result['message']}\n\n"
                                        f"Please check the PM History tab to confirm the completion was recorded.")
                        self.update_status(f"⚠️ PM saved but verification incomplete: {bfm_no}")
                else:
                    # Rollback on failure
                    cursor.execute('ROLLBACK')
                    messagebox.showerror("Error", "Failed to process PM completion. Transaction rolled back.")
                
            except Exception as e:
                # Rollback on exception
                cursor.execute('ROLLBACK')
                raise e

        except Exception as e:
            messagebox.showerror("Error", f"Failed to submit PM completion: {str(e)}")
            import traceback
            print(f"PM Completion Error: {traceback.format_exc()}")
    
    def validate_pm_completion(self, cursor, bfm_no, pm_type, technician, completion_date):
        """Comprehensive validation to prevent duplicate PMs"""
        try:
            issues = []
        
            # Check 1: Same PM type completed recently for this equipment
            cursor.execute('''
                SELECT completion_date, technician_name, id
                FROM pm_completions 
                WHERE bfm_equipment_no = ? AND pm_type = ?
                ORDER BY completion_date DESC LIMIT 1
            ''', (bfm_no, pm_type))
        
            recent_completion = cursor.fetchone()
            if recent_completion:
                last_completion_date, last_technician, completion_id = recent_completion
                try:
                    last_date = datetime.strptime(last_completion_date, '%Y-%m-%d')
                    current_date = datetime.strptime(completion_date, '%Y-%m-%d')
                    days_since = (current_date - last_date).days
                
                    # Different thresholds for different PM types
                    min_days = {
                        'Monthly': 25,      # Monthly PMs should be ~30 days apart
                        'Six Month': 150,   # Six month PMs should be ~180 days apart
                        'Annual': 300       # Annual PMs should be ~365 days apart
                    }
                
                    threshold = min_days.get(pm_type, 7)  # Default 7 days for other types
                
                    if days_since < threshold:
                        issues.append(f"⚠️ DUPLICATE DETECTED: {pm_type} PM for {bfm_no} was completed only {days_since} days ago")
                        issues.append(f"   Previous completion: {last_completion_date} by {last_technician}")
                        issues.append(f"   Minimum interval for {pm_type} PM: {threshold} days")
                    
                except ValueError:
                    # If date parsing fails, flag it as potential issue
                    issues.append(f"⚠️ Date parsing issue with previous completion: {last_completion_date}")

            # Check 2: Same technician completing same equipment too frequently  
            cursor.execute('''
                SELECT COUNT(*) 
                FROM pm_completions 
                WHERE bfm_equipment_no = ? AND technician_name = ?
                AND completion_date >= DATE(?, '-7 days')
            ''', (bfm_no, technician, completion_date))
        
            recent_count = cursor.fetchone()[0]
            if recent_count > 0:
                issues.append(f"⚠️ Same technician ({technician}) completed PM on {bfm_no} within last 7 days")

            # Check 3: Equipment exists and is active
            cursor.execute('SELECT status FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
            equipment_status = cursor.fetchone()
        
            if not equipment_status:
                issues.append(f"❌ Equipment {bfm_no} not found in database")
            elif equipment_status[0] in ['Missing', 'Run to Failure'] and pm_type not in ['CANNOT FIND', 'Run to Failure']:
                issues.append(f"⚠️ Equipment {bfm_no} has status '{equipment_status[0]}' - unusual for {pm_type} PM")

            # Check 4: Scheduled PM exists for this week
            current_week_start = self.get_week_start(datetime.strptime(completion_date, '%Y-%m-%d'))
            cursor.execute('''
                SELECT COUNT(*) FROM weekly_pm_schedules 
                WHERE bfm_equipment_no = ? AND pm_type = ? 
                AND assigned_technician = ? AND week_start_date = ?
            ''', (bfm_no, pm_type, technician, current_week_start.strftime('%Y-%m-%d')))
        
            scheduled_count = cursor.fetchone()[0]
            if scheduled_count == 0 and pm_type in ['Monthly', 'Annual']:
                issues.append(f"ℹ️ No scheduled PM found for this week - completing ahead of schedule")

            # Return validation result
            if issues:
                return {
                    'valid': False,
                    'message': f"Found {len(issues)} potential issue(s):\n\n" + "\n\n".join(issues)
                }
            else:
                return {'valid': True, 'message': 'Validation passed'}
            
        except Exception as e:
            return {
                'valid': False,
                'message': f"Validation error: {str(e)}"
            }

    def verify_pm_completion_saved(self, cursor, bfm_no, pm_type, technician, completion_date):
        """Verify that the PM completion was actually saved to the database"""
        try:
            # Check 1: PM completion record exists
            cursor.execute('''
                SELECT id, completion_date, technician_name, created_date
                FROM pm_completions 
                WHERE bfm_equipment_no = ? AND pm_type = ? AND technician_name = ?
                AND completion_date = ?
                ORDER BY created_date DESC LIMIT 1
            ''', (bfm_no, pm_type, technician, completion_date))
        
            completion_record = cursor.fetchone()
            if not completion_record:
                return {
                    'verified': False,
                    'message': f"❌ PM completion record not found in database"
                }

            # Check 2: Equipment PM dates updated (for normal PMs)
            if pm_type in ['Monthly', 'Six Month', 'Annual']:
                date_field = f'last_{pm_type.lower().replace(" ", "_")}_pm'
                cursor.execute(f'SELECT {date_field} FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
            
                equipment_date = cursor.fetchone()
                if equipment_date and equipment_date[0] != completion_date:
                    return {
                        'verified': False,
                        'message': f"⚠️ Equipment {date_field} not updated correctly. Expected: {completion_date}, Found: {equipment_date[0]}"
                    }

            # Check 3: Weekly schedule updated if applicable
            current_week_start = self.get_week_start(datetime.strptime(completion_date, '%Y-%m-%d'))
            cursor.execute('''
                SELECT status FROM weekly_pm_schedules 
                WHERE bfm_equipment_no = ? AND pm_type = ? 
                AND assigned_technician = ? AND week_start_date = ?
            ''', (bfm_no, pm_type, technician, current_week_start.strftime('%Y-%m-%d')))
        
            schedule_status = cursor.fetchone()
            if schedule_status and schedule_status[0] != 'Completed':
                return {
                    'verified': False,
                    'message': f"⚠️ Weekly schedule not marked as completed. Status: {schedule_status[0]}"
                }

            return {
                'verified': True,
                'message': f"✅ All verification checks passed",
                'completion_id': completion_record[0]
            }
        
        except Exception as e:
            return {
                'verified': False,
                'message': f"Verification error: {str(e)}"
            }

    def process_normal_pm_completion(self, cursor, bfm_no, pm_type, technician, completion_date, 
                                labor_hours, labor_minutes, pm_due_date, special_equipment, notes, next_annual_pm):
        """Process normal PM completion with enhanced error handling"""
        try:
            cursor.execute('''
                INSERT INTO pm_completions 
                (bfm_equipment_no, pm_type, technician_name, completion_date, 
                labor_hours, labor_minutes, pm_due_date, special_equipment, 
                notes, next_annual_pm_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                bfm_no, pm_type, technician, completion_date,
                labor_hours, labor_minutes, pm_due_date, special_equipment,
                notes, next_annual_pm
            ))
        
         
            completion_id = cursor.lastrowid
            if not completion_id:
                raise Exception("Failed to get completion record ID")

            # Update equipment PM dates
            if pm_type == 'Monthly':
                if next_annual_pm: 
                    cursor.execute('''
                        UPDATE equipment SET 
                        last_monthly_pm = ?, 
                        next_monthly_pm = DATE(?, '+30 days'),
                        next_annual_pm = ?,  
                        updated_date = CURRENT_TIMESTAMP
                        WHERE bfm_equipment_no = ?
                    ''', (completion_date, completion_date, next_annual_pm, bfm_no))
                else:
                    cursor.execute('''
                        UPDATE equipment SET 
                        last_monthly_pm = ?, 
                        next_monthly_pm = DATE(?, '+30 days'),
                        updated_date = CURRENT_TIMESTAMP
                        WHERE bfm_equipment_no = ?
                    ''', (completion_date, completion_date, bfm_no))
                    
            elif pm_type == 'Six Month':
                cursor.execute('''
                    UPDATE equipment SET 
                    last_six_month_pm = ?, 
                    next_six_month_pm = DATE(?, '+180 days'),
                    updated_date = CURRENT_TIMESTAMP
                    WHERE bfm_equipment_no = ?
                ''', (completion_date, completion_date, bfm_no))
                
            elif pm_type == 'Annual':
                cursor.execute('''
                    UPDATE equipment SET 
                    last_annual_pm = ?, 
                    next_annual_pm = DATE(?, '+365 days'),
                    updated_date = CURRENT_TIMESTAMP
                    WHERE bfm_equipment_no = ?
                ''', (completion_date, completion_date, bfm_no))

            # Verify equipment update worked
            affected_rows = cursor.rowcount
            if affected_rows != 1:
                raise Exception(f"Equipment update failed - affected {affected_rows} rows instead of 1")

            # Update weekly schedule status if exists
            current_week_start = self.get_week_start(datetime.strptime(completion_date, '%Y-%m-%d'))
            cursor.execute('''
                UPDATE weekly_pm_schedules SET 
                status = 'Completed', 
                completion_date = ?, 
                labor_hours = ?, 
                notes = ?
                WHERE bfm_equipment_no = ? AND pm_type = ? AND assigned_technician = ?
                AND week_start_date = ? AND status = 'Scheduled'
            ''', (completion_date, labor_hours + (labor_minutes/60), notes, 
                bfm_no, pm_type, technician, current_week_start.strftime('%Y-%m-%d')))

            # DEBUG: Check if the update worked
            updated_rows = cursor.rowcount
            print(f"DEBUG: Updated {updated_rows} weekly schedule rows for {bfm_no} - {pm_type} by {technician}")

            print(f"✅ Normal PM completion processed successfully: {bfm_no} - {pm_type}")
            return True
            
        except Exception as e:
            print(f"❌ Error processing normal PM completion: {str(e)}")
            return False
    
    def fix_weekly_schedule_status_flexible(self):
        """Enhanced method to fix weekly schedule status with flexible matching"""
        try:
            cursor = self.conn.cursor()
        
            # First, get all actual completions for the week
            cursor.execute('''
                SELECT bfm_equipment_no, pm_type, technician_name, completion_date,
                    (labor_hours + labor_minutes/60.0) as total_hours, notes
                FROM pm_completions 
                WHERE completion_date BETWEEN '2025-08-25' AND '2025-08-31'
            ''')
        
            completions = cursor.fetchall()
            print(f"Found {len(completions)} actual completions to process")
        
            updated_count = 0
        
            for completion in completions:
                bfm_no, pm_type, technician, comp_date, hours, notes = completion
            
                # Try exact match first
                cursor.execute('''
                    UPDATE weekly_pm_schedules 
                    SET status = 'Completed',
                        completion_date = ?,
                        labor_hours = ?,
                        notes = ?
                    WHERE bfm_equipment_no = ? AND pm_type = ? AND assigned_technician = ?
                    AND week_start_date = '2025-08-25' AND status = 'Scheduled'
                ''', (comp_date, hours, notes, bfm_no, pm_type, technician))
            
                exact_matches = cursor.rowcount
            
                # If no exact match, try equipment + PM type match (without LIMIT)
                if exact_matches == 0:
                    # First check if there's an available scheduled PM for this equipment/PM type
                    cursor.execute('''
                        SELECT id FROM weekly_pm_schedules 
                        WHERE bfm_equipment_no = ? AND pm_type = ?
                        AND week_start_date = '2025-08-25' AND status = 'Scheduled'
                    ''', (bfm_no, pm_type))
                
                    available = cursor.fetchone()
                
                    if available:
                        # Update the first available matching record
                        cursor.execute('''
                            UPDATE weekly_pm_schedules 
                            SET status = 'Completed',
                                completion_date = ?,
                                labor_hours = ?,
                                notes = ?,
                                assigned_technician = ?
                            WHERE id = ?
                        ''', (comp_date, hours, notes, technician, available[0]))
                    
                        flexible_matches = cursor.rowcount
                        if flexible_matches > 0:
                            print(f"Flexible match: {bfm_no} {pm_type} reassigned to {technician}")
                    
                        updated_count += flexible_matches
                    else:
                        print(f"No scheduled PM found for {bfm_no} {pm_type}")
                else:
                    updated_count += exact_matches
                    print(f"Exact match: {bfm_no} {pm_type} by {technician}")
        
            self.conn.commit()
        
            messagebox.showinfo("Success", 
                            f"Processed {len(completions)} completions\n"
                            f"Updated {updated_count} weekly schedule records!")
            print(f"Final result: Updated {updated_count} out of {len(completions)} completions")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fix weekly schedule: {str(e)}")
            print(f"Error: {e}")
    
    
    
    
    def process_cannot_find_pm(self, cursor, bfm_no, technician, completion_date, notes):
        """Process CANNOT FIND PM with validation"""
        try:
            # Get equipment info
            cursor.execute('SELECT description, location FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
            equipment_info = cursor.fetchone()
            description = equipment_info[0] if equipment_info else ''
            location = equipment_info[1] if equipment_info else ''

            # Insert into cannot_find_assets table
            cursor.execute('''
                INSERT INTO cannot_find_assets 
                (bfm_equipment_no, description, location, technician_name, report_date, notes)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (bfm_no, description, location, technician, completion_date, notes))

            # Update equipment status
            cursor.execute('UPDATE equipment SET status = "Missing" WHERE bfm_equipment_no = ?', (bfm_no,))
        
            affected_rows = cursor.rowcount
            if affected_rows != 1:
                raise Exception(f"Equipment status update failed - affected {affected_rows} rows")

            print(f"✅ Cannot Find PM processed: {bfm_no}")
            return True
        
        except Exception as e:
            print(f"❌ Error processing Cannot Find PM: {str(e)}")
            return False

    def process_run_to_failure_pm(self, cursor, bfm_no, technician, completion_date, total_hours, notes):
        """Process Run to Failure PM with validation"""
        try:
            # Get equipment info
            cursor.execute('SELECT description, location FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
            equipment_info = cursor.fetchone()
            description = equipment_info[0] if equipment_info else ''
            location = equipment_info[1] if equipment_info else ''

            # Insert into run_to_failure_assets table
            cursor.execute('''
                INSERT INTO run_to_failure_assets 
                (bfm_equipment_no, description, location, technician_name, completion_date, labor_hours, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (bfm_no, description, location, technician, completion_date, total_hours, notes))

            # Update equipment status and disable all PM types
            cursor.execute('''
                UPDATE equipment SET 
                status = "Run to Failure",
                monthly_pm = 0,
                six_month_pm = 0,
                annual_pm = 0,
                updated_date = CURRENT_TIMESTAMP
                WHERE bfm_equipment_no = ?
            ''', (bfm_no,))
        
            affected_rows = cursor.rowcount
            if affected_rows != 1:
                raise Exception(f"Equipment update failed - affected {affected_rows} rows")

            print(f"✅ Run to Failure PM processed: {bfm_no}")
            return True
        
        except Exception as e:
            print(f"❌ Error processing Run to Failure PM: {str(e)}")
            return False

    # Additional method to add to your class:
    def show_recent_completions_for_equipment(self, bfm_no):
        """Show recent completions for specific equipment - useful for debugging"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT pm_type, technician_name, completion_date, 
                    (labor_hours + labor_minutes/60.0) as total_hours,
                    notes
                FROM pm_completions 
                WHERE bfm_equipment_no = ?
                ORDER BY completion_date DESC LIMIT 10
            ''', (bfm_no,))
        
            completions = cursor.fetchall()
        
            if completions:
                report = f"RECENT PM COMPLETIONS FOR {bfm_no}\n"
                report += "=" * 50 + "\n\n"
            
                for pm_type, tech, date, hours, notes in completions:
                    report += f"• {date} - {pm_type} PM by {tech} ({hours:.1f}h)\n"
                    if notes:
                        report += f"  Notes: {notes[:100]}...\n" if len(notes) > 100 else f"  Notes: {notes}\n"
                    report += "\n"
            
                # Show in a dialog
                dialog = tk.Toplevel(self.root)
                dialog.title(f"PM History - {bfm_no}")
                dialog.geometry("600x400")
            
                text_widget = tk.Text(dialog, wrap='word', font=('Courier', 10))
                text_widget.pack(fill='both', expand=True, padx=10, pady=10)
                text_widget.insert('1.0', report)
                text_widget.config(state='disabled')
            
            else:
                messagebox.showinfo("No History", f"No PM completions found for equipment {bfm_no}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PM history: {str(e)}")
      
    # 7. LOAD CANNOT FIND ASSETS
    def load_cannot_find_assets(self):
        """Load cannot find assets data"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT bfm_equipment_no, description, location, technician_name, report_date, status
                FROM cannot_find_assets 
                WHERE status = 'Missing'
                ORDER BY report_date DESC
            ''')
        
            # Clear existing items
            for item in self.cannot_find_tree.get_children():
                self.cannot_find_tree.delete(item)
        
            # Add cannot find records
            for asset in cursor.fetchall():
                bfm_no, description, location, technician, report_date, status = asset
                self.cannot_find_tree.insert('', 'end', values=(
                    bfm_no, description or '', location or '', technician, report_date, status
                ))
            
        except Exception as e:
            print(f"Error loading cannot find assets: {e}")

    # 8. LOAD RUN TO FAILURE ASSETS
    def load_run_to_failure_assets(self):
        """Enhanced method to load run to failure assets with better data handling"""
        try:
            cursor = self.conn.cursor()
        
            # Get data from both run_to_failure_assets table AND equipment table
            cursor.execute('''
                SELECT DISTINCT
                    COALESCE(rtf.bfm_equipment_no, e.bfm_equipment_no) as bfm_no,
                    COALESCE(rtf.description, e.description) as description,
                    COALESCE(rtf.location, e.location) as location,
                    COALESCE(rtf.technician_name, 'System Change') as technician,
                    COALESCE(rtf.completion_date, e.updated_date, CURRENT_DATE) as completion_date,
                    COALESCE(rtf.labor_hours, 0) as labor_hours,
                    COALESCE(rtf.notes, 'Set via equipment edit') as notes
                FROM equipment e
                LEFT JOIN run_to_failure_assets rtf ON e.bfm_equipment_no = rtf.bfm_equipment_no
                WHERE e.status = 'Run to Failure'
            
                UNION
            
                SELECT 
                    rtf.bfm_equipment_no,
                    rtf.description,
                    rtf.location,
                    rtf.technician_name,
                    rtf.completion_date,
                    rtf.labor_hours,
                    rtf.notes
                FROM run_to_failure_assets rtf
                LEFT JOIN equipment e ON rtf.bfm_equipment_no = e.bfm_equipment_no
                WHERE e.bfm_equipment_no IS NULL OR e.status = 'Run to Failure'
            
                ORDER BY completion_date DESC
            ''')
        
            # Clear existing items
            for item in self.run_to_failure_tree.get_children():
                self.run_to_failure_tree.delete(item)
        
            # Add run to failure records
            for asset in cursor.fetchall():
                bfm_no, description, location, technician, completion_date, hours, notes = asset
                hours_display = f"{hours:.1f}h" if hours else "0.0h"
            
                self.run_to_failure_tree.insert('', 'end', values=(
                    bfm_no,
                    description or 'No description',
                    location or 'Unknown location',
                    technician or 'Unknown',
                    completion_date or '',
                    hours_display
                ))
            
            # Update the count in equipment statistics
            self.update_equipment_statistics()
            
        except Exception as e:
            print(f"Error loading run to failure assets: {e}")
            messagebox.showerror("Error", f"Failed to load Run to Failure assets: {str(e)}")
            
            
    # 9. EXPORT CANNOT FIND TO PDF
    def export_cannot_find_pdf(self):
        """Export Cannot Find assets to PDF"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Cannot_Find_Assets_{timestamp}.pdf"
        
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT bfm_equipment_no, description, location, technician_name, report_date, notes
                FROM cannot_find_assets 
                WHERE status = 'Missing'
                ORDER BY report_date DESC
            ''')
        
            assets = cursor.fetchall()
        
            doc = SimpleDocTemplate(filename, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
        
            # Title
            title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], 
                                    fontSize=18, textColor=colors.darkred, alignment=1)
            story.append(Paragraph("AIRBUS AIT - CANNOT FIND ASSETS REPORT", title_style))
            story.append(Spacer(1, 20))
        
            # Report info
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            story.append(Paragraph(f"Total Missing Assets: {len(assets)}", styles['Normal']))
            story.append(Spacer(1, 20))
        
            if assets:
                # Create table
                data = [['BFM Equipment No.', 'Description', 'Location', 'Reported By', 'Report Date']]
                for asset in assets:
                    bfm_no, description, location, technician, report_date, notes = asset
                    data.append([
                        bfm_no,
                        (description[:30] + '...') if description and len(description) > 30 else (description or ''),
                        location or '',
                        technician,
                        report_date
                    ])
                
                table = Table(data, colWidths=[1.5*inch, 2.5*inch, 1.2*inch, 1.2*inch, 1*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
            
                story.append(table)
            else:
                story.append(Paragraph("No missing assets found.", styles['Normal']))
        
            doc.build(story)
            messagebox.showinfo("Success", f"Cannot Find report exported to: {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Cannot Find report: {str(e)}")

    # 10. EXPORT RUN TO FAILURE TO PDF
    def export_run_to_failure_pdf(self):
        """Export Run to Failure assets to PDF"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Run_to_Failure_Assets_{timestamp}.pdf"
        
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT bfm_equipment_no, description, location, technician_name, completion_date, labor_hours, notes
                FROM run_to_failure_assets 
                ORDER BY completion_date DESC
            ''')
        
            assets = cursor.fetchall()
        
            doc = SimpleDocTemplate(filename, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
        
            # Title
            title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], 
                                    fontSize=18, textColor=colors.darkblue, alignment=1)
            story.append(Paragraph("AIRBUS AIT - RUN TO FAILURE ASSETS REPORT", title_style))
            story.append(Spacer(1, 20))
        
            # Report info
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            story.append(Paragraph(f"Total Run to Failure Assets: {len(assets)}", styles['Normal']))
            story.append(Spacer(1, 20))
        
            if assets:
                # Create table
                data = [['BFM Equipment No.', 'Description', 'Location', 'Completed By', 'Date', 'Hours']]
                for asset in assets:
                    bfm_no, description, location, technician, completion_date, hours, notes = asset
                    data.append([
                        bfm_no,
                        (description[:25] + '...') if description and len(description) > 25 else (description or ''),
                        location or '',
                        technician,
                        completion_date,
                        f"{hours:.1f}h" if hours else ''
                    ])
            
                table = Table(data, colWidths=[1.4*inch, 2.2*inch, 1*inch, 1*inch, 0.8*inch, 0.6*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
            
                story.append(table)
            else:
                story.append(Paragraph("No Run to Failure assets found.", styles['Normal']))
        
            doc.build(story)
            messagebox.showinfo("Success", f"Run to Failure report exported to: {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Run to Failure report: {str(e)}")

    # 11. MARK ASSET AS FOUND
    def mark_asset_found(self):
        """Mark a cannot find asset as found"""
        selected = self.cannot_find_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an asset to mark as found")
            return
    
        item = self.cannot_find_tree.item(selected[0])
        bfm_no = item['values'][0]
    
        try:
            cursor = self.conn.cursor()
            cursor.execute('UPDATE cannot_find_assets SET status = "Found" WHERE bfm_equipment_no = ?', (bfm_no,))
            cursor.execute('UPDATE equipment SET status = "Active" WHERE bfm_equipment_no = ?', (bfm_no,))
            self.conn.commit()
        
            messagebox.showinfo("Success", f"Asset {bfm_no} marked as found and reactivated")
            self.load_cannot_find_assets()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to mark asset as found: {str(e)}")

    # 12. REACTIVATE ASSET
    def reactivate_asset(self):
        """Enhanced method to reactivate a run to failure asset"""
        selected = self.run_to_failure_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an asset to reactivate")
            return
    
        item = self.run_to_failure_tree.item(selected[0])
        bfm_no = item['values'][0]
    
        result = messagebox.askyesno(
            "Confirm Reactivation", 
            f"Reactivate asset {bfm_no} for PM scheduling?\n\n"
            f"This will:\n"
            f"• Set status to Active\n"
            f"• Enable Monthly and Annual PMs\n"
            f"• Remove from Run to Failure list\n"
            f"• Resume normal PM scheduling"
        )
    
        if result:
            try:
                cursor = self.conn.cursor()
            
                # Update equipment status and enable PMs
                cursor.execute('''
                    UPDATE equipment SET 
                    status = 'Active',
                    monthly_pm = 1,
                    six_month_pm = 0,
                    annual_pm = 1,
                    updated_date = CURRENT_TIMESTAMP
                    WHERE bfm_equipment_no = ?
                ''', (bfm_no,))
            
                # Remove from run_to_failure_assets table (optional, for clean data)
                cursor.execute('DELETE FROM run_to_failure_assets WHERE bfm_equipment_no = ?', (bfm_no,))
            
                self.conn.commit()
            
                messagebox.showinfo(
                    "Success", 
                    f"Asset {bfm_no} successfully reactivated!\n\n"
                    f"Status: Active\n"
                    f"PMs Enabled: Monthly, Annual\n"
                    f"Equipment moved back to main equipment list"
                )
            
                # Refresh all displays
                self.refresh_equipment_list()
                self.load_run_to_failure_assets()
                self.update_equipment_statistics()
                self.update_status(f"Reactivated asset {bfm_no}")
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reactivate asset: {str(e)}")
    
    
    def clear_completion_form(self):
        """Clear the PM completion form"""
        self.completion_bfm_var.set('')
        self.pm_type_var.set('')
        self.completion_tech_var.set('')
        self.labor_hours_var.set('0')
        self.labor_minutes_var.set('0')
        self.pm_due_date_var.set('')
        self.special_equipment_var.set('')
        self.notes_text.delete('1.0', 'end')
        self.next_annual_pm_var.set('')
    
    def load_recent_completions(self):
        """Load recent PM completions with debugging"""
        print("DEBUG: load_recent_completions called")
        try:
            cursor = self.conn.cursor()
            print("DEBUG: Database cursor created")
        
            cursor.execute('''
                SELECT completion_date, bfm_equipment_no, pm_type, technician_name, 
                    (labor_hours + labor_minutes/60.0) as total_hours
                FROM pm_completions 
                ORDER BY completion_date DESC, id DESC LIMIT 500
            ''')
        
            completions = cursor.fetchall()
            print(f"DEBUG: Found {len(completions)} completions in database")
        
            # Clear existing items
            for item in self.recent_completions_tree.get_children():
                self.recent_completions_tree.delete(item)
            print("DEBUG: Cleared existing tree items")
        
            # Add recent completions
            for completion in completions:
                completion_date, bfm_no, pm_type, technician, total_hours = completion
                hours_display = f"{total_hours:.1f}h" if total_hours else "0.0h"
            
                self.recent_completions_tree.insert('', 'end', values=(
                    completion_date, bfm_no, pm_type, technician, hours_display
                ))
                print(f"DEBUG: Added {bfm_no} - {pm_type} - {technician}")
        
            print("DEBUG: Successfully loaded recent completions")
            print(f"Refreshed: {len(completions)} recent completions loaded")
        
        except Exception as e:
            print(f"ERROR in load_recent_completions: {e}")
            import traceback
            traceback.print_exc()
    
    def generate_current_week_report(self):
        """Generate report for current week"""
        try:
            week_start = datetime.strptime(self.week_start_var.get(), '%Y-%m-%d')
            week_end = week_start + timedelta(days=6)
            
            cursor = self.conn.cursor()
            
            # Get weekly statistics
            cursor.execute('''
                SELECT 
                    COUNT(*) as total_scheduled,
                    COUNT(CASE WHEN status = 'Completed' THEN 1 END) as total_completed
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
            ''', (week_start.strftime('%Y-%m-%d'),))
            
            total_scheduled, total_completed = cursor.fetchone()
            completion_rate = (total_completed / total_scheduled * 100) if total_scheduled > 0 else 0
            
            # Get technician performance
            cursor.execute('''
                SELECT 
                    assigned_technician,
                    COUNT(*) as assigned,
                    COUNT(CASE WHEN status = 'Completed' THEN 1 END) as completed,
                    AVG(CASE WHEN status = 'Completed' THEN labor_hours END) as avg_hours
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
                GROUP BY assigned_technician
                ORDER BY assigned_technician
            ''', (week_start.strftime('%Y-%m-%d'),))
            
            tech_performance = cursor.fetchall()
            
            # Generate report text
            report = f"WEEKLY PM PERFORMANCE REPORT\n"
            report += f"Week: {week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}\n"
            report += "=" * 80 + "\n\n"
            
            report += f"OVERALL PERFORMANCE:\n"
            report += f"Target PMs for Week: {self.weekly_pm_target}\n"
            report += f"Scheduled PMs: {total_scheduled}\n"
            report += f"Completed PMs: {total_completed}\n"
            report += f"Completion Rate: {completion_rate:.1f}%\n"
            report += f"Remaining PMs: {total_scheduled - total_completed}\n\n"
            
            # Performance status
            if completion_rate >= 95:
                status = "EXCELLENT"
            elif completion_rate >= 85:
                status = "GOOD"
            elif completion_rate >= 75:
                status = "SATISFACTORY"
            else:
                status = "NEEDS IMPROVEMENT"
            
            report += f"PERFORMANCE STATUS: {status}\n\n"
            
            report += "TECHNICIAN PERFORMANCE:\n"
            report += f"{'Technician':<20} {'Assigned':<10} {'Completed':<10} {'Rate':<8} {'Avg Hours':<10}\n"
            report += "-" * 70 + "\n"
            
            # Clear and update technician performance tree
            for item in self.tech_performance_tree.get_children():
                self.tech_performance_tree.delete(item)
            
            for tech_data in tech_performance:
                technician, assigned, completed, avg_hours = tech_data
                tech_rate = (completed / assigned * 100) if assigned > 0 else 0
                avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "N/A"
                
                report += f"{technician:<20} {assigned:<10} {completed:<10} {tech_rate:<7.1f}% {avg_hours_display:<10}\n"
                
                # Add to tree
                self.tech_performance_tree.insert('', 'end', values=(
                    technician, assigned, completed, f"{tech_rate:.1f}%", avg_hours_display
                ))
            
            # Add PM type breakdown
            cursor.execute('''
                SELECT pm_type, 
                       COUNT(*) as scheduled,
                       COUNT(CASE WHEN status = 'Completed' THEN 1 END) as completed
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
                GROUP BY pm_type
            ''', (week_start.strftime('%Y-%m-%d'),))
            
            pm_types = cursor.fetchall()
            
            if pm_types:
                report += "\nPM TYPE BREAKDOWN:\n"
                report += f"{'PM Type':<15} {'Scheduled':<10} {'Completed':<10} {'Rate':<8}\n"
                report += "-" * 45 + "\n"
                
                for pm_type, scheduled, completed in pm_types:
                    pm_rate = (completed / scheduled * 100) if scheduled > 0 else 0
                    report += f"{pm_type:<15} {scheduled:<10} {completed:<10} {pm_rate:<7.1f}%\n"
            
            # Display report
            self.weekly_report_text.delete('1.0', 'end')
            self.weekly_report_text.insert('end', report)
            
            # Save report to database
            cursor.execute('''
                INSERT OR REPLACE INTO weekly_reports 
                (week_start_date, total_scheduled, total_completed, completion_rate, 
                 technician_performance, report_data)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                week_start.strftime('%Y-%m-%d'),
                total_scheduled,
                total_completed,
                completion_rate,
                json.dumps(tech_performance),
                report
            ))
            
            self.conn.commit()
            self.update_status(f"Weekly report generated - {completion_rate:.1f}% completion rate")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate weekly report: {str(e)}")
    
    def generate_monthly_report(self):
        """Generate monthly PM performance report"""
        try:
            current_date = datetime.now()
            month_start = current_date.replace(day=1)
            
            cursor = self.conn.cursor()
            
            # Get monthly statistics from weekly reports
            cursor.execute('''
                SELECT week_start_date, total_scheduled, total_completed, completion_rate
                FROM weekly_reports 
                WHERE week_start_date >= DATE(?, 'start of month')
                ORDER BY week_start_date
            ''', (current_date.strftime('%Y-%m-%d'),))
            
            weekly_data = cursor.fetchall()
            
            # Get monthly PM completions
            cursor.execute('''
                SELECT 
                    pm_type,
                    COUNT(*) as total_completed,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions 
                WHERE completion_date >= DATE(?, 'start of month')
                GROUP BY pm_type
            ''', (current_date.strftime('%Y-%m-%d'),))
            
            monthly_completions = cursor.fetchall()
            
            # Generate monthly report
            report = f"MONTHLY PM PERFORMANCE REPORT\n"
            report += f"Month: {current_date.strftime('%B %Y')}\n"
            report += "=" * 80 + "\n\n"
            
            if weekly_data:
                total_scheduled = sum(row[1] for row in weekly_data)
                total_completed = sum(row[2] for row in weekly_data)
                avg_completion_rate = sum(row[3] for row in weekly_data) / len(weekly_data)
                
                report += f"MONTHLY SUMMARY:\n"
                report += f"Total Weeks Reported: {len(weekly_data)}\n"
                report += f"Total PMs Scheduled: {total_scheduled}\n"
                report += f"Total PMs Completed: {total_completed}\n"
                report += f"Average Completion Rate: {avg_completion_rate:.1f}%\n"
                report += f"Monthly Target ({len(weekly_data)} weeks × {self.weekly_pm_target}): {len(weekly_data) * self.weekly_pm_target}\n\n"
                
                report += "WEEKLY BREAKDOWN:\n"
                report += f"{'Week Starting':<15} {'Scheduled':<10} {'Completed':<10} {'Rate':<8}\n"
                report += "-" * 45 + "\n"
                
                for week_start, scheduled, completed, rate in weekly_data:
                    report += f"{week_start:<15} {scheduled:<10} {completed:<10} {rate:<7.1f}%\n"
            
            if monthly_completions:
                report += "\nPM TYPE PERFORMANCE (Month):\n"
                report += f"{'PM Type':<15} {'Completed':<10} {'Avg Hours':<10}\n"
                report += "-" * 37 + "\n"
                
                for pm_type, completed, avg_hours in monthly_completions:
                    report += f"{pm_type:<15} {completed:<10} {avg_hours:<9.1f}h\n"
            
            # Display report
            self.weekly_report_text.delete('1.0', 'end')
            self.weekly_report_text.insert('end', report)
            
            self.update_status("Monthly report generated")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate monthly report: {str(e)}")
    
    def export_reports(self):
        """Export reports to file"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"PM_Reports_{timestamp}.txt"
            
            content = self.weekly_report_text.get('1.0', 'end-1c')
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            
            messagebox.showinfo("Success", f"Reports exported to: {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export reports: {str(e)}")
    
    
    
    
    def create_cm_dialog(self):
        """Enhanced CM dialog - auto-fills technician name for technician users"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Corrective Maintenance")
        dialog.geometry("600x550")
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Generate CM number
        cm_number = f"CM-{datetime.now().strftime('%Y%m%d')}-{random.randint(1000, 9999)}"
    
        # Form fields
        row = 0
    
        # CM Number
        ttk.Label(dialog, text="CM Number:").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        cm_number_var = tk.StringVar(value=cm_number)
        ttk.Entry(dialog, textvariable=cm_number_var, width=20, state='readonly').grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Equipment
        ttk.Label(dialog, text="BFM Equipment No:").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        bfm_var = tk.StringVar()
        bfm_combo = ttk.Combobox(dialog, textvariable=bfm_var, width=25)
        bfm_combo.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        
        # Populate equipment list
        cursor = self.conn.cursor()
        cursor.execute('SELECT bfm_equipment_no FROM equipment ORDER BY bfm_equipment_no')
        equipment_list = [row[0] for row in cursor.fetchall()]
        bfm_combo['values'] = equipment_list
        row += 1
        
        # CM Date field
        ttk.Label(dialog, text="CM Date (YYYY-MM-DD):").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        cm_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        cm_date_entry = ttk.Entry(dialog, textvariable=cm_date_var, width=15)
        cm_date_entry.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        
        # Add helper text for date format
        ttk.Label(dialog, text="(Examples: 2025-08-04, 2025-12-15)", 
                font=('Arial', 8), foreground='gray').grid(row=row, column=2, sticky='w', padx=5, pady=5)
        row += 1
    
        # Description
        ttk.Label(dialog, text="Description:").grid(row=row, column=0, sticky='nw', padx=10, pady=5)
        description_text = tk.Text(dialog, width=40, height=4)
        description_text.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Priority
        ttk.Label(dialog, text="Priority:").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        priority_var = tk.StringVar(value='Medium')
        priority_combo = ttk.Combobox(dialog, textvariable=priority_var, 
                                    values=['Low', 'Medium', 'High', 'Emergency'], width=15)
        priority_combo.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Assigned Technician - auto-fill for technicians
        ttk.Label(dialog, text="Assigned Technician:").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        assigned_var = tk.StringVar()
    
        if self.current_user_role == 'Technician':
            # Auto-assign to current technician and make read-only
            assigned_var.set(self.user_name)
            assigned_entry = ttk.Entry(dialog, textvariable=assigned_var, width=20, state='readonly')
            assigned_entry.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        else:
            # Manager can assign to anyone
            assigned_combo = ttk.Combobox(dialog, textvariable=assigned_var, 
                                        values=self.technicians, width=20)
            assigned_combo.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
    
        # Rest of the validation and save logic remains the same...
        def validate_and_save_cm():
            """Validate the CM date format and save"""
            try:
                # Validate the date format
                cm_date_input = cm_date_var.get().strip()
            
                if not cm_date_input:
                    messagebox.showerror("Error", "Please enter a CM date")
                    return
            
                # Try to parse the date to validate format
                try:
                    parsed_date = datetime.strptime(cm_date_input, '%Y-%m-%d')
                
                    if parsed_date > datetime.now() + timedelta(days=1):
                        result = messagebox.askyesno("Future Date Warning", 
                                                f"The CM date '{cm_date_input}' is in the future.\n\n"
                                                f"Are you sure this is correct?")
                        if not result:
                            return
                
                    if parsed_date < datetime.now() - timedelta(days=365):
                        result = messagebox.askyesno("Old Date Warning", 
                                                f"The CM date '{cm_date_input}' is more than 1 year ago.\n\n"
                                                f"Are you sure this is correct?")
                        if not result:
                            return
                
                    validated_date = parsed_date.strftime('%Y-%m-%d')
                
                except ValueError:
                    messagebox.showerror("Invalid Date Format", 
                                    f"Please enter the date in YYYY-MM-DD format.\n\n"
                                    f"Examples:\n"
                                    f"• 2025-08-04 (August 4th, 2025)\n"
                                    f"• 2025-12-15 (December 15th, 2025)\n\n"
                                    f"You entered: '{cm_date_input}'")
                    return
            
                # Validate other required fields
                if not bfm_var.get():
                    messagebox.showerror("Error", "Please select equipment")
                    return
                
                if not description_text.get('1.0', 'end-1c').strip():
                    messagebox.showerror("Error", "Please enter a description")
                    return
            
                # Save to database with the manually entered date
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO corrective_maintenance 
                    (cm_number, bfm_equipment_no, description, priority, assigned_technician, created_date)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    cm_number_var.get(),
                    bfm_var.get(),
                    description_text.get('1.0', 'end-1c'),
                    priority_var.get(),
                    assigned_var.get(),
                    validated_date
                ))
                self.conn.commit()
                
                messagebox.showinfo("Success", 
                                f"Corrective Maintenance created successfully!\n\n"
                                f"CM Number: {cm_number_var.get()}\n"
                                f"CM Date: {validated_date}\n"
                                f"Equipment: {bfm_var.get()}\n"
                                f"Assigned to: {assigned_var.get()}")
                dialog.destroy()
                self.load_corrective_maintenance()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create CM: {str(e)}")
    
        def set_date_to_today():
            cm_date_var.set(datetime.now().strftime('%Y-%m-%d'))
    
        def set_date_to_yesterday():
            yesterday = datetime.now() - timedelta(days=1)
            cm_date_var.set(yesterday.strftime('%Y-%m-%d'))
    
        # Date helper buttons
        date_buttons_frame = ttk.Frame(dialog)
        date_buttons_frame.grid(row=2, column=2, sticky='w', padx=5, pady=5)
    
        ttk.Button(date_buttons_frame, text="Today", command=set_date_to_today).pack(side='top', pady=1)
        ttk.Button(date_buttons_frame, text="Yesterday", command=set_date_to_yesterday).pack(side='top', pady=1)
    
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=row, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Create CM", command=validate_and_save_cm).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='left', padx=5)
        
    
    
    
    
    def edit_cm_dialog(self):
        """Edit existing Corrective Maintenance with full functionality"""
        selected = self.cm_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a CM to edit")
            return

        # Get selected CM data
        item = self.cm_tree.item(selected[0])
        cm_number = item['values'][0]

        # Fetch full CM data from database
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT cm_number, bfm_equipment_no, description, priority, assigned_technician, 
                status, created_date, completion_date, labor_hours, notes, root_cause, corrective_action
            FROM corrective_maintenance 
            WHERE cm_number = ?
        ''', (cm_number,))

        cm_data = cursor.fetchone()
        if not cm_data:
            messagebox.showerror("Error", "CM not found in database")
            return

        # Extract CM data
        (orig_cm_number, orig_bfm_no, orig_description, orig_priority, orig_assigned, 
        orig_status, orig_created, orig_completion, orig_hours, orig_notes, 
        orig_root_cause, orig_corrective_action) = cm_data

        # Create edit dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Corrective Maintenance - {cm_number}")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()

        # Main container with scrollbar
        main_canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # CM Information (header)
        header_frame = ttk.LabelFrame(scrollable_frame, text="CM Information", padding=10)
        header_frame.pack(fill='x', padx=10, pady=5)

        row = 0

        # CM Number (read-only)
        ttk.Label(header_frame, text="CM Number:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Label(header_frame, text=orig_cm_number, font=('Arial', 10, 'bold')).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1

        # Equipment (editable)
        ttk.Label(header_frame, text="BFM Equipment No:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        bfm_var = tk.StringVar(value=orig_bfm_no or '')
        bfm_combo = ttk.Combobox(header_frame, textvariable=bfm_var, width=25)
        bfm_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)

        # Populate equipment list
        cursor.execute('SELECT bfm_equipment_no FROM equipment ORDER BY bfm_equipment_no')
        equipment_list = [row[0] for row in cursor.fetchall()]
        bfm_combo['values'] = equipment_list
        row += 1

        # Priority (editable)
        ttk.Label(header_frame, text="Priority:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        priority_var = tk.StringVar(value=orig_priority or 'Medium')
        priority_combo = ttk.Combobox(header_frame, textvariable=priority_var, 
                                values=['Low', 'Medium', 'High', 'Emergency'], width=15)
        priority_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1

        # Assigned Technician (editable)
        ttk.Label(header_frame, text="Assigned Technician:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        assigned_var = tk.StringVar(value=orig_assigned or '')
        assigned_combo = ttk.Combobox(header_frame, textvariable=assigned_var, 
                                values=self.technicians, width=20)
        assigned_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1

        # Status (editable)
        ttk.Label(header_frame, text="Status:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        status_var = tk.StringVar(value=orig_status or 'Open')
        status_combo = ttk.Combobox(header_frame, textvariable=status_var, 
                              values=['Open', 'In Progress', 'Completed', 'Cancelled'], width=15)
        status_combo.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1

        # Description (editable)
        desc_frame = ttk.LabelFrame(scrollable_frame, text="Description", padding=10)
        desc_frame.pack(fill='x', padx=10, pady=5)

        description_text = tk.Text(desc_frame, width=60, height=4)
        description_text.pack(fill='x', padx=5, pady=5)
        description_text.insert('1.0', orig_description or '')

        # Completion Information (if completed)
        completion_frame = ttk.LabelFrame(scrollable_frame, text="Completion Information", padding=10)
        completion_frame.pack(fill='x', padx=10, pady=5)

        comp_row = 0

        # Labor Hours
        ttk.Label(completion_frame, text="Labor Hours:").grid(row=comp_row, column=0, sticky='w', padx=5, pady=5)
        labor_hours_var = tk.StringVar(value=str(orig_hours or ''))
        ttk.Entry(completion_frame, textvariable=labor_hours_var, width=10).grid(row=comp_row, column=1, sticky='w', padx=5, pady=5)
        comp_row += 1

        # Completion Date
        ttk.Label(completion_frame, text="Completion Date:").grid(row=comp_row, column=0, sticky='w', padx=5, pady=5)
        completion_date_var = tk.StringVar(value=orig_completion or '')
        ttk.Entry(completion_frame, textvariable=completion_date_var, width=15).grid(row=comp_row, column=1, sticky='w', padx=5, pady=5)
        comp_row += 1

        # Notes
        notes_frame = ttk.LabelFrame(scrollable_frame, text="Notes", padding=10)
        notes_frame.pack(fill='x', padx=10, pady=5)

        notes_text = tk.Text(notes_frame, width=60, height=4)
        notes_text.pack(fill='x', padx=5, pady=5)
        notes_text.insert('1.0', orig_notes or '')

        # Root Cause
        root_cause_frame = ttk.LabelFrame(scrollable_frame, text="Root Cause Analysis", padding=10)
        root_cause_frame.pack(fill='x', padx=10, pady=5)

        root_cause_text = tk.Text(root_cause_frame, width=60, height=3)
        root_cause_text.pack(fill='x', padx=5, pady=5)
        root_cause_text.insert('1.0', orig_root_cause or '')

        # Corrective Action
        corrective_action_frame = ttk.LabelFrame(scrollable_frame, text="Corrective Action", padding=10)
        corrective_action_frame.pack(fill='x', padx=10, pady=5)

        corrective_action_text = tk.Text(corrective_action_frame, width=60, height=3)
        corrective_action_text.pack(fill='x', padx=5, pady=5)
        corrective_action_text.insert('1.0', orig_corrective_action or '')

        def save_changes():
            try:
                # Validate inputs
                if not description_text.get('1.0', 'end-1c').strip():
                    messagebox.showerror("Error", "Please enter a description")
                    return

                # Update database
                cursor = self.conn.cursor()
                cursor.execute('''
                    UPDATE corrective_maintenance SET
                    bfm_equipment_no = ?,
                    description = ?,
                    priority = ?,
                    assigned_technician = ?,
                    status = ?,
                    labor_hours = ?,
                    completion_date = ?,
                    notes = ?,
                    root_cause = ?,
                    corrective_action = ?
                    WHERE cm_number = ?
                ''', (
                    bfm_var.get(),
                    description_text.get('1.0', 'end-1c'),
                    priority_var.get(),
                    assigned_var.get(),
                    status_var.get(),
                    float(labor_hours_var.get() or 0),
                    completion_date_var.get() if completion_date_var.get() else None,
                    notes_text.get('1.0', 'end-1c'),
                    root_cause_text.get('1.0', 'end-1c'),
                    corrective_action_text.get('1.0', 'end-1c'),
                    orig_cm_number
                ))

                self.conn.commit()
                messagebox.showinfo("Success", f"CM {orig_cm_number} updated successfully!")
                dialog.destroy()
                self.load_corrective_maintenance()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to update CM: {str(e)}")

        def delete_cm():
            result = messagebox.askyesno("Confirm Delete", 
                                    f"Delete CM {orig_cm_number}?\n\n"
                                    f"This action cannot be undone.")
            if result:
                try:
                    cursor = self.conn.cursor()
                    cursor.execute('DELETE FROM corrective_maintenance WHERE cm_number = ?', (orig_cm_number,))
                    self.conn.commit()
                    messagebox.showinfo("Success", f"CM {orig_cm_number} deleted successfully!")
                    dialog.destroy()
                    self.load_corrective_maintenance()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete CM: {str(e)}")

        # Buttons frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Save Changes", command=save_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Delete CM", command=delete_cm).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)

        # Pack the canvas and scrollbar
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Update scroll region
        scrollable_frame.update_idletasks()
        main_canvas.configure(scrollregion=main_canvas.bbox("all"))
    
    def complete_cm_dialog(self):
        """Dialog to complete Corrective Maintenance"""
        selected = self.cm_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a CM to complete")
            return
        
        # Get selected CM data
        item = self.cm_tree.item(selected[0])
        cm_number = item['values'][0]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Complete Corrective Maintenance")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Completion form
        row = 0
        
        ttk.Label(dialog, text=f"Completing CM: {cm_number}").grid(row=row, column=0, columnspan=2, pady=10)
        row += 1
        
        # Labor hours
        ttk.Label(dialog, text="Labor Hours:").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        labor_hours_var = tk.StringVar(value="0")
        ttk.Entry(dialog, textvariable=labor_hours_var, width=10).grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
        
        # Completion notes
        ttk.Label(dialog, text="Completion Notes:").grid(row=row, column=0, sticky='nw', padx=10, pady=5)
        notes_text = tk.Text(dialog, width=40, height=6)
        notes_text.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
        
        # Root cause
        ttk.Label(dialog, text="Root Cause:").grid(row=row, column=0, sticky='nw', padx=10, pady=5)
        root_cause_text = tk.Text(dialog, width=40, height=3)
        root_cause_text.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
        
        # Corrective action
        ttk.Label(dialog, text="Corrective Action:").grid(row=row, column=0, sticky='nw', padx=10, pady=5)
        corrective_action_text = tk.Text(dialog, width=40, height=3)
        corrective_action_text.grid(row=row, column=1, sticky='w', padx=10, pady=5)
        row += 1
        
        def complete_cm():
            try:
                cursor = self.conn.cursor()
                cursor.execute('''
                    UPDATE corrective_maintenance SET
                    status = 'Completed',
                    completion_date = ?,
                    labor_hours = ?,
                    notes = ?,
                    root_cause = ?,
                    corrective_action = ?
                    WHERE cm_number = ?
                ''', (
                    datetime.now().strftime('%Y-%m-%d'),
                    float(labor_hours_var.get() or 0),
                    notes_text.get('1.0', 'end-1c'),
                    root_cause_text.get('1.0', 'end-1c'),
                    corrective_action_text.get('1.0', 'end-1c'),
                    cm_number
                ))
                self.conn.commit()
                messagebox.showinfo("Success", "CM completed successfully!")
                dialog.destroy()
                self.load_corrective_maintenance()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to complete CM: {str(e)}")
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=row, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Complete CM", command=complete_cm).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='left', padx=5)
    
    def load_corrective_maintenance(self):
        """Load corrective maintenance data"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT cm_number, bfm_equipment_no, description, priority, 
                       assigned_technician, status, created_date
                FROM corrective_maintenance 
                ORDER BY created_date DESC
            ''')
            
            # Clear existing items
            for item in self.cm_tree.get_children():
                self.cm_tree.delete(item)
            
            # Add CM records
            for cm in cursor.fetchall():
                cm_number, bfm_no, description, priority, assigned, status, created = cm
                # Truncate description for display
                display_desc = (description[:47] + '...') if len(description) > 50 else description
                self.cm_tree.insert('', 'end', values=(
                    cm_number, bfm_no, display_desc, priority, assigned, status, created
                ))
                
        except Exception as e:
            print(f"Error loading corrective maintenance: {e}")
    
    def refresh_analytics_dashboard(self):
        """Refresh analytics dashboard with current data"""
        try:
            cursor = self.conn.cursor()
            
            # Generate comprehensive analytics
            analytics = "AIT CMMS ANALYTICS DASHBOARD\n"
            analytics += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            analytics += "=" * 80 + "\n\n"
            
            # Equipment statistics
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Active"')
            active_equipment = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE monthly_pm = 1')
            monthly_pm_equipment = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE six_month_pm = 1')
            six_month_pm_equipment = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE annual_pm = 1')
            annual_pm_equipment = cursor.fetchone()[0]
            
            analytics += "EQUIPMENT OVERVIEW:\n"
            analytics += f"Total Active Equipment: {active_equipment}\n"
            analytics += f"Equipment with Monthly PM: {monthly_pm_equipment}\n"
            analytics += f"Equipment with Six Month PM: {six_month_pm_equipment}\n"
            analytics += f"Equipment with Annual PM: {annual_pm_equipment}\n\n"
            
            # PM completion statistics (last 30 days)
            cursor.execute('''
                SELECT COUNT(*) FROM pm_completions 
                WHERE completion_date >= DATE('now', '-30 days')
            ''')
            recent_completions = cursor.fetchone()[0]
            
            cursor.execute('''
                SELECT pm_type, COUNT(*) 
                FROM pm_completions 
                WHERE completion_date >= DATE('now', '-30 days')
                GROUP BY pm_type
            ''')
            pm_type_stats = cursor.fetchall()
            
            analytics += "PM COMPLETION STATISTICS (Last 30 Days):\n"
            analytics += f"Total PM Completions: {recent_completions}\n"
            for pm_type, count in pm_type_stats:
                analytics += f"{pm_type} PMs: {count}\n"
            analytics += "\n"
            
            # Technician performance (last 30 days)
            cursor.execute('''
                SELECT technician_name, 
                       COUNT(*) as completed_pms,
                       AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions 
                WHERE completion_date >= DATE('now', '-30 days')
                GROUP BY technician_name
                ORDER BY completed_pms DESC
            ''')
            tech_stats = cursor.fetchall()
            
            analytics += "TECHNICIAN PERFORMANCE (Last 30 Days):\n"
            analytics += f"{'Technician':<20} {'Completed PMs':<15} {'Avg Hours':<10}\n"
            analytics += "-" * 47 + "\n"
            for tech, completed, avg_hours in tech_stats:
                analytics += f"{tech:<20} {completed:<15} {avg_hours:<9.1f}h\n"
            analytics += "\n"
            
            # CM statistics
            cursor.execute('SELECT COUNT(*) FROM corrective_maintenance WHERE status = "Open"')
            open_cms = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM corrective_maintenance WHERE status = "Completed"')
            completed_cms = cursor.fetchone()[0]
            
            analytics += "CORRECTIVE MAINTENANCE:\n"
            analytics += f"Open CMs: {open_cms}\n"
            analytics += f"Completed CMs: {completed_cms}\n\n"
            
            # Current week performance
            current_week_start = self.get_week_start(datetime.now()).strftime('%Y-%m-%d')
            cursor.execute('''
                SELECT 
                    COUNT(*) as scheduled,
                    COUNT(CASE WHEN status = 'Completed' THEN 1 END) as completed
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
            ''', (current_week_start,))
            
            week_scheduled, week_completed = cursor.fetchone()
            week_rate = (week_completed / week_scheduled * 100) if week_scheduled > 0 else 0
            
            analytics += "CURRENT WEEK PERFORMANCE:\n"
            analytics += f"Scheduled PMs: {week_scheduled}\n"
            analytics += f"Completed PMs: {week_completed}\n"
            analytics += f"Completion Rate: {week_rate:.1f}%\n\n"
            
            # Display analytics
            self.analytics_text.delete('1.0', 'end')
            self.analytics_text.insert('end', analytics)
            
        except Exception as e:
            print(f"Error refreshing analytics dashboard: {e}")
    
    def show_equipment_analytics(self):
        """Show comprehensive equipment analytics in a new dialog window"""
        try:
            # Create analytics dialog
            analytics_dialog = tk.Toplevel(self.root)
            analytics_dialog.title("Equipment Analytics Dashboard")
            analytics_dialog.geometry("1200x800")
            analytics_dialog.transient(self.root)
            analytics_dialog.grab_set()
        
            # Create notebook for different analytics tabs
            analytics_notebook = ttk.Notebook(analytics_dialog)
            analytics_notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
            # Tab 1: Equipment Overview
            overview_frame = ttk.Frame(analytics_notebook)
            analytics_notebook.add(overview_frame, text="Equipment Overview")
        
            # Tab 2: PM Performance Analysis
            pm_performance_frame = ttk.Frame(analytics_notebook)
            analytics_notebook.add(pm_performance_frame, text="PM Performance")
        
            # Tab 3: Location Analysis
            location_frame = ttk.Frame(analytics_notebook)
            analytics_notebook.add(location_frame, text="Location Analysis")
        
            # Tab 4: Technician Workload
            technician_frame = ttk.Frame(analytics_notebook)
            analytics_notebook.add(technician_frame, text="Technician Analysis")
        
            # Generate analytics for each tab
            self.generate_equipment_overview(overview_frame)
            self.generate_pm_performance_analysis(pm_performance_frame)
            self.generate_location_analysis(location_frame)
            self.generate_technician_analysis(technician_frame)
        
            # Add export button
            export_frame = ttk.Frame(analytics_dialog)
            export_frame.pack(side='bottom', fill='x', padx=10, pady=5)
        
            ttk.Button(export_frame, text="Export All Analytics to PDF", 
                    command=lambda: self.export_equipment_analytics_pdf(analytics_dialog)).pack(side='right', padx=5)
            ttk.Button(export_frame, text="Close", 
                    command=analytics_dialog.destroy).pack(side='right', padx=5)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate equipment analytics: {str(e)}")

    def generate_equipment_overview(self, parent_frame):
        """Generate equipment overview analytics"""
        try:
            cursor = self.conn.cursor()
        
            # Create scrollable text area
            text_frame = ttk.Frame(parent_frame)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            overview_text = tk.Text(text_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=overview_text.yview)
            overview_text.configure(yscrollcommand=scrollbar.set)
        
            overview_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
        
            # Generate analytics content
            analytics = "EQUIPMENT ANALYTICS OVERVIEW\n"
            analytics += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            analytics += "=" * 80 + "\n\n"
        
            # Basic equipment statistics
            cursor.execute('SELECT COUNT(*) FROM equipment')
            total_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Active"')
            active_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Missing"')
            missing_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Run to Failure"')
            rtf_equipment = cursor.fetchone()[0]
        
            analytics += "EQUIPMENT STATUS SUMMARY:\n"
            analytics += f"Total Equipment: {total_equipment}\n"
            analytics += f"Active Equipment: {active_equipment} ({active_equipment/total_equipment*100:.1f}%)\n"
            analytics += f"Missing Equipment: {missing_equipment} ({missing_equipment/total_equipment*100:.1f}%)\n"
            analytics += f"Run to Failure: {rtf_equipment} ({rtf_equipment/total_equipment*100:.1f}%)\n\n"
        
            # PM Type Distribution
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE monthly_pm = 1')
            monthly_pm_count = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE six_month_pm = 1')
            six_month_pm_count = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE annual_pm = 1')
            annual_pm_count = cursor.fetchone()[0]
            
            analytics += "PM TYPE REQUIREMENTS:\n"
            analytics += f"Monthly PM Required: {monthly_pm_count} assets ({monthly_pm_count/total_equipment*100:.1f}%)\n"
            analytics += f"Six Month PM Required: {six_month_pm_count} assets ({six_month_pm_count/total_equipment*100:.1f}%)\n"
            analytics += f"Annual PM Required: {annual_pm_count} assets ({annual_pm_count/total_equipment*100:.1f}%)\n\n"
        
            # Location distribution
            cursor.execute('''
                SELECT location, COUNT(*) as count 
                FROM equipment 
                WHERE location IS NOT NULL AND location != ''
                GROUP BY location 
                ORDER BY count DESC 
                LIMIT 10
            ''')
            location_stats = cursor.fetchall()
        
            if location_stats:
                analytics += "TOP 10 EQUIPMENT LOCATIONS:\n"
                analytics += f"{'Location':<20} {'Count':<10} {'Percentage':<12}\n"
                analytics += "-" * 45 + "\n"
            
                for location, count in location_stats:
                    percentage = count / total_equipment * 100
                    analytics += f"{location:<20} {count:<10} {percentage:<11.1f}%\n"
                analytics += "\n"
        
            # Equipment without PM completions (never serviced)
            cursor.execute('''
                SELECT e.bfm_equipment_no, e.description, e.location
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no
                WHERE pc.bfm_equipment_no IS NULL 
                AND e.status = 'Active'
                ORDER BY e.bfm_equipment_no
                LIMIT 20
            ''')
            never_serviced = cursor.fetchall()
        
            if never_serviced:
                analytics += f"EQUIPMENT NEVER SERVICED ({len(never_serviced)} items shown, may be more):\n"
                analytics += f"{'BFM Number':<15} {'Description':<30} {'Location':<15}\n"
                analytics += "-" * 62 + "\n"
            
                for bfm_no, description, location in never_serviced:
                    desc_short = (description[:27] + '...') if description and len(description) > 27 else (description or 'N/A')
                    loc_short = (location[:12] + '...') if location and len(location) > 12 else (location or 'N/A')
                    analytics += f"{bfm_no:<15} {desc_short:<30} {loc_short:<15}\n"
                analytics += "\n"
        
            # Equipment age analysis (based on creation date if available)
            cursor.execute('''
                SELECT 
                    CASE 
                        WHEN created_date >= DATE('now', '-30 days') THEN 'Last 30 days'
                        WHEN created_date >= DATE('now', '-90 days') THEN 'Last 90 days'
                        WHEN created_date >= DATE('now', '-180 days') THEN 'Last 6 months'
                        WHEN created_date >= DATE('now', '-365 days') THEN 'Last year'
                        ELSE 'Over 1 year'
                    END as age_category,
                    COUNT(*) as count
                FROM equipment
                WHERE created_date IS NOT NULL
                GROUP BY age_category
                ORDER BY 
                    CASE 
                        WHEN age_category = 'Last 30 days' THEN 1
                        WHEN age_category = 'Last 90 days' THEN 2
                        WHEN age_category = 'Last 6 months' THEN 3
                        WHEN age_category = 'Last year' THEN 4
                        ELSE 5
                    END
            ''')
            age_stats = cursor.fetchall()
        
            if age_stats:
                analytics += "EQUIPMENT AGE DISTRIBUTION (by creation date):\n"
                analytics += f"{'Age Category':<20} {'Count':<10} {'Percentage':<12}\n"
                analytics += "-" * 45 + "\n"
            
                total_with_dates = sum(count for _, count in age_stats)
                for age_category, count in age_stats:
                    percentage = count / total_with_dates * 100
                    analytics += f"{age_category:<20} {count:<10} {percentage:<11.1f}%\n"
                analytics += "\n"
        
            # Display analytics
            overview_text.insert('end', analytics)
            overview_text.config(state='disabled')
        
        except Exception as e:
            print(f"Error generating equipment overview: {e}")

    def generate_pm_performance_analysis(self, parent_frame):
        """Generate PM performance analytics"""
        try:
            cursor = self.conn.cursor()
            
            text_frame = ttk.Frame(parent_frame)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            pm_text = tk.Text(text_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=pm_text.yview)
            pm_text.configure(yscrollcommand=scrollbar.set)
        
            pm_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
        
            analytics = "PM PERFORMANCE ANALYTICS\n"
            analytics += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            analytics += "=" * 80 + "\n\n"
        
            # PM completion statistics by type
            cursor.execute('''
                SELECT pm_type, COUNT(*) as count, 
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours,
                    MIN(completion_date) as first_completion,
                    MAX(completion_date) as last_completion
                FROM pm_completions 
                GROUP BY pm_type 
                ORDER BY count DESC
            ''')
            pm_type_stats = cursor.fetchall()
        
            if pm_type_stats:
                analytics += "PM COMPLETION STATISTICS BY TYPE:\n"
                analytics += f"{'PM Type':<15} {'Count':<10} {'Avg Hours':<12} {'Date Range':<25}\n"
                analytics += "-" * 65 + "\n"
            
                for pm_type, count, avg_hours, first_date, last_date in pm_type_stats:
                    avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "N/A"
                    date_range = f"{first_date} to {last_date}" if first_date and last_date else "N/A"
                    analytics += f"{pm_type:<15} {count:<10} {avg_hours_display:<12} {date_range:<25}\n"
                analytics += "\n"
        
            # Monthly completion trends (last 12 months)
            cursor.execute('''
                SELECT 
                    strftime('%Y-%m', completion_date) as month,
                    COUNT(*) as completions,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY strftime('%Y-%m', completion_date)
                ORDER BY month DESC
            ''')
            monthly_trends = cursor.fetchall()
        
            if monthly_trends:
                analytics += "MONTHLY PM COMPLETION TRENDS (Last 12 months):\n"
                analytics += f"{'Month':<10} {'Completions':<12} {'Avg Hours':<12}\n"
                analytics += "-" * 36 + "\n"
            
                for month, completions, avg_hours in monthly_trends:
                    avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    analytics += f"{month:<10} {completions:<12} {avg_hours_display:<12}\n"
                analytics += "\n"
        
            # Equipment with overdue PMs
            current_date = datetime.now()
        
            cursor.execute('''
                SELECT e.bfm_equipment_no, e.description, e.location,
                    e.last_monthly_pm, e.last_annual_pm,
                    CASE 
                        WHEN e.last_monthly_pm IS NULL OR DATE(e.last_monthly_pm, '+30 days') < DATE('now') THEN 'Monthly Overdue'
                        WHEN e.last_annual_pm IS NULL OR DATE(e.last_annual_pm, '+365 days') < DATE('now') THEN 'Annual Overdue'
                        ELSE 'Current'
                    END as pm_status
                FROM equipment e
                WHERE e.status = 'Active' 
                AND (
                    (e.monthly_pm = 1 AND (e.last_monthly_pm IS NULL OR DATE(e.last_monthly_pm, '+30 days') < DATE('now')))
                    OR
                    (e.annual_pm = 1 AND (e.last_annual_pm IS NULL OR DATE(e.last_annual_pm, '+365 days') < DATE('now')))
                )
                ORDER BY e.bfm_equipment_no
                LIMIT 25
            ''')
            overdue_equipment = cursor.fetchall()
        
            if overdue_equipment:
                analytics += f"OVERDUE PM EQUIPMENT ({len(overdue_equipment)} items shown):\n"
                analytics += f"{'BFM Number':<15} {'Description':<25} {'Location':<12} {'Status':<15}\n"
                analytics += "-" * 70 + "\n"
                
                for bfm_no, description, location, pm_status in overdue_equipment:
                    desc_short = (description[:22] + '...') if description and len(description) > 22 else (description or 'N/A')
                    loc_short = (location[:9] + '...') if location and len(location) > 9 else (location or 'N/A')
                    analytics += f"{bfm_no:<15} {desc_short:<25} {loc_short:<12} {pm_status:<15}\n"
                analytics += "\n"
        
            # PM frequency analysis
            cursor.execute('''
                SELECT e.bfm_equipment_no, COUNT(pc.id) as pm_count,
                    MIN(pc.completion_date) as first_pm,
                    MAX(pc.completion_date) as last_pm,
                    AVG(pc.labor_hours + pc.labor_minutes/60.0) as avg_hours
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no
                WHERE e.status = 'Active'
                GROUP BY e.bfm_equipment_no
                HAVING pm_count > 0
                ORDER BY pm_count DESC
                LIMIT 15
            ''')
            high_maintenance_equipment = cursor.fetchall()
        
            if high_maintenance_equipment:
                analytics += "TOP 15 MOST SERVICED EQUIPMENT:\n"
                analytics += f"{'BFM Number':<15} {'PM Count':<10} {'First PM':<12} {'Last PM':<12} {'Avg Hours':<10}\n"
                analytics += "-" * 62 + "\n"
            
                for bfm_no, pm_count, first_pm, last_pm, avg_hours in high_maintenance_equipment:
                    avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    analytics += f"{bfm_no:<15} {pm_count:<10} {first_pm or 'N/A':<12} {last_pm or 'N/A':<12} {avg_hours_display:<10}\n"
                analytics += "\n"
        
            pm_text.insert('end', analytics)
            pm_text.config(state='disabled')
        
        except Exception as e:
            print(f"Error generating PM performance analysis: {e}")

    def generate_location_analysis(self, parent_frame):
        """Generate location-based analytics"""
        try:
            cursor = self.conn.cursor()
        
            text_frame = ttk.Frame(parent_frame)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            location_text = tk.Text(text_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=location_text.yview)
            location_text.configure(yscrollcommand=scrollbar.set)
        
            location_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
        
            analytics = "LOCATION-BASED ANALYTICS\n"
            analytics += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            analytics += "=" * 80 + "\n\n"
        
            # Equipment distribution by location
            cursor.execute('''
                SELECT 
                    COALESCE(location, 'Unknown') as location,
                    COUNT(*) as total_equipment,
                    COUNT(CASE WHEN status = 'Active' THEN 1 END) as active,
                    COUNT(CASE WHEN status = 'Missing' THEN 1 END) as missing,
                    COUNT(CASE WHEN status = 'Run to Failure' THEN 1 END) as rtf
                FROM equipment
                GROUP BY COALESCE(location, 'Unknown')
                ORDER BY total_equipment DESC
            ''')
            location_distribution = cursor.fetchall()
        
            if location_distribution:
                analytics += "EQUIPMENT DISTRIBUTION BY LOCATION:\n"
                analytics += f"{'Location':<20} {'Total':<8} {'Active':<8} {'Missing':<8} {'RTF':<8}\n"
                analytics += "-" * 55 + "\n"
            
                for location, total, active, missing, rtf in location_distribution:
                    loc_display = location[:17] + '...' if len(location) > 17 else location
                    analytics += f"{loc_display:<20} {total:<8} {active:<8} {missing:<8} {rtf:<8}\n"
                analytics += "\n"
        
            # PM completion activity by location
            cursor.execute('''
                SELECT 
                    COALESCE(e.location, 'Unknown') as location,
                    COUNT(pc.id) as total_pms,
                    COUNT(CASE WHEN pc.pm_type = 'Monthly' THEN 1 END) as monthly,
                    COUNT(CASE WHEN pc.pm_type = 'Annual' THEN 1 END) as annual,
                    AVG(pc.labor_hours + pc.labor_minutes/60.0) as avg_hours
                FROM pm_completions pc
                JOIN equipment e ON pc.bfm_equipment_no = e.bfm_equipment_no
                WHERE pc.completion_date >= DATE('now', '-90 days')
                GROUP BY COALESCE(e.location, 'Unknown')
                ORDER BY total_pms DESC
            ''')
            location_pm_activity = cursor.fetchall()
        
            if location_pm_activity:
                analytics += "PM ACTIVITY BY LOCATION (Last 90 days):\n"
                analytics += f"{'Location':<20} {'Total PMs':<10} {'Monthly':<8} {'Annual':<8} {'Avg Hours':<10}\n"
                analytics += "-" * 60 + "\n"
            
                for location, total_pms, monthly, annual, avg_hours in location_pm_activity:
                    loc_display = location[:17] + '...' if len(location) > 17 else location
                    avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    analytics += f"{loc_display:<20} {total_pms:<10} {monthly:<8} {annual:<8} {avg_hours_display:<10}\n"
                analytics += "\n"
        
            # Cannot Find assets by location
            cursor.execute('''
                SELECT 
                    COALESCE(location, 'Unknown') as location,
                    COUNT(*) as missing_count,
                    GROUP_CONCAT(bfm_equipment_no, ', ') as missing_assets
                FROM cannot_find_assets
                WHERE status = 'Missing'
                GROUP BY COALESCE(location, 'Unknown')
                ORDER BY missing_count DESC
            ''')
            missing_by_location = cursor.fetchall()
        
            if missing_by_location:
                analytics += "MISSING ASSETS BY LOCATION:\n"
                analytics += f"{'Location':<20} {'Count':<8} {'Equipment Numbers':<50}\n"
                analytics += "-" * 80 + "\n"
            
                for location, count, assets in missing_by_location:
                    loc_display = location[:17] + '...' if len(location) > 17 else location
                    assets_display = assets[:47] + '...' if assets and len(assets) > 47 else (assets or '')
                    analytics += f"{loc_display:<20} {count:<8} {assets_display:<50}\n"
                analytics += "\n"
        
            # Location efficiency analysis
            cursor.execute('''
                SELECT 
                    COALESCE(e.location, 'Unknown') as location,
                    COUNT(DISTINCT e.bfm_equipment_no) as equipment_count,
                    COUNT(pc.id) as pm_completions,
                    ROUND(CAST(COUNT(pc.id) AS FLOAT) / COUNT(DISTINCT e.bfm_equipment_no), 2) as pms_per_equipment
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no 
                    AND pc.completion_date >= DATE('now', '-365 days')
                WHERE e.status = 'Active'
                GROUP BY COALESCE(e.location, 'Unknown')
                HAVING equipment_count >= 3
                ORDER BY pms_per_equipment DESC
            ''')
            location_efficiency = cursor.fetchall()
        
            if location_efficiency:
                analytics += "LOCATION PM EFFICIENCY (PMs per equipment, last year):\n"
                analytics += f"{'Location':<20} {'Equipment':<10} {'PMs':<8} {'PMs/Equipment':<15}\n"
                analytics += "-" * 55 + "\n"
                
                for location, equipment_count, pm_completions, pms_per_equipment in location_efficiency:
                    loc_display = location[:17] + '...' if len(location) > 17 else location
                    analytics += f"{loc_display:<20} {equipment_count:<10} {pm_completions:<8} {pms_per_equipment:<15}\n"
                analytics += "\n"
        
            location_text.insert('end', analytics)
            location_text.config(state='disabled')
        
        except Exception as e:
            print(f"Error generating location analysis: {e}")

    def generate_technician_analysis(self, parent_frame):
        """Generate technician workload and performance analytics"""
        try:
            cursor = self.conn.cursor()
        
            text_frame = ttk.Frame(parent_frame)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            tech_text = tk.Text(text_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=tech_text.yview)
            tech_text.configure(yscrollcommand=scrollbar.set)
        
            tech_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
        
            analytics = "TECHNICIAN PERFORMANCE ANALYTICS\n"
            analytics += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            analytics += "=" * 80 + "\n\n"
        
            # Overall technician performance
            cursor.execute('''
                SELECT 
                    technician_name,
                    COUNT(*) as total_pms,
                    COUNT(CASE WHEN pm_type = 'Monthly' THEN 1 END) as monthly_pms,
                    COUNT(CASE WHEN pm_type = 'Annual' THEN 1 END) as annual_pms,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours,
                    SUM(labor_hours + labor_minutes/60.0) as total_hours,
                    MIN(completion_date) as first_completion,
                    MAX(completion_date) as last_completion
                FROM pm_completions
                GROUP BY technician_name
                ORDER BY total_pms DESC
            ''')
            technician_performance = cursor.fetchall()
        
            if technician_performance:
                analytics += "OVERALL TECHNICIAN PERFORMANCE:\n"
                analytics += f"{'Technician':<20} {'Total PMs':<10} {'Monthly':<8} {'Annual':<8} {'Avg Hrs':<8} {'Total Hrs':<10}\n"
                analytics += "-" * 75 + "\n"
            
                for tech_data in technician_performance:
                    technician, total_pms, monthly, annual, avg_hours, total_hours, first_date, last_date = tech_data
                    tech_display = technician[:17] + '...' if len(technician) > 17 else technician
                    avg_hours_display = f"{avg_hours:.1f}" if avg_hours else "0.0"
                    total_hours_display = f"{total_hours:.1f}" if total_hours else "0.0"
                
                    analytics += f"{tech_display:<20} {total_pms:<10} {monthly:<8} {annual:<8} {avg_hours_display:<8} {total_hours_display:<10}\n"
                analytics += "\n"
        
            # Recent activity (last 30 days)
            cursor.execute('''
                SELECT 
                    technician_name,
                    COUNT(*) as recent_pms,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours,
                    COUNT(DISTINCT bfm_equipment_no) as unique_equipment
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-30 days')
                GROUP BY technician_name
                ORDER BY recent_pms DESC
            ''')
            recent_activity = cursor.fetchall()
        
            if recent_activity:
                analytics += "RECENT ACTIVITY (Last 30 days):\n"
                analytics += f"{'Technician':<20} {'PMs':<6} {'Avg Hours':<10} {'Unique Equipment':<18}\n"
                analytics += "-" * 56 + "\n"
            
                for technician, recent_pms, avg_hours, unique_equipment in recent_activity:
                    tech_display = technician[:17] + '...' if len(technician) > 17 else technician
                    avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                
                    analytics += f"{tech_display:<20} {recent_pms:<6} {avg_hours_display:<10} {unique_equipment:<18}\n"
                analytics += "\n"
        
            # Cannot Find reports by technician
            cursor.execute('''
                SELECT 
                    technician_name,
                    COUNT(*) as cannot_find_count,
                    COUNT(CASE WHEN report_date >= DATE('now', '-30 days') THEN 1 END) as recent_cf
                FROM cannot_find_assets
                WHERE status = 'Missing'
                GROUP BY technician_name
                ORDER BY cannot_find_count DESC
            ''')
            cannot_find_by_tech = cursor.fetchall()
        
            if cannot_find_by_tech:
                analytics += "CANNOT FIND REPORTS BY TECHNICIAN:\n"
                analytics += f"{'Technician':<20} {'Total CF':<10} {'Recent (30d)':<15}\n"
                analytics += "-" * 47 + "\n"
            
                for technician, total_cf, recent_cf in cannot_find_by_tech:
                    tech_display = technician[:17] + '...' if len(technician) > 17 else technician
                    analytics += f"{tech_display:<20} {total_cf:<10} {recent_cf:<15}\n"
                analytics += "\n"
        
            # Workload distribution analysis
            cursor.execute('''
                SELECT 
                    technician_name,
                    strftime('%Y-%m', completion_date) as month,
                    COUNT(*) as monthly_completions,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-6 months')
                GROUP BY technician_name, strftime('%Y-%m', completion_date)
                ORDER BY technician_name, month DESC
            ''')
            monthly_workload = cursor.fetchall()
        
            if monthly_workload:
                analytics += "MONTHLY WORKLOAD DISTRIBUTION (Last 6 months):\n"
            
                # Group by technician
                tech_monthly = {}
                for technician, month, completions, avg_hours in monthly_workload:
                    if technician not in tech_monthly:
                        tech_monthly[technician] = []
                    tech_monthly[technician].append((month, completions, avg_hours))
            
                for technician, monthly_data in tech_monthly.items():
                    tech_display = technician[:17] + '...' if len(technician) > 17 else technician
                    analytics += f"\n{tech_display}:\n"
                    analytics += f"{'  Month':<12} {'PMs':<6} {'Avg Hours':<10}\n"
                    analytics += "  " + "-" * 30 + "\n"
                
                    for month, completions, avg_hours in monthly_data[:6]:  # Show last 6 months
                        avg_hours_display = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                        analytics += f"  {month:<12} {completions:<6} {avg_hours_display:<10}\n"
                analytics += "\n"
        
            # Efficiency metrics
            if technician_performance:
                analytics += "TECHNICIAN EFFICIENCY METRICS:\n"
                analytics += f"{'Technician':<20} {'PMs/Month':<12} {'Hours/PM':<10} {'Productivity':<12}\n"
                analytics += "-" * 60 + "\n"
            
                for tech_data in technician_performance:
                    technician, total_pms, monthly, annual, avg_hours, total_hours, first_date, last_date = tech_data
                
                    # Calculate months active (rough estimate)
                    if first_date and last_date:
                        try:
                            first_dt = datetime.strptime(first_date, '%Y-%m-%d')
                            last_dt = datetime.strptime(last_date, '%Y-%m-%d')
                            months_active = max(1, (last_dt - first_dt).days / 30.44)  # Average days per month
                            pms_per_month = total_pms / months_active
                        except:
                            months_active = 1
                            pms_per_month = total_pms
                    else:
                        pms_per_month = total_pms
                
                    hours_per_pm = avg_hours if avg_hours else 0
                
                    # Productivity score (PMs per month / hours per PM)
                    productivity = pms_per_month / max(hours_per_pm, 0.1) if hours_per_pm > 0 else pms_per_month
                    
                    tech_display = technician[:17] + '...' if len(technician) > 17 else technician
                    analytics += f"{tech_display:<20} {pms_per_month:<11.1f} {hours_per_pm:<9.1f} {productivity:<11.1f}\n"
                analytics += "\n"
        
            tech_text.insert('end', analytics)
            tech_text.config(state='disabled')
        
        except Exception as e:
            print(f"Error generating technician analysis: {e}")
    
    def export_equipment_analytics_pdf(self, parent_dialog):
        """Export all analytics to a comprehensive PDF report"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Equipment_Analytics_Report_{timestamp}.pdf"
        
            # Create PDF document
            doc = SimpleDocTemplate(filename, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
        
            # Title page
            title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], 
                                    fontSize=20, textColor=colors.darkblue, alignment=1)
            story.append(Paragraph("AIT CMMS EQUIPMENT ANALYTICS REPORT", title_style))
            story.append(Spacer(1, 30))
        
            # Report metadata
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            story.append(Paragraph(f"Report ID: {timestamp}", styles['Normal']))
            story.append(Spacer(1, 40))
        
            # Executive Summary
            cursor = self.conn.cursor()
        
            cursor.execute('SELECT COUNT(*) FROM equipment')
            total_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Active"')
            active_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM pm_completions WHERE completion_date >= DATE("now", "-30 days")')
            recent_pms = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM cannot_find_assets WHERE status = "Missing"')
            missing_assets = cursor.fetchone()[0]
        
            story.append(Paragraph("EXECUTIVE SUMMARY", styles['Heading1']))
            summary_data = [
                ['Metric', 'Value', 'Status'],
                ['Total Equipment', str(total_equipment), 'Baseline'],
                ['Active Equipment', str(active_equipment), f'{active_equipment/total_equipment*100:.1f}%'],
                ['PMs Last 30 Days', str(recent_pms), 'Recent Activity'],
                ['Missing Assets', str(missing_assets), 'Attention Needed' if missing_assets > 0 else 'Good']
            ]
        
            summary_table = Table(summary_data, colWidths=[2*inch, 1.5*inch, 2*inch])
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
        
            story.append(summary_table)
            story.append(PageBreak())
        
            # Add detailed sections
            sections = [
                ("Equipment Overview", self.get_equipment_overview_text()),
                ("PM Performance Analysis", self.get_pm_performance_text()),
                ("Location Analysis", self.get_location_analysis_text()),
                ("Technician Analysis", self.get_technician_analysis_text())
            ]
        
            for section_title, section_content in sections:
                story.append(Paragraph(section_title, styles['Heading1']))
                story.append(Spacer(1, 12))
            
                # Split content into paragraphs
                paragraphs = section_content.split('\n\n')
                for paragraph in paragraphs:
                    if paragraph.strip():
                        story.append(Paragraph(paragraph.replace('\n', '<br/>'), styles['Normal']))
                        story.append(Spacer(1, 6))
            
                story.append(PageBreak())
        
            # Build PDF
            doc.build(story)
        
            messagebox.showinfo("Success", f"Analytics report exported to: {filename}")
            self.update_status(f"Equipment analytics exported to {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export analytics: {str(e)}")

    def get_equipment_overview_text(self):
        """Get equipment overview text for PDF export"""
        try:
            cursor = self.conn.cursor()
        
            # Basic statistics
            cursor.execute('SELECT COUNT(*) FROM equipment')
            total_equipment = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE status = "Active"')
            active_equipment = cursor.fetchone()[0]
        
            text = f"Total Equipment: {total_equipment}\n"
            text += f"Active Equipment: {active_equipment} ({active_equipment/total_equipment*100:.1f}%)\n\n"
        
            # PM requirements
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE monthly_pm = 1')
            monthly_count = cursor.fetchone()[0]
            text += f"Equipment requiring Monthly PM: {monthly_count}\n"
        
            cursor.execute('SELECT COUNT(*) FROM equipment WHERE annual_pm = 1')
            annual_count = cursor.fetchone()[0]
            text += f"Equipment requiring Annual PM: {annual_count}\n"
        
            return text
        
        except Exception as e:
            return f"Error generating overview text: {str(e)}"

    def get_pm_performance_text(self):
        """Get PM performance text for PDF export"""
        try:
            cursor = self.conn.cursor()
        
            cursor.execute('SELECT pm_type, COUNT(*) FROM pm_completions GROUP BY pm_type')
            pm_stats = cursor.fetchall()
        
            text = "PM Completion Statistics:\n"
            for pm_type, count in pm_stats:
                text += f"{pm_type}: {count} completions\n"
        
            return text
        
        except Exception as e:
            return f"Error generating PM performance text: {str(e)}"

    def get_location_analysis_text(self):
        """Get location analysis text for PDF export"""
        try:
            cursor = self.conn.cursor()
        
            cursor.execute('''
                SELECT location, COUNT(*) 
                FROM equipment 
                WHERE location IS NOT NULL 
                GROUP BY location 
                ORDER BY COUNT(*) DESC 
                LIMIT 10
            ''')
            location_stats = cursor.fetchall()
        
            text = "Equipment by Location:\n"
            for location, count in location_stats:
                text += f"{location}: {count} assets\n"
        
            return text
        
        except Exception as e:
            return f"Error generating location analysis text: {str(e)}"

    def get_technician_analysis_text(self):
        """Get technician analysis text for PDF export"""
        try:
            cursor = self.conn.cursor()
            
            cursor.execute('''
                SELECT technician_name, COUNT(*) 
                FROM pm_completions 
                GROUP BY technician_name 
                ORDER BY COUNT(*) DESC
            ''')
            tech_stats = cursor.fetchall()
        
            text = "PM Completions by Technician:\n"
            for technician, count in tech_stats:
                text += f"{technician}: {count} PMs completed\n"
        
            return text
        
        except Exception as e:
            return f"Error generating technician analysis text: {str(e)}"
    
    
    def show_pm_trends(self):
        """Comprehensive PM trends analysis with visualizations and insights"""
        try:
            # Create trends analysis dialog
            trends_dialog = tk.Toplevel(self.root)
            trends_dialog.title("PM Trends Analysis Dashboard")
            trends_dialog.geometry("1400x900")
            trends_dialog.transient(self.root)
            trends_dialog.grab_set()

            # Create notebook for different trend views
            trends_notebook = ttk.Notebook(trends_dialog)
            trends_notebook.pack(fill='both', expand=True, padx=10, pady=10)

            # Tab 1: Monthly Completion Trends
            monthly_frame = ttk.Frame(trends_notebook)
            trends_notebook.add(monthly_frame, text="Monthly Trends")

            # Tab 2: Equipment Performance Trends
            equipment_frame = ttk.Frame(trends_notebook)
            trends_notebook.add(equipment_frame, text="Equipment Trends")

            # Tab 3: Technician Performance Trends
            technician_frame = ttk.Frame(trends_notebook)
            trends_notebook.add(technician_frame, text="Technician Trends")

            # Tab 4: PM Type Distribution Trends
            pm_type_frame = ttk.Frame(trends_notebook)
            trends_notebook.add(pm_type_frame, text="PM Type Trends")

            # Generate content for each tab
            self.generate_monthly_trends_analysis(monthly_frame)
            self.generate_equipment_trends_analysis(equipment_frame)
            self.generate_technician_trends_analysis(technician_frame)
            self.generate_pm_type_trends_analysis(pm_type_frame)

            # Add export and close buttons
            button_frame = ttk.Frame(trends_dialog)
            button_frame.pack(side='bottom', fill='x', padx=10, pady=5)

            ttk.Button(button_frame, text="Export Trends to PDF", 
                    command=lambda: self.export_trends_analysis_pdf(trends_dialog)).pack(side='left', padx=5)
            ttk.Button(button_frame, text="Refresh Analysis", 
                    command=lambda: self.refresh_trends_analysis(trends_dialog)).pack(side='left', padx=5)
            ttk.Button(button_frame, text="Close", 
                    command=trends_dialog.destroy).pack(side='right', padx=5)
    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PM trends analysis: {str(e)}")

    def generate_monthly_trends_analysis(self, parent_frame):
        """Generate monthly PM completion trends analysis"""
        try:
            cursor = self.conn.cursor()

            # Create scrollable frame
            canvas = tk.Canvas(parent_frame)
            scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Monthly completion data (last 24 months)
            cursor.execute('''
                SELECT 
                    strftime('%Y-%m', completion_date) as month,
                    COUNT(*) as total_completions,
                    COUNT(CASE WHEN pm_type = 'Monthly' THEN 1 END) as monthly_pms,
                    COUNT(CASE WHEN pm_type = 'Annual' THEN 1 END) as annual_pms,
                    COUNT(CASE WHEN pm_type = 'Six Month' THEN 1 END) as six_month_pms,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours,
                    COUNT(DISTINCT technician_name) as active_technicians,
                    COUNT(DISTINCT bfm_equipment_no) as unique_equipment
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-24 months')
                GROUP BY strftime('%Y-%m', completion_date)
                ORDER BY month ASC
            ''')

            monthly_data = cursor.fetchall()

            # Create text display for trends
            trends_text = tk.Text(scrollable_frame, wrap='word', font=('Courier', 10), height=40)
            text_scrollbar = ttk.Scrollbar(scrollable_frame, orient='vertical', command=trends_text.yview)
            trends_text.configure(yscrollcommand=text_scrollbar.set)

            # Generate trends report
            report = "PM COMPLETION TRENDS ANALYSIS\n"
            report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += "=" * 80 + "\n\n"

            if monthly_data:
                report += "MONTHLY COMPLETION TRENDS (Last 24 months):\n"
                report += "-" * 80 + "\n"
                report += f"{'Month':<10} {'Total':<8} {'Monthly':<8} {'Annual':<8} {'6-Month':<8} {'Avg Hrs':<8} {'Techs':<6} {'Equipment':<10}\n"
                report += "-" * 80 + "\n"

                total_completions = 0
                total_hours = 0
                peak_month = None
                peak_completions = 0
                lowest_month = None
                lowest_completions = float('inf')

                for month_data in monthly_data:
                    month, total, monthly_pms, annual_pms, six_month_pms, avg_hours, techs, equipment = month_data
                    total_completions += total
                    total_hours += avg_hours if avg_hours else 0

                    # Track peak and low months
                    if total > peak_completions:
                        peak_completions = total
                        peak_month = month
                    if total < lowest_completions:
                        lowest_completions = total
                        lowest_month = month

                    avg_hours_str = f"{avg_hours:.1f}" if avg_hours else "0.0"
                    report += f"{month:<10} {total:<8} {monthly_pms:<8} {annual_pms:<8} {six_month_pms:<8} {avg_hours_str:<8} {techs:<6} {equipment:<10}\n"

                # Calculate trends
                avg_monthly_completions = total_completions / len(monthly_data) if monthly_data else 0
                avg_hours_overall = total_hours / len(monthly_data) if monthly_data else 0

                # Recent trend analysis (last 6 months vs previous 6 months)
                recent_6_months = monthly_data[-6:] if len(monthly_data) >= 6 else monthly_data
                previous_6_months = monthly_data[-12:-6] if len(monthly_data) >= 12 else []

                recent_avg = sum(row[1] for row in recent_6_months) / len(recent_6_months) if recent_6_months else 0
                previous_avg = sum(row[1] for row in previous_6_months) / len(previous_6_months) if previous_6_months else 0

                trend_direction = "UP" if recent_avg > previous_avg else "DOWN" if recent_avg < previous_avg else "STABLE"
                trend_percentage = ((recent_avg - previous_avg) / previous_avg * 100) if previous_avg > 0 else 0

                report += "\n" + "=" * 80 + "\n"
                report += "TREND ANALYSIS SUMMARY:\n"
                report += "=" * 80 + "\n"
                report += f"Total Months Analyzed: {len(monthly_data)}\n"
                report += f"Total Completions: {total_completions}\n"
                report += f"Average Completions per Month: {avg_monthly_completions:.1f}\n"
                report += f"Average Hours per PM: {avg_hours_overall:.1f}h\n\n"

                report += f"Peak Performance Month: {peak_month} ({peak_completions} completions)\n"
                report += f"Lowest Performance Month: {lowest_month} ({lowest_completions} completions)\n\n"

                report += f"6-Month Trend Analysis:\n"
                report += f"Recent 6 months average: {recent_avg:.1f} completions/month\n"
                report += f"Previous 6 months average: {previous_avg:.1f} completions/month\n"
                report += f"Trend Direction: {trend_direction} ({trend_percentage:+.1f}%)\n\n"

                # Seasonal analysis
                report += "SEASONAL PATTERNS:\n"
                report += "-" * 40 + "\n"
                seasonal_data = {}
                for month_data in monthly_data:
                    month_str, total = month_data[0], month_data[1]
                    month_num = int(month_str.split('-')[1])
                    season = self.get_season_from_month(month_num)
                    if season not in seasonal_data:
                        seasonal_data[season] = []
                    seasonal_data[season].append(total)
    
                for season, completions in seasonal_data.items():
                    avg_seasonal = sum(completions) / len(completions)
                    report += f"{season:<10}: {avg_seasonal:.1f} avg completions/month\n"

                # Workload distribution analysis
                report += "\nWORKLOAD DISTRIBUTION INSIGHTS:\n"
                report += "-" * 40 + "\n"
            
                # Calculate coefficient of variation for consistency
                if len(monthly_data) > 1:
                    completions_list = [row[1] for row in monthly_data]
                    import statistics
                    std_dev = statistics.stdev(completions_list)
                    cv = (std_dev / avg_monthly_completions) * 100 if avg_monthly_completions > 0 else 0
                
                    consistency_rating = "Very Consistent" if cv < 15 else "Consistent" if cv < 25 else "Variable" if cv < 35 else "Highly Variable"
                    report += f"Workload Consistency: {consistency_rating} (CV: {cv:.1f}%)\n"
                    report += f"Standard Deviation: {std_dev:.1f} completions\n\n"

                # Recommendations
                report += "RECOMMENDATIONS:\n"
                report += "-" * 40 + "\n"
                if trend_direction == "DOWN":
                    report += "• Investigate causes of declining PM completion rates\n"
                    report += "• Consider additional technician training or resources\n"
                    report += "• Review equipment scheduling and assignment processes\n"
                elif trend_direction == "UP":
                    report += "• Excellent performance trend - maintain current practices\n"
                    report += "• Consider documenting successful strategies for replication\n"
            
                if cv > 30:
                    report += "• High variability detected - investigate scheduling consistency\n"
                    report += "• Consider implementing better workload balancing\n"
            
                if avg_hours_overall > 2.0:
                    report += "• Average PM time is high - review procedures for efficiency\n"
                elif avg_hours_overall < 0.5:
                    report += "• Very low average PM time - verify completeness of work\n"

            else:
                report += "No PM completion data found for trend analysis.\n"

            # Display the report
            trends_text.insert('end', report)
            trends_text.config(state='disabled')

            # Pack widgets
            trends_text.pack(side='left', fill='both', expand=True)
            text_scrollbar.pack(side='right', fill='y')

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

        except Exception as e:
            error_label = ttk.Label(parent_frame, text=f"Error generating monthly trends: {str(e)}")
            error_label.pack(pady=20)

    def generate_equipment_trends_analysis(self, parent_frame):
        """Generate equipment-specific PM trends analysis"""
        try:
            cursor = self.conn.cursor()

            # Create text widget for equipment trends
            equipment_text = tk.Text(parent_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(parent_frame, orient='vertical', command=equipment_text.yview)
            equipment_text.configure(yscrollcommand=scrollbar.set)

            report = "EQUIPMENT PM TRENDS ANALYSIS\n"
            report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += "=" * 80 + "\n\n"

            # Most frequently serviced equipment
            cursor.execute('''
                SELECT e.bfm_equipment_no, e.description, e.location,
                       COUNT(pc.id) as total_pms,
                       AVG(pc.labor_hours + pc.labor_minutes/60.0) as avg_hours,
                       MIN(pc.completion_date) as first_pm,
                       MAX(pc.completion_date) as last_pm,
                       COUNT(CASE WHEN pc.completion_date >= DATE('now', '-90 days') THEN 1 END) as recent_pms
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no
                WHERE e.status = 'Active'
                GROUP BY e.bfm_equipment_no, e.description, e.location
                HAVING total_pms > 0
                ORDER BY total_pms DESC
                LIMIT 20
            ''')

            high_maintenance_equipment = cursor.fetchall()

            if high_maintenance_equipment:
                report += "TOP 20 MOST SERVICED EQUIPMENT:\n"
                report += "-" * 80 + "\n"
                report += f"{'Rank':<5} {'BFM No':<12} {'Description':<25} {'Total PMs':<10} {'Avg Hours':<10} {'Recent (90d)':<12}\n"
                report += "-" * 80 + "\n"

                for i, equipment in enumerate(high_maintenance_equipment, 1):
                    bfm_no, description, location, total_pms, avg_hours, first_pm, last_pm, recent_pms = equipment
                    desc_short = (description[:22] + '...') if description and len(description) > 22 else (description or 'N/A')
                    avg_hours_str = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    
                    report += f"{i:<5} {bfm_no:<12} {desc_short:<25} {total_pms:<10} {avg_hours_str:<10} {recent_pms:<12}\n"

            # Equipment with increasing maintenance needs
            cursor.execute('''
                SELECT bfm_equipment_no,
                       COUNT(CASE WHEN completion_date >= DATE('now', '-90 days') THEN 1 END) as last_90_days,
                       COUNT(CASE WHEN completion_date >= DATE('now', '-180 days') AND completion_date < DATE('now', '-90 days') THEN 1 END) as prev_90_days,
                       COUNT(*) as total_pms
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-180 days')
                GROUP BY bfm_equipment_no
                HAVING last_90_days > prev_90_days AND prev_90_days > 0
                ORDER BY (last_90_days - prev_90_days) DESC
                LIMIT 10
            ''')

            increasing_maintenance = cursor.fetchall()

            if increasing_maintenance:
                report += "\n\nEQUIPMENT WITH INCREASING MAINTENANCE NEEDS:\n"
                report += "-" * 60 + "\n"
                report += f"{'BFM No':<15} {'Recent 90d':<12} {'Previous 90d':<12} {'Increase':<10}\n"
                report += "-" * 60 + "\n"

                for equipment in increasing_maintenance:
                    bfm_no, last_90, prev_90, total = equipment
                    increase = last_90 - prev_90
                    report += f"{bfm_no:<15} {last_90:<12} {prev_90:<12} +{increase:<9}\n"

            # Equipment that hasn't been serviced recently
            cursor.execute('''
                SELECT e.bfm_equipment_no, e.description, e.location,
                       MAX(pc.completion_date) as last_pm_date,
                       JULIANDAY('now') - JULIANDAY(MAX(pc.completion_date)) as days_since_last_pm
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no
                WHERE e.status = 'Active' AND e.monthly_pm = 1
                GROUP BY e.bfm_equipment_no, e.description, e.location
                HAVING days_since_last_pm > 60 OR last_pm_date IS NULL
                ORDER BY days_since_last_pm DESC NULLS LAST
                LIMIT 15
            ''')

            neglected_equipment = cursor.fetchall()

            if neglected_equipment:
                report += "\n\nEQUIPMENT REQUIRING ATTENTION (>60 days since last PM):\n"
                report += "-" * 70 + "\n"
                report += f"{'BFM No':<15} {'Description':<25} {'Last PM':<12} {'Days Since':<12}\n"
                report += "-" * 70 + "\n"

                for equipment in neglected_equipment:
                    bfm_no, description, location, last_pm, days_since = equipment
                    desc_short = (description[:22] + '...') if description and len(description) > 22 else (description or 'N/A')
                    last_pm_str = last_pm if last_pm else 'Never'
                    days_str = f"{int(days_since)}" if days_since else 'N/A'
                    
                    report += f"{bfm_no:<15} {desc_short:<25} {last_pm_str:<12} {days_str:<12}\n"

            # Equipment performance by location
            cursor.execute('''
                SELECT COALESCE(e.location, 'Unknown') as location,
                       COUNT(pc.id) as total_pms,
                       COUNT(DISTINCT e.bfm_equipment_no) as equipment_count,
                       AVG(pc.labor_hours + pc.labor_minutes/60.0) as avg_hours,
                       ROUND(CAST(COUNT(pc.id) AS FLOAT) / COUNT(DISTINCT e.bfm_equipment_no), 2) as pms_per_equipment
                FROM equipment e
                LEFT JOIN pm_completions pc ON e.bfm_equipment_no = pc.bfm_equipment_no
                    AND pc.completion_date >= DATE('now', '-365 days')
                WHERE e.status = 'Active'
                GROUP BY COALESCE(e.location, 'Unknown')
                HAVING equipment_count >= 3
                ORDER BY pms_per_equipment DESC
            ''')

            location_performance = cursor.fetchall()

            if location_performance:
                report += "\n\nPM PERFORMANCE BY LOCATION (Last 12 months):\n"
                report += "-" * 70 + "\n"
                report += f"{'Location':<20} {'Equipment':<10} {'Total PMs':<10} {'PMs/Equipment':<15} {'Avg Hours':<10}\n"
                report += "-" * 70 + "\n"

                for location_data in location_performance:
                    location, total_pms, equipment_count, avg_hours, pms_per_equipment = location_data
                    loc_short = (location[:17] + '...') if len(location) > 17 else location
                    avg_hours_str = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    
                    report += f"{loc_short:<20} {equipment_count:<10} {total_pms:<10} {pms_per_equipment:<15} {avg_hours_str:<10}\n"

            equipment_text.insert('end', report)
            equipment_text.config(state='disabled')

            equipment_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

        except Exception as e:
            error_label = ttk.Label(parent_frame, text=f"Error generating equipment trends: {str(e)}")
            error_label.pack(pady=20)

    def generate_technician_trends_analysis(self, parent_frame):
        """Generate technician performance trends analysis"""
        try:
            cursor = self.conn.cursor()

            # Create text widget for technician trends
            tech_text = tk.Text(parent_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(parent_frame, orient='vertical', command=tech_text.yview)
            tech_text.configure(yscrollcommand=scrollbar.set)

            report = "TECHNICIAN PERFORMANCE TRENDS\n"
            report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += "=" * 80 + "\n\n"

            # Monthly performance trends for each technician
            cursor.execute('''
                SELECT technician_name,
                       strftime('%Y-%m', completion_date) as month,
                       COUNT(*) as completions,
                       AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY technician_name, strftime('%Y-%m', completion_date)
                ORDER BY technician_name, month
            ''')

            monthly_tech_data = cursor.fetchall()

            # Organize data by technician
            tech_monthly = {}
            for row in monthly_tech_data:
                tech, month, completions, avg_hours = row
                if tech not in tech_monthly:
                    tech_monthly[tech] = []
                tech_monthly[tech].append((month, completions, avg_hours))

            if tech_monthly:
                report += "MONTHLY PERFORMANCE TRENDS BY TECHNICIAN:\n"
                report += "=" * 80 + "\n"

                for technician, monthly_data in tech_monthly.items():
                    report += f"\n{technician}:\n"
                    report += "-" * 50 + "\n"
                    report += f"{'Month':<10} {'Completions':<12} {'Avg Hours':<10} {'Trend':<10}\n"
                    report += "-" * 50 + "\n"

                    # Calculate trend
                    completions_list = [data[1] for data in monthly_data]
                    if len(completions_list) >= 3:
                        recent_avg = sum(completions_list[-3:]) / 3
                        earlier_avg = sum(completions_list[:-3]) / len(completions_list[:-3]) if len(completions_list) > 3 else sum(completions_list[:3]) / len(completions_list[:3])
                        trend = "↑" if recent_avg > earlier_avg else "↓" if recent_avg < earlier_avg else "→"
                    else:
                        trend = "→"

                    for month, completions, avg_hours in monthly_data[-6:]:  # Show last 6 months
                        avg_hours_str = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                        report += f"{month:<10} {completions:<12} {avg_hours_str:<10} {trend if month == monthly_data[-1][0] else '':<10}\n"

            # Overall technician comparison
            cursor.execute('''
                SELECT technician_name,
                       COUNT(*) as total_completions,
                       AVG(labor_hours + labor_minutes/60.0) as avg_hours_per_pm,
                       COUNT(DISTINCT bfm_equipment_no) as unique_equipment,
                       COUNT(CASE WHEN completion_date >= DATE('now', '-30 days') THEN 1 END) as recent_completions,
                       MIN(completion_date) as first_completion,
                       MAX(completion_date) as last_completion
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY technician_name
                ORDER BY total_completions DESC
            ''')

            tech_comparison = cursor.fetchall()

            if tech_comparison:
                report += "\n\nTECHNICIAN PERFORMANCE COMPARISON (Last 12 months):\n"
                report += "=" * 90 + "\n"
                report += f"{'Technician':<20} {'Total PMs':<10} {'Avg Hrs':<10} {'Equipment':<10} {'Recent 30d':<12} {'Active Period':<15}\n"
                report += "=" * 90 + "\n"

                for tech_data in tech_comparison:
                    tech, total, avg_hours, equipment, recent, first, last = tech_data
                    avg_hours_str = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                    
                    # Calculate active period
                    if first and last:
                        first_date = datetime.strptime(first, '%Y-%m-%d')
                        last_date = datetime.strptime(last, '%Y-%m-%d')
                        active_days = (last_date - first_date).days
                        active_period = f"{active_days}d"
                    else:
                        active_period = "N/A"

                    tech_short = tech[:17] + '...' if len(tech) > 17 else tech
                    report += f"{tech_short:<20} {total:<10} {avg_hours_str:<10} {equipment:<10} {recent:<12} {active_period:<15}\n"

            # Efficiency metrics
            if tech_comparison:
                report += "\n\nEFFICIENCY METRICS:\n"
                report += "-" * 60 + "\n"
                report += f"{'Technician':<20} {'PMs/Day':<10} {'Productivity':<12} {'Specialization':<15}\n"
                report += "-" * 60 + "\n"

                for tech_data in tech_comparison:
                    tech, total, avg_hours, equipment, recent, first, last = tech_data
                
                    # Calculate PMs per day (approximate)
                    if first and last:
                        first_date = datetime.strptime(first, '%Y-%m-%d')
                        last_date = datetime.strptime(last, '%Y-%m-%d')
                        active_days = max(1, (last_date - first_date).days)
                        pms_per_day = total / active_days
                    else:
                        pms_per_day = 0

                    # Productivity score (PMs per hour)
                    productivity = total / (total * (avg_hours if avg_hours else 1)) if avg_hours else total
                
                    # Specialization (unique equipment ratio)
                    specialization = equipment / total if total > 0 else 0
                    spec_rating = "High" if specialization > 0.8 else "Medium" if specialization > 0.5 else "Low"

                    tech_short = tech[:17] + '...' if len(tech) > 17 else tech
                    report += f"{tech_short:<20} {pms_per_day:<9.2f} {productivity:<11.2f} {spec_rating:<15}\n"

            tech_text.insert('end', report)
            tech_text.config(state='disabled')

            tech_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

        except Exception as e:
            error_label = ttk.Label(parent_frame, text=f"Error generating technician trends: {str(e)}")
            error_label.pack(pady=20)

    def generate_pm_type_trends_analysis(self, parent_frame):
        """Generate PM type distribution and trends analysis"""
        try:
            cursor = self.conn.cursor()

            # Create text widget for PM type trends
            pm_type_text = tk.Text(parent_frame, wrap='word', font=('Courier', 10))
            scrollbar = ttk.Scrollbar(parent_frame, orient='vertical', command=pm_type_text.yview)
            pm_type_text.configure(yscrollcommand=scrollbar.set)

            report = "PM TYPE TRENDS ANALYSIS\n"
            report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += "=" * 80 + "\n\n"

            # Monthly PM type distribution
            cursor.execute('''
                SELECT strftime('%Y-%m', completion_date) as month,
                       pm_type,
                       COUNT(*) as completions,
                       AVG(labor_hours + labor_minutes/60.0) as avg_hours
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY strftime('%Y-%m', completion_date), pm_type
                ORDER BY month, pm_type
            ''')

            monthly_pm_type_data = cursor.fetchall()

            # Organize by month
            monthly_pm_types = {}
            for row in monthly_pm_type_data:
                month, pm_type, completions, avg_hours = row
                if month not in monthly_pm_types:
                    monthly_pm_types[month] = {}
                monthly_pm_types[month][pm_type] = (completions, avg_hours)

            if monthly_pm_types:
                report += "MONTHLY PM TYPE DISTRIBUTION:\n"
                report += "=" * 80 + "\n"
                report += f"{'Month':<10} {'Monthly':<10} {'Annual':<10} {'Six Month':<12} {'Other':<8} {'Total':<8}\n"
                report += "=" * 80 + "\n"

                pm_type_totals = {}
                for month, pm_types in monthly_pm_types.items():
                    monthly_count = pm_types.get('Monthly', (0, 0))[0]
                    annual_count = pm_types.get('Annual', (0, 0))[0]
                    six_month_count = pm_types.get('Six Month', (0, 0))[0]
                    other_count = sum(data[0] for pm_type, data in pm_types.items() 
                                    if pm_type not in ['Monthly', 'Annual', 'Six Month'])
                    total_count = monthly_count + annual_count + six_month_count + other_count

                    # Track totals
                    for pm_type, (count, _) in pm_types.items():
                        pm_type_totals[pm_type] = pm_type_totals.get(pm_type, 0) + count

                    report += f"{month:<10} {monthly_count:<10} {annual_count:<10} {six_month_count:<12} {other_count:<8} {total_count:<8}\n"

            # Overall PM type statistics
            cursor.execute('''
                SELECT pm_type,
                    COUNT(*) as total_completions,
                    AVG(labor_hours + labor_minutes/60.0) as avg_hours,
                       MIN(completion_date) as first_completion,
                       MAX(completion_date) as last_completion,
                       COUNT(DISTINCT technician_name) as technicians_involved,
                       COUNT(DISTINCT bfm_equipment_no) as equipment_serviced
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY pm_type
                ORDER BY total_completions DESC
            ''')

            pm_type_stats = cursor.fetchall()

            if pm_type_stats:
                report += "\n\nPM TYPE PERFORMANCE SUMMARY (Last 12 months):\n"
                report += "=" * 90 + "\n"
                report += f"{'PM Type':<15} {'Total':<8} {'Avg Hours':<10} {'Technicians':<12} {'Equipment':<10} {'Period':<15}\n"
                report += "=" * 90 + "\n"

                total_all_pms = sum(row[1] for row in pm_type_stats)
            
                for pm_data in pm_type_stats:
                    pm_type, total, avg_hours, first, last, techs, equipment = pm_data
                    percentage = (total / total_all_pms * 100) if total_all_pms > 0 else 0
                    avg_hours_str = f"{avg_hours:.1f}h" if avg_hours else "0.0h"
                
                    # Calculate period
                    if first and last:
                        first_date = datetime.strptime(first, '%Y-%m-%d')
                        last_date = datetime.strptime(last, '%Y-%m-%d')
                        period_days = (last_date - first_date).days
                        period_str = f"{period_days}d"
                    else:
                        period_str = "N/A"

                    report += f"{pm_type:<15} {total:<8} {avg_hours_str:<10} {techs:<12} {equipment:<10} {period_str:<15}\n"
                    report += f"{'':>15} ({percentage:.1f}%)\n"

            # PM type efficiency analysis
            if pm_type_stats:
                report += "\n\nPM TYPE EFFICIENCY ANALYSIS:\n"
                report += "-" * 60 + "\n"
            
                # Calculate efficiency metrics
                for pm_data in pm_type_stats:
                    pm_type, total, avg_hours, first, last, techs, equipment = pm_data
                
                    # Equipment coverage (how many unique equipment per PM)
                    coverage = equipment / total if total > 0 else 0
                
                    # Time efficiency rating
                    if avg_hours:
                        if avg_hours <= 1.0:
                            efficiency = "Excellent"
                        elif avg_hours <= 1.5:
                            efficiency = "Good"
                        elif avg_hours <= 2.5:
                            efficiency = "Average"
                        else:
                            efficiency = "Needs Review"
                    else:
                        efficiency = "Unknown"
                
                    report += f"{pm_type} PM Analysis:\n"
                    report += f"  • Average completion time: {avg_hours:.1f}h ({efficiency})\n" if avg_hours else f"  • Average completion time: Unknown\n"
                    report += f"  • Equipment coverage ratio: {coverage:.2f}\n"
                    report += f"  • Technician utilization: {techs} different technicians\n"
                
                    # Frequency analysis
                    if first and last and total > 1:
                        first_date = datetime.strptime(first, '%Y-%m-%d')
                        last_date = datetime.strptime(last, '%Y-%m-%d')
                        total_days = (last_date - first_date).days
                        avg_days_between = total_days / (total - 1) if total > 1 else 0
                        report += f"  • Average interval: {avg_days_between:.1f} days between completions\n"
                
                    report += "\n"

            # Seasonal PM type patterns
            cursor.execute('''
                SELECT 
                    CASE 
                        WHEN CAST(strftime('%m', completion_date) AS INTEGER) IN (12, 1, 2) THEN 'Winter'
                        WHEN CAST(strftime('%m', completion_date) AS INTEGER) IN (3, 4, 5) THEN 'Spring'
                        WHEN CAST(strftime('%m', completion_date) AS INTEGER) IN (6, 7, 8) THEN 'Summer'
                        WHEN CAST(strftime('%m', completion_date) AS INTEGER) IN (9, 10, 11) THEN 'Fall'
                    END as season,
                    pm_type,
                    COUNT(*) as completions
                FROM pm_completions
                WHERE completion_date >= DATE('now', '-12 months')
                GROUP BY season, pm_type
                ORDER BY season, pm_type
            ''')

            seasonal_data = cursor.fetchall()

            if seasonal_data:
                report += "SEASONAL PM TYPE PATTERNS:\n"
                report += "-" * 50 + "\n"
            
                # Organize by season
                seasons = {}
                for row in seasonal_data:
                    season, pm_type, completions = row
                    if season not in seasons:
                        seasons[season] = {}
                    seasons[season][pm_type] = completions

                for season in ['Winter', 'Spring', 'Summer', 'Fall']:
                    if season in seasons:
                        report += f"\n{season}:\n"
                        season_total = sum(seasons[season].values())
                        for pm_type, count in seasons[season].items():
                            percentage = (count / season_total * 100) if season_total > 0 else 0
                            report += f"  {pm_type}: {count} ({percentage:.1f}%)\n"

            # Recommendations based on PM type analysis
            report += "\n\nPM TYPE RECOMMENDATIONS:\n"
            report += "=" * 50 + "\n"
        
            if pm_type_stats:
                # Find the most and least efficient PM types
                sorted_by_hours = sorted(pm_type_stats, key=lambda x: x[2] if x[2] else 0)
                most_efficient = sorted_by_hours[0] if sorted_by_hours else None
                least_efficient = sorted_by_hours[-1] if sorted_by_hours else None
                
                if most_efficient and least_efficient and most_efficient[2] and least_efficient[2]:
                    if most_efficient[0] != least_efficient[0]:
                        report += f"• Most efficient PM type: {most_efficient[0]} ({most_efficient[2]:.1f}h avg)\n"
                        report += f"• Least efficient PM type: {least_efficient[0]} ({least_efficient[2]:.1f}h avg)\n"
                        report += f"• Consider reviewing procedures for {least_efficient[0]} PMs\n\n"
            
                # Check for imbalanced distribution
                monthly_pms = next((row[1] for row in pm_type_stats if row[0] == 'Monthly'), 0)
                annual_pms = next((row[1] for row in pm_type_stats if row[0] == 'Annual'), 0)
            
                if monthly_pms > 0 and annual_pms > 0:
                    ratio = monthly_pms / annual_pms
                    if ratio > 15:
                        report += "• High Monthly-to-Annual PM ratio detected\n"
                        report += "• Consider whether some Monthly PMs could be converted to Annual\n\n"
                    elif ratio < 3:
                        report += "• Low Monthly-to-Annual PM ratio detected\n"
                        report += "• Verify Monthly PM scheduling is adequate\n\n"
            
                # Check for types with long completion times
                long_pm_types = [row for row in pm_type_stats if row[2] and row[2] > 3.0]
                if long_pm_types:
                    report += "• PM types with long completion times (>3h):\n"
                    for pm_type, total, avg_hours, _, _, _, _ in long_pm_types:
                        report += f"  - {pm_type}: {avg_hours:.1f}h average\n"
                    report += "• Review these procedures for potential optimization\n\n"

            pm_type_text.insert('end', report)
            pm_type_text.config(state='disabled')

            pm_type_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

        except Exception as e:
            error_label = ttk.Label(parent_frame, text=f"Error generating PM type trends: {str(e)}")
            error_label.pack(pady=20)

    def get_season_from_month(self, month_num):
        """Helper function to get season from month number"""
        if month_num in [12, 1, 2]:
            return "Winter"
        elif month_num in [3, 4, 5]:
            return "Spring"
        elif month_num in [6, 7, 8]:
            return "Summer"
        else:
            return "Fall"

    def export_trends_analysis_pdf(self, parent_dialog):
        """Export trends analysis to PDF"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"PM_Trends_Analysis_{timestamp}.pdf"

            # Create PDF document
            doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=36, leftMargin=36,
                                topMargin=36, bottomMargin=36)
        
            story = []
            styles = getSampleStyleSheet()

            # Title page
            title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], 
                                    fontSize=20, textColor=colors.darkblue, alignment=1)
            story.append(Paragraph("AIT CMMS PM TRENDS ANALYSIS REPORT", title_style))
            story.append(Spacer(1, 30))

            # Report metadata
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            story.append(Paragraph(f"Report ID: {timestamp}", styles['Normal']))
            story.append(Spacer(1, 20))

            # Executive Summary
            cursor = self.conn.cursor()
        
            # Get summary statistics
            cursor.execute('SELECT COUNT(*) FROM pm_completions WHERE completion_date >= DATE("now", "-12 months")')
            total_pms_year = cursor.fetchone()[0]
        
            cursor.execute('SELECT COUNT(*) FROM pm_completions WHERE completion_date >= DATE("now", "-30 days")')
            total_pms_month = cursor.fetchone()[0]
        
            cursor.execute('SELECT AVG(labor_hours + labor_minutes/60.0) FROM pm_completions WHERE completion_date >= DATE("now", "-12 months")')
            avg_hours = cursor.fetchone()[0] or 0

            story.append(Paragraph("EXECUTIVE SUMMARY", styles['Heading1']))
            summary_text = f"""
            This comprehensive PM trends analysis covers the last 12 months of preventive maintenance activities.
        
            Key Metrics:
            • Total PM Completions (12 months): {total_pms_year}
            • Recent PM Completions (30 days): {total_pms_month}
            • Average PM Duration: {avg_hours:.1f} hours
            • Analysis Period: {(datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')} to {datetime.now().strftime('%Y-%m-%d')}
        
            This report provides insights into monthly completion trends, equipment performance patterns,
            technician productivity analysis, and PM type distribution to support data-driven maintenance decisions.
            """
        
            story.append(Paragraph(summary_text, styles['Normal']))
            story.append(PageBreak())

            # Add key findings sections
            story.append(Paragraph("DETAILED ANALYSIS", styles['Heading1']))
            story.append(Paragraph("The following sections provide comprehensive trends analysis across multiple dimensions of PM performance.", styles['Normal']))
            story.append(Spacer(1, 20))

            # Note about data sources
            story.append(Paragraph("Data Sources and Methodology", styles['Heading2']))
            methodology_text = """
            This analysis is based on PM completion records from the AIT CMMS database. 
            All calculations use standardized date formats and validated completion records.
            Trends are calculated using statistical methods appropriate for time series data.
            """
            story.append(Paragraph(methodology_text, styles['Normal']))

            # Build PDF
            doc.build(story)

            messagebox.showinfo("Success", f"PM trends analysis exported to: {filename}")
            self.update_status(f"PM trends analysis exported to {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export trends analysis: {str(e)}")

    def refresh_trends_analysis(self, parent_dialog):
        """Refresh the trends analysis with current data"""
        try:
            # Destroy and recreate the dialog
            parent_dialog.destroy()
            self.show_pm_trends()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh trends analysis: {str(e)}")
    
    
    
    
    
    
    
    
    def export_analytics(self):
        """Export analytics to file"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"AIT_CMMS_Analytics_{timestamp}.txt"
            
            content = self.analytics_text.get('1.0', 'end-1c')
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            
            messagebox.showinfo("Success", f"Analytics exported to: {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export analytics: {str(e)}")
    
    # Replace your existing import_equipment_csv method with this enhanced version

    def import_equipment_csv(self):
        """Import equipment data from CSV file with PM dates"""
        file_path = filedialog.askopenfilename(
            title="Select Equipment CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
    
        if file_path:
            try:
                # Show column mapping dialog first
                self.show_csv_mapping_dialog(file_path)
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import CSV file: {str(e)}")

    # Replace your show_csv_mapping_dialog method with this fixed version

    def show_csv_mapping_dialog(self, file_path):
        """Show dialog to map CSV columns to database fields"""
    
        try:
            # Read CSV to get column headers
            df = pd.read_csv(file_path, encoding='cp1252', nrows=5)  # Just read first 5 rows to see structure
            csv_columns = list(df.columns)
        
            dialog = tk.Toplevel(self.root)
            dialog.title("Map CSV Columns to Database Fields")
            dialog.geometry("700x600")  # Made it larger
            dialog.transient(self.root)
            dialog.grab_set()
        
            # Main container with scrollbar
            main_canvas = tk.Canvas(dialog)
            scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=main_canvas.yview)
            scrollable_frame = ttk.Frame(main_canvas)
        
            scrollable_frame.bind(
                "<Configure>",
                lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
            )
        
            main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            main_canvas.configure(yscrollcommand=scrollbar.set)
        
            # Instructions
            ttk.Label(scrollable_frame, text="Map your CSV columns to the correct database fields:", 
                    font=('Arial', 12, 'bold')).pack(pady=10)
        
            # Create mapping frame
            mapping_frame = ttk.Frame(scrollable_frame)
            mapping_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
            # Column mappings
            mappings = {}
        
            # Database fields that can be mapped
            db_fields = [
                ("SAP Material No", "sap_material_no"),
                ("BFM Equipment No", "bfm_equipment_no"), 
                ("Description", "description"),
                ("Tool ID/Drawing No", "tool_id_drawing_no"),
                ("Location", "location"),
                ("Master LIN", "master_lin"),
                ("Last Monthly PM (YYYY-MM-DD)", "last_monthly_pm"),
                ("Last Six Month PM (YYYY-MM-DD)", "last_six_month_pm"),
                ("Last Annual PM (YYYY-MM-DD)", "last_annual_pm"),
                ("Monthly PM Required (1/0 or Y/N)", "monthly_pm"),
                ("Six Month PM Required (1/0 or Y/N)", "six_month_pm"),
                ("Annual PM Required (1/0 or Y/N)", "annual_pm")
            ]
        
            # Add "None" option to CSV columns
            csv_options = ["(Not in CSV)"] + csv_columns
        
            row = 0
            for field_name, field_key in db_fields:
                ttk.Label(mapping_frame, text=field_name + ":").grid(row=row, column=0, sticky='w', pady=2)
            
                mapping_var = tk.StringVar()
                combo = ttk.Combobox(mapping_frame, textvariable=mapping_var, values=csv_options, width=30)
                combo.grid(row=row, column=1, padx=10, pady=2)
            
                # Try to auto-match common column names
                for csv_col in csv_columns:
                    csv_lower = csv_col.lower()
                    if field_key == 'sap_material_no' and 'sap' in csv_lower:
                        mapping_var.set(csv_col)
                        break
                    elif field_key == 'bfm_equipment_no' and 'bfm' in csv_lower:
                        mapping_var.set(csv_col)
                        break
                    elif field_key == 'description' and 'description' in csv_lower:
                        mapping_var.set(csv_col)
                        break
                    elif field_key == 'location' and 'location' in csv_lower:
                        mapping_var.set(csv_col)
                        break
                    elif field_key == 'master_lin' and 'lin' in csv_lower:
                        mapping_var.set(csv_col)
                        break
            
                mappings[field_key] = mapping_var
                row += 1
        
            # Show sample data
            sample_frame = ttk.LabelFrame(scrollable_frame, text="Sample Data from Your CSV", padding=10)
            sample_frame.pack(fill='x', padx=20, pady=10)
        
            sample_text = tk.Text(sample_frame, height=6, width=80)
            sample_text.pack()
            sample_text.insert('1.0', df.to_string())
            sample_text.config(state='disabled')
        
            def process_import():
                """Process the import with mapped columns"""
                try:
                    # Get the full CSV data
                    full_df = pd.read_csv(file_path, encoding='cp1252')
                    full_df.columns = full_df.columns.str.strip()
                
                    cursor = self.conn.cursor()
                    imported_count = 0
                    error_count = 0
                
                    for index, row in full_df.iterrows():
                        try:
                            # Extract mapped data
                            data = {}
                            for field_key, mapping_var in mappings.items():
                                csv_column = mapping_var.get()
                                if csv_column != "(Not in CSV)" and csv_column in full_df.columns:
                                    value = row[csv_column]
                                    if pd.isna(value):
                                        data[field_key] = None
                                    else:
                                        # Handle different data types
                                        if field_key in ['monthly_pm', 'six_month_pm', 'annual_pm']:
                                            # Convert Y/N or 1/0 to boolean
                                            if str(value).upper() in ['Y', 'YES', '1', 'TRUE']:
                                                data[field_key] = 1
                                            else:
                                                data[field_key] = 0
                                        elif field_key in ['last_monthly_pm', 'last_six_month_pm', 'last_annual_pm']:
                                            # Handle date fields
                                            try:
                                                # Try to parse date
                                                parsed_date = pd.to_datetime(value).strftime('%Y-%m-%d')
                                                data[field_key] = parsed_date
                                            except:
                                                data[field_key] = None
                                        else:
                                            data[field_key] = str(value)
                                else:
                                    # Set defaults for unmapped fields
                                    if field_key in ['monthly_pm', 'six_month_pm', 'annual_pm']:
                                        data[field_key] = 1  # Default to requiring all PM types
                                    else:
                                        data[field_key] = None
                        
                            # Only import if BFM number exists
                            if data.get('bfm_equipment_no'):
                                cursor.execute('''
                                    INSERT OR REPLACE INTO equipment 
                                    (sap_material_no, bfm_equipment_no, description, tool_id_drawing_no, location, 
                                    master_lin, monthly_pm, six_month_pm, annual_pm, last_monthly_pm, 
                                    last_six_month_pm, last_annual_pm, next_monthly_pm, next_six_month_pm, next_annual_pm)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
                                        CASE WHEN ? IS NOT NULL THEN DATE(?, '+30 days') ELSE NULL END,
                                        CASE WHEN ? IS NOT NULL THEN DATE(?, '+180 days') ELSE NULL END,
                                        CASE WHEN ? IS NOT NULL THEN DATE(?, '+365 days') ELSE NULL END)
                                ''', (
                                    data.get('sap_material_no'),
                                    data.get('bfm_equipment_no'),
                                    data.get('description'),
                                    data.get('tool_id_drawing_no'),
                                    data.get('location'),
                                    data.get('master_lin'),
                                    data.get('monthly_pm', 1),
                                    data.get('six_month_pm', 1),
                                    data.get('annual_pm', 1),
                                    data.get('last_monthly_pm'),
                                    data.get('last_six_month_pm'),
                                    data.get('last_annual_pm'),
                                    data.get('last_monthly_pm'),
                                    data.get('last_monthly_pm'),
                                    data.get('last_six_month_pm'),
                                    data.get('last_six_month_pm'),
                                    data.get('last_annual_pm'),
                                    data.get('last_annual_pm')
                                ))
                                imported_count += 1
                            else:
                                error_count += 1
                            
                        except Exception as e:
                            print(f"Error importing row {index}: {e}")
                            error_count += 1
                            continue
                
                    self.conn.commit()
                    dialog.destroy()
                
                    # Show results
                    result_msg = f"Import completed!\n\n"
                    result_msg += f"✅ Successfully imported: {imported_count} records\n"
                    if error_count > 0:
                        result_msg += f"⚠️ Skipped (errors): {error_count} records\n"
                    result_msg += f"\nTotal processed: {imported_count + error_count} records"
                
                    messagebox.showinfo("Import Results", result_msg)
                    self.refresh_equipment_list()
                    self.update_status(f"Imported {imported_count} equipment records")
                
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to process import: {str(e)}")
        
            def cancel_import():
                    """Cancel the import process"""
                    dialog.destroy()
        
            # 🎯 BUTTONS FRAME - This was missing!
            button_frame = ttk.Frame(scrollable_frame)
            button_frame.pack(side='bottom', fill='x', padx=20, pady=20)
        
            # Import button (green)
            import_button = ttk.Button(button_frame, text="✅ Import with These Mappings", 
                                    command=process_import)
            import_button.pack(side='left', padx=10)
        
            # Cancel button
            cancel_button = ttk.Button(button_frame, text="❌ Cancel", 
                                    command=cancel_import)
            cancel_button.pack(side='right', padx=10)
        
            # Pack the canvas and scrollbar
            main_canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        
            # Make the dialog modal
            dialog.focus_set()
            dialog.grab_set()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV file: {str(e)}")
            return
    
    
    
    def add_equipment_dialog(self):
        """Dialog to add new equipment"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Equipment")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form fields
        fields = [
            ("SAP Material No:", tk.StringVar()),
            ("BFM Equipment No:", tk.StringVar()),
            ("Description:", tk.StringVar()),
            ("Tool ID/Drawing No:", tk.StringVar()),
            ("Location:", tk.StringVar()),
            ("Master LIN:", tk.StringVar())
        ]
        
        entries = {}
        
        for i, (label, var) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=i, column=0, sticky='w', padx=10, pady=5)
            entry = ttk.Entry(dialog, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[label] = var
        
        # PM type checkboxes
        pm_frame = ttk.LabelFrame(dialog, text="PM Types", padding=10)
        pm_frame.grid(row=len(fields), column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        monthly_var = tk.BooleanVar(value=True)
        six_month_var = tk.BooleanVar(value=True)
        annual_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(pm_frame, text="Monthly PM", variable=monthly_var).pack(anchor='w')
        ttk.Checkbutton(pm_frame, text="Six Month PM", variable=six_month_var).pack(anchor='w')
        ttk.Checkbutton(pm_frame, text="Annual PM", variable=annual_var).pack(anchor='w')
        
        def save_equipment():
            try:
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO equipment 
                    (sap_material_no, bfm_equipment_no, description, tool_id_drawing_no, 
                     location, master_lin, monthly_pm, six_month_pm, annual_pm)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    entries["SAP Material No:"].get(),
                    entries["BFM Equipment No:"].get(),
                    entries["Description:"].get(),
                    entries["Tool ID/Drawing No:"].get(),
                    entries["Location:"].get(),
                    entries["Master LIN:"].get(),
                    monthly_var.get(),
                    six_month_var.get(),
                    annual_var.get()
                ))
                self.conn.commit()
                messagebox.showinfo("Success", "Equipment added successfully!")
                dialog.destroy()
                self.refresh_equipment_list()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add equipment: {str(e)}")
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Save", command=save_equipment).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='left', padx=5)
    
    def edit_equipment_dialog(self):
        """Enhanced dialog to edit existing equipment with Run to Failure option"""
        selected = self.equipment_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select equipment to edit")
            return
    
        # Get selected equipment data
        item = self.equipment_tree.item(selected[0])
        bfm_no = item['values'][1]  # BFM Equipment No.
    
        # Fetch full equipment data
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
        equipment_data = cursor.fetchone()
    
        if not equipment_data:
            messagebox.showerror("Error", "Equipment not found in database")
            return
    
        # Create edit dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Equipment")
        dialog.geometry("500x500")  # Made taller for additional options
        dialog.transient(self.root)
        dialog.grab_set()
    
        # Pre-populate fields
        fields = [
            ("SAP Material No:", tk.StringVar(value=equipment_data[1] or '')),
            ("BFM Equipment No:", tk.StringVar(value=equipment_data[2] or '')),
            ("Description:", tk.StringVar(value=equipment_data[3] or '')),
            ("Tool ID/Drawing No:", tk.StringVar(value=equipment_data[4] or '')),
            ("Location:", tk.StringVar(value=equipment_data[5] or '')),
            ("Master LIN:", tk.StringVar(value=equipment_data[6] or ''))
        ]
    
        entries = {}
    
        for i, (label, var) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=i, column=0, sticky='w', padx=10, pady=5)
            entry = ttk.Entry(dialog, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[label] = var
    
        # PM type checkboxes and Run to Failure option
        pm_frame = ttk.LabelFrame(dialog, text="PM Types & Equipment Status", padding=10)
        pm_frame.grid(row=len(fields), column=0, columnspan=2, padx=10, pady=10, sticky='ew')
    
        # Current equipment status
        current_status = equipment_data[16] or 'Active'  # Status field
    
        # PM checkboxes (disabled if currently Run to Failure)
        monthly_var = tk.BooleanVar(value=bool(equipment_data[7]))
        six_month_var = tk.BooleanVar(value=bool(equipment_data[8]))
        annual_var = tk.BooleanVar(value=bool(equipment_data[9]))
    
        monthly_cb = ttk.Checkbutton(pm_frame, text="Monthly PM", variable=monthly_var)
        monthly_cb.pack(anchor='w')
    
        six_month_cb = ttk.Checkbutton(pm_frame, text="Six Month PM", variable=six_month_var)
        six_month_cb.pack(anchor='w')
    
        annual_cb = ttk.Checkbutton(pm_frame, text="Annual PM", variable=annual_var)
        annual_cb.pack(anchor='w')
    
        # Separator
        ttk.Separator(pm_frame, orient='horizontal').pack(fill='x', pady=10)
    
        # Run to Failure option
        run_to_failure_var = tk.BooleanVar(value=(current_status == 'Run to Failure'))
        rtf_cb = ttk.Checkbutton(pm_frame, text="🔧 Set as Run to Failure Equipment", 
                                variable=run_to_failure_var,
                                command=lambda: toggle_pm_options())
        rtf_cb.pack(anchor='w', pady=5)
    
        # Status info
        status_label = ttk.Label(pm_frame, text=f"Current Status: {current_status}", 
                                font=('Arial', 9, 'italic'))
        status_label.pack(anchor='w', pady=2)
    
        def toggle_pm_options():
            """Enable/disable PM options based on Run to Failure selection"""
            if run_to_failure_var.get():
                # Disable PM options when Run to Failure is selected
                monthly_cb.config(state='disabled')
                six_month_cb.config(state='disabled') 
                annual_cb.config(state='disabled')
                monthly_var.set(False)
                six_month_var.set(False)
                annual_var.set(False)
                status_label.config(text="Status: Will be set to Run to Failure", foreground='red')
            else:
                # Enable PM options when Run to Failure is not selected
                monthly_cb.config(state='normal')
                six_month_cb.config(state='normal')
                annual_cb.config(state='normal')
                status_label.config(text="Status: Will be set to Active", foreground='green')
    
        # Initialize the toggle state
        toggle_pm_options()
    
        def update_equipment():
            """Update equipment with enhanced Run to Failure handling"""
            try:
                cursor = self.conn.cursor()
            
                # Determine new status
                new_status = 'Run to Failure' if run_to_failure_var.get() else 'Active'
            
                # Update equipment record
                cursor.execute('''
                    UPDATE equipment SET
                    sap_material_no = ?, description = ?, tool_id_drawing_no = ?, 
                    location = ?, master_lin = ?, monthly_pm = ?, six_month_pm = ?, annual_pm = ?,
                    status = ?, updated_date = CURRENT_TIMESTAMP
                    WHERE bfm_equipment_no = ?
                ''', (
                    entries["SAP Material No:"].get(),
                    entries["Description:"].get(),
                    entries["Tool ID/Drawing No:"].get(),
                    entries["Location:"].get(),
                    entries["Master LIN:"].get(),
                    monthly_var.get() and not run_to_failure_var.get(),  # Disable PMs if RTF
                    six_month_var.get() and not run_to_failure_var.get(),
                    annual_var.get() and not run_to_failure_var.get(),
                    new_status,
                    bfm_no
                ))
            
                # If changing TO Run to Failure, add entry to run_to_failure_assets table
                if run_to_failure_var.get() and current_status != 'Run to Failure':
                    cursor.execute('''
                        INSERT OR REPLACE INTO run_to_failure_assets 
                        (bfm_equipment_no, description, location, technician_name, completion_date, labor_hours, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        bfm_no,
                        entries["Description:"].get(),
                        entries["Location:"].get(),
                        'System Change',  # Default technician
                        datetime.now().strftime('%Y-%m-%d'),
                        0.0,
                        f'Equipment manually set to Run to Failure status via equipment edit dialog'
                    ))
            
                # If changing FROM Run to Failure back to Active, remove from run_to_failure_assets
                elif not run_to_failure_var.get() and current_status == 'Run to Failure':
                    cursor.execute('DELETE FROM run_to_failure_assets WHERE bfm_equipment_no = ?', (bfm_no,))
            
                self.conn.commit()
            
                # Show appropriate success message
                if run_to_failure_var.get():
                    success_msg = f"Equipment {bfm_no} updated successfully!\n\nStatus changed to: Run to Failure\n"
                    success_msg += "- All PM requirements disabled\n"
                    success_msg += "- Equipment moved to Run to Failure tab\n"
                    success_msg += "- No future PMs will be scheduled"
                else:
                    success_msg = f"Equipment {bfm_no} updated successfully!\n\nStatus: Active"
            
                messagebox.showinfo("Success", success_msg)
                dialog.destroy()
            
                # Refresh all relevant displays
                self.refresh_equipment_list()
                self.load_run_to_failure_assets()  # Refresh Run to Failure tab
                self.update_equipment_statistics()  # Update statistics
            
                # Update status bar
                if run_to_failure_var.get():
                    self.update_status(f"Equipment {bfm_no} set to Run to Failure")
                else:
                    self.update_status(f"Equipment {bfm_no} reactivated from Run to Failure")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update equipment: {str(e)}")
    
        # Buttons with enhanced styling
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=15)
    
        update_btn = ttk.Button(button_frame, text="✅ Update Equipment", command=update_equipment)
        update_btn.pack(side='left', padx=10)
    
        cancel_btn = ttk.Button(button_frame, text="❌ Cancel", command=dialog.destroy)
        cancel_btn.pack(side='left', padx=5)
    
        # Add warning label for Run to Failure
        warning_frame = ttk.Frame(dialog)
        warning_frame.grid(row=len(fields)+2, column=0, columnspan=2, padx=10, pady=5)
    
        warning_text = "⚠️ Run to Failure equipment will not be scheduled for regular PMs"
        warning_label = ttk.Label(warning_frame, text=warning_text, 
                             font=('Arial', 8, 'italic'), foreground='orange')
        warning_label.pack()
    
    def refresh_equipment_list(self):
        """Refresh equipment list display"""
        try:
            self.load_equipment_data()
        
            # Clear existing items
            for item in self.equipment_tree.get_children():
                self.equipment_tree.delete(item)
        
            # Add equipment to tree
            for equipment in self.equipment_data:
                if len(equipment) >= 9:
                    self.equipment_tree.insert('', 'end', values=(
                        equipment[1] or '',  # SAP
                        equipment[2] or '',  # BFM
                        equipment[3] or '',  # Description
                        equipment[5] or '',  # Location
                        equipment[6] or '',  # Master LIN
                        'Yes' if equipment[7] else 'No',  # Monthly PM
                        'Yes' if equipment[8] else 'No',  # Six Month PM
                        'Yes' if equipment[9] else 'No',  # Annual PM
                        equipment[16] or 'Active'  # Status
                    ))
        
            # Update statistics
            self.update_equipment_statistics()
        
            # Update status
            self.update_status(f"Equipment list refreshed - {len(self.equipment_data)} items")
        
        except Exception as e:
            print(f"Error refreshing equipment list: {e}")
            messagebox.showerror("Error", f"Failed to refresh equipment list: {str(e)}")
    
    def filter_equipment_list(self, *args):
        """Filter equipment list based on search term"""
        search_term = self.equipment_search_var.get().lower()
        
        # Clear existing items
        for item in self.equipment_tree.get_children():
            self.equipment_tree.delete(item)
        
        # Add filtered equipment
        for equipment in self.equipment_data:
            if len(equipment) >= 9:
                # Check if search term matches any field
                searchable_fields = [
                    equipment[1] or '',  # SAP
                    equipment[2] or '',  # BFM
                    equipment[3] or '',  # Description
                    equipment[5] or '',  # Location
                    equipment[6] or ''   # Master LIN
                ]
                
                if not search_term or any(search_term in field.lower() for field in searchable_fields):
                    self.equipment_tree.insert('', 'end', values=(
                        equipment[1] or '',  # SAP
                        equipment[2] or '',  # BFM
                        equipment[3] or '',  # Description
                        equipment[5] or '',  # Location
                        equipment[6] or '',  # Master LIN
                        'Yes' if equipment[7] else 'No',  # Monthly PM
                        'Yes' if equipment[8] else 'No',  # Six Month PM
                        'Yes' if equipment[9] else 'No',  # Annual PM
                        equipment[16] or 'Active'  # Status
                    ))
    
    def export_equipment_list(self):
        """Export equipment list to CSV"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export Equipment List",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if file_path:
                cursor = self.conn.cursor()
                cursor.execute('SELECT * FROM equipment ORDER BY bfm_equipment_no')
                equipment_data = cursor.fetchall()
                
                # Create DataFrame
                columns = ['ID', 'SAP Material No', 'BFM Equipment No', 'Description', 
                          'Tool ID/Drawing No', 'Location', 'Master LIN', 'Monthly PM', 
                          'Six Month PM', 'Annual PM', 'Last Monthly PM', 'Last Six Month PM', 
                          'Last Annual PM', 'Next Monthly PM', 'Next Six Month PM', 
                          'Next Annual PM', 'Status', 'Created Date', 'Updated Date']
                
                df = pd.DataFrame(equipment_data, columns=columns)
                df.to_csv(file_path, index=False)
                
                messagebox.showinfo("Success", f"Equipment list exported to {file_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export equipment list: {str(e)}")
    
    def load_equipment_data(self):
        """Load equipment data from database"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT * FROM equipment ORDER BY bfm_equipment_no')
            self.equipment_data = cursor.fetchall()
        except Exception as e:
            print(f"Error loading equipment data: {e}")
            self.equipment_data = []
    
    def generate_weekly_assignments(self):
        """
        SIMPLIFIED weekly assignment generation
        """
        try:
            week_start = datetime.strptime(self.week_start_var.get(), '%Y-%m-%d')
            week_end = week_start + timedelta(days=6)
            cursor = self.conn.cursor()
        
            print(f"DEBUG: Generating assignments for week {week_start.strftime('%Y-%m-%d')}")
        
            # Clear existing assignments for this week
            cursor.execute('DELETE FROM weekly_pm_schedules WHERE week_start_date = ?', 
                        (week_start.strftime('%Y-%m-%d'),))
        
            # Get all active equipment
            cursor.execute('''
                SELECT e.bfm_equipment_no, e.description, e.monthly_pm, e.annual_pm,
                    e.last_monthly_pm, e.last_annual_pm
                FROM equipment e
                WHERE e.status = 'Active' 
                AND e.status != 'Run to Failure' 
                AND e.status != 'Missing'
                ORDER BY e.bfm_equipment_no
            ''')
            equipment_list = cursor.fetchall()
        
            # Generate PM assignments with SIMPLIFIED logic
            pm_assignments = []
            assigned_assets = set()
        
            for equipment in equipment_list:
                (bfm_no, description, monthly, annual, last_monthly, last_annual) = equipment
            
                # Skip if already assigned
                if bfm_no in assigned_assets:
                    continue
            
                assigned_pm = None
            
                # SIMPLIFIED PRIORITY LOGIC:
                # 1. If never done, prioritize Monthly
                # 2. If Monthly is overdue, assign Monthly
                # 3. If Annual is significantly overdue, assign Annual
            
                if monthly and (not last_monthly or self.is_pm_overdue(last_monthly, 30)):
                    assigned_pm = 'Monthly'
                elif annual and (not last_annual or self.is_pm_overdue(last_annual, 365)):
                    # Only assign Annual if Monthly is current or doesn't exist
                    if not monthly or (last_monthly and not self.is_pm_overdue(last_monthly, 30)):
                        assigned_pm = 'Annual'
            
                if assigned_pm:
                    pm_assignments.append((bfm_no, assigned_pm, description))
                    assigned_assets.add(bfm_no)
                    print(f"DEBUG: Assigned {bfm_no} - {assigned_pm} PM")
        
            print(f"DEBUG: Total assignments: {len(pm_assignments)}")
        
            # Distribute among technicians (simplified)
            total_pms = min(self.weekly_pm_target, len(pm_assignments))
            pms_per_tech = total_pms // len(self.technicians)
            extra_pms = total_pms % len(self.technicians)
        
            assignment_index = 0
        
            for tech_index, technician in enumerate(self.technicians):
                tech_pm_count = pms_per_tech + (1 if tech_index < extra_pms else 0)
            
                for _ in range(tech_pm_count):
                    if assignment_index < len(pm_assignments):
                        bfm_no, pm_type, description = pm_assignments[assignment_index]
                    
                        # Schedule throughout the week
                        day_offset = (assignment_index % 5)
                        scheduled_date = week_start + timedelta(days=day_offset)
                    
                        cursor.execute('''
                            INSERT INTO weekly_pm_schedules 
                            (week_start_date, bfm_equipment_no, pm_type, assigned_technician, scheduled_date)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (
                            week_start.strftime('%Y-%m-%d'),
                            bfm_no,
                            pm_type,
                            technician,
                            scheduled_date.strftime('%Y-%m-%d')
                        ))
                    
                        assignment_index += 1
        
            self.conn.commit()
        
            # Show results
            messagebox.showinfo("Scheduling Complete", 
                            f"Generated {total_pms} PM assignments for week {week_start.strftime('%Y-%m-%d')}\n\n"
                            f"Unique assets: {len(assigned_assets)}")
        
            # Refresh displays
            self.refresh_technician_schedules()
            self.update_status(f"Generated {total_pms} PM assignments")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate assignments: {str(e)}")
            import traceback
            traceback.print_exc()

    def is_pm_overdue(self, last_pm_date, frequency_days):
        """
        Simple overdue check
        """
        if not last_pm_date:
            return True  # Never done = overdue
        
        try:
            standardizer = DateStandardizer(self.conn)
            parsed_date = standardizer.parse_date_flexible(last_pm_date)
        
            if parsed_date:
                last_date = datetime.strptime(parsed_date, '%Y-%m-%d')
                days_since = (datetime.now() - last_date).days
                return days_since >= frequency_days * 0.8  # 80% of frequency = overdue
            else:
                return True  # Can't parse = assume overdue
            
        except Exception:
            return True  # Error = assume overdue

    def check_pm_scheduling_status_comprehensive(self, cursor, bfm_no, pm_type, last_pm_date, next_pm_date, 
                              frequency_days, week_start, week_end):
        """
        COMPREHENSIVE validation against ALL completed PMs
        """
        try:
            current_date = datetime.now()
        
            status = {
                'schedule': False,
                'conflict': False,
                'reason': '',
                'last_date': last_pm_date,
                'days_since': None,
                'days_overdue': None
            }
        
            print(f"DEBUG: Comprehensive check {bfm_no} {pm_type} PM")
        
            # Check 1: Get ALL recent PM completions for this equipment (any type)
            cursor.execute('''
                SELECT pm_type, completion_date, technician_name
                FROM pm_completions 
                WHERE bfm_equipment_no = ?
                AND completion_date >= DATE(?, '-30 days')  -- Check last 30 days
                ORDER BY completion_date DESC
            ''', (bfm_no, current_date.strftime('%Y-%m-%d')))
        
            all_recent_completions = cursor.fetchall()
        
            # Check if THIS specific PM type was completed recently
            for completed_pm_type, completion_date, technician in all_recent_completions:
                if completed_pm_type == pm_type:
                    try:
                        comp_date_obj = datetime.strptime(completion_date, '%Y-%m-%d')
                        days_since_completion = (current_date - comp_date_obj).days
                    
                        # Don't reschedule if completed within minimum interval
                        min_interval = max(frequency_days * 0.75, 14)  # 75% of frequency or 14 days minimum
                    
                        if days_since_completion < min_interval:
                            status['conflict'] = True
                            status['reason'] = f"{pm_type} PM completed {days_since_completion} days ago by {technician} (min interval: {min_interval} days)"
                            return status
                        
                    except Exception as e:
                        print(f"DEBUG: Error parsing completion date {completion_date}: {e}")
        
            # Check 2: Cross-validation between PM types
            if pm_type == 'Annual':
                # If Annual PM is being scheduled, check if Monthly was completed very recently
                for completed_pm_type, completion_date, technician in all_recent_completions:
                    if completed_pm_type == 'Monthly':
                        try:
                            comp_date_obj = datetime.strptime(completion_date, '%Y-%m-%d')
                            days_since_monthly = (current_date - comp_date_obj).days
                        
                            # If Monthly was completed less than 7 days ago, don't schedule Annual yet
                            if days_since_monthly < 7:
                                status['conflict'] = True
                                status['reason'] = f"Annual PM blocked - Monthly PM completed {days_since_monthly} days ago by {technician}"
                                return status
                            
                        except Exception as e:
                            print(f"DEBUG: Error parsing monthly completion date: {e}")
        
            elif pm_type == 'Monthly':
                # If Monthly PM is being scheduled, check if Annual was completed recently
                for completed_pm_type, completion_date, technician in all_recent_completions:
                    if completed_pm_type == 'Annual':
                        try:
                            comp_date_obj = datetime.strptime(completion_date, '%Y-%m-%d')
                            days_since_annual = (current_date - comp_date_obj).days
                        
                            # If Annual was completed less than 30 days ago, Monthly might not be needed
                            if days_since_annual < 30:
                                status['conflict'] = True
                                status['reason'] = f"Monthly PM blocked - Annual PM completed {days_since_annual} days ago by {technician}"
                                return status
                            
                        except Exception as e:
                            print(f"DEBUG: Error parsing annual completion date: {e}")
        
            # Check 3: Validate against equipment table dates
            if last_pm_date:
                try:
                    standardizer = DateStandardizer(self.conn)
                    parsed_last_date = standardizer.parse_date_flexible(last_pm_date)
                
                    if parsed_last_date:
                        last_date_obj = datetime.strptime(parsed_last_date, '%Y-%m-%d')
                        days_since_last = (current_date - last_date_obj).days
                        status['days_since'] = days_since_last
                    
                        # Check if actually due
                        if days_since_last >= frequency_days * 0.8:  # 80% of frequency
                            status['days_overdue'] = days_since_last - frequency_days
                        
                            # Additional check: make sure completion records align
                            equipment_date_conflict = False
                            for completed_pm_type, completion_date, technician in all_recent_completions:
                                if completed_pm_type == pm_type:
                                    try:
                                        comp_date_obj = datetime.strptime(completion_date, '%Y-%m-%d')
                                        if comp_date_obj > last_date_obj:
                                            print(f"DEBUG: Equipment table out of sync - completion record newer than equipment date")
                                            equipment_date_conflict = True
                                            break
                                    except:
                                        pass
                        
                            if equipment_date_conflict:
                                status['conflict'] = True
                                status['reason'] = f"{pm_type} PM data inconsistency - equipment table vs completion records"
                                return status
                        
                            status['schedule'] = True
                            status['reason'] = f"{pm_type} PM due (completed {days_since_last} days ago, overdue by {status['days_overdue']} days)"
                            return status
                        else:
                            days_until_due = frequency_days - days_since_last
                            if days_until_due <= 14:  # Allow up to 2 weeks early
                                status['schedule'] = True
                                status['reason'] = f"{pm_type} PM due soon ({days_until_due} days early)"
                                return status
                            else:
                                status['conflict'] = True
                                status['reason'] = f"{pm_type} PM not due for {days_until_due} days"
                                return status
                
                except Exception as e:
                    print(f"DEBUG: Error processing equipment table date: {e}")
        
            # Check 4: Never done equipment - prioritize Monthly
            if not last_pm_date:
                if pm_type == 'Monthly':
                    status['schedule'] = True
                    status['reason'] = f"Monthly PM never completed - FIRST TIME PRIORITY"
                    status['days_since'] = 9999
                    return status
                elif pm_type == 'Annual':
                    # Check if Monthly is also never done
                    cursor.execute('SELECT last_monthly_pm FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
                    monthly_result = cursor.fetchone()
                    if monthly_result and not monthly_result[0]:
                        status['conflict'] = True
                        status['reason'] = f"Annual PM blocked - Monthly PM should be completed first for new equipment"
                        return status
                    else:
                        status['schedule'] = True
                        status['reason'] = f"Annual PM never completed - HIGH PRIORITY"
                        status['days_since'] = 9999
                        return status
        
            # Default: allow scheduling if no conflicts found
            status['schedule'] = True
            status['reason'] = f"{pm_type} PM cleared for scheduling"
            return status
        
        except Exception as e:
            print(f"ERROR: Exception in comprehensive check for {bfm_no}: {e}")
            # On error, don't schedule to be safe
            status['conflict'] = True
            status['reason'] = f"{pm_type} PM validation error - blocked for safety"
            return status


    def validate_against_recent_completions(self, week_start):
        """
        Specific validation against your recent completions from 2025-09-04
        """
        try:
            cursor = self.conn.cursor()
        
            # Get all completions from the problematic date
            cursor.execute('''
                SELECT bfm_equipment_no, pm_type, technician_name
                FROM pm_completions 
                WHERE completion_date = '2025-09-04'
                ORDER BY bfm_equipment_no
            ''')
        
            sept_4_completions = cursor.fetchall()
        
            if not sept_4_completions:
                return []  # No completions on that date
        
            print(f"DEBUG: Found {len(sept_4_completions)} completions on 2025-09-04")
        
            # Check if any of these are being rescheduled
            conflicts = []
            cursor.execute('''
                SELECT bfm_equipment_no, pm_type 
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
            ''', (week_start.strftime('%Y-%m-%d'),))
        
            current_schedule = cursor.fetchall()
        
            for scheduled_bfm, scheduled_pm_type in current_schedule:
                for completed_bfm, completed_pm_type, technician in sept_4_completions:
                    if (scheduled_bfm == completed_bfm and 
                        (scheduled_pm_type == completed_pm_type or 
                        (scheduled_pm_type == 'Annual' and completed_pm_type in ['Monthly', 'Annual']))):
                    
                        conflicts.append({
                            'bfm_no': scheduled_bfm,
                            'scheduled_type': scheduled_pm_type,
                            'completed_type': completed_pm_type,
                            'technician': technician
                        })
        
            return conflicts
        
        except Exception as e:
            print(f"ERROR: Error checking recent completions: {e}")
            return []


    def add_comprehensive_validation_to_generate_weekly_assignments(self):
        """
        Add this validation call to your generate_weekly_assignments method
        """
        # Add this code BEFORE the assignment generation in generate_weekly_assignments
    
        # Validate against recent completions
        week_start = datetime.strptime(self.week_start_var.get(), '%Y-%m-%d')
        completion_conflicts = self.validate_against_recent_completions(week_start)
        
        if completion_conflicts:
            conflict_msg = f"POTENTIAL DUPLICATE SCHEDULING DETECTED:\n\n"
            conflict_msg += f"The following equipment was completed on 2025-09-04 but is being rescheduled:\n\n"
        
            for conflict in completion_conflicts[:10]:  # Show first 10
                conflict_msg += f"• {conflict['bfm_no']}: Scheduling {conflict['scheduled_type']} PM, but {conflict['completed_type']} PM was completed by {conflict['technician']}\n"
        
            if len(completion_conflicts) > 10:
                conflict_msg += f"\n... and {len(completion_conflicts) - 10} more conflicts"
        
            result = messagebox.askyesno(
                "Duplicate Scheduling Warning",
                f"{conflict_msg}\n\n"
                f"Do you want to proceed anyway?\n\n"
                f"Click 'No' to cancel and review the issues.",
                icon='warning'
            )
        
            if not result:
                self.update_status("Scheduling cancelled due to duplicate PM conflicts")
                return False
    
        return True  # Continue with scheduling


    def validate_weekly_schedule_before_generation(self):
        """
        SIMPLIFIED validation that's less restrictive
        """
        try:
            week_start = datetime.strptime(self.week_start_var.get(), '%Y-%m-%d')
            cursor = self.conn.cursor()
        
            validation_issues = []
        
            # Check 1: Equipment with very recent completions (last 3 days only)
            cursor.execute('''
                SELECT pc.bfm_equipment_no, pc.pm_type, pc.completion_date, pc.technician_name
                FROM pm_completions pc
                WHERE pc.completion_date >= DATE(?, '-3 days')  -- Only last 3 days
                ORDER BY pc.completion_date DESC
                LIMIT 20
            ''', (week_start.strftime('%Y-%m-%d'),))
        
            recent_completions = cursor.fetchall()
        
            if recent_completions:
                validation_issues.append("VERY RECENT COMPLETIONS (Last 3 days):")
                for bfm_no, pm_type, comp_date, tech in recent_completions[:10]:
                    validation_issues.append(f"  • {bfm_no} - {pm_type} PM completed {comp_date} by {tech}")
        
            # Check 2: Already scheduled for this exact week
            cursor.execute('''
                SELECT ws.bfm_equipment_no, ws.pm_type, ws.assigned_technician, ws.status
                FROM weekly_pm_schedules ws
                WHERE ws.week_start_date = ?
            ''', (week_start.strftime('%Y-%m-%d'),))
        
            current_week_schedules = cursor.fetchall()
        
            if current_week_schedules:
                validation_issues.append(f"\nALREADY SCHEDULED FOR THIS WEEK ({len(current_week_schedules)} items):")
                validation_issues.append("  This will clear existing schedules and regenerate.")
        
            # Show simplified validation results
            if validation_issues:
                issues_text = "\n".join(validation_issues)
                result = messagebox.askyesno(
                    "Quick Validation Check",
                    f"{issues_text}\n\n"
                    f"Continue with scheduling?\n\n"
                    f"Note: This will proceed with enhanced duplicate prevention.",
                    icon='info'
                )
            
                if result:
                    # IMPORTANT: Actually call the scheduling method
                    self.generate_weekly_assignments()
            
                return result
            else:
                messagebox.showinfo("Validation Passed", "No major conflicts detected. Proceeding with scheduling.")
                # IMPORTANT: Actually call the scheduling method
                self.generate_weekly_assignments()
                return True
            
        except Exception as e:
            messagebox.showerror("Validation Error", f"Error during validation: {str(e)}")
            return False


    def verify_no_duplicates(self):
        """Verification method to check for duplicate asset assignments"""
        try:
            week_start = self.week_start_var.get()
            cursor = self.conn.cursor()
        
            # Check for duplicate asset assignments in current week
            cursor.execute('''
                SELECT bfm_equipment_no, COUNT(*) as assignment_count, 
                    GROUP_CONCAT(pm_type || ' (' || assigned_technician || ')') as assignments
                FROM weekly_pm_schedules 
                WHERE week_start_date = ?
                GROUP BY bfm_equipment_no
                HAVING COUNT(*) > 1
            ''', (week_start,))
        
            duplicates = cursor.fetchall()
        
            if duplicates:
                error_msg = f"DUPLICATE ASSIGNMENTS FOUND!\n\n"
                for bfm_no, count, assignments in duplicates:
                    error_msg += f"Asset {bfm_no}: {count} assignments\n"
                    error_msg += f"  Details: {assignments}\n\n"
            
                messagebox.showerror("Duplicate Assignments Detected", error_msg)
                return False
            else:
                messagebox.showinfo("Verification Passed", 
                                f"No duplicate asset assignments found for week {week_start}")
                return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to verify assignments: {str(e)}")
            return False
    
    def is_pm_due(self, last_pm_date, frequency_days, current_week_start, bfm_no=None, pm_type=None):
        """
        Enhanced PM due date checking with comprehensive validation
    
        Args:
            last_pm_date: Last completion date string
            frequency_days: PM frequency in days (30, 180, 365)
            current_week_start: Week start date for scheduling
            bfm_no: Equipment number (optional, for enhanced checking)
            pm_type: PM type (optional, for enhanced checking)
    
        Returns:
            bool: True if PM is due and safe to schedule
        """
        try:
            # Basic check - never done
            if not last_pm_date:
                return True
        
            # Parse the last PM date with flexible parsing
            standardizer = DateStandardizer(self.conn)
            parsed_date = standardizer.parse_date_flexible(last_pm_date)
        
            if not parsed_date:
                print(f"DEBUG: Could not parse date '{last_pm_date}' for {bfm_no} {pm_type}")
                return True  # If can't parse, assume it's due
        
            # Calculate days since last PM
            last_date = datetime.strptime(parsed_date, '%Y-%m-%d')
            current_date = datetime.now()
            days_since = (current_date - last_date).days
        
            # Enhanced validation if equipment details provided
            if bfm_no and pm_type:
                cursor = self.conn.cursor()
            
                # Check if completed too recently (minimum 75% of frequency)
                min_interval = max(frequency_days * 0.75, 14)
                if days_since < min_interval:
                    print(f"DEBUG: {bfm_no} {pm_type} PM too recent - {days_since} days (min: {min_interval})")
                    return False
            
                # Check recent completion records for verification
                cursor.execute('''
                    SELECT completion_date 
                    FROM pm_completions 
                    WHERE bfm_equipment_no = ? AND pm_type = ?
                    ORDER BY completion_date DESC LIMIT 1
                ''', (bfm_no, pm_type))
            
                recent_completion = cursor.fetchone()
                if recent_completion and recent_completion[0]:
                    try:
                        comp_date = datetime.strptime(recent_completion[0], '%Y-%m-%d')
                        days_since_completion = (current_date - comp_date).days
                    
                        # If completion record is more recent than equipment table, use it
                        if comp_date > last_date:
                            days_since = days_since_completion
                            print(f"DEBUG: Using completion record for {bfm_no} - {days_since} days since completion")
                
                    except Exception as e:
                        print(f"DEBUG: Error parsing completion date: {e}")
            
                # Check if already scheduled recently
                cursor.execute('''
                    SELECT week_start_date, status 
                    FROM weekly_pm_schedules 
                    WHERE bfm_equipment_no = ? AND pm_type = ?
                    AND week_start_date >= DATE(?, '-21 days')
                    ORDER BY week_start_date DESC LIMIT 1
                ''', (bfm_no, pm_type, current_week_start.strftime('%Y-%m-%d')))
            
                recent_schedule = cursor.fetchone()
                if recent_schedule:
                    week_date, status = recent_schedule
                    if status == 'Completed':
                        print(f"DEBUG: {bfm_no} {pm_type} PM already completed in week {week_date}")
                        return False
                    elif status == 'Scheduled' and week_date != current_week_start.strftime('%Y-%m-%d'):
                        print(f"DEBUG: {bfm_no} {pm_type} PM already scheduled for week {week_date}")
                        return False
            
            # Standard due date logic
            next_due_date = last_date + timedelta(days=frequency_days)
            week_end = current_week_start + timedelta(days=6)
        
            # PM is due if next due date is within or before current week
            is_due = next_due_date <= week_end
        
            if is_due:
                overdue_days = (current_date - next_due_date).days
                print(f"DEBUG: {bfm_no or 'Equipment'} {pm_type or 'PM'} is due (overdue by {overdue_days} days)")
        
            return is_due
        
        except Exception as e:
            print(f"ERROR: Exception in is_pm_due for {bfm_no} {pm_type}: {e}")
            # On error, assume it's due to be safe (but log the error)
            return True


    def is_pm_due_simple(self, last_pm_date, frequency_days, current_week_start):
        """
        Simplified version for backward compatibility
        Use this if you don't want to update all existing calls to is_pm_due
        """
        return self.is_pm_due(last_pm_date, frequency_days, current_week_start)
    
    def refresh_technician_schedules(self):
        """Refresh all technician schedule displays"""
        week_start = self.week_start_var.get()
        
        for technician, tree in self.technician_trees.items():
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
            
            # Load scheduled PMs for this technician
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT ws.bfm_equipment_no, e.description, ws.pm_type, ws.scheduled_date, ws.status
                FROM weekly_pm_schedules ws
                JOIN equipment e ON ws.bfm_equipment_no = e.bfm_equipment_no
                WHERE ws.assigned_technician = ? AND ws.week_start_date = ?
                ORDER BY ws.scheduled_date
            ''', (technician, week_start))
            
            assignments = cursor.fetchall()
            
            for assignment in assignments:
                bfm_no, description, pm_type, scheduled_date, status = assignment
                tree.insert('', 'end', values=(bfm_no, description, pm_type, scheduled_date, status))
    
    def print_weekly_pm_forms(self):
        """Generate and print PM forms for the week"""
        try:
            week_start = self.week_start_var.get()
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Create directory for PM forms
            forms_dir = f"PM_Forms_Week_{week_start}_{timestamp}"
            os.makedirs(forms_dir, exist_ok=True)
            
            cursor = self.conn.cursor()
            
            # Generate forms for each technician
            for technician in self.technicians:
                cursor.execute('''
                    SELECT ws.bfm_equipment_no, e.sap_material_no, e.description, e.tool_id_drawing_no,
                           e.location, e.master_lin, ws.pm_type, ws.scheduled_date
                    FROM weekly_pm_schedules ws
                    JOIN equipment e ON ws.bfm_equipment_no = e.bfm_equipment_no
                    WHERE ws.assigned_technician = ? AND ws.week_start_date = ?
                    ORDER BY ws.scheduled_date
                ''', (technician, week_start))
                
                assignments = cursor.fetchall()
                
                if assignments:
                    # Create PDF for this technician
                    filename = os.path.join(forms_dir, f"{technician.replace(' ', '_')}_PM_Forms.pdf")
                    self.create_pm_forms_pdf(filename, technician, assignments)
            
            messagebox.showinfo("Success", f"PM forms generated in directory: {forms_dir}")
            self.update_status(f"PM forms generated for week {week_start}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PM forms: {str(e)}")
    
    def create_pm_forms_pdf(self, filename, technician, assignments):
        """Create PDF with PM forms for a technician - ENHANCED WITH CUSTOM TEMPLATES"""
        try:
            doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=36, leftMargin=36,  # Reduced margins
                                topMargin=36, bottomMargin=36)

            styles = getSampleStyleSheet()
            story = []

            # Custom styles for better text wrapping
            cell_style = ParagraphStyle(
                'CellStyle',
                parent=styles['Normal'],
                fontSize=8,
                leading=10,
                wordWrap='LTR'
            )

            header_cell_style = ParagraphStyle(
                'HeaderCellStyle',
                parent=styles['Normal'],
                fontSize=9,
                fontName='Helvetica-Bold',
                leading=11,
                wordWrap='LTR'
            )

            # AIT Logo style
            logo_style = ParagraphStyle(
                'LogoStyle',
                parent=styles['Heading1'],
                fontSize=18,
                fontName='Helvetica-Bold',
                alignment=1,
                textColor=colors.red
            )

            company_style = ParagraphStyle(
                'CompanyStyle',
                parent=styles['Heading1'],
                fontSize=14,
                fontName='Helvetica-Bold',
                alignment=1,
                textColor=colors.darkblue
            )

            print(f"DEBUG: Creating PDF for {technician}")
            print(f"DEBUG: Total assignments: {len(assignments)}")

            for i, assignment in enumerate(assignments):
                print(f"DEBUG: Processing assignment {i}: {assignment}")
            
                # Safety check for assignment data
                if not assignment or len(assignment) < 8:
                    print(f"DEBUG: Skipping invalid assignment {i}")
                    continue

                # ASSIGN VARIABLES FIRST - CRITICAL TO DO THIS EARLY
                bfm_no, sap_no, description, tool_id, location, master_lin, pm_type, scheduled_date = assignment
            
                # Add None checks for all variables
                bfm_no = bfm_no or ''
                sap_no = sap_no or ''
                description = description or ''
                tool_id = tool_id or ''
                location = location or ''
                master_lin = master_lin or ''
                pm_type = pm_type or 'Monthly'
                scheduled_date = scheduled_date or ''
            
                print(f"DEBUG: Processing {bfm_no} - {pm_type}")

                # =================== CUSTOM TEMPLATE INTEGRATION ===================
                print(f"DEBUG: About to call get_pm_template_for_equipment for {bfm_no}, {pm_type}")

                try:
                    custom_template = self.get_pm_template_for_equipment(bfm_no, pm_type)
                    print(f"DEBUG: get_pm_template_for_equipment returned: {custom_template}")
                
                    if custom_template:
                        print("DEBUG: Custom template found, extracting data...")
                        checklist_items = custom_template.get('checklist_items', [])
                        estimated_hours = custom_template.get('estimated_hours', 1.0)
                        special_instructions = custom_template.get('special_instructions', '')
                        safety_notes = custom_template.get('safety_notes', '')
                    
                        # Safety check for checklist items
                        if not checklist_items or not isinstance(checklist_items, list):
                            print("DEBUG: Invalid custom checklist_items, using default")
                            checklist_items = [
                                'Special Equipment Used (List):',
                                'Validate your maintenance with Date / Stamp / Hours',
                                'Refer to drawing when performing maintenance',
                                'Make sure all instruments are properly calibrated',
                                'Make sure tool is properly identified',
                                'Make sure all mobile mechanisms move fluidly',
                                'Visually inspect the welds',
                                'Take note of any anomaly or defect (create a CM if needed)',
                                'Check all screws. Tighten if needed.',
                                'Check the pins for wear',
                                'Make sure all tooling is secured to the equipment with cable',
                                'Ensure all tags (BFM and SAP) are applied and securely fastened',
                                'All documentation are picked up from work area',
                                'All parts and tools have been picked up',
                                'Workspace has been cleaned up',
                                'Dry runs have been performed (tests, restarts, etc.)',
                                "Ensure that AIT Sticker is applied"
                            ]
                    
                        document_name = f"Custom_PM_Template_{pm_type}"
                        document_revision = "C1"
                        print(f"DEBUG: Using custom template for {bfm_no} - {pm_type}")
                    else:
                        print("DEBUG: No custom template, using defaults...")
                        # Use default checklist items
                        checklist_items = [
                            'Special Equipment Used (List):',
                            'Validate your maintenance with Date / Stamp / Hours',
                            'Refer to drawing when performing maintenance',
                            'Make sure all instruments are properly calibrated',
                            'Make sure tool is properly identified',
                            'Make sure all mobile mechanisms move fluidly',
                            'Visually inspect the welds',
                            'Take note of any anomaly or defect (create a CM if needed)',
                            'Check all screws. Tighten if needed.',
                            'Check the pins for wear',
                            'Make sure all tooling is secured to the equipment with cable',
                            'Ensure all tags (BFM and SAP) are applied and securely fastened',
                            'All documentation are picked up from work area',
                            'All parts and tools have been picked up',
                            'Workspace has been cleaned up',
                            'Dry runs have been performed (tests, restarts, etc.)',
                            "Ensure that AIT Sticker is applied"
                        ]
                        estimated_hours = 1.0
                        special_instructions = ''
                        safety_notes = "Always be aware of both Airbus and AIT safety policies and ensure safety policies are followed."
                        document_name = 'Preventive_Maintenance_Form'
                        document_revision = 'A2'
                        print(f"DEBUG: Using default template for {bfm_no} - {pm_type}")
                    
                except Exception as e:
                    print(f"DEBUG: Exception in template section: {e}")
                    # Fallback to basic defaults
                    checklist_items = [
                        'Special Equipment Used (List):',
                        'Validate your maintenance with Date / Stamp / Hours',
                        'Refer to drawing when performing maintenance',
                        'Make sure all instruments are properly calibrated',
                        'Make sure tool is properly identified',
                        'Make sure all mobile mechanisms move fluidly',
                        'Visually inspect the welds',
                        'Take note of any anomaly or defect (create a CM if needed)',
                        'Check all screws. Tighten if needed.',
                        'Check the pins for wear',
                        'Make sure all tooling is secured to the equipment with cable',
                        'Ensure all tags (BFM and SAP) are applied and securely fastened',
                        'All documentation are picked up from work area',
                        'All parts and tools have been picked up',
                        'Workspace has been cleaned up',
                        'Dry runs have been performed (tests, restarts, etc.)',
                        "Ensure that AIT Sticker is applied"
                    ]
                    estimated_hours = 1.0
                    special_instructions = ''
                    safety_notes = "Always be aware of both Airbus and AIT safety policies and ensure safety policies are followed."
                    document_name = 'Preventive_Maintenance_Form'
                    document_revision = 'A2'

                print(f"DEBUG: Final checklist_items: {len(checklist_items)} items")
                # ====================================================================

                # Get the last PM date for this equipment and PM type from database
                cursor = self.conn.cursor()
                last_pm_date = ""

                try:
                    if pm_type == 'Monthly':
                        cursor.execute('SELECT last_monthly_pm FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
                    elif pm_type == 'Six Month':
                        cursor.execute('SELECT last_six_month_pm FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
                    elif pm_type == 'Annual':
                        cursor.execute('SELECT last_annual_pm FROM equipment WHERE bfm_equipment_no = ?', (bfm_no,))
                    else:
                        cursor.execute('''
                            SELECT completion_date FROM pm_completions 
                            WHERE bfm_equipment_no = ? AND pm_type = ? 
                            ORDER BY completion_date DESC LIMIT 1
                        ''', (bfm_no, pm_type))

                    result = cursor.fetchone()
                    if result and result[0]:
                        raw_date = str(result[0]).strip()
                        last_pm_date = ""

                        if raw_date:
                            # Try multiple date formats
                            date_formats = [
                                '%m/%d/%y',      # 8/14/25
                                '%m/%d/%Y',      # 8/14/2025  
                                '%Y-%m-%d',      # 2025-08-14
                                '%m-%d-%y',      # 8-14-25
                                '%m-%d-%Y'       # 8-14-2025
                            ]

                            for date_format in date_formats:
                                try:
                                    date_obj = datetime.strptime(raw_date, date_format)
            
                                    # Handle 2-digit years (assume 20xx if < 50, 19xx if >= 50)
                                    if date_obj.year < 1950:
                                        date_obj = date_obj.replace(year=date_obj.year + 2000)
            
                                    last_pm_date = date_obj.strftime('%m/%d/%Y')  # Always output as MM/DD/YYYY
                                    break  # Successfully parsed, exit the loop
                                except ValueError:
                                    continue  # Try next format

                            # If no format worked, use the raw date as-is
                            if not last_pm_date:
                                last_pm_date = raw_date
        
                except Exception as e:
                    print(f"Error getting last PM date for {bfm_no}: {e}")
                    last_pm_date = ""

                # Add page break between forms (except for first)
                if i > 0:
                    story.append(PageBreak())

                # Header with company logo
                try:
                    from reportlab.platypus import Image
                    # Use the correct path to your logo file
                    logo_path = r"C:\Users\stu15olen\Desktop\AIT_CMMS\img\ait_logo.png"  # Update with your actual logo filename
    
                    if os.path.exists(logo_path):
                        # Create centered logo
                        logo_image = Image(logo_path, width=4*inch, height=1.2*inch)
        
                        # Center the logo in a table
                        logo_data = [[logo_image]]
                        logo_table = Table(logo_data, colWidths=[7*inch])
                        logo_table.setStyle(TableStyle([
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('TOPPADDING', (0, 0), (-1, -1), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
                        ]))
        
                        story.append(logo_table)
                    else:
                        # Fallback to text if logo file not found
                        print(f"Logo file not found at: {logo_path}")
                        story.append(Paragraph("AIT - BUILDING THE FUTURE OF AEROSPACE", company_style))
                        story.append(Spacer(1, 15))

                except Exception as e:
                    print(f"Could not load logo: {e}")
                    # Fallback to text header
                    story.append(Paragraph("AIT - BUILDING THE FUTURE OF AEROSPACE", company_style))
                    story.append(Spacer(1, 15))
    
                # Equipment information table
                equipment_data = [
                    [
                        Paragraph('(SAP) Material Number:', header_cell_style), 
                        Paragraph(str(sap_no), cell_style), 
                        Paragraph('Tool ID / Drawing Number:', header_cell_style), 
                        Paragraph(str(tool_id), cell_style)
                    ],
                    [
                        Paragraph('(BFM) Equipment Number:', header_cell_style), 
                        Paragraph(str(bfm_no), cell_style), 
                        Paragraph('Description of Equipment:', header_cell_style), 
                        Paragraph(str(description), cell_style)
                    ],
                    [
                        Paragraph('Date of Last PM:', header_cell_style), 
                        Paragraph(str(last_pm_date), cell_style), 
                        Paragraph('Location of Equipment:', header_cell_style), 
                        Paragraph(str(location), cell_style)
                    ],
                    [
                        Paragraph('Maintenance Technician:', header_cell_style), 
                        Paragraph(str(technician), cell_style), 
                        Paragraph('PM Cycle:', header_cell_style), 
                        Paragraph(str(pm_type), cell_style)
                    ],
                    [
                        Paragraph('Estimated Hours:', header_cell_style), 
                        Paragraph(f'{estimated_hours:.1f}h', cell_style), 
                        Paragraph('Date of PM Completion:', header_cell_style), 
                        Paragraph('', cell_style)
                    ],
                    [
                        Paragraph('Signature of Technician:', header_cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('Template Type:', header_cell_style), 
                        Paragraph('Custom' if custom_template else 'Standard', cell_style)
                    ]
                    ]
        
                # Add safety notes if they exist
                if safety_notes and safety_notes.strip():
                        equipment_data.append([
                        Paragraph(f'Safety: {safety_notes}', cell_style), 
                        '', '', ''
                ])
        
                # Add special instructions if they exist
                if special_instructions and special_instructions.strip():
                        equipment_data.append([
                        Paragraph(f'Special Instructions: {special_instructions}', cell_style), 
                        '', '', ''
                ])
        
                # Add print date
                equipment_data.append([
                    Paragraph(f'Printed: {datetime.now().strftime("%m/%d/%Y")}', cell_style), 
                    '', '', ''
                ])

                # Create equipment table
                # Create equipment table
                equipment_table = Table(equipment_data, colWidths=[1.8*inch, 1.7*inch, 1.8*inch, 1.7*inch])

                # Build table style commands without None values
                table_style_commands = [
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('SPAN', (0, -1), (-1, -1)),  # Always span printed date
                ]

                # Add conditional spans only if they exist
                if safety_notes and safety_notes.strip():
                    table_style_commands.append(('SPAN', (0, -3), (-1, -3)))

                if special_instructions and special_instructions.strip():
                    table_style_commands.append(('SPAN', (0, -2), (-1, -2)))

                equipment_table.setStyle(TableStyle(table_style_commands))

                story.append(equipment_table)
                story.append(Spacer(1, 15))

                # PM checklist table
                checklist_header = "CUSTOM PM CHECKLIST:" if custom_template else "PM CHECKLIST:"
                checklist_data = [
                    [
                        Paragraph('', header_cell_style), 
                        Paragraph(checklist_header, header_cell_style), 
                        Paragraph('', header_cell_style), 
                        Paragraph('Complete', header_cell_style), 
                        Paragraph('Labor Time', header_cell_style)
                    ]
                ]

                # Add checklist items
                for idx, item in enumerate(checklist_items, 1):
                    checklist_data.append([
                        Paragraph(str(idx), cell_style), 
                        Paragraph(str(item), cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('', cell_style)
                    ])

                # Create checklist table
                checklist_table = Table(checklist_data, colWidths=[0.3*inch, 4.2*inch, 0.4*inch, 0.7*inch, 1.4*inch])
                checklist_table.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('LEFTPADDING', (0, 0), (-1, -1), 2),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                    ('TOPPADDING', (0, 0), (-1, -1), 2),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ]))

                story.append(checklist_table)
                story.append(Spacer(1, 15))

                # Notes and completion section
                completion_data = [
                    [
                        Paragraph('Notes from Technician:', header_cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('Next Annual PM Date:', header_cell_style)
                    ],
                    [
                        Paragraph('', cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('', cell_style)
                    ],
                    [
                        Paragraph('All Data Entered Into System:', header_cell_style), 
                        Paragraph('', cell_style), 
                        Paragraph('Total Time', header_cell_style)
                    ],
                    [
                        Paragraph('Document Name', header_cell_style), 
                        Paragraph('Revision', header_cell_style), 
                        Paragraph('', cell_style)
                    ],
                    [
                        Paragraph(document_name, cell_style), 
                        Paragraph(document_revision, cell_style), 
                        Paragraph('', cell_style)
                    ]
                ]

                completion_table = Table(completion_data, colWidths=[2.8*inch, 2.2*inch, 2*inch])
                completion_table.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ]))

                story.append(completion_table)

            # Build PDF
            print(f"DEBUG: Building PDF with {len(story)} elements")
            doc.build(story)
            print(f"DEBUG: PDF created successfully: {filename}")

        except Exception as e:
            print(f"Error creating PM forms PDF: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def export_weekly_schedule(self):
        """Export weekly schedule to Excel"""
        try:
            week_start = self.week_start_var.get()
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Weekly_PM_Schedule_{week_start}_{timestamp}.xlsx"
            
            cursor = self.conn.cursor()
            cursor.execute('''
                SELECT ws.assigned_technician, ws.bfm_equipment_no, e.description, 
                       ws.pm_type, ws.scheduled_date, ws.status
                FROM weekly_pm_schedules ws
                JOIN equipment e ON ws.bfm_equipment_no = e.bfm_equipment_no
                WHERE ws.week_start_date = ?
                ORDER BY ws.assigned_technician, ws.scheduled_date
            ''', (week_start,))
            
            schedule_data = cursor.fetchall()
            
            # Create DataFrame
            df = pd.DataFrame(schedule_data, columns=[
                'Technician', 'BFM Equipment No', 'Description', 'PM Type', 'Scheduled Date', 'Status'
            ])
            
            # Export to Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Weekly Schedule', index=False)
                
                # Create summary sheet
                summary_data = []
                for tech in self.technicians:
                    tech_count = len(df[df['Technician'] == tech])
                    summary_data.append([tech, tech_count])
                
                summary_df = pd.DataFrame(summary_data, columns=['Technician', 'Assigned PMs'])
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            messagebox.showinfo("Success", f"Weekly schedule exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export weekly schedule: {str(e)}")
    
    def create_pm_history_search_tab(self):
        """PM History Search tab for comprehensive equipment completion information"""
        self.pm_history_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pm_history_frame, text="PM History Search")
    
        # Search controls
        search_controls_frame = ttk.LabelFrame(self.pm_history_frame, text="Search Equipment PM History", padding=15)
        search_controls_frame.pack(fill='x', padx=10, pady=5)
    
        # Search input
        search_input_frame = ttk.Frame(search_controls_frame)
        search_input_frame.pack(fill='x', pady=5)
    
        ttk.Label(search_input_frame, text="Search:").pack(side='left', padx=5)
        self.history_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_input_frame, textvariable=self.history_search_var, width=30)
        search_entry.pack(side='left', padx=5)
    
        ttk.Button(search_input_frame, text="Search", command=self.search_pm_history_simple).pack(side='left', padx=5)
        ttk.Button(search_input_frame, text="Clear", command=self.clear_search_simple).pack(side='left', padx=5)
    
        # Results display
        results_frame = ttk.LabelFrame(self.pm_history_frame, text="Search Results", padding=10)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
        # Results tree
        self.history_search_tree = ttk.Treeview(results_frame,
                                            columns=('BFM No', 'SAP No', 'Description', 'PM Type', 'Technician', 'Date', 'Hours'),
                                            show='headings')
    
        for col in ('BFM No', 'SAP No', 'Description', 'PM Type', 'Technician', 'Date', 'Hours'):
            self.history_search_tree.heading(col, text=col)
            self.history_search_tree.column(col, width=120)
    
        self.history_search_tree.pack(fill='both', expand=True)

    def search_pm_history_simple(self):
        """Simple PM history search"""
        try:
            search_term = self.history_search_var.get().lower()
            cursor = self.conn.cursor()
        
            if search_term:
                cursor.execute('''
                    SELECT pc.bfm_equipment_no, e.sap_material_no, e.description, 
                        pc.pm_type, pc.technician_name, pc.completion_date,
                        (pc.labor_hours + pc.labor_minutes/60.0) as total_hours
                    FROM pm_completions pc
                    LEFT JOIN equipment e ON pc.bfm_equipment_no = e.bfm_equipment_no
                    WHERE LOWER(pc.bfm_equipment_no) LIKE ? 
                    OR LOWER(e.description) LIKE ?
                    OR LOWER(pc.technician_name) LIKE ?
                    ORDER BY pc.completion_date DESC LIMIT 50
                ''', (f'%{search_term}%', f'%{search_term}%', f'%{search_term}%'))
            else:
                cursor.execute('''
                    SELECT pc.bfm_equipment_no, e.sap_material_no, e.description, 
                        pc.pm_type, pc.technician_name, pc.completion_date,
                        (pc.labor_hours + pc.labor_minutes/60.0) as total_hours
                    FROM pm_completions pc
                    LEFT JOIN equipment e ON pc.bfm_equipment_no = e.bfm_equipment_no
                    ORDER BY pc.completion_date DESC LIMIT 20
                ''')
        
            results = cursor.fetchall()
        
            # Clear existing
            for item in self.history_search_tree.get_children():
                self.history_search_tree.delete(item)
        
            # Add results
            for result in results:
                bfm_no, sap_no, description, pm_type, technician, date, hours = result
                hours_display = f"{hours:.1f}h" if hours else "0.0h"
            
                self.history_search_tree.insert('', 'end', values=(
                    bfm_no or '', sap_no or '', description or '', 
                    pm_type or '', technician or '', date or '', hours_display
                ))
        except Exception as e:
            print(f"Search error: {e}")

    def clear_search_simple(self):
        """Clear search"""
        self.history_search_var.set('')
        self.search_pm_history_simple()
    
    
    
    

# Main application startup
if __name__ == "__main__":
    root = tk.Tk()
    app = AITCMMSSystem(root)
    root.mainloop()
