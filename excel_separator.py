import pandas as pd
import os
from tkinter import Tk, filedialog, Label, ttk, scrolledtext, PhotoImage
from tkinter.messagebox import showinfo, showerror
import time
from datetime import datetime
import sys

class ExcelSeparator:
    def __init__(self):
        self.root = Tk()
        self.root.title("Excel Sheet Separator")
        self.root.geometry("900x700")
        
        # Configure the style
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use 'clam' theme as base
        
        # Configure custom styles
        self.style.configure('Header.TLabel', 
                           font=('Segoe UI', 24, 'bold'),
                           padding=10)
        
        self.style.configure('SubHeader.TLabel',
                           font=('Segoe UI', 10),
                           foreground='#666666',
                           padding=5)
        
        self.style.configure('Modern.TButton',
                           font=('Segoe UI', 10),
                           padding=10)
        
        self.style.configure('Progress.Horizontal.TProgressbar',
                           background='#2ecc71',
                           troughcolor='#f0f0f0',
                           borderwidth=0,
                           thickness=15)

        # Set window background
        self.root.configure(bg='#f5f6fa')
        
        self.setup_ui()

    def setup_ui(self):
        # Create main container with padding
        container = ttk.Frame(self.root, padding="20 20 20 20")
        container.pack(fill='both', expand=True)
        
        # Header section
        header_frame = ttk.Frame(container)
        header_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(header_frame, 
                 text="Excel Sheet Separator",
                 style='Header.TLabel').pack(anchor='w')
        
        ttk.Label(header_frame,
                 text="Split your complex Excel sheets with ease",
                 style='SubHeader.TLabel').pack(anchor='w')

        # Buttons section with modern styling
        button_frame = ttk.Frame(container)
        button_frame.pack(fill='x', pady=(0, 20))
        
        select_btn = ttk.Button(button_frame,
                              text="Select Excel File",
                              style='Modern.TButton',
                              command=self.process_excel)
        select_btn.pack(side='left', padx=(0, 10))
        
        clear_btn = ttk.Button(button_frame,
                             text="Clear Log",
                             style='Modern.TButton',
                             command=self.clear_log)
        clear_btn.pack(side='left', padx=(0, 10))
        
        exit_btn = ttk.Button(button_frame,
                            text="Exit",
                            style='Modern.TButton',
                            command=self.root.quit)
        exit_btn.pack(side='left')

        # Progress section
        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill='x', pady=(0, 20))
        
        self.progress = ttk.Progressbar(progress_frame,
                                      style='Progress.Horizontal.TProgressbar',
                                      length=300,
                                      mode='determinate')
        self.progress.pack(side='left', padx=(0, 10), fill='x', expand=True)
        
        self.status_label = ttk.Label(progress_frame,
                                    text="Ready",
                                    font=('Segoe UI', 10))
        self.status_label.pack(side='left', padx=(0, 10))

        # Log section with custom styling
        log_frame = ttk.LabelFrame(container,
                                 text="Processing Log",
                                 padding="10 10 10 10")
        log_frame.pack(fill='both', expand=True)
        
        self.log_window = scrolledtext.ScrolledText(
            log_frame,
            height=20,
            font=('Consolas', 10),
            background='#ffffff',
            foreground='#2c3e50'
        )
        self.log_window.pack(fill='both', expand=True)
        
        # Status bar
        status_frame = ttk.Frame(container)
        status_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Label(status_frame,
                 text=" 2024 Excel Sheet Separator",
                 font=('Segoe UI', 8),
                 foreground='#666666').pack(side='left')

    def clear_log(self):
        self.log_window.delete(1.0, 'end')
        self.log("Log cleared")

    def log(self, message, level='INFO'):
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Color coding for different message levels
        colors = {
            'INFO': '#2c3e50',    # Dark blue
            'SUCCESS': '#27ae60',  # Green
            'WARNING': '#f39c12',  # Orange
            'ERROR': '#c0392b'     # Red
        }
        
        self.log_window.tag_config(level, foreground=colors.get(level, '#2c3e50'))
        
        log_entry = f"[{timestamp}] {level}: {message}\n"
        self.log_window.insert('end', log_entry, level)
        self.log_window.see('end')
        self.root.update_idletasks()

    def update_progress(self, value, message):
        self.progress['value'] = value
        self.status_label.config(text=message)
        
        # Log with appropriate level based on progress
        if value == 100:
            self.log(message, 'SUCCESS')
        elif value == 0:
            self.log(message, 'INFO')
        else:
            self.log(message, 'INFO')
        
        self.root.update_idletasks()

    def process_excel(self):
        # Open file dialog
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return

        try:
            # Reset progress bar and log
            self.progress['value'] = 0
            self.log("\n" + "="*50)
            self.log(f"Starting new processing job")
            self.update_progress(0, "Reading file...")
            
            # Read the file based on extension
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext == '.csv':
                self.log(f"Reading CSV file: {file_path}")
                df = pd.read_csv(file_path)
            else:
                self.log(f"Reading Excel file: {file_path}")
                df = pd.read_excel(file_path)

            total_rows = len(df)
            self.log(f"Total rows found: {total_rows}")
            self.log(f"Columns found: {df.columns.tolist()}")
            
            if df.empty:
                self.log("Error: Empty file detected")
                showerror("Error", "The selected file is empty!")
                return

            # Get the first column name
            first_col = df.columns[0]
            
            # Initialize variables
            tables = {}
            current_title = None
            processed_rows = 0

            # Process each row
            start_time = time.time()
            
            # First pass to identify all titles
            self.update_progress(10, "Identifying tables...")
            titles = []
            for index, row in df.iterrows():
                first_cell = str(row[first_col]).strip()
                if pd.notna(first_cell) and first_cell != '' and all(pd.isna(row[col]) or str(row[col]).strip() == '' 
                                       for col in df.columns if col != first_col):
                    titles.append(first_cell)
                    self.log(f"Found table title: {first_cell}")
            
            self.log(f"Total tables found: {len(titles)}")
            
            # Initialize all tables
            for title in titles:
                tables[title] = []

            # Process rows
            self.update_progress(20, "Processing rows...")
            current_title = None
            for index, row in df.iterrows():
                first_cell = str(row[first_col]).strip()
                
                # Check if this is a title row
                if first_cell in titles:
                    current_title = first_cell
                elif current_title:
                    # This is a data row for the current title
                    if any(pd.notna(row[col]) and str(row[col]).strip() != '' 
                          for col in df.columns if col != first_col):
                        row_dict = {}
                        for col in df.columns:
                            value = row[col]
                            if pd.isna(value) or str(value).strip() == '':
                                row_dict[col] = ''
                            else:
                                row_dict[col] = value
                        tables[current_title].append(row_dict)
                
                processed_rows += 1
                if processed_rows % 100 == 0:  # Update progress every 100 rows
                    progress = 20 + (60 * processed_rows / total_rows)
                    elapsed_time = time.time() - start_time
                    rows_per_second = processed_rows / elapsed_time
                    remaining_rows = total_rows - processed_rows
                    estimated_remaining_time = remaining_rows / rows_per_second if rows_per_second > 0 else 0
                    status_msg = f"Processing rows... {processed_rows}/{total_rows} (~{estimated_remaining_time:.1f}s remaining)"
                    self.update_progress(progress, status_msg)
                    self.log(f"Processing speed: {rows_per_second:.1f} rows/second")

            # Create output filename
            output_dir = os.path.dirname(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}_separated.xlsx")

            # Create a new Excel writer object
            self.update_progress(80, "Writing output file...")
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write each table to a separate sheet
                tables_processed = 0
                total_tables = len(tables)
                
                for title, rows in tables.items():
                    if rows:  # Only process tables that have rows
                        # Create DataFrame from rows
                        table_df = pd.DataFrame(rows)
                        self.log(f"Creating sheet '{title}' with {len(table_df)} rows and {len(table_df.columns)} columns")
                        
                        # Get the last three columns
                        all_columns = list(table_df.columns)
                        last_three_cols = all_columns[-3:]
                        
                        # Convert last three columns to numeric, removing any currency symbols and commas
                        for col in last_three_cols:
                            self.log(f"Converting column '{col}' to currency format")
                            table_df[col] = table_df[col].apply(lambda x: str(x).replace('$', '').replace(',', '') if pd.notna(x) else '')
                            table_df[col] = pd.to_numeric(table_df[col], errors='coerce')
                        
                        # More thorough sheet name sanitization
                        sheet_name = str(title)
                        # Replace invalid characters with underscore
                        invalid_chars = [':', '\\', '/', '?', '*', '[', ']', "'"]
                        for char in invalid_chars:
                            sheet_name = sheet_name.replace(char, '_')
                        # Ensure sheet name doesn't exceed Excel's 31 character limit
                        sheet_name = sheet_name[:31]
                        # Ensure sheet name isn't empty
                        if not sheet_name or sheet_name.isspace():
                            sheet_name = f"Table_{len(tables)}"
                        
                        # Write the table to the sheet
                        table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Get the worksheet to apply formatting
                        worksheet = writer.sheets[sheet_name]
                        
                        # Convert column index to Excel column letters
                        def get_column_letter(col_idx):
                            result = ''
                            while col_idx >= 0:
                                result = chr(65 + (col_idx % 26)) + result
                                col_idx = (col_idx // 26) - 1
                            return result
                        
                        # Apply currency formatting to last three columns
                        for col_idx, col in enumerate(all_columns):
                            if col in last_three_cols:
                                col_letter = get_column_letter(col_idx)
                                # Format all cells in this column (excluding header)
                                for row in range(2, len(table_df) + 2):  # +2 because Excel is 1-based and we have a header
                                    cell = f"{col_letter}{row}"
                                    worksheet[cell].number_format = '"$"#,##0.00'
                        
                        tables_processed += 1
                        progress = 80 + (19 * tables_processed / total_tables)
                        self.update_progress(progress, f"Writing table {tables_processed}/{total_tables}")

            self.update_progress(100, "Complete!")
            self.log(f"Processing complete! Output saved to: {output_path}")
            self.log("="*50 + "\n")
            showinfo("Success", f"File has been processed!\nOutput saved as: {output_path}")
            self.status_label.config(text="Ready")

        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            showerror("Error", error_msg)
            print(f"Error details: {str(e)}")
            self.status_label.config(text="Error occurred")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelSeparator()
    app.run()
