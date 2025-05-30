import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import threading
import queue
from macros.RefFinder import UNPAYWALL_EMAIL, process_references_backend

# --- GUI Application ---
class App:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Open Access Reference Checker")

        self.input_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.processing_thread = None
        self.gui_update_queue = queue.Queue()

        # --- Top frame for controls ---
        controls_frame = tk.Frame(self.root)
        controls_frame.pack(fill=tk.X, padx=10, pady=(10,0))

        tk.Button(controls_frame, text="Select Input File (.txt)", command=self.select_input_file).pack(side=tk.LEFT, padx=5, pady=5)
        self.input_file_label = tk.Label(controls_frame, textvariable=self.input_file_path, relief=tk.GROOVE, width=30, anchor='w', justify=tk.LEFT) # Adjusted width
        self.input_file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        tk.Button(controls_frame, text="Select Output Folder", command=self.select_output_folder).pack(side=tk.LEFT, padx=5, pady=5)
        self.output_folder_label = tk.Label(controls_frame, textvariable=self.output_folder_path, relief=tk.GROOVE, width=30, anchor='w', justify=tk.LEFT) # Adjusted width
        self.output_folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        self.start_button = tk.Button(controls_frame, text="Start Processing", command=self.start_processing_thread, font=("Arial", 12, "bold"), bg="lightblue")
        self.start_button.pack(side=tk.LEFT, pady=5, padx=(10,5), ipady=5) # Adjusted padding

        # --- NEW: Button to copy summary ---
        self.copy_summary_button = tk.Button(controls_frame, text="Copy Summary", command=self.copy_summary_to_clipboard)
        self.copy_summary_button.pack(side=tk.LEFT, padx=5, pady=5, ipady=2)


        # --- Main PanedWindow to divide summary and log ---
        self.paned_window = tk.PanedWindow(self.root, orient=tk.VERTICAL, sashrelief=tk.RAISED)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Summary Treeview Frame (Top pane of PanedWindow) ---
        summary_frame = tk.Frame(self.paned_window, relief=tk.GROOVE, borderwidth=2)
        self.paned_window.add(summary_frame, height=300) # Adjusted initial height

        summary_label = tk.Label(summary_frame, text="Live Summary:", font=("Arial", 10, "bold"))
        summary_label.pack(pady=(5,0), anchor='w', padx=5)
        
        self.summary_tree = ttk.Treeview(summary_frame, columns=("DOI", "TitleInfo", "URL"), show="headings") # Renamed StatusInfo to TitleInfo
        self.summary_tree.heading("DOI", text="DOI")
        self.summary_tree.heading("TitleInfo", text="Title -- Status / File / Note") # Combined for better display
        self.summary_tree.heading("URL", text="Relevant URL")
        
        self.summary_tree.column("DOI", width=150, anchor='w', stretch=tk.NO)
        self.summary_tree.column("TitleInfo", width=400, anchor='w') # Main info column
        self.summary_tree.column("URL", width=250, anchor='w', stretch=tk.NO)


        summary_vsb = ttk.Scrollbar(summary_frame, orient="vertical", command=self.summary_tree.yview)
        summary_hsb = ttk.Scrollbar(summary_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=summary_vsb.set, xscrollcommand=summary_hsb.set)
        
        summary_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        summary_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.summary_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tree_categories = {
            'downloaded_pdf': self.summary_tree.insert("", tk.END, iid="cat_downloaded", text="✅ Downloaded PDFs", open=True),
            'open_access_no_pdf_link': self.summary_tree.insert("", tk.END, iid="cat_oa_no_link", text="🔑 OA (No PDF/Download Failed)", open=True),
            'paywalled': self.summary_tree.insert("", tk.END, iid="cat_paywalled", text="💰 Paywalled", open=True),
            'unpaywall_data_unavailable': self.summary_tree.insert("", tk.END, iid="cat_unavailable", text="⚠️ Unpaywall/SS Data Failed", open=True),
            'doi_not_found': self.summary_tree.insert("", tk.END, iid="cat_doi_not_found", text="❌ DOI Not Found", open=True)
        }
        self.summary_tree.config(show="tree headings")


        # --- Log Area Frame (Bottom pane of PanedWindow) ---
        log_frame = tk.Frame(self.paned_window, relief=tk.GROOVE, borderwidth=2)
        self.paned_window.add(log_frame, height=200) 

        log_label = tk.Label(log_frame, text="Detailed Log:", font=("Arial", 10, "bold"))
        log_label.pack(pady=(5,0), anchor='w', padx=5)
        self.log_text_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text_area.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)
        self.log_text_area.configure(state='disabled')

        self.check_gui_queue()

    def select_input_file(self):
        path = filedialog.askopenfilename(
            title="Select the .txt file with references",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*")),
            initialdir=os.getcwd()
        )
        if path:
            self.input_file_path.set(os.path.normpath(path))

    def select_output_folder(self):
        path = filedialog.askdirectory(
            title="Select Folder to Save Downloaded PDFs",
            initialdir=os.getcwd()
        )
        if path:
            self.output_folder_path.set(os.path.normpath(path))

    def update_log_area(self, message):
        self.log_text_area.configure(state='normal')
        self.log_text_area.insert(tk.END, message + "\n")
        self.log_text_area.see(tk.END) 
        self.log_text_area.configure(state='disabled')

    def update_summary_treeview(self, category_key, item_data):
        parent_iid = self.tree_categories.get(category_key)
        if parent_iid:
            doi_val = item_data.get('doi', 'N/A')
            title_val = item_data.get('title', 'No Title')
            status_info_val = item_data.get('status_info', 'N/A')
            url_val = item_data.get('url_info', 'N/A')
            
            # Combine title and status for the second column for better readability
            title_status_combined = f"{title_val[:60]}{'...' if len(title_val) > 60 else ''} -- {status_info_val}"

            # Insert item: text is usually for the first "tree" column if not using 'headings' show option only.
            # Since we use 'headings', the first defined column by 'columns=' is #0.
            # However, Treeview with 'show="tree headings"' expects 'text' for the first column of an item.
            # Let's make the item's 'text' the DOI, and 'values' for the defined columns.
            # Our columns are "DOI", "TitleInfo", "URL".
            # So 'text' will be the primary identifier (like a short title or DOI for the item itself in tree column)
            # and 'values' will be for the columns DOI, TitleInfo, URL.
            # Let's adjust Treeview insertion: use item's title as 'text' in the tree.
            
            # Current setup: `columns=("DOI", "TitleInfo", "URL")`, `show="headings"` for tree.
            # And categories are inserted with `text="Category Name"`.
            # Child items are inserted with `text=doi_val, values=(doi_val, title_status_combined, url_val)`.
            # This means the first column (#0 or 'tree' column) for the child will show `doi_val`.
            # And then the `values` tuple maps to the columns "DOI", "TitleInfo", "URL".
            # This is slightly redundant having DOI twice. Let's fix for clarity.

            # New approach:
            # 'text' for item can be a short version of title.
            # 'values' for columns: DOI, Status, URL
            item_text_display = (title_val[:40] + '...') if len(title_val) > 40 else title_val

            # Values for the columns as defined: "DOI", "TitleInfo", "URL"
            # The 'TitleInfo' column will now hold the status. The title is in the tree's item text.
            tree_values = (
                doi_val,
                status_info_val, # This is now the 'Status / File / Note'
                url_val
            )

            new_item_id = self.summary_tree.insert(parent_iid, tk.END, text=item_text_display, values=tree_values)
            self.summary_tree.see(new_item_id)

    def check_gui_queue(self):
        try:
            while True:
                queued_item = self.gui_update_queue.get_nowait()
                if isinstance(queued_item, dict) and 'type' in queued_item:
                    if queued_item['type'] == 'log':
                        self.update_log_area(queued_item['message'])
                    elif queued_item['type'] == 'summary_add':
                        self.update_summary_treeview(queued_item['category_key'], queued_item['item_data'])
                else:
                    self.update_log_area(str(queued_item))
        except queue.Empty:
            pass
        self.root.after(100, self.check_gui_queue)

    def gui_update_callback(self, item_to_log):
        self.gui_update_queue.put(item_to_log)

    def start_processing_thread(self):
        input_file = self.input_file_path.get()
        output_folder = self.output_folder_path.get()

        if not input_file:
            messagebox.showerror("Error", "Please select an input file.")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return
        
        if UNPAYWALL_EMAIL == 'YOUR_EMAIL@example.com' or not UNPAYWALL_EMAIL:
             messagebox.showerror("Error", "Please set your Unpaywall email at the top of the script (UNPAYWALL_EMAIL variable).")
             return

        for category_iid_key in self.tree_categories: # Use the key to get iid
            parent_iid = self.tree_categories[category_iid_key]
            for item_iid in self.summary_tree.get_children(parent_iid):
                self.summary_tree.delete(item_iid)
        
        self.log_text_area.configure(state='normal')
        self.log_text_area.delete(1.0, tk.END)
        self.log_text_area.configure(state='disabled')
        
        self.gui_update_callback({'type': 'log', 'message': "Starting processing..."})
        
        self.start_button.config(text="Processing...", state=tk.DISABLED)
        self.copy_summary_button.config(state=tk.DISABLED) # Disable copy button during processing

        self.processing_thread = threading.Thread(
            target=self.run_backend_task,
            args=(input_file, output_folder),
            daemon=True
        )
        self.processing_thread.start()

    def run_backend_task(self, input_file, output_folder):
        try:
            references_from_file = []
            with open(input_file, 'r', encoding='utf-8') as f:
                references_from_file = [line.strip() for line in f if line.strip()]
            
            if not references_from_file:
                self.gui_update_callback({'type': 'log', 'message': f"No references found in '{input_file}' or the file is empty."})
                # self.processing_finished() # This will be called in finally
                return

            process_references_backend(references_from_file, output_folder, self.gui_update_callback)

        except FileNotFoundError:
            self.gui_update_callback({'type': 'log', 'message': f"🚨 Error: Input file '{input_file}' not found."})
        except Exception as e:
            self.gui_update_callback({'type': 'log', 'message': f"🚨 An unexpected error occurred in backend processing: {e}"})
            import traceback
            self.gui_update_callback({'type': 'log', 'message': traceback.format_exc()})
        finally:
            self.processing_finished()
            
    def processing_finished(self):
        self.root.after(0, lambda: self.start_button.config(text="Start Processing", state=tk.NORMAL))
        self.root.after(0, lambda: self.copy_summary_button.config(state=tk.NORMAL)) # Re-enable copy button
        self.gui_update_callback({'type': 'log', 'message': "\n--- All Processing Finished. Check summary above. ---"})

    # --- NEW METHOD TO COPY SUMMARY ---
    def copy_summary_to_clipboard(self):
        summary_lines = []
        try:
            for category_key, parent_iid in self.tree_categories.items():
                category_text = self.summary_tree.item(parent_iid, 'text')
                summary_lines.append(f"\n{category_text}")
                
                children = self.summary_tree.get_children(parent_iid)
                if not children:
                    summary_lines.append("\t(No items in this category)")
                    continue

                for child_iid in children:
                    item_text = self.summary_tree.item(child_iid, 'text') # This is the short title
                    item_values = self.summary_tree.item(child_iid, 'values') # (DOI, StatusInfo, URL)
                    
                    # Ensure item_values is a list/tuple and has enough elements
                    doi_val = item_values[0] if len(item_values) > 0 else "N/A"
                    status_info_val = item_values[1] if len(item_values) > 1 else "N/A"
                    url_val = item_values[2] if len(item_values) > 2 else "N/A"

                    summary_lines.append(f"\t- Title: {item_text}")
                    summary_lines.append(f"\t  DOI: {doi_val}")
                    summary_lines.append(f"\t  Status/Note: {status_info_val}")
                    summary_lines.append(f"\t  URL: {url_val}")
            
            clipboard_text = "\n".join(summary_lines)
            
            if not clipboard_text.strip():
                self.gui_update_callback({'type': 'log', 'message': "Summary is empty. Nothing to copy."})
                return

            self.root.clipboard_clear()
            self.root.clipboard_append(clipboard_text)
            self.root.update() # To make it available to other apps immediately
            self.gui_update_callback({'type': 'log', 'message': "Summary copied to clipboard!"})
            messagebox.showinfo("Copied", "Summary copied to clipboard!")

        except Exception as e:
            self.gui_update_callback({'type': 'log', 'message': f"Error copying summary: {e}"})
            messagebox.showerror("Error", f"Could not copy summary to clipboard: {e}")
