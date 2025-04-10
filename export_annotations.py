import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import os
import json
from joppy.client_api import ClientApi
import xml.etree.ElementTree as ET
import win32api
import string
import requests
import socket
from urllib.parse import urljoin
from datetime import datetime

class KoboToJoplinApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kobo to Joplin Annotation Exporter")
        
        # Load configuration
        self.config = self.load_config()
        
        # Check Joplin service and API token
        if not self.check_joplin_service():
            messagebox.showerror("Error", "Could not connect to Joplin Web Clipper service. Please make sure Joplin is running and the Web Clipper is enabled.")
            return
            
        if not self.validate_api_token():
            messagebox.showerror("Error", "Invalid Joplin API token. Please check your configuration.")
            return
        
        # Initialize Joplin API
        self.joplin = ClientApi(token=self.config['joplin_api_token'])
        
        # Setup UI
        self.setup_ui()
        
        # Store selected annotations
        self.selected_annotations = []
        
        # Detect Kobo devices
        self.detect_kobo_devices()
        
    def check_joplin_service(self):
        """Check if Joplin Web Clipper service is running."""
        try:
            base_url = self.config['web_clipper']['url']
            port = self.config['web_clipper']['port']
            url = f"{base_url}:{port}/ping"
            
            # First check if the port is open
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()
            
            if result != 0:
                return False
                
            # Then check if the service responds
            response = requests.get(url, timeout=5)
            return response.status_code == 200
        except:
            return False
            
    def validate_api_token(self):
        """Validate the Joplin API token."""
        try:
            base_url = self.config['web_clipper']['url']
            port = self.config['web_clipper']['port']
            url = f"{base_url}:{port}/notes"
            params = {'token': self.config['joplin_api_token']}
            
            response = requests.get(url, params=params, timeout=5)
            return response.status_code == 200
        except:
            return False
            
    def load_config(self):
        """Load configuration from config.json file."""
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
                
            # Ensure web_clipper section exists
            if 'web_clipper' not in config:
                config['web_clipper'] = {
                    'url': 'http://localhost',
                    'port': 41184
                }
                with open(config_path, 'w') as f:
                    json.dump(config, f, indent=4)
                    
            return config
        except FileNotFoundError:
            # Create default config if it doesn't exist
            default_config = {
                "joplin_api_token": "",
                "notebook_id": "",
                "web_clipper": {
                    "url": "http://localhost",
                    "port": 41184
                }
            }
            with open(config_path, 'w') as f:
                json.dump(default_config, f, indent=4)
            return default_config
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid config.json file")
            return {
                "joplin_api_token": "",
                "notebook_id": "",
                "web_clipper": {
                    "url": "http://localhost",
                    "port": 41184
                }
            }
        
    def detect_kobo_devices(self):
        """Detect connected Kobo devices and update the dropdown."""
        self.kobo_devices = []
        self.device_paths = {}
        
        # Get all drive letters
        drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
        
        for drive in drives:
            try:
                # Check for Kobo device signature
                if os.path.exists(os.path.join(drive, ".kobo")):
                    # Get device name
                    device_name = win32api.GetVolumeInformation(drive)[0]
                    if not device_name:
                        device_name = "Kobo Device"
                    self.kobo_devices.append(device_name)
                    self.device_paths[device_name] = drive
            except:
                continue
        
        # Update dropdown
        self.device_dropdown['values'] = self.kobo_devices
        if self.kobo_devices:
            self.device_dropdown.set(self.kobo_devices[0])
            self.load_books()
        else:
            self.device_dropdown.set("No Kobo device detected")
            
    def setup_ui(self):
        """Setup the user interface."""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Device selection
        device_frame = ttk.Frame(main_frame)
        device_frame.pack(fill=tk.X, pady=5)
        ttk.Label(device_frame, text="Select Kobo Device:").pack(side=tk.LEFT, padx=5)
        self.device_dropdown = ttk.Combobox(device_frame, state="readonly")
        self.device_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.device_dropdown.bind('<<ComboboxSelected>>', lambda e: self.load_books())
        
        # Create paned window for the two panels
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Top panel - Books
        books_frame = ttk.Frame(paned)
        paned.add(books_frame, weight=1)
        
        # Books tree view
        books_tree_frame = ttk.Frame(books_frame)
        books_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbars for books tree
        books_scroll_y = ttk.Scrollbar(books_tree_frame)
        books_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        books_scroll_x = ttk.Scrollbar(books_tree_frame, orient=tk.HORIZONTAL)
        books_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Create books tree view
        self.books_tree = ttk.Treeview(books_tree_frame, 
                                      columns=('Book', 'Author', 'Annotations'),
                                      show='headings',
                                      yscrollcommand=books_scroll_y.set,
                                      xscrollcommand=books_scroll_x.set)
        
        # Configure scrollbars
        books_scroll_y.config(command=self.books_tree.yview)
        books_scroll_x.config(command=self.books_tree.xview)
        
        # Define columns for books
        self.books_tree.heading('Book', text='Book')
        self.books_tree.heading('Author', text='Author')
        self.books_tree.heading('Annotations', text='Annotations')
        
        # Set column widths for books
        self.books_tree.column('Book', width=200)
        self.books_tree.column('Author', width=150)
        self.books_tree.column('Annotations', width=100)
        
        self.books_tree.pack(expand=True, fill=tk.BOTH)
        
        # Add selection mode for books
        self.books_tree['selectmode'] = 'browse'
        self.books_tree.bind('<<TreeviewSelect>>', self.on_book_selected)
        
        # Bottom panel - Annotations
        annotations_frame = ttk.Frame(paned)
        paned.add(annotations_frame, weight=2)
        
        # Annotations tree view
        annotations_tree_frame = ttk.Frame(annotations_frame)
        annotations_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbars for annotations tree
        annotations_scroll_y = ttk.Scrollbar(annotations_tree_frame)
        annotations_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        annotations_scroll_x = ttk.Scrollbar(annotations_tree_frame, orient=tk.HORIZONTAL)
        annotations_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Create annotations tree view
        self.tree = ttk.Treeview(annotations_tree_frame, 
                                columns=('Book', 'Author', 'Annotation', 'Date', 'BookmarkID', 'Type'),
                                show='headings',
                                yscrollcommand=annotations_scroll_y.set,
                                xscrollcommand=annotations_scroll_x.set)
        
        # Configure scrollbars
        annotations_scroll_y.config(command=self.tree.yview)
        annotations_scroll_x.config(command=self.tree.xview)
        
        # Define columns for annotations
        self.tree.heading('Book', text='Book', command=lambda: self.treeview_sort_column('Book', False))
        self.tree.heading('Author', text='Author', command=lambda: self.treeview_sort_column('Author', False))
        self.tree.heading('Annotation', text='Annotation', command=lambda: self.treeview_sort_column('Annotation', False))
        self.tree.heading('Date', text='Date', command=lambda: self.treeview_sort_column('Date', False))
        
        # Hide the BookmarkID and Type columns as they're for internal use
        self.tree.column('BookmarkID', width=0, stretch=tk.NO)
        self.tree.column('Type', width=0, stretch=tk.NO)
        
        # Set column widths for annotations
        self.tree.column('Book', width=200)
        self.tree.column('Author', width=150)
        self.tree.column('Annotation', width=400)
        self.tree.column('Date', width=150)
        
        self.tree.pack(expand=True, fill=tk.BOTH)
        
        # Add selection mode for annotations
        self.tree['selectmode'] = 'extended'
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Export to Joplin", command=self.export_to_joplin).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Settings", command=self.open_settings).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
    def treeview_sort_column(self, col, reverse):
        # Get all items from the tree
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        # Sort the items
        items.sort(reverse=reverse)
        
        # Rearrange items in sorted positions
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
            
        # Toggle the sort direction for the next time
        self.tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))
        
    def load_books(self):
        """Load books and their annotation counts from the Kobo database."""
        try:
            # Clear existing items
            for item in self.books_tree.get_children():
                self.books_tree.delete(item)
            
            # Get selected device
            selected_device = self.device_dropdown.get()
            if not selected_device:
                return
            
            # Connect to database
            db_path = os.path.join(self.device_paths[selected_device], ".kobo", "KoboReader.sqlite")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Query books and their annotation counts
            query = """
                SELECT 
                    BookContent.Title,
                    BookContent.Attribution,
                    COUNT(Bookmark.BookmarkID) as AnnotationCount
                FROM Bookmark
                JOIN Content ON Bookmark.ContentID = Content.ContentID
                JOIN Content as BookContent ON Content.BookID = BookContent.ContentID
                WHERE (Bookmark.Text IS NOT NULL OR Bookmark.Type = 'markup')
                GROUP BY BookContent.Title, BookContent.Attribution
                ORDER BY BookContent.Title
            """
            
            cursor.execute(query)
            books = cursor.fetchall()
            
            # Add books to tree view
            for book in books:
                book_title = book[0] or "Unknown Title"
                author = book[1] or "Unknown Author"
                annotation_count = book[2]
                
                self.books_tree.insert('', 'end', values=(
                    book_title,
                    author,
                    annotation_count
                ))
            
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load books: {str(e)}")
            print(f"Error details: {str(e)}")
            
    def on_book_selected(self, event):
        """Handle book selection and load its annotations."""
        selected_items = self.books_tree.selection()
        if not selected_items:
            return
            
        # Clear existing annotations
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Get selected book details
        values = self.books_tree.item(selected_items[0])['values']
        book_title = values[0]
        author = values[1]
        
        # Load annotations for this book
        self.load_annotations_for_book(book_title, author)
        
    def load_annotations_for_book(self, book_title, author):
        """Load annotations for a specific book."""
        try:
            # Get selected device
            selected_device = self.device_dropdown.get()
            if not selected_device:
                return
            
            # Connect to database
            db_path = os.path.join(self.device_paths[selected_device], ".kobo", "KoboReader.sqlite")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Query annotations for this book
            query = """
                SELECT 
                    Bookmark.Text,
                    Bookmark.DateCreated,
                    Bookmark.BookmarkID,
                    Bookmark.Type,
                    CASE 
                        WHEN Bookmark.Type != 'markup' AND Chapter.Title IS NOT NULL THEN Chapter.Title
                        ELSE ''
                    END as ChapterTitle
                FROM Bookmark
                JOIN Content ON Bookmark.ContentID = Content.ContentID
                JOIN Content as BookContent ON Content.BookID = BookContent.ContentID
                LEFT JOIN Content as Chapter ON Content.ContentID = Chapter.ContentID 
                    AND Chapter.Title LIKE '%-%'
                    AND CAST(SUBSTR(Chapter.Title, INSTR(Chapter.Title, '-') + 1) AS INTEGER) IS NOT NULL
                WHERE BookContent.Title = ? 
                AND BookContent.Attribution = ?
                AND (Bookmark.Text IS NOT NULL OR Bookmark.Type = 'markup')
                ORDER BY Bookmark.DateCreated DESC
            """
            
            cursor.execute(query, (book_title, author))
            annotations = cursor.fetchall()
            
            # Add annotations to tree view
            for annotation in annotations:
                text = annotation[0] or ""
                date_created = annotation[1]
                bookmark_id = annotation[2]
                annotation_type = annotation[3]
                chapter_title = annotation[4]
                
                # Format date
                try:
                    if isinstance(date_created, str):
                        date_obj = datetime.fromisoformat(date_created)
                    else:
                        date_obj = datetime.fromtimestamp(int(date_created))
                    formatted_date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                except (ValueError, TypeError):
                    formatted_date = "Unknown Date"
                
                # For markup annotations, show a placeholder text
                if annotation_type == 'markup':
                    text = "[Markup annotation]"
                
                # Add to tree view
                self.tree.insert('', 'end', values=(
                    book_title,
                    author,
                    text,
                    formatted_date,
                    bookmark_id,
                    annotation_type
                ))
            
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load annotations: {str(e)}")
            print(f"Error details: {str(e)}")
            
    def export_to_joplin(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select annotations to export")
            return
            
        # Group annotations by book
        annotations_by_book = {}
        for item in selected_items:
            values = self.tree.item(item)['values']
            book_title = values[0]
            author = values[1]
            bookmark_id = values[4]
            
            if (book_title, author) not in annotations_by_book:
                annotations_by_book[(book_title, author)] = []
            annotations_by_book[(book_title, author)].append(bookmark_id)
            
        # Process each book's annotations
        for (book_title, author), bookmark_ids in annotations_by_book.items():
            # Get the original annotation data from the database
            db_path = os.path.join(self.device_paths[self.device_dropdown.get()], ".kobo", "KoboReader.sqlite")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Get all annotations for this book in one query
            query = """
                SELECT 
                    Bookmark.Text,
                    Bookmark.BookmarkID,
                    Bookmark.ContentID,
                    CASE 
                        WHEN Bookmark.Type = 'markup' THEN 'Markup'
                        WHEN Chapter.Title IS NOT NULL THEN Chapter.Title
                        ELSE ''
                    END as ChapterTitle,
                    Bookmark.Type,
                    Bookmark.DateCreated
                FROM Bookmark
                JOIN Content ON Bookmark.ContentID = Content.ContentID
                JOIN Content as BookContent ON Content.BookID = BookContent.ContentID
                LEFT JOIN Content as Chapter ON Chapter.ContentID LIKE Bookmark.ContentID || '-%'
                    AND CAST(SUBSTR(Chapter.ContentID, INSTR(Chapter.ContentID, '-') + 1) AS INTEGER) IS NOT NULL
                WHERE BookContent.Title = ? 
                AND BookContent.Attribution = ?
                AND Bookmark.BookmarkID IN ({})
                ORDER BY Bookmark.DateCreated
            """.format(','.join(['?'] * len(bookmark_ids)))
            
            cursor.execute(query, [book_title, author] + bookmark_ids)
            results = cursor.fetchall()
            
            if results:
                # Check if note exists
                note_title = f"{book_title} - {author}"
                notes = self.joplin.search_all(query=note_title, type_="note")
                existing_note = notes[0] if notes else None
                
                # Prepare content for all annotations
                all_content = []
                
                for result in results:
                    annotation_text = result[0] or ""
                    bookmark_id = result[1]
                    bookmark_content_id = result[2]
                    chapter_title = result[3] or ""
                    annotation_type = result[4] or ""
                    date_created = result[5]
                    
                    # Format the date for sorting
                    try:
                        if isinstance(date_created, str):
                            date_obj = datetime.fromisoformat(date_created)
                        else:
                            date_obj = datetime.fromtimestamp(int(date_created))
                        formatted_date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                    except (ValueError, TypeError):
                        formatted_date = "Unknown Date"
                    
                    # Prepare content for this annotation
                    content = []
                    
                    # Add chapter information
                    if chapter_title:
                        content.append(f"## {chapter_title}\n")
                    
                    # Add timestamp
                    content.append(f"### {formatted_date}\n")
                    
                    # Handle markup annotations
                    if annotation_type == 'markup' and bookmark_id:
                        markup_dir = os.path.join(self.device_paths[self.device_dropdown.get()], ".kobo", "markups")
                        svg_path = os.path.join(markup_dir, f"{bookmark_id}.svg")
                        jpg_path = os.path.join(markup_dir, f"{bookmark_id}.jpg")
                        
                        markup_file = None
                        if os.path.exists(svg_path):
                            markup_file = svg_path
                        elif os.path.exists(jpg_path):
                            markup_file = jpg_path
                        
                        if markup_file:
                            try:
                                # Add the resource
                                resource_id = self.joplin.add_resource(
                                    filename=markup_file,
                                    title=os.path.basename(markup_file)
                                )
                                
                                # Create the content with timestamp
                                markup_content = f"![Markup](:/{resource_id})"
                                if annotation_text and 'Associated Text:' in annotation_text:
                                    associated_text = annotation_text.split('Associated Text:')[1].strip()
                                    markup_content += f"\n{associated_text}"
                                
                                content.append(markup_content)
                                
                                # Link the resource to the note
                                if existing_note:
                                    self.joplin.add_resource_to_note(
                                        resource_id=resource_id,
                                        note_id=existing_note.id
                                    )
                                
                            except Exception as e:
                                print(f"Debug: Error uploading markup file: {str(e)}")
                                content.append("[Error attaching markup file]")
                    else:
                        # Regular annotation - wrap in code block
                        if annotation_text:  # Only add if there's actual text
                            content.append(f"```\n{annotation_text}\n```")
                    
                    # Join content if we have any content
                    if content:
                        all_content.append((formatted_date, '\n'.join(content)))
                
                try:
                    if existing_note:
                        # Update existing note with all new content
                        existing_content = existing_note.body
                        for timestamp, content in all_content:
                            existing_content = self.insert_content_in_order(existing_content, content, timestamp)
                        
                        self.joplin.modify_note(
                            id_=existing_note.id,
                            body=existing_content
                        )
                    else:
                        # Create new note with all content
                        final_content = []
                        for timestamp, content in all_content:
                            final_content.append(content)
                        
                        self.joplin.add_note(
                            title=note_title,
                            body='\n\n---\n\n'.join(final_content),
                            parent_id=self.config['notebook_id']
                        )
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to export note: {str(e)}")
                    print(f"Error details: {str(e)}")
                    continue
            else:
                print(f"Debug: No annotations found for {book_title} - {author}")
                continue
                
        messagebox.showinfo("Success", "Annotations exported successfully!")
        
    def insert_content_in_order(self, existing_content, new_content, timestamp):
        """Insert new content in chronological order within the existing note content."""
        if not existing_content:
            return f"Timestamp: {timestamp}\n{new_content}"
            
        # Split the existing content into sections
        sections = existing_content.split('\n\n---\n\n')
        
        # Create a list of (timestamp, content) tuples
        content_with_timestamps = []
        
        for section in sections:
            if not section:
                continue
                
            # Try to find a timestamp in the section
            section_timestamp = None
            for line in section.split('\n'):
                if line and line.startswith('Timestamp: '):
                    section_timestamp = line.replace('Timestamp: ', '')
                    break
            
            content_with_timestamps.append((section_timestamp or "Unknown Date", section))
        
        # Add the new content with its timestamp
        content_with_timestamps.append((timestamp, f"Timestamp: {timestamp}\n{new_content}"))
        
        # Sort by timestamp
        content_with_timestamps.sort(key=lambda x: x[0])
        
        # Rebuild the content
        return '\n\n---\n\n'.join(content[1] for content in content_with_timestamps)

    def open_settings(self):
        """Open the settings dialog."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("400x300")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # Create main frame
        main_frame = ttk.Frame(settings_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Joplin API Token
        ttk.Label(main_frame, text="Joplin API Token:").grid(row=0, column=0, sticky=tk.W, pady=5)
        api_token_var = tk.StringVar(value=self.config.get('joplin_api_token', ''))
        api_token_entry = ttk.Entry(main_frame, textvariable=api_token_var, width=40)
        api_token_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Notebook ID
        ttk.Label(main_frame, text="Notebook ID:").grid(row=1, column=0, sticky=tk.W, pady=5)
        notebook_id_var = tk.StringVar(value=self.config.get('notebook_id', ''))
        notebook_id_entry = ttk.Entry(main_frame, textvariable=notebook_id_var, width=40)
        notebook_id_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Web Clipper URL
        ttk.Label(main_frame, text="Web Clipper URL:").grid(row=2, column=0, sticky=tk.W, pady=5)
        web_clipper_url_var = tk.StringVar(value=self.config.get('web_clipper', {}).get('url', 'http://localhost'))
        web_clipper_url_entry = ttk.Entry(main_frame, textvariable=web_clipper_url_var, width=40)
        web_clipper_url_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Web Clipper Port
        ttk.Label(main_frame, text="Web Clipper Port:").grid(row=3, column=0, sticky=tk.W, pady=5)
        web_clipper_port_var = tk.StringVar(value=str(self.config.get('web_clipper', {}).get('port', 41184)))
        web_clipper_port_entry = ttk.Entry(main_frame, textvariable=web_clipper_port_var, width=40)
        web_clipper_port_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        def save_settings():
            """Save the settings and update the configuration."""
            self.config['joplin_api_token'] = api_token_var.get()
            self.config['notebook_id'] = notebook_id_var.get()
            self.config['web_clipper'] = {
                'url': web_clipper_url_var.get(),
                'port': int(web_clipper_port_var.get())
            }
            
            # Save to file
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            with open(config_path, 'w') as f:
                json.dump(self.config, f, indent=4)
            
            # Reinitialize Joplin API with new token
            self.joplin = ClientApi(token=self.config['joplin_api_token'])
            
            settings_window.destroy()
            messagebox.showinfo("Success", "Settings saved successfully!")
        
        ttk.Button(button_frame, text="Save", command=save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=settings_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)

if __name__ == "__main__":
    root = tk.Tk()
    app = KoboToJoplinApp(root)
    root.mainloop()
