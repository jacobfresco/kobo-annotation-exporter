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
import sys
import ebooklib
from ebooklib import epub
from PIL import Image, ImageTk
import io
import tempfile
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration
import cairosvg
import base64
import re
from bs4 import BeautifulSoup
import cssutils
import logging
import imgkit
import shutil

# Disable cssutils logging
cssutils.log.setLevel(logging.CRITICAL)

def check_dependencies():
    """Check if all required dependencies are installed and accessible."""
    missing_deps = []
    download_links = {
        'wkhtmltopdf': 'https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.msvc2015-win64.exe',
        'Python packages': 'pip install -r requirements.txt'
    }
    
    # Check Python packages
    required_packages = {
        'joppy': 'joppy',
        'win32api': 'pywin32',
        'requests': 'requests',
        'ebooklib': 'ebooklib',
        'PIL': 'Pillow',
        'bs4': 'beautifulsoup4',
        'cssutils': 'cssutils',
        'imgkit': 'imgkit'
    }
    
    for module, package in required_packages.items():
        try:
            __import__(module)
        except ImportError:
            missing_deps.append(f"Python package '{package}' is not installed")
    
    # Check wkhtmltopdf installation
    wkhtmltopdf_path = shutil.which('wkhtmltopdf')
    if not wkhtmltopdf_path:
        missing_deps.append("wkhtmltopdf is not installed or not in PATH")
    
    return missing_deps, download_links

class KoboToJoplinApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kobo to Joplin Annotation Exporter")
        
        # Load chapter formats configuration
        self.chapter_formats = self.load_chapter_formats()
        if not self.chapter_formats:
            messagebox.showerror("Error", "Failed to load chapter formats configuration")
            self.root.destroy()
            return
        
        # Check dependencies first
        missing_deps, download_links = check_dependencies()
        if missing_deps:
            error_msg = "The following dependencies are missing:\n\n" + "\n".join(missing_deps)
            error_msg += "\n\nInstallation instructions:\n"
            error_msg += "1. For Python packages, run: " + download_links['Python packages'] + "\n"
            error_msg += "2. For wkhtmltopdf:\n"
            error_msg += "   a. Download and install from: " + download_links['wkhtmltopdf'] + "\n"
            error_msg += "   b. Add wkhtmltopdf to PATH:\n"
            error_msg += "      - Open System Properties (Win + Pause/Break)\n"
            error_msg += "      - Click 'Advanced system settings'\n"
            error_msg += "      - Click 'Environment Variables'\n"
            error_msg += "      - Under 'System variables', find and select 'Path'\n"
            error_msg += "      - Click 'Edit' and add 'C:\\Program Files\\wkhtmltopdf\\bin'\n"
            error_msg += "      - Click 'OK' on all windows\n"
            error_msg += "   c. No restart required - just close and reopen this application\n"
            error_msg += "\nAfter installing, please restart the application."
            
            # Create a more detailed error window
            error_window = tk.Toplevel(self.root)
            error_window.title("Missing Dependencies")
            error_window.geometry("700x500")  # Made window larger to fit instructions
            
            # Add text widget with scrollbar
            frame = ttk.Frame(error_window)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            text_widget = tk.Text(frame, wrap=tk.WORD)
            scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Insert error message
            text_widget.insert(tk.END, error_msg)
            text_widget.configure(state='disabled')  # Make read-only
            
            # Add buttons
            button_frame = ttk.Frame(error_window)
            button_frame.pack(fill=tk.X, pady=5)
            
            def open_download_link():
                import webbrowser
                webbrowser.open(download_links['wkhtmltopdf'])
            
            ttk.Button(button_frame, text="Download wkhtmltopdf", 
                      command=open_download_link).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Close", 
                      command=lambda: [error_window.destroy(), self.root.destroy()]).pack(side=tk.LEFT, padx=5)
            
            # Make the error window modal
            error_window.transient(self.root)
            error_window.grab_set()
            self.root.wait_window(error_window)
            return
        
        # Set window icon
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        icon_path = os.path.join(base_path, 'icon.ico')
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        
        # Load configuration
        self.config = self.load_config()
        if not self.config:  # If config loading failed
            self.root.destroy()
            return
        
        # Check Joplin service and API token
        if not self.check_joplin_service():
            messagebox.showerror("Error", "Could not connect to Joplin Web Clipper service. Please make sure Joplin is running and the Web Clipper is enabled.")
            self.root.destroy()
            return
            
        if not self.validate_api_token():
            messagebox.showerror("Error", "Invalid Joplin API token. Please check your configuration.")
            self.root.destroy()
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
        """Load configuration from config.json or create from default if not exists"""
        try:
            # Get the directory where the script or executable is located
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = os.path.dirname(sys.executable)
            else:
                # Running as script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            config_path = os.path.join(base_path, 'config.json')
            default_config_path = os.path.join(base_path, 'config.json.default')
            
            print(f"Looking for config at: {config_path}")  # Debug print
            
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    print("Successfully loaded config.json")  # Debug print
                    return config
            elif os.path.exists(default_config_path):
                with open(default_config_path, 'r') as f:
                    config = json.load(f)
                    print("Using default config")  # Debug print
                    return config
            else:
                print("No config files found")  # Debug print
                messagebox.showerror("Error", "No configuration file found. Please create config.json or config.json.default")
                self.root.destroy()
                return None
        except Exception as e:
            print(f"Error loading config: {str(e)}")  # Debug print
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
            self.root.destroy()
            return None
        
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
        
        self.export_button = ttk.Button(button_frame, text="Export to Joplin", command=self.export_to_joplin)
        self.export_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Settings", command=self.open_settings).pack(side=tk.LEFT, padx=5)
        
        # Bind selection event to update button text
        self.tree.bind('<<TreeviewSelect>>', self.update_export_button_text)
        
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
            
    def locate_epub_file(self, book_title, author):
        """Locate the EPUB file for a given book on the Kobo device."""
        try:
            selected_device = self.device_dropdown.get()
            if not selected_device:
                return None

            # Common paths where Kobo stores books
            possible_paths = [
                os.path.join(self.device_paths[selected_device], "Digital Editions"),
                os.path.join(self.device_paths[selected_device], "Books"),
                os.path.join(self.device_paths[selected_device], "eBooks")
            ]

            # Search for the EPUB file
            for base_path in possible_paths:
                if not os.path.exists(base_path):
                    continue

                # Walk through all subdirectories
                for root, dirs, files in os.walk(base_path):
                    for file in files:
                        if file.lower().endswith('.epub'):
                            # Try to match the book by title and author
                            try:
                                epub_path = os.path.join(root, file)
                                book = epub.read_epub(epub_path)
                                
                                # Get book metadata
                                metadata = book.get_metadata('DC', 'title')
                                if not metadata:
                                    continue
                                    
                                epub_title = metadata[0][0]
                                epub_author = book.get_metadata('DC', 'creator')
                                epub_author = epub_author[0][0] if epub_author else ""
                                
                                # Compare titles and authors (case-insensitive)
                                if (book_title.lower() in epub_title.lower() or 
                                    epub_title.lower() in book_title.lower()) and \
                                   (not author or not epub_author or 
                                    author.lower() in epub_author.lower() or 
                                    epub_author.lower() in author.lower()):
                                    return epub_path
                            except Exception as e:
                                print(f"Error reading EPUB file {file}: {str(e)}")
                                continue

            return None
        except Exception as e:
            print(f"Error locating EPUB file: {str(e)}")
            return None

    def get_page_image(self, epub_path, content_id, chapter_progress=None):
        """Extract a specific page from an EPUB file as an image."""
        try:
            print(f"\n=== Page Image Generation Debug ===")
            print(f"Content ID: {content_id}")
            print(f"Chapter Progress: {chapter_progress}")
            print(f"EPUB Path: {epub_path}")
            
            # Read the EPUB file
            book = epub.read_epub(epub_path)
            
            # Determine if this is a KEPUB
            is_kepub = False
            for format_config in self.chapter_formats['kepub_formats']:
                if format_config['path_marker'] in content_id:
                    is_kepub = True
                    break
            
            print(f"Is KEPUB: {is_kepub}")
            
            # Parse the content ID to get the chapter and position
            try:
                # Check for KEPUB formats first
                for format_config in self.chapter_formats['kepub_formats']:
                    if format_config['path_marker'] in content_id:
                        print(f"\nProcessing KEPUB content ID: {content_id}")
                        # Extract the chapter number using the configured pattern
                        chapter_match = re.search(format_config['chapter_pattern'], content_id)
                        if chapter_match:
                            chapter_num = int(chapter_match.group(1))
                            position = 0  # Position is not available in KEPUB format
                            print(f"Found chapter number: {chapter_num}")
                            
                            # For KEPUB, we need to find the specific chapter file
                            chapter_path = None
                            print("Searching for chapter file...")
                            
                            # Get all document items and sort them by their href
                            doc_items = []
                            print("Available chapters in EPUB:")
                            for item in book.get_items():
                                if item.get_type() == ebooklib.ITEM_DOCUMENT:
                                    href = item.get_name()
                                    print(f"Found document: {href}")
                                    # Check if the href matches any KEPUB format pattern
                                    if any(re.search(fmt['chapter_pattern'], href, re.IGNORECASE) for fmt in self.chapter_formats['kepub_formats']):
                                        doc_items.append(item)
                                        print(f"Added KEPUB chapter: {href}")
                            
                            print(f"Total KEPUB chapters found: {len(doc_items)}")
                            
                            # Sort by the chapter number in the filename
                            def get_chapter_number(item):
                                href = item.get_name()
                                for fmt in self.chapter_formats['kepub_formats']:
                                    match = re.search(fmt['chapter_pattern'], href, re.IGNORECASE)
                                    if match:
                                        return int(match.group(1))
                                print(f"No chapter number found in {href}")
                                return 0
                            
                            doc_items.sort(key=get_chapter_number)
                            
                            # Find the chapter by its number
                            chapter = None
                            print(f"Looking for chapter with number {chapter_num}")
                            for item in doc_items:
                                href = item.get_name()
                                for fmt in self.chapter_formats['kepub_formats']:
                                    match = re.search(fmt['chapter_pattern'], href, re.IGNORECASE)
                                    if match:
                                        current_num = int(match.group(1))
                                        print(f"Checking chapter {current_num}")
                                        if current_num == chapter_num:
                                            chapter = item
                                            print(f"Found matching chapter: {href}")
                                            break
                                if chapter:
                                    break
                            
                            if chapter:
                                print(f"Found chapter: {chapter.get_name()}")
                            else:
                                print(f"Could not find chapter with number {chapter_num}")
                                return None
                                
                            print(f"Successfully loaded chapter content")
                            break
                else:
                    # Handle regular EPUB format
                    print(f"\nProcessing regular EPUB content ID: {content_id}")
                    
                    # Special handling for OEBPS/part format
                    if 'OEBPS/part' in content_id:
                        chapter_match = re.search(r'part(\d+)\.xhtml', content_id)
                        if chapter_match:
                            chapter_num = int(chapter_match.group(1))
                            position = 0
                            print(f"Found OEBPS chapter number: {chapter_num}")
                            
                            # Get all document items
                            doc_items = []
                            print("Searching for OEBPS chapter files...")
                            for item in book.get_items():
                                if item.get_type() == ebooklib.ITEM_DOCUMENT:
                                    href = item.get_name()
                                    print(f"Found document: {href}")
                                    if 'OEBPS/part' in href:
                                        doc_items.append(item)
                                        print(f"Added OEBPS chapter: {href}")
                            
                            print(f"Total OEBPS chapters found: {len(doc_items)}")
                            
                            # Sort by the chapter number in the filename
                            def get_chapter_number(item):
                                href = item.get_name()
                                match = re.search(r'part(\d+)\.xhtml', href)
                                if match:
                                    return int(match.group(1))
                                print(f"No chapter number found in {href}")
                                return 0
                            
                            doc_items.sort(key=get_chapter_number)
                            
                            # Find the chapter by its number
                            chapter = None
                            print(f"Looking for OEBPS chapter with number {chapter_num}")
                            for item in doc_items:
                                href = item.get_name()
                                match = re.search(r'part(\d+)\.xhtml', href)
                                if match:
                                    current_num = int(match.group(1))
                                    print(f"Checking chapter {current_num}")
                                    if current_num == chapter_num:
                                        chapter = item
                                        print(f"Found matching chapter: {href}")
                                        break
                            
                            if chapter:
                                print(f"Found chapter: {chapter.get_name()}")
                            else:
                                print(f"Could not find OEBPS chapter with number {chapter_num}")
                                return None
                            
                            print(f"Successfully loaded OEBPS chapter content")
                        else:
                            print(f"Could not parse OEBPS chapter number from content ID")
                            return None
                    else:
                        # Handle other EPUB formats
                        for format_config in self.chapter_formats['epub_formats']:
                            if format_config['path_marker'] in content_id:
                                parts = content_id.split(format_config['path_marker'])
                                if len(parts) > 1:
                                    chapter_info = parts[1]
                                    chapter_match = re.search(format_config['chapter_pattern'], chapter_info)
                                    if chapter_match:
                                        chapter_num = int(chapter_match.group(1))
                                        position = int(chapter_match.group(2)) if chapter_match.group(2) else 0
                                        
                                        # Get all document items and sort them
                                        doc_items = [item for item in book.get_items() if item.get_type() == ebooklib.ITEM_DOCUMENT]
                                        doc_items.sort(key=lambda x: x.get_name())
                                        
                                        # Find the chapter by its position in the sorted list
                                        if 0 <= chapter_num - 1 < len(doc_items):
                                            chapter = doc_items[chapter_num - 1]
                                            print(f"Found chapter: {chapter.get_name()}")
                                        else:
                                            print(f"Chapter number {chapter_num} out of range")
                                            return None
                                        break
            except (ValueError, IndexError) as e:
                print(f"Invalid content ID format: {content_id}, error: {str(e)}")
                return None
            
            if not chapter:
                print("No chapter found for the given content ID")
                return None
            
            # Get the chapter content as HTML
            print("\nGetting chapter content as HTML...")
            html_content = chapter.get_content().decode('utf-8')
            print(f"HTML content length: {len(html_content)}")
            
            # Parse HTML with BeautifulSoup
            print("Parsing HTML with BeautifulSoup...")
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Extract and inline CSS
            styles = soup.find_all('style')
            css_text = ''
            for style in styles:
                css_text += style.string + '\n'
            
            # Create a temporary directory for assets
            with tempfile.TemporaryDirectory() as temp_dir:
                # Process images in the HTML
                print("Processing images in HTML...")
                for img in soup.find_all('img'):
                    if img.get('src'):
                        # Get the image data from the EPUB
                        img_path = img['src']
                        img_item = book.get_item_with_href(img_path)
                        if img_item:
                            # Convert image to base64
                            img_data = img_item.get_content()
                            img_b64 = base64.b64encode(img_data).decode('utf-8')
                            img['src'] = f"data:image/png;base64,{img_b64}"
                
                # Get font settings from config
                font_family = self.config.get('font', {}).get('family', 'Arial')
                font_size = self.config.get('font', {}).get('size', 16)
                use_kobo_font = self.config.get('font', {}).get('use_kobo_font', False)
                kobo_font = self.config.get('font', {}).get('kobo_font', '')
                
                print(f"\nUsing font settings:")
                print(f"  Font family: {font_family}")
                print(f"  Font size: {font_size}")
                print(f"  Use Kobo font: {use_kobo_font}")
                print(f"  Kobo font: {kobo_font}")
                
                # If using Kobo font, try to find it in the device's fonts directory
                if use_kobo_font and kobo_font:
                    selected_device = self.device_dropdown.get()
                    if selected_device:
                        kobo_font_path = os.path.join(self.device_paths[selected_device], "fonts", kobo_font)
                        if os.path.exists(kobo_font_path):
                            # Create @font-face rule for the Kobo font
                            font_face = f"""
                            @font-face {{
                                font-family: 'KoboFont';
                                src: url('file:///{kobo_font_path}') format('truetype');
                            }}
                            """
                            css_text = font_face + css_text
                            font_family = 'KoboFont'
                
                # Create a complete HTML document with proper styling
                print("\nCreating HTML document...")
                html_doc = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <meta name="viewport" content="width=800">
                    <style>
                        @page {{
                            size: 800px 1800px;
                            margin: 0;
                        }}
                        html {{
                            width: 800px;
                            margin: 0;
                            padding: 0;
                        }}
                        body {{
                            margin: 0;
                            padding: 20px;
                            font-family: {font_family}, sans-serif;
                            font-size: {font_size}px;
                            line-height: 1.4;
                            color: #333;
                            width: 760px;
                            background: transparent;
                        }}
                        img {{
                            max-width: 100%;
                            height: auto;
                        }}
                        p {{
                            margin: 0 0 1em 0;
                        }}
                        h1, h2, h3, h4, h5, h6 {{
                            margin: 1em 0 0.5em 0;
                            font-family: {font_family}, sans-serif;
                            font-size: {font_size * 1.2}px;
                        }}
                        {css_text}
                    </style>
                </head>
                <body>
                    <div id="content">
                        {soup.body.decode_contents() if soup.body else soup.decode_contents()}
                    </div>
                </body>
                </html>
                """

                # Calculate the vertical offset for cropping
                page_height = 1800  # px, as set in your CSS
                if chapter_progress is not None:
                    print(f"\n=== Page Position Calculation ===")
                    print(f"Raw chapter_progress from DB: {chapter_progress}")
                    
                    # Convert chapter_progress to float if it's a string
                    if isinstance(chapter_progress, str):
                        try:
                            chapter_progress = float(chapter_progress)
                            print(f"Converted chapter_progress to float: {chapter_progress}")
                        except ValueError:
                            print(f"Error converting chapter_progress to float: {chapter_progress}")
                            chapter_progress = 0.0
                    
                    # Ensure chapter_progress is between 0 and 1
                    chapter_progress = max(0.0, min(1.0, chapter_progress))
                    print(f"Normalized chapter_progress: {chapter_progress}")
                    
                    # Calculate total height of the chapter content
                    # First render the full chapter to get its actual height
                    options = {
                        'format': 'png',
                        'encoding': 'UTF-8',
                        'width': 800,
                        'height': 10000,  # Use a large height to get full content
                        'enable-local-file-access': None,
                        'disable-smart-width': None,
                        'quality': 100,
                        'quiet': None,
                        'log-level': 'info'
                    }
                    
                    # Create temporary files for HTML and PNG
                    html_path = os.path.join(temp_dir, 'page.html')
                    full_png_path = os.path.join(temp_dir, 'full_page.png')
                    
                    # Write HTML to file
                    with open(html_path, 'w', encoding='utf-8') as f:
                        f.write(html_doc)
                    
                    # Convert HTML to PNG to get full height
                    imgkit.from_file(html_path, full_png_path, options=options)
                    full_img = Image.open(full_png_path)
                    total_height = full_img.height
                    print(f"Total chapter height: {total_height}px")
                    
                    # Calculate target position based on chapter progress
                    if is_kepub:
                        # For KEPUB, we need to handle the page counting differently
                        # KEPUBs use a fixed number of pages per chapter
                        # We'll use the chapter progress to determine which page to show
                        print("Using KEPUB page calculation")
                        # Calculate the number of pages in the chapter
                        num_pages = (total_height + page_height - 1) // page_height
                        print(f"Number of pages in chapter: {num_pages}")
                        
                        # Calculate which page to show based on chapter progress
                        target_page = int(chapter_progress * num_pages)
                        print(f"Target page number: {target_page}")
                        
                        # Calculate the crop position for that page
                        crop_y = target_page * page_height
                        print(f"KEPUB crop_y position: {crop_y}px")
                    else:
                        # For regular EPUB, use the continuous position calculation
                        print("Using EPUB continuous position calculation")
                        target_position = int(chapter_progress * total_height)
                        print(f"Target position in chapter: {target_position}px")
                        
                        # Calculate which page this position falls on
                        page_number = target_position // page_height
                        position_in_page = target_position % page_height
                        print(f"Page number in chapter: {page_number}")
                        print(f"Position within page: {position_in_page}px")
                        
                        # Calculate crop position to show the target position in the middle of the viewport
                        crop_y = max(0, target_position - (page_height // 2))
                        print(f"Initial crop_y position: {crop_y}px")
                    
                    # Adjust crop_y to ensure we don't go beyond the total height
                    max_crop_y = max(0, total_height - page_height)
                    crop_y = min(crop_y, max_crop_y)
                    print(f"Final adjusted crop_y position: {crop_y}px")
                    
                    # Crop the full image to get the specific page
                    crop_box = (0, crop_y, 800, min(crop_y + page_height, total_height))
                    print(f"Cropping image with box: {crop_box}")
                    img = full_img.crop(crop_box)
                    print(f"Cropped image size: {img.size}")
                    
                    # Clean up temporary files
                    try:
                        os.unlink(full_png_path)
                    except Exception as e:
                        print(f"Warning: Could not delete temporary file {full_png_path}: {str(e)}")
                    
                    return img, crop_y, total_height
                else:
                    print("\nNo chapter progress value available, rendering first page")
                    # If no chapter progress, just render the first page
                    options = {
                        'format': 'png',
                        'encoding': 'UTF-8',
                        'width': 800,
                        'height': page_height,
                        'enable-local-file-access': None,
                        'disable-smart-width': None,
                        'quality': 100,
                        'quiet': None,
                        'log-level': 'info'
                    }
                    
                    # Create temporary files for HTML and PNG
                    html_path = os.path.join(temp_dir, 'page.html')
                    png_path = os.path.join(temp_dir, 'page.png')
                    
                    # Write HTML to file
                    with open(html_path, 'w', encoding='utf-8') as f:
                        f.write(html_doc)
                    
                    # Convert HTML to PNG
                    imgkit.from_file(html_path, png_path, options=options)
                    img = Image.open(png_path)
                    
                    # Clean up temporary files
                    try:
                        os.unlink(png_path)
                    except Exception as e:
                        print(f"Warning: Could not delete temporary file {png_path}: {str(e)}")
                    
                    return img, 0, page_height
                
        except Exception as e:
            print(f"Error extracting page from EPUB: {str(e)}")
            # Return a placeholder image with error message
            img = Image.new('RGB', (800, 1800), color='white')
            return img, 0, 1800

    def merge_markup_with_page(self, markup_path, page_image, chapter_progress=None, crop_y=0, total_height=0):
        """Merge a markup image with a page image."""
        try:
            # Convert SVG to PNG if needed
            if markup_path.lower().endswith('.svg'):
                # Read SVG file
                with open(markup_path, 'rb') as f:
                    svg_data = f.read()
                
                # Create a temporary file for the PNG
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                    temp_path = temp_file.name
                
                # Convert SVG to PNG using cairosvg
                cairosvg.svg2png(bytestring=svg_data, write_to=temp_path)
                
                # Open the converted PNG
                markup_img = Image.open(temp_path)
                
                # Clean up temporary file
                try:
                    os.unlink(temp_path)
                except Exception as e:
                    print(f"Warning: Could not delete temporary file {temp_path}: {str(e)}")
            else:
                # Open the markup image directly if it's not SVG
                markup_img = Image.open(markup_path)
            
            # Convert both images to RGBA if they aren't already
            markup_img = markup_img.convert('RGBA')
            page_image = page_image.convert('RGBA')
            
            # Get the dimensions
            page_width, page_height = page_image.size
            markup_width, markup_height = markup_img.size
            
            print(f"\n=== Markup Positioning Debug ===")
            print(f"Page dimensions: {page_width}x{page_height}")
            print(f"Markup dimensions: {markup_width}x{markup_height}")
            
            # Calculate the scaling factor needed to match the page width
            scale_factor = page_width / markup_width
            print(f"Scale factor: {scale_factor}")
            
            # Resize the markup image to match the page width while maintaining aspect ratio
            new_markup_height = int(markup_height * scale_factor)
            markup_img = markup_img.resize((page_width, new_markup_height), Image.Resampling.LANCZOS)
            print(f"Resized markup dimensions: {page_width}x{new_markup_height}")
            
            # If we have chapter progress, we need to adjust the markup position
            if chapter_progress is not None and total_height > 0:
                print(f"\nMarkup Position Calculation:")
                print(f"Chapter progress: {chapter_progress}")
                print(f"Total height: {total_height}")
                print(f"Crop Y: {crop_y}")
                
                # Calculate the scale factor between the markup and the page content
                content_scale = new_markup_height / total_height
                print(f"Content scale: {content_scale}")
                
                # Calculate the markup offset based on the page's crop position
                markup_offset = int(crop_y * content_scale)
                print(f"Initial markup offset: {markup_offset}")
                
                # For KEPUB, we need to handle the markup position differently
                # The markup should be positioned relative to the current page
                if markup_path.lower().endswith('.svg'):
                    print("Using KEPUB markup positioning")
                    # For KEPUB, the markup is already positioned correctly for the current page
                    # We just need to ensure it's within the page bounds
                    markup_offset = max(0, min(markup_offset, new_markup_height - page_height))
                else:
                    print("Using EPUB markup positioning")
                    # For EPUB, we need to adjust the markup position based on the page's crop position
                    markup_offset = int(crop_y * content_scale)
                
                print(f"Final markup offset: {markup_offset}")
                
                # Crop the markup image to match the page height
                crop_box = (0, markup_offset, page_width, markup_offset + page_height)
                print(f"Crop box: {crop_box}")
                markup_img = markup_img.crop(crop_box)
                print(f"Cropped markup dimensions: {markup_img.size}")
            
            # Create a new image with the same size as the page
            merged_img = Image.new('RGBA', page_image.size, (255, 255, 255, 0))
            
            # Paste the page image
            merged_img.paste(page_image, (0, 0))
            
            # Paste the markup on top with alpha compositing
            merged_img = Image.alpha_composite(merged_img, markup_img)
            
            # Convert back to RGB for saving
            merged_img = merged_img.convert('RGB')
            
            return merged_img
            
        except Exception as e:
            print(f"Error merging markup with page: {str(e)}")
            return None

    def get_annotation_position(self, db_path, bookmark_id):
        """Get the exact position information for an annotation from the Kobo database."""
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Query to get position information
            query = """
                SELECT 
                    Bookmark.ContentID,
                    Bookmark.Annotation,
                    Bookmark.Text,
                    Content.ContentID as EpubPath,
                    Bookmark.ChapterProgress,
                    Bookmark.VolumeId,
                    Bookmark.StartContainerPath,
                    Bookmark.EndContainerPath,
                    Bookmark.StartOffset,
                    Bookmark.EndOffset
                FROM Bookmark
                JOIN Content ON Bookmark.ContentID = Content.ContentID
                WHERE Bookmark.BookmarkID = ?
            """
            
            cursor.execute(query, (bookmark_id,))
            result = cursor.fetchone()
            
            if result:
                content_id = result[0]
                annotation = result[1]
                text = result[2]
                epub_path = result[3]
                ChapterProgress = result[4]
                volume_id = result[5]
                start_container = result[6]
                end_container = result[7]
                start_offset = result[8]
                end_offset = result[9]
                
                print(f"Database values:")  # Debug log
                print(f"  ContentID: {content_id}")
                print(f"  ChapterProgress: {ChapterProgress}")
                print(f"  VolumeId: {volume_id}")
                print(f"  StartContainerPath: {start_container}")
                print(f"  EndContainerPath: {end_container}")
                print(f"  StartOffset: {start_offset}")
                print(f"  EndOffset: {end_offset}")
                
                # Parse the content ID to get position
                try:
                    # Check for KEPUB formats first
                    for format_config in self.chapter_formats['kepub_formats']:
                        if format_config['path_marker'] in content_id:
                            # Extract the chapter number using the configured pattern
                            chapter_match = re.search(format_config['chapter_pattern'], content_id)
                            if chapter_match:
                                chapter_num = int(chapter_match.group(1))
                                position = 0  # Position is not available in KEPUB format
                                
                                # Extract the base EPUB path
                                epub_path = content_id.split(format_config['epub_path_split'])[0]
                                
                                return {
                                    'chapter_num': chapter_num,
                                    'position': position,
                                    'content_id': content_id,
                                    'epub_path': epub_path,
                                    'annotation': annotation,
                                    'text': text,
                                    'ChapterProgress': ChapterProgress,
                                    'start_container': start_container,
                                    'end_container': end_container,
                                    'start_offset': start_offset,
                                    'end_offset': end_offset
                                }
                    
                    # If not a KEPUB format, check EPUB formats
                    for format_config in self.chapter_formats['epub_formats']:
                        if format_config['path_marker'] in content_id:
                            parts = content_id.split(format_config['path_marker'])
                            if len(parts) > 1:
                                chapter_info = parts[1]
                                chapter_match = re.search(format_config['chapter_pattern'], chapter_info)
                                if chapter_match:
                                    chapter_num = int(chapter_match.group(1))
                                    position = int(chapter_match.group(2)) if chapter_match.group(2) else 0
                                    epub_path = parts[0]
                                    
                                    return {
                                        'chapter_num': chapter_num,
                                        'position': position,
                                        'content_id': content_id,
                                        'epub_path': epub_path,
                                        'annotation': annotation,
                                        'text': text,
                                        'ChapterProgress': ChapterProgress
                                    }
                    
                    # Special handling for OEBPS/partXXXX.xhtml format
                    if 'OEBPS/part' in content_id:
                        chapter_match = re.search(r'part(\d+)\.xhtml', content_id)
                        if chapter_match:
                            chapter_num = int(chapter_match.group(1))
                            position = 0
                            epub_path = content_id.split('!!')[0]
                            
                            return {
                                'chapter_num': chapter_num,
                                'position': position,
                                'content_id': content_id,
                                'epub_path': epub_path,
                                'annotation': annotation,
                                'text': text,
                                'ChapterProgress': ChapterProgress
                            }
                    
                    print(f"Could not match content ID to any known format: {content_id}")
                    return None
                    
                except (ValueError, IndexError) as e:
                    print(f"Error parsing content ID: {content_id}, error: {str(e)}")
                    return None
                    
            return None
            
        except Exception as e:
            print(f"Error getting annotation position: {str(e)}")
            return None
        finally:
            if conn:
                conn.close()

    def preview_combined_image(self, image, bookmark_id):
        """Show a preview window for the combined image."""
        print(f"Creating preview window for bookmark {bookmark_id}...")  # Debug log
        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"Preview - Bookmark {bookmark_id}")
        
        # Set window icon
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        icon_path = os.path.join(base_path, 'icon.ico')
        if os.path.exists(icon_path):
            preview_window.iconbitmap(icon_path)
        
        # Make the preview window modal and keep it on top
        preview_window.transient(self.root)
        preview_window.grab_set()
        
        # Get screen dimensions
        screen_width = preview_window.winfo_screenwidth()
        screen_height = preview_window.winfo_screenheight()
        
        # Calculate window size based on image dimensions
        img_width, img_height = image.size
        # Use 90% of screen width as maximum width
        max_width = int(screen_width * 0.9)
        # Use 75% of screen height for image and 10% for buttons
        max_image_height = int(screen_height * 0.75)
        button_height = int(screen_height * 0.1)
        
        # Calculate scale factors for both width and height
        width_scale = max_width / img_width
        height_scale = max_image_height / img_height
        
        # Use the smaller scale to ensure both dimensions fit
        scale = min(width_scale, height_scale)
        
        # Calculate final window dimensions
        window_width = int(img_width * scale)
        window_height = int(img_height * scale) + button_height
        
        # Calculate position to center the window
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Create main frame with padding
        main_frame = ttk.Frame(preview_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add buttons in a frame at the top
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        
        def export_and_close():
            print(f"Exporting bookmark {bookmark_id} to Joplin...")  # Debug log
            try:
                # Get book title and author from the database
                db_path = os.path.join(self.device_paths[self.device_dropdown.get()], ".kobo", "KoboReader.sqlite")
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # Query to get book title and author
                query = """
                    SELECT 
                        BookContent.Title,
                        BookContent.Attribution
                    FROM Bookmark
                    JOIN Content ON Bookmark.ContentID = Content.ContentID
                    JOIN Content as BookContent ON Content.BookID = BookContent.ContentID
                    WHERE Bookmark.BookmarkID = ?
                """
                
                cursor.execute(query, (bookmark_id,))
                result = cursor.fetchone()
                conn.close()
                
                if not result:
                    raise Exception("Could not find book information")
                
                book_title = result[0] or "Unknown Title"
                author = result[1] or "Unknown Author"
                
                
                # Save the image to a temporary file
                temp_path = os.path.join(tempfile.gettempdir(), f"preview_{bookmark_id}.png")
                print(f"Saving image to: {temp_path}")  # Debug log
                image.save(temp_path)
                
                print("Adding image as resource to Joplin...")  # Debug log
                # Add the image as a resource
                resource_id = self.joplin.add_resource(
                    filename=temp_path,
                    title=f"Markup with Page {bookmark_id}"
                )
                
                print("Creating note in Joplin...")  # Debug log
                # Create a new note with the image using the same title format as other annotations
                self.joplin.add_note(
                    title=f"{book_title} - {author}",
                    body=f"![Markup with Page](:/{resource_id})",
                    parent_id=self.config['notebook_id']
                )
                
                print("Export completed successfully")  # Debug log
                preview_window.destroy()
                
            except Exception as e:
                print(f"Error during export: {str(e)}")  # Debug log
                preview_window.destroy()
                # Show error message in the main window
                self.root.after(100, lambda: messagebox.showerror("Error", f"Failed to export to Joplin: {str(e)}"))
            finally:
                # Clean up temporary file
                if os.path.exists(temp_path):
                    print(f"Cleaning up temporary file: {temp_path}")  # Debug log
                    os.remove(temp_path)
        
        def save_image():
            """Save the image to a file."""
            try:
                # Get the default filename from the bookmark ID
                default_filename = f"annotation_{bookmark_id}.png"
                
                # Ask user for save location
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
                    initialfile=default_filename
                )
                
                if file_path:  # If user didn't cancel
                    print(f"Saving image to: {file_path}")  # Debug log
                    image.save(file_path)
                    messagebox.showinfo("Success", "Image saved successfully!")
            except Exception as e:
                print(f"Error saving image: {str(e)}")  # Debug log
                messagebox.showerror("Error", f"Failed to save image: {str(e)}")
        
        # Create buttons with more padding and make them more visible
        export_button = ttk.Button(button_frame, text="Export to Joplin", 
                                 command=export_and_close, padding=5)
        export_button.pack(side=tk.LEFT, padx=5)
        
        save_button = ttk.Button(button_frame, text="Save Image", 
                               command=save_image, padding=5)
        save_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = ttk.Button(button_frame, text="Cancel", 
                                 command=preview_window.destroy, padding=5)
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Create canvas frame
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create canvas with scrollbars
        canvas = tk.Canvas(canvas_frame, width=window_width, height=int(img_height * scale))
        scrollbar_y = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar_x = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        
        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Pack scrollbars and canvas
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        print("Converting image for preview...")  # Debug log
        # Convert PIL image to PhotoImage
        # Resize image to match window size
        resized_image = image.resize((window_width, int(img_height * scale)), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(resized_image)
        
        # Add image to canvas
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        canvas.image = photo  # Keep a reference
        
        # Configure canvas scrolling region
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        print("Preview window created successfully")  # Debug log

    def export_to_joplin(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select annotations to export")
            return
            
        print("Starting export process...")  # Debug log
            
        # Group annotations by book
        annotations_by_book = {}
        for item in selected_items:
            values = self.tree.item(item)['values']
            book_title = values[0]
            author = values[1]
            bookmark_id = values[4]
            annotation_type = values[5]  # Get the annotation type
            
            print(f"Processing annotation: {bookmark_id}, type: {annotation_type}")  # Debug log
            
            if (book_title, author) not in annotations_by_book:
                annotations_by_book[(book_title, author)] = []
            annotations_by_book[(book_title, author)].append((bookmark_id, annotation_type))
            
        # Process each book's annotations
        for (book_title, author), bookmark_info in annotations_by_book.items():
            print(f"\nProcessing book: {book_title} by {author}")  # Debug log
            
            # Get the original annotation data from the database
            db_path = os.path.join(self.device_paths[self.device_dropdown.get()], ".kobo", "KoboReader.sqlite")
            
            # Prepare content for all annotations
            all_content = []
            export_success = False  # Track if any export was successful
            
            for bookmark_id, annotation_type in bookmark_info:
                print(f"\nProcessing bookmark: {bookmark_id}, type: {annotation_type}")  # Debug log
                
                # Get position information for this annotation
                position_info = self.get_annotation_position(db_path, bookmark_id)
                if not position_info:
                    print(f"Could not get position information for bookmark {bookmark_id}")
                    continue
                
                print(f"Position info: {position_info}")  # Debug log
                
                # Get the page image with the correct position
                # Strip /mnt/onboard prefix and ensure .epub extension
                epub_path = position_info['epub_path'].replace('/mnt/onboard/', '')
                # Remove any extensions after .epub
                if '.epub.' in epub_path:
                    epub_path = epub_path[:epub_path.find('.epub.') + 5]
                # Ensure it ends with .epub
                if not epub_path.endswith('.epub'):
                    epub_path = epub_path + '.epub'
                
                epub_path = os.path.join(self.device_paths[self.device_dropdown.get()], epub_path)
                print(f"Looking for EPUB at: {epub_path}")  # Debug log
                
                # Handle markup annotations
                if annotation_type == 'markup':
                    print("Processing markup annotation...")  # Debug log
                    markup_dir = os.path.join(self.device_paths[self.device_dropdown.get()], ".kobo", "markups")
                    svg_path = os.path.join(markup_dir, f"{bookmark_id}.svg")
                    jpg_path = os.path.join(markup_dir, f"{bookmark_id}.jpg")
                    
                    print(f"Looking for markup files:")  # Debug log
                    print(f"SVG: {svg_path} - Exists: {os.path.exists(svg_path)}")
                    print(f"JPG: {jpg_path} - Exists: {os.path.exists(jpg_path)}")
                    
                    markup_file = None
                    if os.path.exists(svg_path):
                        markup_file = svg_path
                    elif os.path.exists(jpg_path):
                        markup_file = jpg_path
                    
                    if markup_file and os.path.exists(epub_path):
                        print("Found markup file and EPUB, getting page image...")  # Debug log
                        # Get the page image
                        page_image, crop_y, total_height = self.get_page_image(epub_path, position_info['content_id'], position_info.get('ChapterProgress'))
                        if page_image:
                            print("Got page image, merging with markup...")  # Debug log
                            # Merge markup with page image
                            merged_image = self.merge_markup_with_page(markup_file, page_image, position_info.get('ChapterProgress'), crop_y, total_height)
                            if merged_image:
                                print("Successfully merged images, showing preview...")  # Debug log
                                # Show preview window
                                self.preview_combined_image(merged_image, bookmark_id)
                                export_success = True  # Mark that we had a successful export
                                continue
                            else:
                                print("Failed to merge images")  # Debug log
                        else:
                            print("Failed to get page image")  # Debug log
                    else:
                        print(f"Missing files - Markup: {markup_file}, EPUB: {epub_path}")  # Debug log
                
                # If we couldn't create a merged image or it's not a markup, fall back to text annotation
                if position_info['text']:
                    print("Falling back to text annotation...")  # Debug log
                    all_content.append((
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        f"```\n{position_info['text']}\n```"
                    ))
            
            # Create or update the note in Joplin for text annotations
            if all_content:
                print(f"\nCreating/updating note for text annotations...")  # Debug log
                note_title = f"{book_title} - {author}"
                notes = self.joplin.search_all(query=note_title, type_="note")
                
                # Only consider the note if it's in our configured notebook
                existing_note = None
                for note in notes:
                    if note.parent_id == self.config['notebook_id']:
                        existing_note = note
                        break
                
                try:
                    if existing_note:
                        print("Updating existing note...")  # Debug log
                        # Update existing note with all new content
                        existing_content = existing_note.body
                        for timestamp, content in all_content:
                            existing_content = self.insert_content_in_order(existing_content, content, timestamp)
                        
                        self.joplin.modify_note(
                            id_=existing_note.id,
                            body=existing_content
                        )
                    else:
                        print("Creating new note...")  # Debug log
                        # Create new note with all content
                        final_content = []
                        for timestamp, content in all_content:
                            final_content.append(content)
                        
                        self.joplin.add_note(
                            title=note_title,
                            body='\n\n---\n\n'.join(final_content),
                            parent_id=self.config['notebook_id']
                        )
                    export_success = True  # Mark that we had a successful export
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to export note: {str(e)}")
                    print(f"Error details: {str(e)}")
                    continue
            
            # Show success message only if we had any successful exports
            if export_success:
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
        settings_window.geometry("400x500")  # Made window taller to fit new settings
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
        
        # Font Settings Frame
        font_frame = ttk.LabelFrame(main_frame, text="Font Settings", padding="5")
        font_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Use Kobo Font Checkbox
        use_kobo_font_var = tk.BooleanVar(value=self.config.get('font', {}).get('use_kobo_font', False))
        use_kobo_font_check = ttk.Checkbutton(font_frame, text="Use Kobo Font", variable=use_kobo_font_var)
        use_kobo_font_check.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Font Family (Windows)
        ttk.Label(font_frame, text="Windows Font:").grid(row=1, column=0, sticky=tk.W, pady=5)
        font_family_var = tk.StringVar(value=self.config.get('font', {}).get('family', 'Arial'))
        font_family_entry = ttk.Entry(font_frame, textvariable=font_family_var, width=40)
        font_family_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Kobo Font Name
        ttk.Label(font_frame, text="Kobo Font Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        kobo_font_var = tk.StringVar(value=self.config.get('font', {}).get('kobo_font', ''))
        kobo_font_entry = ttk.Entry(font_frame, textvariable=kobo_font_var, width=40)
        kobo_font_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Font Size
        ttk.Label(font_frame, text="Font Size:").grid(row=3, column=0, sticky=tk.W, pady=5)
        font_size_var = tk.StringVar(value=str(self.config.get('font', {}).get('size', 16)))
        font_size_entry = ttk.Entry(font_frame, textvariable=font_size_var, width=40)
        font_size_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        def save_settings():
            """Save the settings and update the configuration."""
            try:
                # Validate font size
                font_size = int(font_size_var.get())
                if font_size < 8 or font_size > 72:
                    raise ValueError("Font size must be between 8 and 72")
                
                self.config['joplin_api_token'] = api_token_var.get()
                self.config['notebook_id'] = notebook_id_var.get()
                self.config['web_clipper'] = {
                    'url': web_clipper_url_var.get(),
                    'port': int(web_clipper_port_var.get())
                }
                self.config['font'] = {
                    'use_kobo_font': use_kobo_font_var.get(),
                    'family': font_family_var.get(),
                    'kobo_font': kobo_font_var.get(),
                    'size': font_size
                }
                
                # Save to file
                config_path = os.path.join(os.path.dirname(__file__), 'config.json')
                with open(config_path, 'w') as f:
                    json.dump(self.config, f, indent=4)
                
                # Reinitialize Joplin API with new token
                self.joplin = ClientApi(token=self.config['joplin_api_token'])
                
                settings_window.destroy()
                messagebox.showinfo("Success", "Settings saved successfully!")
            except ValueError as e:
                messagebox.showerror("Error", str(e))
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
        
        ttk.Button(button_frame, text="Save", command=save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=settings_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        font_frame.columnconfigure(1, weight=1)

    def load_chapter_formats(self):
        """Load chapter formats configuration from JSON file."""
        try:
            # Get the directory where the script or executable is located
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = os.path.dirname(sys.executable)
            else:
                # Running as script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            config_path = os.path.join(base_path, 'chapter_formats.json')
            
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    return json.load(f)
            else:
                print(f"Chapter formats configuration not found at: {config_path}")
                return None
        except Exception as e:
            print(f"Error loading chapter formats: {str(e)}")
            return None

    def update_export_button_text(self, event=None):
        """Update the export button text based on selected annotation type."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.export_button.configure(text="Export to Joplin", state="normal")
            return
            
        # Check if we have mixed annotation types
        has_markup = False
        has_other = False
        
        for item in selected_items:
            values = self.tree.item(item)['values']
            if values:
                if values[5] == 'markup':  # Type is in the 6th column
                    has_markup = True
                else:
                    has_other = True
                    
                # If we have both types, disable the button
                if has_markup and has_other:
                    self.export_button.configure(text="Export to Joplin", state="disabled")
                    return
        
        # If we only have markup, show Preview Image
        if has_markup:
            self.export_button.configure(text="Preview Image", state="normal")
        # If we only have other types, show Export to Joplin
        else:
            self.export_button.configure(text="Export to Joplin", state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = KoboToJoplinApp(root)
    root.mainloop()