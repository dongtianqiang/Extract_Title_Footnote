import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import os
import sys
import time
import traceback
from typing import Callable, Optional

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Add the current directory to Python path to import modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

class CancellationToken:
    """Token to signal cancellation of long-running operations"""
    def __init__(self):
        self._cancelled = False
        self._callbacks = []
    
    def cancel(self):
        """Signal that operation should be cancelled"""
        self._cancelled = True
        # Notify all callbacks
        for callback in self._callbacks:
            try:
                callback()
            except Exception:
                pass
    
    def is_cancelled(self) -> bool:
        """Check if cancellation has been requested"""
        return self._cancelled
    
    def register_callback(self, callback: Callable):
        """Register a callback to be called when cancellation occurs"""
        self._callbacks.append(callback)

class ExtractTitleFootnoteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extract Title Footnote")
        self.root.geometry("800x750")
        self.root.minsize(600, 500)
        
        # Configure dark theme colors
        self.bg_color = "#1e1e2e"
        self.fg_color = "#d4d4d4"
        self.accent_color = "#007acc"
        self.disabled_color = "#3c3c4a"
        self.cancel_color = "#ff4444"  # Red color for cancel button
        
        self.root.configure(bg=self.bg_color)
        
        # Load and set window icon
        self._load_icon()
        
        # Cancellation support
        self.cancellation_token: Optional[CancellationToken] = None
        
        # README window reference
        self.readme_window = None
        
        # Create tabbed interface
        self.create_tab_interface()
        
        # Status variables
        self.processing = False
        self.process_thread = None
        self.monitor_thread = None
        
        # RTF processing state
        self.rtf_processing = False
        self.rtf_process_thread = None
        self.rtf_monitor_thread = None
        self.rtf_cancellation_token: Optional[CancellationToken] = None
    
    def _get_icon_path(self):
        """Get icon file path - search in multiple locations"""
        # Search for icon files in priority order
        icon_filenames = ['favicon32.ico']
        
        for icon_name in icon_filenames:
            # Priority 1: Current directory (development)
            current_dir = os.path.join(os.getcwd(), icon_name)
            if os.path.exists(current_dir):
                return current_dir
            
            # Priority 2: Executable directory (packaged)
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.join(os.path.dirname(sys.executable), icon_name)
                if os.path.exists(exe_dir):
                    return exe_dir
            
            # Priority 3: PyInstaller temp directory (_MEIPASS)
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                pyinstaller_dir = os.path.join(sys._MEIPASS, icon_name)
                if os.path.exists(pyinstaller_dir):
                    return pyinstaller_dir
        
        return None
    
    def _load_icon(self):
        """Load and apply window icons using optimized method"""
        icon_path = self._get_icon_path()
        if icon_path:
            try:
                # Set default icon for taskbar and Alt+Tab
                self.root.iconbitmap(default=icon_path)
                # Set window icon for title bar
                self.root.iconbitmap(icon_path)
            except Exception as e:
                print(f"Warning: Could not set window icon: {e}")
    
    def create_tab_interface(self):
        """Create the tabbed interface with Shell Processor and RTF Processor tabs"""
        # Main container for tabs
        self.main_container = tk.Frame(self.root, bg=self.bg_color)
        self.main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)
        
        # Top navigation bar (tab bar)
        self.tab_bar = tk.Frame(self.main_container, bg=self.bg_color, height=30)
        self.tab_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.tab_bar.grid_columnconfigure(0, weight=0)  # README button - no expansion
        self.tab_bar.grid_columnconfigure(1, weight=1)  # Shell Processor tab - expand
        self.tab_bar.grid_columnconfigure(2, weight=1)  # RTF Processor tab - no expansion
        
        # README button (top-left, before Shell Processor tab)
        self.readme_btn = tk.Button(
            self.tab_bar,
            text="ℹ README",
            font=("Segoe UI", 9, "bold"),
            command=self.show_readme,
            bg=self.disabled_color,
            fg=self.fg_color,
            relief="flat",
            pady=4,
            padx=8,
            cursor="hand2"
        )
        self.readme_btn.grid(row=0, column=0, sticky=tk.W, padx=(5, 10), pady=3)
        
        # Shell Processor tab
        self.shell_tab_btn = tk.Button(
            self.tab_bar, 
            text="Shell Processor",
            font=("Segoe UI", 10, "bold"),
            command=lambda: self.show_tab("shell"),
            bg=self.accent_color,
            fg="white",
            relief="flat",
            pady=6
        )
        self.shell_tab_btn.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=0, pady=0)
        
        # RTF Processor tab
        self.rtf_tab_btn = tk.Button(
            self.tab_bar,
            text="RTF Processor",
            font=("Segoe UI", 10, "bold"),
            command=lambda: self.show_tab("rtf"),
            bg=self.disabled_color,
            fg=self.fg_color,
            relief="flat",
            pady=6
        )
        self.rtf_tab_btn.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=0, pady=0)
        
        # Create content frames for each tab
        self.shell_content_frame = tk.Frame(self.main_container, bg=self.bg_color)
        self.rtf_content_frame = tk.Frame(self.main_container, bg=self.bg_color)
        
        # Build Shell Processor tab content
        self.build_shell_processor_tab()
        
        # Build RTF Processor tab content
        self.build_rtf_processor_tab()
        
        # Show Shell Processor tab by default
        self.show_tab("shell")

    def _get_readme_path(self):
        """Get README file path - works in both development and packaged modes"""
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # Running in frozen mode - use PyInstaller temp directory
            return os.path.join(sys._MEIPASS, 'README_V1.0.md')
        else:
            # Running in development - use current file directory
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'README_V1.0.md')

    def show_readme(self):
        """Show README content in a popup window"""
        # If window already exists, bring it to front
        if self.readme_window is not None:
            try:
                self.readme_window.focus_force()
                self.readme_window.lift()
                self.readme_window.attributes('-topmost', True)
                self.readme_window.attributes('-topmost', False)
                return
            except:
                # Window was closed unexpectedly, reset reference
                self.readme_window = None
        
        # Create new README window
        readme_win = tk.Toplevel(self.root)
        self.readme_window = readme_win
        
        # Configure window
        readme_win.title("README - 标题脚注提取工具")
        readme_win.geometry("700x600")
        readme_win.configure(bg=self.bg_color)
        
        # Set window icon
        icon_path = self._get_icon_path()
        if icon_path:
            try:
                readme_win.iconbitmap(icon_path)
            except Exception:
                pass
        
        # Main container
        main_frame = tk.Frame(readme_win, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Text area with scrollbar
        text_widget = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#2d2d3b",
            fg=self.fg_color,
            insertbackground=self.fg_color,
            relief="flat",
            borderwidth=0
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Load README content
        try:
            readme_path = self._get_readme_path()
            if os.path.exists(readme_path):
                with open(readme_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                text_widget.insert(tk.END, content)
            else:
                text_widget.insert(tk.END, "README file not found: README_V1.0.md")
        except Exception as e:
            text_widget.insert(tk.END, f"Error loading README: {str(e)}")
        
        # Make text widget read-only
        text_widget.config(state='disabled')
        
        # Handle window close
        def on_close():
            self.readme_window = None
            readme_win.destroy()
        
        readme_win.protocol("WM_DELETE_WINDOW", on_close)
        
        # Bring window to front
        readme_win.after(100, lambda: self._focus_readme_window(readme_win))
    
    def _focus_readme_window(self, win):
        """Helper to focus README window"""
        try:
            win.focus_force()
            win.lift()
            win.attributes('-topmost', True)
            win.attributes('-topmost', False)
        except:
            pass
    
    def close_readme_window(self):
        """Close the README window"""
        if self.readme_window:
            try:
                self.readme_window.destroy()
            except:
                pass
            finally:
                self.readme_window = None
    
    def show_tab(self, tab_name):
        """Switch between tabs"""
        # Hide all content frames
        self.shell_content_frame.grid_forget()
        self.rtf_content_frame.grid_forget()
        
        # Reset tab button styles
        self.shell_tab_btn.config(bg=self.disabled_color, fg=self.fg_color)
        self.rtf_tab_btn.config(bg=self.disabled_color, fg=self.fg_color)
        
        # Show selected tab and highlight button
        if tab_name == "shell":
            self.shell_content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            self.shell_tab_btn.config(bg=self.accent_color, fg="white")
        elif tab_name == "rtf":
            self.rtf_content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            self.rtf_tab_btn.config(bg=self.accent_color, fg="white")
    
    def _create_styled_entry(self, parent, textvariable=None, **kwargs):
        """Create an entry widget with consistent dark theme styling"""
        entry = tk.Entry(
            parent,
            textvariable=textvariable,
            font=("Segoe UI", 10),
            bg="#2d2d3b",
            fg=self.fg_color,
            insertbackground=self.fg_color,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#444455",
            **kwargs
        )
        return entry
    
    def _create_styled_button(self, parent, text, command, bg=None, fg="white", **kwargs):
        """Create a button with consistent styling"""
        if bg is None:
            bg = self.accent_color
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 10),
            relief="flat",
            height=1,
            bg=bg,
            fg=fg,
            **kwargs
        )
        return btn
    
    def _add_log_entry(self, text_widget, label, message):
        """Generic method to add log entry to any text widget"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        text_widget.insert(tk.END, log_entry)
        text_widget.see(tk.END)
        
        # Update logs count
        lines = int(text_widget.index('end-1c').split('.')[0])
        label.config(text=f"Logs ({lines} entries)")
    
    def _copy_logs_to_clipboard(self, text_widget, log_func):
        """Generic method to copy logs to clipboard"""
        try:
            # If there's a selection, copy that
            if text_widget.tag_ranges("sel"):
                selected_text = text_widget.get("sel.first", "sel.last")
            else:
                # Copy all logs
                selected_text = text_widget.get(1.0, tk.END).rstrip()
            
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            log_func("Logs copied to clipboard")
        except Exception as e:
            log_func(f"Error copying logs: {str(e)}")
    
    def build_shell_processor_tab(self):
        
        # Main frame for Shell Processor
        main_frame = tk.Frame(self.shell_content_frame, bg=self.bg_color, padx=10, pady=10)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.shell_content_frame.grid_columnconfigure(0, weight=1)
        self.shell_content_frame.grid_rowconfigure(2, weight=1)
        
        # Configure grid weights
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)
        
        # Input frame
        input_frame = tk.Frame(main_frame, bg=self.bg_color)
        input_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        input_frame.grid_columnconfigure(0, weight=6)  # Max Footnote Columns: 60%
        input_frame.grid_columnconfigure(1, weight=4)  # Project ID: 40%
        input_frame.grid_columnconfigure(2, weight=0)  # Browse button: fixed width
        
        # Make input_frame expand to fill main_frame
        main_frame.grid_columnconfigure(0, weight=1)
        
        # DOCX file path input
        docx_label = tk.Label(input_frame, text="Shell File Path", 
                              bg=self.bg_color, 
                              fg=self.fg_color,
                              font=("Segoe UI", 10, "bold"))
        docx_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.docx_path_var = tk.StringVar()
        self.docx_entry = self._create_styled_entry(input_frame, textvariable=self.docx_path_var)
        self.docx_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        browse_btn = tk.Button(input_frame, text="Browse...", 
                               command=self.browse_file,
                               bg=self.accent_color,
                               fg="white",
                               relief="flat",
                               font=("Segoe UI", 10),
                               height=1)
        browse_btn.grid(row=1, column=2, padx=(10, 0), pady=(0, 10), sticky=tk.W)
        
        docx_hint = tk.Label(input_frame, text="Enter the full path to the shell file to process", 
                             bg=self.bg_color, 
                             fg="#9e9e9e",
                             font=("Segoe UI", 8))
        docx_hint.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(0, 15))
        
        # Max footnote columns input
        columns_label = tk.Label(input_frame, text="Max Footnote Columns", 
                                 bg=self.bg_color, 
                                 fg=self.fg_color,
                                 font=("Segoe UI", 10, "bold"))
        columns_label.grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        
        self.max_columns_var = tk.StringVar(value="7")
        self.columns_entry = self._create_styled_entry(input_frame, textvariable=self.max_columns_var)
        self.columns_entry.grid(row=4, column=0, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        
        columns_hint = tk.Label(input_frame, text="Range: 1 - 10, default: 7", 
                                bg=self.bg_color, 
                                fg="#9e9e9e",
                                font=("Segoe UI", 8))
        columns_hint.grid(row=5, column=0, sticky=tk.W, pady=(0, 15))
        
        # Project ID input
        project_label = tk.Label(input_frame, text="Project ID", 
                                 bg=self.bg_color, 
                                 fg=self.fg_color,
                                 font=("Segoe UI", 10, "bold"))
        project_label.grid(row=3, column=1, sticky=tk.W, pady=(0, 5))
        
        self.project_id_var = tk.StringVar()
        self.project_id_entry = self._create_styled_entry(input_frame, textvariable=self.project_id_var)
        self.project_id_entry.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        project_hint = tk.Label(input_frame, text="Used as the output filename prefix", 
                                bg=self.bg_color, 
                                fg="#9e9e9e",
                                font=("Segoe UI", 8))
        project_hint.grid(row=5, column=1, columnspan=2, sticky=tk.W, pady=(0, 15))
        
        # Custom footnote keywords section
        keywords_frame = tk.Frame(input_frame, bg=self.bg_color)
        keywords_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        keywords_frame.grid_columnconfigure(1, weight=1)
        
        keywords_label = tk.Label(keywords_frame, text="Custom Footnote Keywords", 
                                 bg=self.bg_color, 
                                 fg=self.fg_color,
                                 font=("Segoe UI", 10, "bold"))
        keywords_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        # Frame to hold keyword entries
        self.keywords_container = tk.Frame(keywords_frame, bg=self.bg_color)
        self.keywords_container.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E))
        # Make column 0 expandable like the Shell File Path input
        self.keywords_container.grid_columnconfigure(0, weight=1)  # Allow entry to expand
        
        # Store keyword entries
        self.keyword_vars = []
        self.keyword_entries = []
        self.remove_buttons = []  # Store remove buttons
        
        # Add keyword hint
        keywords_hint = tk.Label(keywords_frame, text="Add custom keywords (default 'programming'/'programmer' always active, case-insensitive)", 
                                bg=self.bg_color, 
                                fg="#9e9e9e",
                                font=("Segoe UI", 8))
        keywords_hint.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # Add Keyword button
        self.add_keyword_btn = tk.Button(keywords_frame, text="+ Add Keyword",
                                        command=self.add_keyword_entry,
                                        bg=self.disabled_color,
                                        fg=self.fg_color,
                                        relief="flat",
                                        font=("Segoe UI", 10),
                                        height=1,
                                        anchor="w")
        self.add_keyword_btn.grid(row=3, column=0, sticky=tk.W, pady=(10, 0))
        
        # Add default keyword entry (after all UI elements are created)
        self.add_keyword_entry()
        
        # Buttons frame
        buttons_frame = tk.Frame(main_frame, bg=self.bg_color)
        buttons_frame.grid(row=2, column=0, columnspan=2, pady=(20, 0))
        
        self.cancel_btn = tk.Button(buttons_frame, text="Cancel", 
                                    command=self.cancel_processing,
                                    bg=self.disabled_color,
                                    fg=self.fg_color,
                                    relief="flat",
                                    font=("Segoe UI", 10),
                                    height=1,
                                    state="disabled")
        self.cancel_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.confirm_btn = tk.Button(buttons_frame, text="Confirm", 
                                     command=self.start_processing,
                                     bg=self.disabled_color,
                                     fg=self.fg_color,
                                     relief="flat",
                                     font=("Segoe UI", 10),
                                     height=1)
        self.confirm_btn.grid(row=0, column=1)
        
        # Enable/disable confirm button based on inputs
        self.docx_path_var.trace_add("write", self.update_confirm_button_state)
        self.max_columns_var.trace_add("write", self.update_confirm_button_state)
        self.project_id_var.trace_add("write", self.update_confirm_button_state)
        
        # Trace keyword variables
        self.setup_keyword_traces()
        
        # Initialize button states
        self.update_keyword_button_states()
        
        # Logs frame
        logs_frame = tk.Frame(main_frame, bg=self.bg_color)
        logs_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        # Configure grid weights for proper expansion
        logs_frame.grid_rowconfigure(0, weight=0)  # Log label row - no expansion
        logs_frame.grid_rowconfigure(1, weight=1)  # Log text area row - expandable
        logs_frame.grid_columnconfigure(0, weight=1)  # Column 0 - expandable
        
        # Store logs label as instance variable
        self.logs_label = tk.Label(logs_frame, text="Logs (0 entries)", 
                                  bg=self.bg_color, 
                                  fg=self.fg_color,
                                  font=("Segoe UI", 10, "bold"))
        self.logs_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # Log text area with scrollbar
        self.logs_text = scrolledtext.ScrolledText(logs_frame, 
                                                  wrap=tk.WORD,
                                                  font=("Consolas", 9),
                                                  bg="#1e1e2e",
                                                  fg="#d4d4d4",
                                                  insertbackground="#d4d4d4",
                                                  relief="flat",
                                                  borderwidth=0)
        # Configure the text widget to expand with the container
        self.logs_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Set minimum size for the logs text area to ensure visibility
        self.logs_text.configure(height=15)  # Minimum 15 lines visible
        
        # Add copy functionality to logs
        self.logs_text.bind("<Control-c>", self.copy_logs)
        
    def build_rtf_processor_tab(self):
        """Build the RTF Processor tab content"""
        # Main frame for RTF Processor
        main_frame = tk.Frame(self.rtf_content_frame, bg=self.bg_color, padx=10, pady=10)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.rtf_content_frame.grid_columnconfigure(0, weight=1)
        self.rtf_content_frame.grid_rowconfigure(1, weight=1)
        
        # Configure grid weights
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Input frame
        input_frame = tk.Frame(main_frame, bg=self.bg_color)
        input_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        input_frame.grid_columnconfigure(1, weight=1)
        
        # LOT file path input
        lot_label = tk.Label(input_frame, text="LOT File Path", 
                              bg=self.bg_color, 
                              fg=self.fg_color,
                              font=("Segoe UI", 10, "bold"))
        lot_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.lot_path_var = tk.StringVar()
        self.lot_entry = self._create_styled_entry(input_frame, textvariable=self.lot_path_var)
        self.lot_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        browse_lot_btn = tk.Button(input_frame, text="Browse...", 
                               command=self.browse_lot_file,
                               bg=self.accent_color,
                               fg="white",
                               relief="flat",
                               font=("Segoe UI", 10),
                               height=1)
        browse_lot_btn.grid(row=1, column=2, padx=(10, 0), pady=(0, 10), sticky=tk.E)
        
        lot_hint = tk.Label(input_frame, text="Enter the full path to the LOT file to process (RTF files must be in the same folder)", 
                             bg=self.bg_color, 
                             fg="#9e9e9e",
                             font=("Segoe UI", 8))
        lot_hint.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(0, 15))
        
        # Project ID input
        project_label = tk.Label(input_frame, text="Project ID", 
                                 bg=self.bg_color, 
                                 fg=self.fg_color,
                                 font=("Segoe UI", 10, "bold"))
        project_label.grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        
        self.rtf_project_id_var = tk.StringVar()
        self.rtf_project_id_entry = self._create_styled_entry(input_frame, textvariable=self.rtf_project_id_var)
        self.rtf_project_id_entry.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        project_hint = tk.Label(input_frame, text="Used as the output filename prefix", 
                                bg=self.bg_color, 
                                fg="#9e9e9e",
                                font=("Segoe UI", 8))
        project_hint.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(0, 15))
        
        # Buttons frame
        buttons_frame = tk.Frame(main_frame, bg=self.bg_color)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(20, 0))
        
        self.rtf_cancel_btn = tk.Button(buttons_frame, text="Cancel", 
                                    command=self.cancel_rtf_processing,
                                    bg=self.disabled_color,
                                    fg=self.fg_color,
                                    relief="flat",
                                    font=("Segoe UI", 10),
                                    height=1,
                                    state="disabled")
        self.rtf_cancel_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.rtf_confirm_btn = tk.Button(buttons_frame, text="Confirm", 
                                     command=self.start_rtf_processing,
                                     bg=self.disabled_color,
                                     fg=self.fg_color,
                                     relief="flat",
                                     font=("Segoe UI", 10),
                                     height=1)
        self.rtf_confirm_btn.grid(row=0, column=1)
        
        # Enable/disable confirm button based on inputs
        self.lot_path_var.trace_add("write", self.update_rtf_confirm_button_state)
        self.rtf_project_id_var.trace_add("write", self.update_rtf_confirm_button_state)
        
        # Initialize button state
        self.update_rtf_confirm_button_state()
        
        # Logs frame
        logs_frame = tk.Frame(main_frame, bg=self.bg_color)
        logs_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        logs_frame.grid_rowconfigure(0, weight=0)
        logs_frame.grid_rowconfigure(1, weight=1)
        logs_frame.grid_columnconfigure(0, weight=1)
        
        self.rtf_logs_label = tk.Label(logs_frame, text="Logs (0 entries)", 
                                  bg=self.bg_color, 
                                  fg=self.fg_color,
                                  font=("Segoe UI", 10, "bold"))
        self.rtf_logs_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.rtf_logs_text = scrolledtext.ScrolledText(logs_frame, 
                                                  wrap=tk.WORD,
                                                  font=("Consolas", 9),
                                                  bg="#1e1e2e",
                                                  fg="#d4d4d4",
                                                  insertbackground="#d4d4d4",
                                                  relief="flat",
                                                  borderwidth=0)
        self.rtf_logs_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.rtf_logs_text.configure(height=15)
        self.rtf_logs_text.bind("<Control-c>", self.copy_rtf_logs)
        
        # RTF processing state is initialized in __init__
    
    def add_keyword_entry(self):
        """Add a new keyword entry field"""
        row = len(self.keyword_vars)
        
        # Keyword variable
        keyword_var = tk.StringVar()
        self.keyword_vars.append(keyword_var)
        
        # Entry widget (adaptive width like Shell File Path)
        keyword_entry = self._create_styled_entry(self.keywords_container, textvariable=keyword_var)
        keyword_entry.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.keyword_entries.append(keyword_entry)
        
        # Remove button (on the right side of the entry)
        remove_btn = tk.Button(self.keywords_container, text="−",
                              command=lambda r=row: self.remove_keyword_entry(r),
                              bg=self.disabled_color,
                              fg=self.fg_color,
                              relief="flat",
                              font=("Segoe UI", 10),
                              width=3,
                              height=1,
                              state="disabled" if row == 0 else "normal")
        remove_btn.grid(row=row, column=1, padx=(5, 0), pady=(0, 5), sticky=tk.E)
        self.remove_buttons.append(remove_btn)
        
        # Bind trace to update button states
        keyword_var.trace_add("write", self.update_keyword_button_states)
        
        # Focus on the new entry
        keyword_entry.focus()
        
        # Update button states
        self.update_keyword_button_states()
        
    def remove_keyword_entry(self, index):
        """Remove a keyword entry field"""
        if len(self.keyword_vars) <= 1:
            return  # Keep at least one entry
            
        # Remove widgets
        self.keyword_entries[index].destroy()
        self.remove_buttons[index].destroy()
        
        # Remove from lists
        del self.keyword_vars[index]
        del self.keyword_entries[index]
        del self.remove_buttons[index]
        
        # Re-grid remaining entries and buttons
        for i, entry in enumerate(self.keyword_entries):
            entry.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
            
            # Update remove button position and command
            self.remove_buttons[i].config(command=lambda r=i: self.remove_keyword_entry(r))
            self.remove_buttons[i].grid(row=i, column=1, padx=(5, 0), pady=(0, 5), sticky=tk.E)
        
        # Update button states
        self.update_keyword_button_states()
            
    def get_custom_keywords(self):
        """Get list of custom keywords from entries"""
        keywords = []
        for var in self.keyword_vars:
            keyword = var.get().strip()
            if keyword:
                keywords.append(keyword)
        return keywords
        
    def setup_keyword_traces(self):
        """Setup traces for all keyword variables"""
        def on_keyword_change(*args):
            self.update_confirm_button_state()
            self.update_keyword_button_states()
        
        for var in self.keyword_vars:
            var.trace_add("write", on_keyword_change)
    
    def update_keyword_button_states(self, *args):
        """Update the state of add/remove buttons based on keyword entries"""
        # "+ Add Keyword" button is always enabled (no restriction on empty fields)
        if hasattr(self, 'add_keyword_btn') and self.add_keyword_btn:
            self.add_keyword_btn.config(bg=self.accent_color, fg="white", state="normal")
        
        # Update remove button states
        for i, btn in enumerate(self.remove_buttons):
            if len(self.keyword_vars) > 1:  # Allow removal if more than one entry
                btn.config(state="normal", bg="#ff4444", fg="white")  # Red color for remove
            else:
                btn.config(state="disabled", bg=self.disabled_color, fg=self.fg_color)
        
        # Also update confirm button state if it exists
        if hasattr(self, 'confirm_btn') and self.confirm_btn:
            self.update_confirm_button_state()
        
    def browse_file(self):
        """Open file dialog to select DOCX file"""
        file_path = filedialog.askopenfilename(
            title="Select Shell File",
            filetypes=[("DOCX files", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            self.docx_path_var.set(file_path)
    
    def browse_lot_file(self):
        """Open file dialog to select LOT file"""
        file_path = filedialog.askopenfilename(
            title="Select LOT File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.lot_path_var.set(file_path)
    
    def update_confirm_button_state(self, *args):
        """Update the state of the confirm button based on input validity"""
        # Skip if confirm button hasn't been created yet
        if not hasattr(self, 'confirm_btn') or not self.confirm_btn:
            return
            
        docx_path = self.docx_path_var.get().strip()
        max_columns = self.max_columns_var.get().strip()
        project_id = self.project_id_var.get().strip()
        custom_keywords = self.get_custom_keywords()
        
        # Check if all required fields are filled
        if docx_path and max_columns and project_id and not self.processing and not self.rtf_processing:
            # Check if max_columns is a valid number between 1-10
            try:
                columns_val = int(max_columns)
                if 1 <= columns_val <= 10:
                    # Even if no custom keywords are provided, we still allow processing
                    # (custom keywords are optional)
                    self.confirm_btn.config(bg=self.accent_color, fg="white")
                    self.confirm_btn.config(state="normal")
                    return
            except ValueError:
                pass
        
        # Disable confirm button if any field is invalid
        self.confirm_btn.config(bg=self.disabled_color, fg=self.fg_color)
        self.confirm_btn.config(state="disabled")
    
    def start_processing(self):
        """Start the document processing in a separate thread"""
        if self.processing or self.rtf_processing:
            # Don't start if either processor is already running
            if self.rtf_processing:
                self.add_log("Error: RTF Processor is currently running. Please wait for it to complete.")
            return
            
        docx_path = self.docx_path_var.get().strip()
        max_columns = int(self.max_columns_var.get().strip())
        project_id = self.project_id_var.get().strip()
        custom_keywords = self.get_custom_keywords()
        
        # Validate inputs
        if not os.path.isfile(docx_path):
            self.add_log(f"Error: File '{docx_path}' does not exist.")
            return
            
        if not 1 <= max_columns <= 10:
            self.add_log("Error: Max footnote columns must be between 1 and 10.")
            return
            
        if not project_id:
            self.add_log("Error: Project ID is required.")
            return
        
        # Reset logs
        self.logs_text.delete(1.0, tk.END)
        self.add_log("Starting processing...")
        
        # Create cancellation token
        self.cancellation_token = CancellationToken()
        
        # Start processing in a separate thread
        self.processing = True
        self.process_thread = threading.Thread(
            target=self.run_processing,
            args=(docx_path, max_columns, project_id, custom_keywords, self.cancellation_token),
            daemon=True
        )
        self.process_thread.start()
        
        # Update button states for both processors
        self.update_all_processor_buttons()
        
        # Ensure cancel button shows as active (red) and is clickable
        self.cancel_btn.config(state="normal", bg="#ff4444", fg="white")
        
        # Start monitoring thread
        self.monitor_thread = threading.Thread(target=self.monitor_processing, daemon=True)
        self.monitor_thread.start()
    
    def run_processing(self, docx_path, max_columns, project_id, custom_keywords, cancellation_token):
        """Run the actual processing in a background thread"""
        try:
            # Import the main processing function dynamically to avoid circular imports
            module = __import__('main', fromlist=['process_document'])
            process_document = getattr(module, 'process_document')
            
            # Get workspace directory
            workspace = os.path.dirname(docx_path)
            
            # Determine output file names based on project ID
            if project_id:
                output_file = os.path.join(workspace, f"{project_id}_TF_Contents.xlsx")
                auxiliary_file = os.path.join(workspace, f"{project_id}_LOT.xlsx")
            else:
                output_file = os.path.join(workspace, "TF_Contents.xlsx")
                auxiliary_file = os.path.join(workspace, "LOT.xlsx")
            
            # Add log about starting
            self.add_log(f"Processing document: {docx_path}")
            self.add_log(f"Max footnote columns: {max_columns}")
            self.add_log(f"Project ID: {project_id}")
            self.add_log(f"Custom keywords: {custom_keywords if custom_keywords else 'None (using defaults)'}")
            self.add_log(f"TF_Contents file: {output_file}")
            self.add_log(f"LOT file: {auxiliary_file}")
            
            # Redirect stdout to capture print outputs
            old_stdout = sys.stdout
            sys.stdout = self.LogRedirector(self.add_log)
            
            try:
                # Pass cancellation token and custom keywords to processing function
                result = process_document(docx_path, output_file, max_columns, workspace, project_id, custom_keywords, cancellation_token)
            finally:
                sys.stdout = old_stdout
            
            # Check if cancelled
            if cancellation_token.is_cancelled():
                self.add_log("Processing was cancelled by user.")
                return
            
            # Add results to log
            self.add_log(f"Processing completed successfully")
            self.add_log(f"Process returned: {result}")
            
        except Exception as e:
            if not cancellation_token.is_cancelled():
                self.add_log(f"Error during processing: {str(e)}")
                self.add_log(traceback.format_exc())
        finally:
            self.processing = False
    
    def monitor_processing(self):
        """Monitor the processing thread and update UI"""
        while self.processing and self.process_thread and self.process_thread.is_alive():
            time.sleep(0.1)
            # Check for cancellation
            if self.cancellation_token and self.cancellation_token.is_cancelled():
                break
        
        # Processing finished or cancelled, update UI
        self.root.after(100, self.on_processing_complete)
    
    def on_processing_complete(self):
        """Called when processing is complete"""
        # Reset button states using the validation logic
        self.update_all_processor_buttons()
        self.cancel_btn.config(state="disabled", bg=self.disabled_color, fg=self.fg_color)
        if self.processing:  # Only show completion message if not cancelled
            self.add_log("Processing finished.")
        self.processing = False
        self.cancellation_token = None
    
    def cancel_processing(self):
        """Cancel the ongoing processing"""
        if self.processing and self.cancellation_token:
            self.add_log("Cancelling processing...")
            self.cancellation_token.cancel()
            # Give some time for graceful shutdown
            if self.process_thread and self.process_thread.is_alive():
                self.process_thread.join(timeout=2)
            # Reset button states
            self.update_all_processor_buttons()
            self.cancel_btn.config(state="disabled", bg=self.disabled_color, fg=self.fg_color)
            self.add_log("Processing cancelled.")
            self.processing = False
            self.cancellation_token = None
    
    def add_log(self, message):
        """Add a log message to the logs text area"""
        # Update the text widget from the main thread
        self.root.after(0, lambda: self._add_log_entry(self.logs_text, self.logs_label, message))
    
    def copy_logs(self, event=None):
        """Copy selected text or all logs to clipboard"""
        self._copy_logs_to_clipboard(self.logs_text, self.add_log)
    
    def update_rtf_confirm_button_state(self, *args):
        """Update the state of the RTF confirm button based on input validity"""
        # Skip if confirm button hasn't been created yet
        if not hasattr(self, 'rtf_confirm_btn') or not self.rtf_confirm_btn:
            return
            
        lot_path = self.lot_path_var.get().strip()
        project_id = self.rtf_project_id_var.get().strip()
        
        # Check if all required fields are filled
        if lot_path and project_id and not self.processing and not self.rtf_processing:
            self.rtf_confirm_btn.config(bg=self.accent_color, fg="white")
            self.rtf_confirm_btn.config(state="normal")
            return
        
        # Disable confirm button if any field is invalid
        self.rtf_confirm_btn.config(bg=self.disabled_color, fg=self.fg_color)
        self.rtf_confirm_btn.config(state="disabled")

    def update_all_processor_buttons(self):
        """Update button states for both processors based on processing status and input validity"""
        # Update Shell Processor buttons
        if self.processing or self.rtf_processing:
            # Shell processor is running - disable its confirm button
            if hasattr(self, 'confirm_btn') and self.confirm_btn:
                self.confirm_btn.config(state="disabled", bg=self.disabled_color, fg="#666666", cursor="arrow")
            if hasattr(self, 'rtf_confirm_btn') and self.rtf_confirm_btn:
                    self.rtf_confirm_btn.config(state="disabled", bg=self.disabled_color, fg="#666666", cursor="arrow")
        else:
            # Shell processor is not running - update based on input validation
            self.update_confirm_button_state()
            self.update_rtf_confirm_button_state()
        
        # Update RTF Processor buttons
        if self.rtf_processing:
            # RTF processor is running - disable its confirm button
            if hasattr(self, 'rtf_confirm_btn') and self.rtf_confirm_btn:
                self.rtf_confirm_btn.config(state="disabled", bg=self.disabled_color, fg="#666666", cursor="arrow")
        else:
            # RTF processor is not running - update based on input validation
            self.update_rtf_confirm_button_state()
        
        # If one processor is running, disable the other processor's confirm button with clear visual feedback
        if self.processing:
            # Shell processor is running - disable RTF confirm button
            if hasattr(self, 'rtf_confirm_btn') and self.rtf_confirm_btn:
                self.rtf_confirm_btn.config(state="disabled", bg=self.disabled_color, fg="#666666", cursor="arrow")
        
        if self.rtf_processing:
            # RTF processor is running - disable Shell confirm button
            if hasattr(self, 'confirm_btn') and self.confirm_btn:
                self.confirm_btn.config(state="disabled", bg=self.disabled_color, fg="#666666", cursor="arrow")
    
    def start_rtf_processing(self):
        """Start the RTF processing in a separate thread"""
        if self.rtf_processing or self.processing:
            # Don't start if either processor is already running
            if self.processing:
                self.add_rtf_log("Error: Shell Processor is currently running. Please wait for it to complete.")
            return
            
        lot_path = self.lot_path_var.get().strip()
        project_id = self.rtf_project_id_var.get().strip()
        
        # Validate inputs
        if not os.path.isfile(lot_path):
            self.add_rtf_log(f"Error: File '{lot_path}' does not exist.")
            return
            
        if not project_id:
            self.add_rtf_log("Error: Project ID is required.")
            return
        
        # Reset logs
        self.rtf_logs_text.delete(1.0, tk.END)
        self.add_rtf_log("Starting RTF processing...")
        
        # Create cancellation token
        self.rtf_cancellation_token = CancellationToken()
        
        # Start processing in a separate thread
        self.rtf_processing = True
        self.rtf_process_thread = threading.Thread(
            target=self.run_rtf_processing,
            args=(lot_path, project_id, self.rtf_cancellation_token),
            daemon=True
        )
        self.rtf_process_thread.start()
        
        # Update button states for both processors
        self.update_all_processor_buttons()
        
        # Ensure cancel button shows as active (red) and is clickable
        self.rtf_cancel_btn.config(state="normal", bg="#ff4444", fg="white")
        
        # Start monitoring thread
        self.rtf_monitor_thread = threading.Thread(target=self.monitor_rtf_processing, daemon=True)
        self.rtf_monitor_thread.start()
    
    def run_rtf_processing(self, lot_path, project_id, cancellation_token):
        """Run the actual RTF processing in a background thread"""
        try:
            # Import the RTF processing function
            from process_rtf_content import process_lot_and_merge_rtf
            
            # Get LOT file directory
            lot_dir = os.path.dirname(lot_path)
            
            # Build the output filename
            output_file = os.path.join(lot_dir, f"{project_id}_rtf_title_footnote.rtf")
            
            # Add log about starting
            self.add_rtf_log(f"Processing LOT file: {lot_path}")
            self.add_rtf_log(f"Project ID: {project_id}")
            self.add_rtf_log(f"RTF directory: {lot_dir}")
            self.add_rtf_log(f"Output file: {output_file}")
            
            # Redirect stdout to capture print outputs
            old_stdout = sys.stdout
            sys.stdout = self.LogRedirector(self.add_rtf_log)
            
            try:
                # Call the processing function with cancellation token
                result = process_lot_and_merge_rtf(lot_path, output_file, cancellation_token)
            finally:
                sys.stdout = old_stdout
            
            # Check if cancelled
            if cancellation_token and cancellation_token.is_cancelled():
                self.add_rtf_log("RTF processing was cancelled by user.")
                return
            
            # Add results to log
            if result:
                self.add_rtf_log(f"RTF processing completed successfully")
                self.add_rtf_log(f"Output file: {output_file}")
            else:
                self.add_rtf_log(f"RTF processing failed")
            
        except Exception as e:
            if not cancellation_token or not cancellation_token.is_cancelled():
                self.add_rtf_log(f"Error during RTF processing: {str(e)}")
                self.add_rtf_log(traceback.format_exc())
            else:
                self.add_rtf_log("RTF processing was cancelled by user.")
        finally:
            self.rtf_processing = False
    
    def monitor_rtf_processing(self):
        """Monitor the RTF processing thread and update UI"""
        while self.rtf_processing and self.rtf_process_thread and self.rtf_process_thread.is_alive():
            time.sleep(0.1)
            # Check for cancellation
            if self.rtf_cancellation_token and self.rtf_cancellation_token.is_cancelled():
                break
        
        # Processing finished or cancelled, update UI
        self.root.after(100, self.on_rtf_processing_complete)
    
    def on_rtf_processing_complete(self):
        """Called when RTF processing is complete"""
        # Reset button states using the validation logic
        self.update_all_processor_buttons()
        self.rtf_cancel_btn.config(state="disabled", bg=self.disabled_color, fg=self.fg_color)
        if self.rtf_processing:  # Only show completion message if not cancelled
            self.add_rtf_log("RTF processing finished.")
        self.rtf_processing = False
        self.rtf_cancellation_token = None
    
    def cancel_rtf_processing(self):
        """Cancel the ongoing RTF processing"""
        if self.rtf_processing and self.rtf_cancellation_token:
            self.add_rtf_log("Cancelling RTF processing...")
            self.rtf_cancellation_token.cancel()
            # Give some time for graceful shutdown
            if self.rtf_process_thread and self.rtf_process_thread.is_alive():
                self.rtf_process_thread.join(timeout=2)
            # Reset button states
            self.update_all_processor_buttons()
            self.rtf_cancel_btn.config(state="disabled", bg=self.disabled_color, fg=self.fg_color)
            self.add_rtf_log("RTF processing cancelled.")
            self.rtf_processing = False
            self.rtf_cancellation_token = None
    
    def add_rtf_log(self, message):
        """Add a log message to the RTF logs text area"""
        # Update the text widget from the main thread
        self.root.after(0, lambda: self._add_log_entry(self.rtf_logs_text, self.rtf_logs_label, message))
    
    def copy_rtf_logs(self, event=None):
        """Copy selected text or all RTF logs to clipboard"""
        self._copy_logs_to_clipboard(self.rtf_logs_text, self.add_rtf_log)
    
    class LogRedirector:
        """Redirector class to capture print statements"""
        def __init__(self, log_func):
            self.log_func = log_func
        
        def write(self, message):
            if message.strip():
                self.log_func(message.strip())
        
        def flush(self):
            pass

def main():
    root = tk.Tk()
    
    # Hide window during initialization to prevent blank flash
    root.withdraw()
    
    # Set dark theme
    root.tk_setPalette(background='#1e1e2e', foreground='#d4d4d4')
    
    # Create app instance (icon is loaded in ExtractTitleFootnoteGUI.__init__)
    app = ExtractTitleFootnoteGUI(root)
    
    # Center the window
    window_width = 800
    window_height = 750
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Handle window close to also close README window
    def on_main_window_close():
        app.close_readme_window()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_main_window_close)
    
    # Show window after all components are ready
    root.deiconify()
    
    root.mainloop()

if __name__ == "__main__":
    main()