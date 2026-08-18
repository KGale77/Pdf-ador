import os
import sys
import threading
import queue
from tkinter import filedialog, messagebox
import customtkinter as ctk
from converter import convert_to_pdf, SUPPORTED_EXTS, WORD_EXTS, EXCEL_EXTS, PPT_EXTS, find_libreoffice

# Theme configuration
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Custom Color Palette
BG_COLOR = "#0f172a"          # Slate 900
CARD_BG = "#1e293b"           # Slate 800
ACCENT_COLOR = "#6366f1"      # Indigo 500
ACCENT_HOVER = "#4f46e5"      # Indigo 600
TEXT_MAIN = "#f8fafc"         # Slate 50
TEXT_MUTED = "#94a3b8"        # Slate 400

STATUS_COLORS = {
    "Pending": ("#475569", "#cbd5e1"),      # (Bg, Text) Slate 600
    "Converting": ("#d97706", "#fef3c7"),   # (Bg, Text) Amber 600
    "Success": ("#059669", "#d1fae5"),      # (Bg, Text) Emerald 600
    "Error": ("#e11d48", "#ffe4e6")          # (Bg, Text) Rose 600
}

class FileRow(ctk.CTkFrame):
    """A row inside the scrollable file list representing a single file."""
    def __init__(self, master, file_path, on_remove, **kwargs):
        super().__init__(master, fg_color=CARD_BG, corner_radius=8, height=60, **kwargs)
        self.file_path = file_path
        self.on_remove = on_remove
        
        # Grid layout for elements
        self.grid_columnconfigure(0, weight=0) # File Type Emoji
        self.grid_columnconfigure(1, weight=1) # File Name & Path
        self.grid_columnconfigure(2, weight=0) # File Size
        self.grid_columnconfigure(3, weight=0) # Status Badge
        self.grid_columnconfigure(4, weight=0) # Delete Button
        
        # 1. Type Emoji
        ext = os.path.splitext(file_path)[1].lower()
        if ext in WORD_EXTS:
            icon = "📝"
            icon_color = "#3b82f6"  # Word Blue
        elif ext in EXCEL_EXTS:
            icon = "📊"
            icon_color = "#10b981"  # Excel Green
        elif ext in PPT_EXTS:
            icon = "📈"
            icon_color = "#f97316"  # PPT Orange
        else:
            icon = "📄"
            icon_color = TEXT_MUTED
            
        self.icon_label = ctk.CTkLabel(
            self, text=icon, font=("Segoe UI", 22), width=40
        )
        self.icon_label.grid(row=0, column=0, padx=(12, 8), pady=10, sticky="w")
        
        # 2. File Name & Path
        file_name = os.path.basename(file_path)
        folder_name = os.path.basename(os.path.dirname(file_path))
        
        self.text_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.text_frame.grid(row=0, column=1, padx=4, pady=8, sticky="w")
        
        self.name_label = ctk.CTkLabel(
            self.text_frame, text=file_name, font=("Segoe UI", 13, "bold"),
            text_color=TEXT_MAIN, anchor="w"
        )
        self.name_label.pack(anchor="w")
        
        self.path_label = ctk.CTkLabel(
            self.text_frame, text=f".../{folder_name}", font=("Segoe UI", 11),
            text_color=TEXT_MUTED, anchor="w"
        )
        self.path_label.pack(anchor="w")
        
        # 3. File Size
        size_str = "0 KB"
        try:
            size_bytes = os.path.getsize(file_path)
            if size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.1f} KB"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
        except Exception:
            pass
            
        self.size_label = ctk.CTkLabel(
            self, text=size_str, font=("Segoe UI", 11), text_color=TEXT_MUTED
        )
        self.size_label.grid(row=0, column=2, padx=15, pady=10, sticky="e")
        
        # 4. Status Badge
        self.status_badge = ctk.CTkLabel(
            self, text="Pendiente", font=("Segoe UI", 11, "bold"),
            fg_color=STATUS_COLORS["Pending"][0],
            text_color=STATUS_COLORS["Pending"][1],
            corner_radius=6, height=22, width=90
        )
        self.status_badge.grid(row=0, column=3, padx=10, pady=10, sticky="e")
        
        # 5. Delete Button
        self.delete_btn = ctk.CTkButton(
            self, text="✕", font=("Segoe UI", 12, "bold"),
            fg_color="transparent", hover_color="#ef4444", text_color=TEXT_MUTED,
            width=28, height=28, corner_radius=14,
            command=lambda: self.on_remove(self)
        )
        self.delete_btn.grid(row=0, column=4, padx=(5, 12), pady=10, sticky="e")

    def update_status(self, status, error_msg=""):
        """Updates the visual status of this row."""
        bg, text_col = STATUS_COLORS.get(status, STATUS_COLORS["Pending"])
        
        # Status labels mapping
        status_labels = {
            "Pending": "Pendiente",
            "Converting": "Procesando",
            "Success": "Convertido",
            "Error": "Error"
        }
        
        self.status_badge.configure(
            fg_color=bg,
            text_color=text_col,
            text=status_labels.get(status, status)
        )
        
        if status == "Error" and error_msg:
            # If error, show tooltip or error indicator
            self.name_label.configure(text_color="#f43f5e")
            self.path_label.configure(text=f"Error: {error_msg[:40]}...", text_color="#f43f5e")
        elif status == "Success":
            self.name_label.configure(text_color="#10b981")
            self.path_label.configure(text_color=TEXT_MUTED)
            # Remove delete button since it's already converted
            self.delete_btn.grid_forget()

class PdfadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Pdf-ador - PDF Converter")
        self.geometry("900x650")
        self.minsize(800, 600)
        self.configure(fg_color=BG_COLOR)
        
        # Data state
        self.queue_files = [] # List of strings (paths)
        self.row_widgets = {} # Maps path -> FileRow widget
        self.conversion_states = {} # Maps path -> status string
        self.is_converting = False
        
        # Find LibreOffice path for fallbacks
        self.libreoffice_detected_path = find_libreoffice()
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        # Configure layout grids
        self.grid_rowconfigure(0, weight=0) # Header
        self.grid_rowconfigure(1, weight=1) # Main Content Area (Queue / Dashboard)
        self.grid_rowconfigure(2, weight=0) # Config Panel
        self.grid_rowconfigure(3, weight=0) # Footer Status
        self.grid_columnconfigure(0, weight=1)
        
        # ==================== HEADER ====================
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=30, pady=(20, 10), sticky="ew")
        
        self.title_label = ctk.CTkLabel(
            self.header_frame, text="PDF-ador 📄⚡", 
            font=("Segoe UI", 26, "bold"), text_color=TEXT_MAIN
        )
        self.title_label.pack(side="left")
        
        self.subtitle_label = ctk.CTkLabel(
            self.header_frame, text="Conversor portable a PDF", 
            font=("Segoe UI", 12, "italic"), text_color=TEXT_MUTED
        )
        self.subtitle_label.pack(side="left", padx=15, pady=(8, 0))
        
        self.stats_label = ctk.CTkLabel(
            self.header_frame, text="", font=("Segoe UI", 12, "bold"), text_color=TEXT_MUTED
        )
        self.stats_label.pack(side="right", pady=(8, 0))
        
        # ==================== MAIN CONTENT ====================
        # This frame holds either the Dashboard (when empty) or the Queue List
        self.main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content_frame.grid(row=1, column=0, padx=30, pady=10, sticky="nsew")
        self.main_content_frame.grid_rowconfigure(0, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)
        
        # 1. EMPTY STATE DASHBOARD CARD
        self.dashboard_card = ctk.CTkFrame(self.main_content_frame, fg_color=CARD_BG, corner_radius=16)
        self.dashboard_card.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.dashboard_card.grid_columnconfigure(0, weight=1)
        self.dashboard_card.grid_rowconfigure(0, weight=1)
        
        dashboard_inner = ctk.CTkFrame(self.dashboard_card, fg_color="transparent")
        dashboard_inner.grid(row=0, column=0, padx=40, pady=40)
        
        self.dash_icon = ctk.CTkLabel(
            dashboard_inner, text="📂", font=("Segoe UI", 64)
        )
        self.dash_icon.pack(pady=10)
        
        self.dash_title = ctk.CTkLabel(
            dashboard_inner, text="Carga tus documentos de Office",
            font=("Segoe UI", 20, "bold"), text_color=TEXT_MAIN
        )
        self.dash_title.pack(pady=5)
        
        self.dash_desc = ctk.CTkLabel(
            dashboard_inner, 
            text="Selecciona archivos o carpetas completas.\nSoporta formatos de Word, Excel y PowerPoint.",
            font=("Segoe UI", 13), text_color=TEXT_MUTED, justify="center"
        )
        self.dash_desc.pack(pady=(5, 20))
        
        dash_buttons_frame = ctk.CTkFrame(dashboard_inner, fg_color="transparent")
        dash_buttons_frame.pack()
        
        self.dash_add_files_btn = ctk.CTkButton(
            dash_buttons_frame, text="📄 Seleccionar Archivos", font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, height=42, width=180,
            command=self.browse_files
        )
        self.dash_add_files_btn.pack(side="left", padx=10)
        
        self.dash_add_folder_btn = ctk.CTkButton(
            dash_buttons_frame, text="📁 Seleccionar Carpeta", font=("Segoe UI", 13, "bold"),
            fg_color="#334155", hover_color="#475569", height=42, width=180,
            command=self.browse_folder
        )
        self.dash_add_folder_btn.pack(side="left", padx=10)
        
        # 2. QUEUE LIST FRAME (Shown when files are added)
        self.queue_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        
        # Header of queue list (Controls inside queue)
        self.queue_control_bar = ctk.CTkFrame(self.queue_frame, fg_color="transparent")
        self.queue_control_bar.pack(fill="x", pady=(0, 10))
        
        self.q_title = ctk.CTkLabel(
            self.queue_control_bar, text="Cola de Conversión", font=("Segoe UI", 15, "bold"), text_color=TEXT_MAIN
        )
        self.q_title.pack(side="left")
        
        self.q_clear_btn = ctk.CTkButton(
            self.queue_control_bar, text="🗑️ Limpiar Todo", font=("Segoe UI", 11, "bold"),
            fg_color="transparent", hover_color="#ef4444", text_color=TEXT_MUTED,
            height=28, width=100, command=self.clear_queue
        )
        self.q_clear_btn.pack(side="right", padx=5)
        
        self.q_add_more_btn = ctk.CTkButton(
            self.queue_control_bar, text="➕ Agregar Más", font=("Segoe UI", 11, "bold"),
            fg_color="#334155", hover_color="#475569", text_color=TEXT_MAIN,
            height=28, width=110, command=self.browse_files
        )
        self.q_add_more_btn.pack(side="right", padx=5)
        
        # Scrollable container for rows
        self.scrollable_files_frame = ctk.CTkScrollableFrame(
            self.queue_frame, fg_color="transparent", label_text=""
        )
        self.scrollable_files_frame.pack(fill="both", expand=True)
        
        # ==================== CONFIG & ACTION PANEL ====================
        self.config_panel = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        self.config_panel.grid(row=2, column=0, padx=30, pady=10, sticky="ew")
        
        # Column weights inside config panel
        self.config_panel.grid_columnconfigure(0, weight=1) # Dest folder configuration
        self.config_panel.grid_columnconfigure(1, weight=0) # Engine configuration
        self.config_panel.grid_columnconfigure(2, weight=0) # Action button
        
        # Left side: Destination Selection
        self.dest_frame = ctk.CTkFrame(self.config_panel, fg_color="transparent")
        self.dest_frame.grid(row=0, column=0, padx=20, pady=15, sticky="nsew")
        
        self.dest_title = ctk.CTkLabel(
            self.dest_frame, text="Carpeta de Salida", font=("Segoe UI", 13, "bold"), text_color=TEXT_MAIN
        )
        self.dest_title.pack(anchor="w")
        
        self.dest_option = ctk.StringVar(value="same")
        
        self.radio_same = ctk.CTkRadioButton(
            self.dest_frame, text="Misma carpeta de los archivos originales",
            value="same", variable=self.dest_option, font=("Segoe UI", 12),
            command=self.toggle_custom_dest_ui
        )
        self.radio_same.pack(anchor="w", pady=(8, 4))
        
        self.radio_custom = ctk.CTkRadioButton(
            self.dest_frame, text="Carpeta personalizada...",
            value="custom", variable=self.dest_option, font=("Segoe UI", 12),
            command=self.toggle_custom_dest_ui
        )
        self.radio_custom.pack(anchor="w", pady=4)
        
        self.custom_dest_frame = ctk.CTkFrame(self.dest_frame, fg_color="transparent")
        # Hidden by default
        
        self.entry_custom_path = ctk.CTkEntry(
            self.custom_dest_frame, placeholder_text="Selecciona carpeta...",
            font=("Segoe UI", 11), height=28
        )
        self.entry_custom_path.pack(side="left", fill="x", expand=True, padx=(20, 8))
        
        self.btn_browse_dest = ctk.CTkButton(
            self.custom_dest_frame, text="Buscar...", font=("Segoe UI", 11, "bold"),
            fg_color="#334155", hover_color="#475569", height=28, width=70,
            command=self.browse_output_dir
        )
        self.btn_browse_dest.pack(side="left")
        
        # Center: Engine Selector
        self.engine_frame = ctk.CTkFrame(self.config_panel, fg_color="transparent")
        self.engine_frame.grid(row=0, column=1, padx=20, pady=15, sticky="ns")
        
        self.engine_title = ctk.CTkLabel(
            self.engine_frame, text="Motor de Conversión", font=("Segoe UI", 13, "bold"), text_color=TEXT_MAIN
        )
        self.engine_title.pack(anchor="w", pady=(0, 8))
        
        self.engine_segmented = ctk.CTkSegmentedButton(
            self.engine_frame, values=["Auto", "MS Office", "LibreOffice"],
            font=("Segoe UI", 11, "bold"), command=self.on_engine_change
        )
        self.engine_segmented.set("Auto")
        self.engine_segmented.pack(pady=2)
        
        # Engine Info / Alert Text
        self.engine_info = ctk.CTkLabel(
            self.engine_frame, text="COM + LibreOffice Fallback", font=("Segoe UI", 10), text_color=TEXT_MUTED
        )
        self.engine_info.pack(anchor="w", pady=5)
        
        # Right side: Convert Button & Progress
        self.action_frame = ctk.CTkFrame(self.config_panel, fg_color="transparent")
        self.action_frame.grid(row=0, column=2, padx=20, pady=15, sticky="ns")
        
        self.progress_bar = ctk.CTkProgressBar(
            self.action_frame, width=180, height=8, fg_color="#334155", progress_color=ACCENT_COLOR
        )
        self.progress_bar.set(0)
        # Hidden by default
        
        self.convert_btn = ctk.CTkButton(
            self.action_frame, text="⚡ Convertir a PDF", font=("Segoe UI", 14, "bold"),
            fg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, height=45, width=180,
            command=self.start_conversion
        )
        self.convert_btn.pack(side="bottom", pady=(5, 0))
        
        # ==================== FOOTER ====================
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.grid(row=3, column=0, padx=30, pady=(5, 15), sticky="ew")
        
        # Display detected engine statuses
        office_status = "Instalado" if sys.platform == "win32" else "No Soportado"
        # We can check if office is available. Let's do a quick try-import/dispatch check:
        self.has_office = False
        if sys.platform == "win32":
            try:
                import win32com.client
                # We won't dispatch here because it opens the application, we just check if module exists
                self.has_office = True
            except ImportError:
                office_status = "No Instalado"
                
        lo_status = "Instalado" if self.libreoffice_detected_path else "No Detectado"
        
        self.status_footer = ctk.CTkLabel(
            self.footer_frame, 
            text=f"Sistemas:  Office COM [{office_status}]  |  LibreOffice [{lo_status}]",
            font=("Segoe UI", 11), text_color=TEXT_MUTED
        )
        self.status_footer.pack(side="left")
        
        self.footer_note = ctk.CTkLabel(
            self.footer_frame, text="Versión Portable v1.0", font=("Segoe UI", 10), text_color=TEXT_MUTED
        )
        self.footer_note.pack(side="right")
        
        self.update_content_view()

    # ==================== UI STATE MANAGEMENT ====================
    def update_content_view(self):
        """Switches between Empty state and Queue list based on files in queue."""
        if not self.queue_files:
            self.queue_frame.pack_forget()
            self.dashboard_card.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
            self.convert_btn.configure(state="disabled")
            self.stats_label.configure(text="")
        else:
            self.dashboard_card.grid_forget()
            self.queue_frame.pack(fill="both", expand=True)
            self.convert_btn.configure(state="normal")
            self.update_stats()

    def update_stats(self):
        """Recalculates conversion stats and updates label."""
        total = len(self.queue_files)
        success = sum(1 for p in self.queue_files if self.conversion_states.get(p) == "Success")
        failed = sum(1 for p in self.queue_files if self.conversion_states.get(p) == "Error")
        converting = sum(1 for p in self.queue_files if self.conversion_states.get(p) == "Converting")
        
        if total > 0:
            pct = (success + failed) / total
            self.stats_label.configure(
                text=f"Procesados: {success + failed}/{total} | Completado: {pct * 100:.0f}%"
            )
            if self.is_converting:
                self.progress_bar.set(pct)
        else:
            self.stats_label.configure(text="")

    def toggle_custom_dest_ui(self):
        """Displays or hides the Custom Destination Folder field."""
        if self.dest_option.get() == "custom":
            self.custom_dest_frame.pack(fill="x", anchor="w", pady=(5, 0))
        else:
            self.custom_dest_frame.pack_forget()

    def on_engine_change(self, value):
        """Updates text description when engine changes."""
        if value == "Auto":
            self.engine_info.configure(text="COM + LibreOffice Fallback")
        elif value == "MS Office":
            self.engine_info.configure(text="Automatización Office COM Nativo")
        elif value == "LibreOffice":
            self.engine_info.configure(text="Conversor LibreOffice Portable/Headless")

    # ==================== QUEUE OPERATIONS ====================
    def add_to_queue(self, paths):
        """Adds unique supported files to conversion list."""
        added_any = False
        for path in paths:
            path = os.path.abspath(path)
            if path in self.queue_files:
                continue
                
            ext = os.path.splitext(path)[1].lower()
            if ext in SUPPORTED_EXTS:
                self.queue_files.append(path)
                self.conversion_states[path] = "Pending"
                
                # Create Row UI
                row = FileRow(self.scrollable_files_frame, path, on_remove=self.remove_from_queue)
                row.pack(fill="x", pady=4, padx=5)
                self.row_widgets[path] = row
                added_any = True
                
        if added_any:
            self.update_content_view()

    def remove_from_queue(self, row_widget):
        """Removes a file from queue list."""
        if self.is_converting:
            return # Block removal during conversion
            
        path = row_widget.file_path
        if path in self.queue_files:
            self.queue_files.remove(path)
            
        if path in self.conversion_states:
            del self.conversion_states[path]
            
        if path in self.row_widgets:
            self.row_widgets[path].destroy()
            del self.row_widgets[path]
            
        self.update_content_view()

    def clear_queue(self):
        """Resets the entire queue list."""
        if self.is_converting:
            return
            
        for row in self.row_widgets.values():
            row.destroy()
            
        self.queue_files.clear()
        self.row_widgets.clear()
        self.conversion_states.clear()
        self.update_content_view()

    # ==================== DIALOG BROWSER TRIGGERS ====================
    def browse_files(self):
        """Spawns standard file dialog for file selection."""
        # Convert SUPPORTED_EXTS to tkinter filetypes tuple
        filetypes = [
            ("Archivos de MS Office", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt"),
            ("Documentos de Word", "*.docx *.doc"),
            ("Planillas de Excel", "*.xlsx *.xls"),
            ("Presentaciones PowerPoint", "*.pptx *.ppt"),
            ("Todos los archivos", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos de Office",
            filetypes=filetypes
        )
        if files:
            self.add_to_queue(files)

    def browse_folder(self):
        """Spawns directory selector and scans recursively for files."""
        dir_path = filedialog.askdirectory(title="Seleccionar carpeta con archivos de Office")
        if not dir_path:
            return
            
        # Scan folder for supported documents
        found_paths = []
        for root, _, files in os.walk(dir_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_EXTS:
                    found_paths.append(os.path.join(root, file))
                    
        if found_paths:
            self.add_to_queue(found_paths)
        else:
            messagebox.showinfo(
                "Búsqueda completa", 
                "No se encontraron archivos de Office compatibles (.docx, .doc, .xlsx, .xls, .pptx, .ppt) en la carpeta seleccionada."
            )

    def browse_output_dir(self):
        """Spawns directory selector for output files."""
        dir_path = filedialog.askdirectory(title="Seleccionar carpeta de destino para PDFs")
        if dir_path:
            self.entry_custom_path.delete(0, "end")
            self.entry_custom_path.insert(0, os.path.abspath(dir_path))

    # ==================== CONVERSION ENGINE EXECUTION ====================
    def start_conversion(self):
        """Prepares state and fires background execution thread."""
        if self.is_converting or not self.queue_files:
            return
            
        # Check Custom output path validity
        custom_out = None
        if self.dest_option.get() == "custom":
            custom_out = self.entry_custom_path.get().strip()
            if not custom_out or not os.path.exists(custom_out):
                messagebox.showerror("Carpeta inválida", "La carpeta de salida especificada no es válida o no existe.")
                return
                
        # Engine choice
        engine = self.engine_segmented.get().lower()
        if engine == "auto":
            engine = "auto"
        elif engine == "ms office":
            engine = "office"
        elif engine == "libreoffice":
            engine = "libreoffice"
            
        # Lock UI
        self.is_converting = True
        self.convert_btn.configure(state="disabled", text="Procesando...")
        self.q_clear_btn.configure(state="disabled")
        self.q_add_more_btn.configure(state="disabled")
        self.dash_add_files_btn.configure(state="disabled")
        self.dash_add_folder_btn.configure(state="disabled")
        
        # Display Progress UI elements
        self.progress_bar.pack(side="top", pady=5)
        self.progress_bar.set(0)
        
        # Reset completed status from previous runs so we can re-process if needed
        for path in self.queue_files:
            state = self.conversion_states.get(path)
            if state != "Success":
                self.conversion_states[path] = "Pending"
                self.row_widgets[path].update_status("Pending")
                
        # Start execution thread
        threading.Thread(
            target=self.run_conversion_worker, 
            args=(engine, custom_out),
            daemon=True
        ).start()

    def run_conversion_worker(self, engine, custom_output_dir):
        """Worker thread executing conversions sequentially."""
        total_files = len(self.queue_files)
        success_count = 0
        failed_count = 0
        
        for index, path in enumerate(self.queue_files):
            # Check if it was already successfully converted in previous click
            if self.conversion_states.get(path) == "Success":
                success_count += 1
                continue
                
            # Update UI to Converting
            self.conversion_states[path] = "Converting"
            self.safe_update_row_status(path, "Converting")
            self.safe_update_stats()
            
            # Setup output PDF path
            input_name_no_ext = os.path.splitext(os.path.basename(path))[0]
            
            if custom_output_dir:
                output_path = os.path.join(custom_output_dir, f"{input_name_no_ext}.pdf")
            else:
                output_path = os.path.join(os.path.dirname(path), f"{input_name_no_ext}.pdf")
                
            try:
                # Convert file!
                convert_to_pdf(
                    input_path=path,
                    output_path=output_path,
                    engine=engine,
                    libreoffice_path=self.libreoffice_detected_path
                )
                
                # Mark Success
                self.conversion_states[path] = "Success"
                self.safe_update_row_status(path, "Success")
                success_count += 1
            except Exception as e:
                # Mark Failed
                error_msg = str(e)
                self.conversion_states[path] = "Error"
                self.safe_update_row_status(path, "Error", error_msg)
                failed_count += 1
                
            self.safe_update_stats()
            
        # Conversion completed!
        self.after(0, self.finish_conversion, success_count, failed_count, custom_output_dir)

    def finish_conversion(self, success, failed, custom_output_dir):
        """Fires when worker thread completes."""
        self.is_converting = False
        
        # Unlock UI
        self.convert_btn.configure(state="normal", text="⚡ Convertir a PDF")
        self.q_clear_btn.configure(state="normal")
        self.q_add_more_btn.configure(state="normal")
        self.dash_add_files_btn.configure(state="normal")
        self.dash_add_folder_btn.configure(state="normal")
        
        # Hide progress bar
        self.progress_bar.pack_forget()
        
        # Summary message
        total = success + failed
        if failed == 0:
            msg = f"Se convirtieron con éxito {success} de {total} archivos a PDF."
            messagebox.showinfo("Conversión Finalizada", msg)
        else:
            msg = f"Conversión finalizada.\n\nExitosos: {success}\nFallidos: {failed}\n\nRevisa los detalles en cada archivo de la lista."
            messagebox.showerror("Conversión Completa con Errores", msg)
            
        # Offer to open output folder
        if success > 0:
            open_folder = messagebox.askyesno(
                "Abrir carpeta", 
                "¿Deseas abrir la carpeta de destino donde se guardaron los PDFs?"
            )
            if open_folder:
                # Determine folder path to open
                folder_to_open = None
                if custom_output_dir:
                    folder_to_open = custom_output_dir
                elif self.queue_files:
                    # Open folder of the first file
                    folder_to_open = os.path.dirname(self.queue_files[0])
                    
                if folder_to_open and os.path.exists(folder_to_open):
                    os.startfile(folder_to_open)

    # Thread-safe UI update wrappers
    def safe_update_row_status(self, path, status, error_msg=""):
        self.after(0, lambda: self.row_widgets[path].update_status(status, error_msg))

    def safe_update_stats(self):
        self.after(0, self.update_stats)

if __name__ == "__main__":
    app = PdfadorApp()
    app.mainloop()
