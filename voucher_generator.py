#!/usr/bin/env python3
"""
Alyssa's Voucher Generator
GUI app with live preview for creating personalized drink vouchers.
Repository: https://github.com/riconanci/TicketGen
"""

import csv
import os
import math
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import traceback


class VoucherGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Alyssa's Voucher Generator")
        self.root.geometry("720x930")
        self.root.resizable(False, False)
        
        # Variables
        self.csv_path = None
        self.image_path = None
        self.attendees = []
        self.voucher_image = None
        self.image_aspect_ratio = 1.71  # Default ratio
        
        # Blanks mode
        self.blanks_mode = tk.IntVar(value=0)
        self.extra_text_var = tk.StringVar(value="")
        self.blank_pages_var = tk.StringVar(value="1")
        
        # Title settings
        self.title_var = tk.StringVar(value="DRINK VOUCHER")
        self.title_font_size_var = tk.StringVar(value="10")
        self.title_bold_var = tk.IntVar(value=1)
        self.title_color = "#000000"
        self.title_outline_var = tk.IntVar(value=0)
        self.title_underline_var = tk.IntVar(value=0)
        
        # Name settings
        self.name_font_size_var = tk.StringVar(value="14")
        self.name_bold_var = tk.IntVar(value=1)
        self.name_color = "#000000"
        self.name_outline_var = tk.IntVar(value=0)
        self.name_underline_var = tk.IntVar(value=0)
        self.swap_names_var = tk.IntVar(value=0)  # Swap first/last name order
        
        # Layout variables
        self.orientation_var = tk.StringVar(value="Portrait")
        self.ticket_width_var = tk.StringVar(value="3")
        self.ticket_height_var = tk.StringVar(value="1.75")
        self.tickets_per_attendee_var = tk.StringVar(value="5")
        self.align_top_left_var = tk.IntVar(value=1)
        self.batch_mode_var = tk.IntVar(value=1)  # Group tickets by attendee (default on)
        self.cutting_guides_var = tk.IntVar(value=0)  # Dotted cutting lines (default off)
        self.bw_mode_var = tk.IntVar(value=0)  # Black and white mode (default off)
        
        # Preview mode
        self.preview_mode = tk.StringVar(value="ticket")
        
        # Draggable text positions (percentage from center, negative=up, positive=down)
        self.title_y_pos = -0.25
        self.name_y_pos = 0.08
        
        # Dragging state
        self.dragging = None
        self.drag_start_y = 0
        self.drag_start_pos = 0
        self.preview_ticket_height = 0
        self.preview_offset_y = 0
        
        self.setup_ui()
        self.update_valid_sizes()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="8")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === BLANKS AND ABOUT BUTTONS (top-right) ===
        about_row = ttk.Frame(main_frame)
        about_row.pack(fill=tk.X, pady=(0, 4))
        self.about_btn = ttk.Button(about_row, text="About", command=self.toggle_about, bootstyle="secondary-outline")
        self.about_btn.pack(side=tk.RIGHT)
        self.blanks_btn = ttk.Button(about_row, text="Blanks", command=self.toggle_blanks_mode, bootstyle="warning-outline")
        self.blanks_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
        # Track about window
        self.about_window = None
        
        # === FILE SELECTION ===
        file_frame = ttk.Labelframe(main_frame, text="Step 1: Select Files", padding="6")
        file_frame.pack(fill=tk.X, pady=(0, 4))
        
        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X)
        
        self.csv_btn = ttk.Button(file_row, text="Select CSV", command=self.select_csv, bootstyle="success")
        self.csv_btn.pack(side=tk.LEFT)
        self.csv_label = ttk.Label(file_row, text="No file", foreground="gray")
        self.csv_label.pack(side=tk.LEFT, padx=(8, 15))
        
        self.img_btn = ttk.Button(file_row, text="Select Image", command=self.select_image, bootstyle="info")
        self.img_btn.pack(side=tk.LEFT)
        self.img_label = ttk.Label(file_row, text="No file", foreground="gray")
        self.img_label.pack(side=tk.LEFT, padx=(8, 15))
        
        self.bw_check = ttk.Checkbutton(file_row, text="B&W", 
                                         variable=self.bw_mode_var, command=self.on_bw_change)
        self.bw_check.pack(side=tk.LEFT)
        
        # === PREVIEW ===
        preview_frame = ttk.Labelframe(main_frame, text="Preview (drag title/name to reposition)", padding="4")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 1))
        
        toggle_row = ttk.Frame(preview_frame)
        toggle_row.pack(pady=(0, 2))
        
        self.ticket_btn = ttk.Button(toggle_row, text="Single Ticket", command=lambda: self.set_preview_mode("ticket"),
                                      bootstyle="info")
        self.ticket_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.layout_btn = ttk.Button(toggle_row, text="Page Layout", command=lambda: self.set_preview_mode("layout"),
                                      bootstyle="secondary")
        self.layout_btn.pack(side=tk.LEFT)
        
        self.preview_canvas = tk.Canvas(preview_frame, width=500, height=240, bg="white", relief="flat", bd=1, 
                                         highlightthickness=1, highlightbackground="#dee2e6")
        self.preview_canvas.pack(pady=2)
        self.preview_canvas.create_text(250, 120, text="Select an image to see preview", fill="gray")
        
        self.preview_canvas.bind("<Button-1>", self.on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        self.layout_info_label = ttk.Label(preview_frame, text="", foreground="blue")
        self.layout_info_label.pack()
        
        self.calc_info_label = ttk.Label(preview_frame, text="", foreground="blue")
        self.calc_info_label.pack()
        
        # === TEXT SETTINGS ===
        text_frame = ttk.Labelframe(main_frame, text="Step 2: Text Settings", padding="6")
        text_frame.pack(fill=tk.X, pady=(0, 4))
        text_frame.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Title text entry row
        self.title_text_row = ttk.Frame(text_frame)
        self.title_text_row.pack(fill=tk.X, pady=(0, 4))
        self.title_text_row.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        ttk.Label(self.title_text_row, text="Title Text:", width=10).pack(side=tk.LEFT)
        title_entry = ttk.Entry(self.title_text_row, textvariable=self.title_var, width=25)
        title_entry.pack(side=tk.LEFT)
        title_entry.bind('<KeyRelease>', lambda e: self.on_step2_interact())
        title_entry.bind('<FocusIn>', lambda e: self.on_step2_interact())
        ttk.Label(self.title_text_row, text="(leave empty for no title)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        
        # Extra text entry row (for blanks mode)
        self.extra_text_row = ttk.Frame(text_frame)
        self.extra_text_row.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        ttk.Label(self.extra_text_row, text="Extra Text:", width=10).pack(side=tk.LEFT)
        extra_entry = ttk.Entry(self.extra_text_row, textvariable=self.extra_text_var, width=25)
        extra_entry.pack(side=tk.LEFT)
        extra_entry.bind('<KeyRelease>', lambda e: self.on_step2_interact())
        extra_entry.bind('<FocusIn>', lambda e: self.on_step2_interact())
        ttk.Label(self.extra_text_row, text="(optional single line)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        # Note: extra_text_row is NOT packed by default - shown only in blanks mode
        
        # Use grid for parameter alignment
        params_frame = ttk.Frame(text_frame)
        params_frame.pack(fill=tk.X)
        params_frame.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Header row
        ttk.Label(params_frame, text="", width=10).grid(row=0, column=0)
        ttk.Label(params_frame, text="Size", width=6).grid(row=0, column=1, padx=(0, 8))
        ttk.Label(params_frame, text="Color", width=6).grid(row=0, column=2, padx=(0, 8))
        ttk.Label(params_frame, text="Bold", width=5).grid(row=0, column=3, padx=(0, 8))
        ttk.Label(params_frame, text="Outline", width=7).grid(row=0, column=4, padx=(0, 8))
        ttk.Label(params_frame, text="Underline", width=9).grid(row=0, column=5)
        
        # Title row
        ttk.Label(params_frame, text="Title:", width=10).grid(row=1, column=0, sticky='w', pady=4)
        
        title_size = ttk.Combobox(params_frame, textvariable=self.title_font_size_var, 
                                   values=["8", "9", "10", "11", "12", "13", "14", "16", "18", "20", "24"], width=5, state="readonly")
        title_size.grid(row=1, column=1, padx=(0, 8), pady=4)
        title_size.bind('<<ComboboxSelected>>', lambda e: self.on_step2_interact())
        title_size.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.title_color_canvas = tk.Canvas(params_frame, width=28, height=22, 
                                             highlightthickness=1, highlightbackground="#999", cursor="hand2")
        self.title_color_canvas.grid(row=1, column=2, padx=(0, 8), pady=4)
        self.title_color_canvas.create_rectangle(0, 0, 29, 23, fill=self.title_color, outline="")
        self.title_color_canvas.bind('<Button-1>', lambda e: self.pick_title_color())
        
        self.title_bold_check = ttk.Checkbutton(params_frame, text="", variable=self.title_bold_var, 
                                                 command=self.on_step2_interact, bootstyle="round-toggle")
        self.title_bold_check.grid(row=1, column=3, padx=(0, 8), pady=4)
        self.title_bold_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.title_outline_check = ttk.Checkbutton(params_frame, text="", variable=self.title_outline_var, 
                                                    command=self.on_step2_interact, bootstyle="round-toggle")
        self.title_outline_check.grid(row=1, column=4, padx=(0, 8), pady=4)
        self.title_outline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.title_underline_check = ttk.Checkbutton(params_frame, text="", variable=self.title_underline_var, 
                                                      command=self.on_step2_interact, bootstyle="round-toggle")
        self.title_underline_check.grid(row=1, column=5, pady=4)
        self.title_underline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Name row (label changes to "Extra:" in blanks mode)
        self.name_row_label = ttk.Label(params_frame, text="Name:", width=10)
        self.name_row_label.grid(row=2, column=0, sticky='w', pady=4)
        
        name_size = ttk.Combobox(params_frame, textvariable=self.name_font_size_var, 
                                  values=["8", "9", "10", "11", "12", "13", "14", "16", "18", "20", "24"], width=5, state="readonly")
        name_size.grid(row=2, column=1, padx=(0, 8), pady=4)
        name_size.bind('<<ComboboxSelected>>', lambda e: self.on_step2_interact())
        name_size.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.name_color_canvas = tk.Canvas(params_frame, width=28, height=22, 
                                            highlightthickness=1, highlightbackground="#999", cursor="hand2")
        self.name_color_canvas.grid(row=2, column=2, padx=(0, 8), pady=4)
        self.name_color_canvas.create_rectangle(0, 0, 29, 23, fill=self.name_color, outline="")
        self.name_color_canvas.bind('<Button-1>', lambda e: self.pick_name_color())
        
        self.name_bold_check = ttk.Checkbutton(params_frame, text="", variable=self.name_bold_var, 
                                                command=self.on_step2_interact, bootstyle="round-toggle")
        self.name_bold_check.grid(row=2, column=3, padx=(0, 8), pady=4)
        self.name_bold_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.name_outline_check = ttk.Checkbutton(params_frame, text="", variable=self.name_outline_var, 
                                                   command=self.on_step2_interact, bootstyle="round-toggle")
        self.name_outline_check.grid(row=2, column=4, padx=(0, 8), pady=4)
        self.name_outline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.name_underline_check = ttk.Checkbutton(params_frame, text="", variable=self.name_underline_var, 
                                                     command=self.on_step2_interact, bootstyle="round-toggle")
        self.name_underline_check.grid(row=2, column=5, pady=4)
        self.name_underline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Swap names row (hidden in blanks mode)
        self.swap_row = ttk.Frame(text_frame)
        self.swap_row.pack(fill=tk.X, pady=(4, 0))
        self.swap_row.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        ttk.Label(self.swap_row, text="", width=10).pack(side=tk.LEFT)
        self.swap_names_check = ttk.Checkbutton(self.swap_row, text="Swap First/Last name order", 
                                                 variable=self.swap_names_var, command=self.on_step2_interact)
        self.swap_names_check.pack(side=tk.LEFT)
        self.swap_names_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # === LAYOUT SETTINGS ===
        layout_frame = ttk.Labelframe(main_frame, text="Step 3: Ticket Size & Layout", padding="6")
        layout_frame.pack(fill=tk.X, pady=(0, 4))
        layout_frame.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # Orientation row
        orient_row = ttk.Frame(layout_frame)
        orient_row.pack(fill=tk.X, pady=2)
        orient_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        ttk.Label(orient_row, text="Page:", width=10).pack(side=tk.LEFT)
        portrait_rb = ttk.Radiobutton(orient_row, text="Portrait (8.5×11)", variable=self.orientation_var, 
                                       value="Portrait", command=self.on_orientation_change)
        portrait_rb.pack(side=tk.LEFT, padx=(0, 15))
        portrait_rb.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        landscape_rb = ttk.Radiobutton(orient_row, text="Landscape (11×8.5)", variable=self.orientation_var, 
                                        value="Landscape", command=self.on_orientation_change)
        landscape_rb.pack(side=tk.LEFT)
        landscape_rb.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # Tickets per attendee row (MOVED BEFORE ticket size)
        tpa_row = ttk.Frame(layout_frame)
        tpa_row.pack(fill=tk.X, pady=2)
        tpa_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        self.tickets_label = ttk.Label(tpa_row, text="Tickets:", width=10)
        self.tickets_label.pack(side=tk.LEFT)
        self.tpa_combo = ttk.Combobox(tpa_row, textvariable=self.tickets_per_attendee_var,
                                       values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], 
                                       width=5, state="readonly")
        self.tpa_combo.pack(side=tk.LEFT)
        self.tpa_combo.bind('<<ComboboxSelected>>', lambda e: self.on_step3_interact())
        self.tpa_combo.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        self.attendee_info_label = ttk.Label(tpa_row, text="", bootstyle="info")
        self.attendee_info_label.pack(side=tk.LEFT, padx=(15, 0))
        
        # Pages field (for blanks mode only)
        self.pages_label = ttk.Label(tpa_row, text="Pages:", width=10)
        self.pages_combo = ttk.Combobox(tpa_row, textvariable=self.blank_pages_var,
                                        values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "15", "20"], 
                                        width=5, state="readonly")
        self.pages_combo.bind('<<ComboboxSelected>>', lambda e: self.on_step3_interact())
        self.pages_combo.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        # Note: pages_label and pages_combo are NOT packed by default - shown only in blanks mode
        
        # Size row (MOVED AFTER tickets)
        size_row = ttk.Frame(layout_frame)
        size_row.pack(fill=tk.X, pady=2)
        size_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        ttk.Label(size_row, text="Ticket Size:", width=10).pack(side=tk.LEFT)
        self.width_combo = ttk.Combobox(size_row, textvariable=self.ticket_width_var, width=5, state="readonly")
        self.width_combo.pack(side=tk.LEFT, padx=(0, 2))
        self.width_combo.bind('<<ComboboxSelected>>', lambda e: self.on_size_change())
        self.width_combo.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        ttk.Label(size_row, text="×").pack(side=tk.LEFT, padx=2)
        self.height_combo = ttk.Combobox(size_row, textvariable=self.ticket_height_var, width=5, state="readonly")
        self.height_combo.pack(side=tk.LEFT, padx=(0, 2))
        self.height_combo.bind('<<ComboboxSelected>>', lambda e: self.on_size_change())
        self.height_combo.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        ttk.Label(size_row, text="in").pack(side=tk.LEFT, padx=(0, 10))
        
        self.grid_info_label = ttk.Label(size_row, text="", bootstyle="info")
        self.grid_info_label.pack(side=tk.LEFT)
        
        # Align top-left row
        align_row = ttk.Frame(layout_frame)
        align_row.pack(fill=tk.X, pady=2)
        align_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        ttk.Label(align_row, text="", width=10).pack(side=tk.LEFT)
        self.align_check = ttk.Checkbutton(align_row, text="Align Top-Left (easier cutting)", 
                                            variable=self.align_top_left_var, command=self.on_align_change)
        self.align_check.pack(side=tk.LEFT)
        self.align_check.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # Batch mode row (hidden in blanks mode)
        self.batch_row = ttk.Frame(layout_frame)
        self.batch_row.pack(fill=tk.X, pady=2)
        self.batch_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        ttk.Label(self.batch_row, text="", width=10).pack(side=tk.LEFT)
        self.batch_check = ttk.Checkbutton(self.batch_row, text="Group tickets by attendee (easier distribution)", 
                                            variable=self.batch_mode_var, command=self.on_batch_change)
        self.batch_check.pack(side=tk.LEFT)
        self.batch_check.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # Cutting guides row
        cutting_row = ttk.Frame(layout_frame)
        cutting_row.pack(fill=tk.X, pady=2)
        cutting_row.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        ttk.Label(cutting_row, text="", width=10).pack(side=tk.LEFT)
        self.cutting_check = ttk.Checkbutton(cutting_row, text="Add dotted cutting lines", 
                                              variable=self.cutting_guides_var)
        self.cutting_check.pack(side=tk.LEFT)
        self.cutting_check.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # === GENERATE ===
        generate_frame = ttk.Frame(main_frame)
        generate_frame.pack(fill=tk.X, pady=6)
        
        self.generate_btn = ttk.Button(generate_frame, text="GENERATE PDF", command=self.generate_pdf,
                                        bootstyle="danger", width=18)
        self.generate_btn.pack(ipady=6)
        self.generate_btn.configure(state="disabled")
        
        # Configure button font using style
        style = ttk.Style()
        style.configure("danger.TButton", font=("Segoe UI", 12, "bold"))
        
        self.status_label = ttk.Label(generate_frame, text="Select CSV and image to get started", 
                                       foreground="gray", font=("Segoe UI", 11))
        self.status_label.pack(pady=(6, 10))
    
    def on_step2_interact(self):
        self.set_preview_mode("ticket")
        self.update_preview()
    
    def toggle_about(self):
        """Toggle About dialog"""
        # If window exists and is still open, close it
        if self.about_window is not None:
            try:
                self.about_window.destroy()
            except:
                pass
            self.about_window = None
            self.about_btn.configure(bootstyle="secondary-outline")
            return
        
        # Open the window
        self.about_btn.configure(bootstyle="secondary")
        self.about_window = tk.Toplevel(self.root)
        self.about_window.overrideredirect(True)  # Remove title bar and X button
        self.about_window.transient(self.root)
        
        # Window size
        win_w, win_h = 280, 190
        
        # Center on parent
        x = self.root.winfo_x() + 230
        y = self.root.winfo_y() + 250
        self.about_window.geometry(f"{win_w}x{win_h}+{x}+{y}")
        
        # Use Canvas to draw background (bypasses ttkbootstrap theming)
        canvas = tk.Canvas(self.about_window, width=win_w, height=win_h, 
                          highlightthickness=0, bd=0)
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # Draw border (green)
        canvas.create_rectangle(0, 0, win_w, win_h, fill="#27ae60", outline="#27ae60")
        
        # Draw inner background (light mint)
        border = 3
        canvas.create_rectangle(border, border, win_w-border, win_h-border, 
                               fill="#e8f8f0", outline="#e8f8f0")
        
        # Draw top accent stripe
        canvas.create_rectangle(border, border, win_w-border, border+6, 
                               fill="#27ae60", outline="#27ae60")
        
        # Draw text directly on canvas
        center_x = win_w // 2
        
        canvas.create_text(center_x, 55, text="Alyssa's", 
                          font=("Segoe UI", 22, "bold"), fill="#1e8449")
        
        canvas.create_text(center_x, 90, text="Voucher Generator", 
                          font=("Segoe UI", 13), fill="#2e7d4a")
        
        canvas.create_text(center_x, 130, text="version 1.0", 
                          font=("Segoe UI", 9), fill="#5a9a6e")
        
        canvas.create_text(center_x, 155, text="made with love - R", 
                          font=("Segoe UI", 9, "italic"), fill="#5a9a6e")
        
        # Make sure window appears on top
        self.about_window.lift()
        self.about_window.grab_set()
        
        # Click anywhere to close
        def close_about(e=None):
            self.about_window.destroy()
            self.about_window = None
            self.about_btn.configure(bootstyle="secondary-outline")
        
        canvas.bind("<Button-1>", close_about)
        self.about_window.bind("<Button-1>", close_about)
    
    def toggle_blanks_mode(self):
        """Toggle between normal mode and blanks mode"""
        if self.blanks_mode.get():
            # Turn OFF blanks mode - restore normal UI
            self.blanks_mode.set(0)
            self.blanks_btn.configure(bootstyle="warning-outline")
            
            # Show CSV elements - need to repack in order
            # First unpack image button and label temporarily
            self.img_btn.pack_forget()
            self.img_label.pack_forget()
            self.bw_check.pack_forget()
            
            # Now pack everything in correct order
            self.csv_btn.pack(side=tk.LEFT)
            self.csv_label.pack(side=tk.LEFT, padx=(8, 15))
            self.img_btn.pack(side=tk.LEFT)
            self.img_label.pack(side=tk.LEFT, padx=(8, 15))
            self.bw_check.pack(side=tk.LEFT)
            
            # Hide Extra Text row, show Swap row
            self.extra_text_row.pack_forget()
            self.swap_row.pack(fill=tk.X, pady=(4, 0))
            
            # Change label back to "Name:"
            self.name_row_label.configure(text="Name:")
            
            # Show Tickets field, hide Pages, show attendee info and batch row
            self.pages_label.pack_forget()
            self.pages_combo.pack_forget()
            self.tickets_label.pack(side=tk.LEFT)
            self.tpa_combo.pack(side=tk.LEFT)
            self.attendee_info_label.pack(side=tk.LEFT, padx=(15, 0))
            self.batch_row.pack(fill=tk.X, pady=2)
            
        else:
            # Turn ON blanks mode
            self.blanks_mode.set(1)
            self.blanks_btn.configure(bootstyle="warning")
            
            # Hide CSV elements
            self.csv_btn.pack_forget()
            self.csv_label.pack_forget()
            
            # Show Extra Text row after title_text_row, hide Swap row
            self.extra_text_row.pack(fill=tk.X, pady=(0, 4), after=self.title_text_row)
            self.swap_row.pack_forget()
            
            # Change label to "Extra:"
            self.name_row_label.configure(text="Extra:")
            
            # Hide Tickets field, show Pages, hide attendee info and batch row
            self.tickets_label.pack_forget()
            self.tpa_combo.pack_forget()
            self.attendee_info_label.pack_forget()
            self.pages_label.pack(side=tk.LEFT)
            self.pages_combo.pack(side=tk.LEFT)
            self.batch_row.pack_forget()
        
        self.check_ready()
        self.update_preview()
    
    def on_step3_interact(self):
        """Called when orientation, tickets, or align changes - auto-fits to image"""
        self.set_preview_mode("layout")
        self.auto_fit_to_image()
        self.update_preview()
    
    def on_size_change(self):
        """Called when width/height manually changed - no auto-fit"""
        self.set_preview_mode("layout")
        self.update_valid_sizes()
        self.update_preview()
    
    def on_align_change(self):
        """Called when align checkbox changed - no auto-fit, just update preview"""
        self.set_preview_mode("layout")
        self.update_preview()
    
    def on_orientation_change(self):
        """Called when orientation changed - update valid sizes but don't auto-fit"""
        self.set_preview_mode("layout")
        self.update_valid_sizes()
        self.update_preview()
    
    def on_batch_change(self):
        """Called when batch mode checkbox changed"""
        self.set_preview_mode("layout")
        self.update_preview()
    
    def on_bw_change(self):
        """Called when black & white checkbox changed"""
        self.update_preview()
    
    def get_processed_image(self):
        """Get the voucher image with B&W filter applied if enabled"""
        if not self.voucher_image:
            return None
        
        img = self.voucher_image.copy()
        if self.bw_mode_var.get():
            # Convert to grayscale, then back to RGB/RGBA for compatibility
            if img.mode == 'RGBA':
                # Preserve alpha channel
                r, g, b, a = img.split()
                gray = img.convert('L')
                img = Image.merge('RGBA', (gray, gray, gray, a))
            else:
                gray = img.convert('L')
                img = Image.merge('RGB', (gray, gray, gray))
        return img
    
    def pick_title_color(self):
        self.set_preview_mode("ticket")
        # Default palette position to hue 120 (green area)
        initial = "#40C040"  # Hue 120 green
        color = colorchooser.askcolor(color=initial, title="Choose Title Color")
        if color[1]:
            self.title_color = color[1]
            self.title_color_canvas.delete("all")
            self.title_color_canvas.create_rectangle(0, 0, 29, 23, fill=self.title_color, outline="")
            self.update_preview()
    
    def pick_name_color(self):
        self.set_preview_mode("ticket")
        # Default palette position to hue 120 (green area)
        initial = "#40C040"  # Hue 120 green
        color = colorchooser.askcolor(color=initial, title="Choose Name Color")
        if color[1]:
            self.name_color = color[1]
            self.name_color_canvas.delete("all")
            self.name_color_canvas.create_rectangle(0, 0, 29, 23, fill=self.name_color, outline="")
            self.update_preview()
    
    def hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def on_canvas_click(self, event):
        if self.preview_mode.get() != "ticket" or not self.voucher_image:
            return
        
        canvas_height = 240
        ticket_top = self.preview_offset_y
        ticket_bottom = ticket_top + self.preview_ticket_height
        
        if event.y < ticket_top or event.y > ticket_bottom:
            return
        
        relative_y = (event.y - ticket_top) / self.preview_ticket_height - 0.5
        
        if abs(relative_y - self.title_y_pos) < 0.12:
            self.dragging = "title"
            self.drag_start_y = event.y
            self.drag_start_pos = self.title_y_pos
            self.preview_canvas.config(cursor="sb_v_double_arrow")
        elif abs(relative_y - self.name_y_pos) < 0.15:
            self.dragging = "name"
            self.drag_start_y = event.y
            self.drag_start_pos = self.name_y_pos
            self.preview_canvas.config(cursor="sb_v_double_arrow")
    
    def on_canvas_drag(self, event):
        if not self.dragging or self.preview_ticket_height == 0:
            return
        
        delta_y = (event.y - self.drag_start_y) / self.preview_ticket_height
        new_pos = max(-0.45, min(0.45, self.drag_start_pos + delta_y))
        
        if self.dragging == "title":
            self.title_y_pos = new_pos
        else:
            self.name_y_pos = new_pos
        
        self.update_preview()
    
    def on_canvas_release(self, event):
        self.dragging = None
        self.preview_canvas.config(cursor="")
        
    def set_preview_mode(self, mode):
        self.preview_mode.set(mode)
        if mode == "ticket":
            self.ticket_btn.configure(bootstyle="info")
            self.layout_btn.configure(bootstyle="secondary")
        else:
            self.ticket_btn.configure(bootstyle="secondary")
            self.layout_btn.configure(bootstyle="info")
        self.update_preview()
    
    def update_valid_sizes(self):
        page_w, page_h = self.get_page_dimensions()
        page_w_in, page_h_in = page_w / inch, page_h / inch
        
        all_widths = ["1.5", "2", "2.5", "2.75", "3", "3.5", "4", "4.25", "5.5"]
        all_heights = ["1", "1.25", "1.5", "1.75", "2", "2.5", "2.75", "3", "3.5", "4", "5.5"]
        
        valid_widths = [w for w in all_widths if float(w) <= page_w_in]
        valid_heights = [h for h in all_heights if float(h) <= page_h_in]
        
        self.width_combo['values'] = valid_widths
        self.height_combo['values'] = valid_heights
        
        if self.ticket_width_var.get() not in valid_widths and valid_widths:
            self.ticket_width_var.set("3" if "3" in valid_widths else valid_widths[-1])
        if self.ticket_height_var.get() not in valid_heights and valid_heights:
            self.ticket_height_var.set("1.75" if "1.75" in valid_heights else valid_heights[-1])
        
        try:
            cols = int(page_w_in / float(self.ticket_width_var.get()))
            rows = int(page_h_in / float(self.ticket_height_var.get()))
            self.grid_info_label.configure(text=f"({cols}×{rows} grid)")
        except:
            pass
        
    def get_page_dimensions(self):
        if self.orientation_var.get() == "Landscape":
            return landscape(letter)
        return letter
    
    def get_ticket_dimensions(self):
        return float(self.ticket_width_var.get()) * inch, float(self.ticket_height_var.get()) * inch
    
    def calculate_grid(self):
        page_w, page_h = self.get_page_dimensions()
        ticket_w, ticket_h = self.get_ticket_dimensions()
        
        cols = max(1, int(page_w // ticket_w))
        rows = max(1, int(page_h // ticket_h))
        
        tpa = int(self.tickets_per_attendee_var.get())
        
        if self.batch_mode_var.get():
            # Batch mode ON: group tickets by attendee in rows
            rows_per_att = math.ceil(tpa / cols)
            att_per_page = max(1, rows // rows_per_att)
        else:
            # Batch mode OFF: fill entire page, attendees not grouped
            total_slots = cols * rows
            att_per_page = max(1, total_slots // tpa)
            rows_per_att = math.ceil(tpa / cols)  # Still needed for some calculations
        
        return cols, rows, att_per_page, rows_per_att
    
    def calculate_total_pages(self):
        if not self.attendees:
            return 0
        cols, rows, att_per_page, rows_per_att = self.calculate_grid()
        return math.ceil(len(self.attendees) / att_per_page)
    
    def update_calc_display(self):
        cols, rows, att_per_page, rows_per_att = self.calculate_grid()
        tpa = int(self.tickets_per_attendee_var.get())
        
        tw, th = self.ticket_width_var.get(), self.ticket_height_var.get()
        
        if self.blanks_mode.get():
            # Blanks mode: show tickets per page and total
            pages = int(self.blank_pages_var.get())
            tickets_per_page = cols * rows
            total_tickets = tickets_per_page * pages
            self.layout_info_label.configure(text=f"Ticket: {tw}\" × {th}\"  |  {tickets_per_page} per page")
            self.calc_info_label.configure(text=f"{pages} pages × {tickets_per_page} = {total_tickets} blank tickets")
        else:
            # Normal mode
            total_pages = self.calculate_total_pages()
            self.layout_info_label.configure(text=f"Ticket: {tw}\" × {th}\"  |  {tpa} per person")
            
            if self.attendees:
                self.attendee_info_label.configure(text=f"({att_per_page}/page, {total_pages} pages)")
                self.calc_info_label.configure(
                    text=f"{len(self.attendees)} attendees × {tpa} = {len(self.attendees)*tpa} tickets ({total_pages} pages)"
                )
            else:
                self.attendee_info_label.configure(text=f"({att_per_page} attendee(s) per page)")
                self.calc_info_label.configure(text="")
        
    def select_csv(self):
        path = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV", "*.csv"), ("All", "*.*")])
        if path:
            self.csv_path = path
            self.attendees = self.read_attendees(path)
            self.csv_label.configure(text=f"{os.path.basename(path)[:15]} ({len(self.attendees)} attendees)", foreground="")
            self.check_ready()
            self.update_preview()
            
    def select_image(self):
        path = filedialog.askopenfilename(title="Select Image", filetypes=[("Images", "*.jpg *.jpeg *.png *.gif *.bmp"), ("All", "*.*")])
        if path:
            self.image_path = path
            self.img_label.configure(text=os.path.basename(path)[:20], foreground="")
            self.voucher_image = Image.open(path)
            self.image_aspect_ratio = self.voucher_image.width / self.voucher_image.height
            # Auto-fit to image ratio on load
            self.auto_fit_to_image()
            self.check_ready()
            self.update_preview()
    
    def auto_fit_to_image(self):
        """Find best ticket dimensions to match image aspect ratio"""
        if not self.voucher_image:
            return
        
        img_ratio = self.image_aspect_ratio
        page_w, page_h = self.get_page_dimensions()
        page_w_in, page_h_in = page_w / inch, page_h / inch
        
        all_widths = [1.5, 2, 2.5, 2.75, 3, 3.5, 4, 4.25, 5.5]
        all_heights = [1, 1.25, 1.5, 1.75, 2, 2.5, 2.75, 3, 3.5, 4, 5.5]
        
        best_match = None
        best_diff = float('inf')
        
        for w in all_widths:
            if w > page_w_in:
                continue
            for h in all_heights:
                if h > page_h_in:
                    continue
                ticket_ratio = w / h
                diff = abs(ticket_ratio - img_ratio)
                if diff < best_diff:
                    best_diff = diff
                    best_match = (w, h)
        
        if best_match:
            self.ticket_width_var.set(str(best_match[0]))
            self.ticket_height_var.set(str(best_match[1]))
            self.update_valid_sizes()
            
    def check_ready(self):
        if self.blanks_mode.get():
            # Blanks mode: only need image
            if self.image_path:
                self.generate_btn.configure(state="normal")
                self.status_label.configure(text="✓ Ready!", foreground="#28a745")
            else:
                self.generate_btn.configure(state="disabled")
                self.status_label.configure(text="Select an image to get started", foreground="gray")
        else:
            # Normal mode: need CSV and image
            if self.csv_path and self.image_path and self.attendees:
                self.generate_btn.configure(state="normal")
                self.status_label.configure(text="✓ Ready!", foreground="#28a745")
            else:
                self.generate_btn.configure(state="disabled")
                self.status_label.configure(text="Select CSV and image to get started", foreground="gray")
            
    def read_attendees(self, path):
        attendees = []
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                for row in csv.reader(f):
                    if row and row[0].strip():
                        name = f"{row[0].strip()}, {row[1].strip()}" if len(row) >= 2 and row[1].strip() else row[0].strip()
                        attendees.append(name)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV:\n{e}")
        return attendees
    
    def parse_name(self, full_name):
        if ',' in full_name:
            parts = full_name.split(',', 1)
            first = parts[1].strip() if len(parts) > 1 else ""
            last = parts[0].strip()
        else:
            first = full_name.strip()
            last = ""
        
        # Swap if checkbox is checked
        if self.swap_names_var.get():
            first, last = last, first
        
        return first, last
            
    def update_preview(self):
        self.update_calc_display()
        self.check_ready()
        if not self.voucher_image:
            return
        if self.preview_mode.get() == "ticket":
            self.update_ticket_preview()
        else:
            self.update_layout_preview()
    
    def resize_image_to_fill(self, img, tw, th):
        """Resize image to exactly fill target dimensions (stretch to fit)"""
        return img.resize((tw, th), Image.LANCZOS)
    
    def draw_text_with_outline(self, draw, pos, text, font, fill_color, outline=False, underline=False):
        x, y = pos
        if outline:
            for dx in [-2, -1, 0, 1, 2]:
                for dy in [-2, -1, 0, 1, 2]:
                    if dx != 0 or dy != 0:
                        draw.text((x + dx, y + dy), text, font=font, fill="white")
        draw.text((x, y), text, font=font, fill=fill_color)
        
        if underline:
            bbox = draw.textbbox((x, y), text, font=font)
            line_y = bbox[3] + 2  # 2 pixels below text
            draw.line([(bbox[0], line_y), (bbox[2], line_y)], fill=fill_color, width=2)
    
    def get_preview_font(self, size, bold=False):
        """Get PIL font for preview"""
        try:
            font_file = "arialbd.ttf" if bold else "arial.ttf"
            return ImageFont.truetype(font_file, size)
        except:
            # Try Linux fonts
            if bold:
                font_file = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
            else:
                font_file = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
            try:
                return ImageFont.truetype(font_file, size)
            except:
                return ImageFont.load_default()
    
    def update_ticket_preview(self):
        if not self.voucher_image:
            return
        
        try:
            if self.blanks_mode.get():
                # Blanks mode: use extra_text as single line (empty strings if not set)
                extra = self.extra_text_var.get().strip()
                first, last = extra, ""  # Extra text on first line, nothing on second
            else:
                first, last = self.parse_name(self.attendees[0]) if self.attendees else ("First", "Last")
            
            tw_in, th_in = float(self.ticket_width_var.get()), float(self.ticket_height_var.get())
            aspect = tw_in / th_in
            
            canvas_w, canvas_h = 500, 240
            max_w, max_h = 470, 210
            
            if aspect > max_w / max_h:
                pw, ph = max_w, int(max_w / aspect)
            else:
                ph, pw = max_h, int(max_h * aspect)
            
            self.preview_ticket_height = ph
            self.preview_offset_y = (canvas_h - ph) // 2
            
            # Stretch image to fill entire ticket, composite onto white background (matches PDF)
            processed_img = self.get_processed_image()
            stretched = processed_img.resize((pw, ph), Image.LANCZOS)
            ticket = Image.new('RGB', (pw, ph), '#FFFFFF')
            if stretched.mode == 'RGBA':
                ticket.paste(stretched, (0, 0), stretched)  # Use alpha as mask
            else:
                ticket.paste(stretched, (0, 0))
            
            draw = ImageDraw.Draw(ticket)
            
            scale = min(pw / 216, ph / 126)
            
            # Title font
            title_size = max(8, int(int(self.title_font_size_var.get()) * scale * 1.8))
            title_font = self.get_preview_font(title_size, self.title_bold_var.get())
            
            # Name font
            name_size = max(10, int(int(self.name_font_size_var.get()) * scale * 1.8))
            name_font = self.get_preview_font(name_size, self.name_bold_var.get())
            
            cx, cy = pw // 2, ph // 2
            
            # Draw title
            title = self.title_var.get()
            title_bottom = cy  # Track where title ends for name positioning
            if title.strip():
                title_y_center = cy + int(self.title_y_pos * ph)
                # Get bbox at origin to calculate dimensions
                bbox = draw.textbbox((0, 0), title.upper(), font=title_font)
                tw_text = bbox[2] - bbox[0]
                th_text = bbox[3] - bbox[1]
                # Position text centered
                tx = cx - tw_text // 2 - bbox[0]  # Adjust for left offset
                ty = title_y_center - th_text // 2 - bbox[1]  # Adjust for top offset
                self.draw_text_with_outline(draw, (tx, ty), title.upper(), title_font, 
                                           self.title_color, self.title_outline_var.get(),
                                           self.title_underline_var.get())
                # Get actual bbox after positioning
                actual_bbox = draw.textbbox((tx, ty), title.upper(), font=title_font)
                draw.rectangle([actual_bbox[0] - 2, actual_bbox[1] - 2, 
                               actual_bbox[2] + 2, actual_bbox[3] + 2], outline="#2196F3", width=2)
                title_bottom = actual_bbox[3]
            
            # Draw name - First name on top, Last name below (or single line in blanks mode)
            name_y_center = cy + int(self.name_y_pos * ph)
            
            if first and not last:
                # Single line mode (blanks mode with extra text)
                first_bbox_0 = draw.textbbox((0, 0), first, font=name_font)
                first_w = first_bbox_0[2] - first_bbox_0[0]
                first_h = first_bbox_0[3] - first_bbox_0[1]
                
                first_x = cx - first_w // 2 - first_bbox_0[0]
                first_y = name_y_center - first_h // 2 - first_bbox_0[1]
                self.draw_text_with_outline(draw, (first_x, first_y), first, name_font, 
                                            self.name_color, self.name_outline_var.get(),
                                            self.name_underline_var.get())
                first_actual = draw.textbbox((first_x, first_y), first, font=name_font)
                
                # Draw box around single line
                draw.rectangle([first_actual[0] - 2, first_actual[1] - 2, 
                               first_actual[2] + 2, first_actual[3] + 2], outline="#4CAF50", width=2)
            elif first or last:
                # Two line mode (normal mode)
                # Get dimensions for both names
                first_bbox_0 = draw.textbbox((0, 0), first, font=name_font)
                first_w = first_bbox_0[2] - first_bbox_0[0]
                first_h = first_bbox_0[3] - first_bbox_0[1]
                
                last_bbox_0 = draw.textbbox((0, 0), last, font=name_font)
                last_w = last_bbox_0[2] - last_bbox_0[0]
                last_h = last_bbox_0[3] - last_bbox_0[1]
                
                # Total height with gap
                line_gap = int(4 * scale)
                total_name_height = first_h + line_gap + last_h
                
                # Position first name
                first_x = cx - first_w // 2 - first_bbox_0[0]
                first_y = name_y_center - total_name_height // 2 - first_bbox_0[1]
                self.draw_text_with_outline(draw, (first_x, first_y), first, name_font, 
                                            self.name_color, self.name_outline_var.get(),
                                            self.name_underline_var.get())
                first_actual = draw.textbbox((first_x, first_y), first, font=name_font)
                
                # Position last name
                last_x = cx - last_w // 2 - last_bbox_0[0]
                last_y = first_actual[3] + line_gap - last_bbox_0[1]
                self.draw_text_with_outline(draw, (last_x, last_y), last, name_font, 
                                            self.name_color, self.name_outline_var.get(),
                                            self.name_underline_var.get())
                last_actual = draw.textbbox((last_x, last_y), last, font=name_font)
                
                # Draw box around both name lines
                box_left = min(first_actual[0], last_actual[0]) - 2
                box_right = max(first_actual[2], last_actual[2]) + 2
                draw.rectangle([box_left, first_actual[1] - 2, box_right, last_actual[3] + 2], outline="#4CAF50", width=2)
            
            self.preview_photo = ImageTk.PhotoImage(ticket)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image((canvas_w - pw)//2, (canvas_h - ph)//2, anchor=tk.NW, image=self.preview_photo)
            
            # Legend
            self.preview_canvas.create_rectangle(10, canvas_h-20, 20, canvas_h-10, fill="#2196F3", outline="#2196F3")
            self.preview_canvas.create_text(24, canvas_h-15, anchor=tk.W, text="Title", fill="#333", font=("Arial", 8))
            self.preview_canvas.create_rectangle(60, canvas_h-20, 70, canvas_h-10, fill="#4CAF50", outline="#4CAF50")
            legend_text = "Extra" if self.blanks_mode.get() else "Name"
            self.preview_canvas.create_text(74, canvas_h-15, anchor=tk.W, text=legend_text, fill="#333", font=("Arial", 8))
            
        except Exception as e:
            print(f"Preview error: {e}")
            traceback.print_exc()
    
    def update_layout_preview(self):
        try:
            page_w, page_h = self.get_page_dimensions()
            cols, rows, att_per_page, rows_per_att = self.calculate_grid()
            tpa = int(self.tickets_per_attendee_var.get())
            ticket_w, ticket_h = self.get_ticket_dimensions()
            
            canvas_w, canvas_h = 500, 240
            scale = min((canvas_w - 40) / page_w, (canvas_h - 40) / page_h)
            
            pw, ph = int(page_w * scale), int(page_h * scale)
            tw, th = int(ticket_w * scale), int(ticket_h * scale)
            
            page = Image.new('RGB', (pw, ph), 'white')
            draw = ImageDraw.Draw(page)
            
            try:
                font = ImageFont.truetype("arial.ttf", 8)
            except:
                font = ImageFont.load_default()
            
            mini = None
            if self.voucher_image and tw > 10 and th > 10:
                processed_img = self.get_processed_image()
                stretched = processed_img.resize((tw, th), Image.LANCZOS)
                mini = Image.new('RGB', (tw, th), '#FFFFFF')
                if stretched.mode == 'RGBA':
                    mini.paste(stretched, (0, 0), stretched)
                else:
                    mini.paste(stretched, (0, 0))
            
            gw, gh = cols * tw, rows * th
            
            if self.align_top_left_var.get():
                ox, oy = 0, 0
            else:
                ox, oy = (pw - gw) // 2, (ph - gh) // 2
            
            if self.blanks_mode.get():
                # Blanks mode: all tickets are the same (no attendee numbers)
                for row in range(rows):
                    for col in range(cols):
                        x, y = ox + col * tw, oy + row * th
                        if mini:
                            page.paste(mini, (x, y))
                        draw.rectangle([x, y, x+tw-1, y+th-1], outline='#999')
            elif self.batch_mode_var.get():
                # Batch mode ON: group tickets by attendee
                for row in range(rows):
                    att_row = row // rows_per_att
                    row_in_att = row % rows_per_att
                    curr_att = att_row + 1
                    
                    if curr_att > att_per_page:
                        break
                    
                    for col in range(cols):
                        x, y = ox + col * tw, oy + row * th
                        ticket_idx = row_in_att * cols + col
                        
                        if ticket_idx < tpa:
                            if mini:
                                page.paste(mini, (x, y))
                            draw.rectangle([x, y, x+tw-1, y+th-1], outline='#999')
                            label = f"{curr_att}"
                            # Center of ticket
                            cx, cy = x + tw // 2, y + th // 2
                            # Get text size for background box
                            bbox = draw.textbbox((cx, cy), label, font=font, anchor='mm')
                            draw.rectangle([bbox[0] - 2, bbox[1] - 2, bbox[2] + 2, bbox[3] + 2], 
                                          fill='white', outline='#666')
                            draw.text((cx, cy), label, fill='#333', font=font, anchor='mm')
                        else:
                            draw.rectangle([x, y, x+tw-1, y+th-1], fill='#e0e0e0', outline='#ccc')
            else:
                # Batch mode OFF: fill page sequentially
                ticket_count = 0
                total_tickets_on_page = att_per_page * tpa
                
                for row in range(rows):
                    for col in range(cols):
                        if ticket_count >= total_tickets_on_page:
                            break
                        
                        x, y = ox + col * tw, oy + row * th
                        curr_att = (ticket_count // tpa) + 1
                        
                        if mini:
                            page.paste(mini, (x, y))
                        draw.rectangle([x, y, x+tw-1, y+th-1], outline='#999')
                        label = f"{curr_att}"
                        # Center of ticket
                        cx, cy = x + tw // 2, y + th // 2
                        # Get text size for background box
                        bbox = draw.textbbox((cx, cy), label, font=font, anchor='mm')
                        draw.rectangle([bbox[0] - 2, bbox[1] - 2, bbox[2] + 2, bbox[3] + 2], 
                                      fill='white', outline='#666')
                        draw.text((cx, cy), label, fill='#333', font=font, anchor='mm')
                        
                        ticket_count += 1
            
            draw.rectangle([0, 0, pw-1, ph-1], outline='#333', width=2)
            
            self.preview_photo = ImageTk.PhotoImage(page)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image((canvas_w-pw)//2, (canvas_h-ph)//2, anchor=tk.NW, image=self.preview_photo)
            
        except Exception as e:
            print(f"Layout error: {e}")
            traceback.print_exc()
    
    def generate_pdf(self):
        if self.blanks_mode.get():
            # Blanks mode: only need image
            if not self.image_path:
                messagebox.showwarning("Missing", "Select an image first.")
                return
        else:
            # Normal mode: need CSV and image
            if not self.csv_path or not self.image_path or not self.attendees:
                messagebox.showwarning("Missing", "Select CSV and image first.")
                return
        
        default_name = "blank_vouchers.pdf" if self.blanks_mode.get() else "vouchers.pdf"
        output = filedialog.asksaveasfilename(title="Save PDF", defaultextension=".pdf", 
                                               filetypes=[("PDF", "*.pdf")], initialfile=default_name)
        if not output:
            return
        
        self.status_label.configure(text="Generating PDF...", foreground="#17a2b8")
        self.root.update()
        
        try:
            if self.blanks_mode.get():
                self.create_blanks_pdf(output)
                pages = int(self.blank_pages_var.get())
                cols, rows, _, _ = self.calculate_grid()
                total_tickets = cols * rows * pages
                self.status_label.configure(text=f"✓ Created {total_tickets} blank vouchers on {pages} pages!", foreground="#28a745")
                messagebox.showinfo("Success", f"Created {total_tickets} blank vouchers!\n{pages} pages\n\nSaved to:\n{output}")
            else:
                self.create_pdf(output)
                tpa = int(self.tickets_per_attendee_var.get())
                total_pages = self.calculate_total_pages()
                self.status_label.configure(text=f"✓ Created {len(self.attendees)*tpa} vouchers on {total_pages} pages!", foreground="#28a745")
                messagebox.showinfo("Success", f"Created {len(self.attendees)*tpa} vouchers!\n{total_pages} pages\n\nSaved to:\n{output}")
        except Exception as e:
            self.status_label.configure(text="Error creating PDF", foreground="#dc3545")
            messagebox.showerror("Error", f"Could not create PDF:\n{e}")
            traceback.print_exc()
            
    def create_pdf(self, output):
        page_w, page_h = self.get_page_dimensions()
        ticket_w, ticket_h = self.get_ticket_dimensions()
        cols, rows, att_per_page, rows_per_att = self.calculate_grid()
        tpa = int(self.tickets_per_attendee_var.get())
        
        gw, gh = cols * ticket_w, rows * ticket_h
        
        if self.align_top_left_var.get():
            ox, oy = 0, 0
        else:
            ox, oy = (page_w - gw) / 2, (page_h - gh) / 2
        
        # Prepare image - stretch to fill exact dimensions, composite onto white
        dpi = 3
        iw, ih = int(ticket_w * dpi), int(ticket_h * dpi)
        processed_img = self.get_processed_image()
        stretched = processed_img.resize((iw, ih), Image.LANCZOS)
        ticket_img = Image.new('RGB', (iw, ih), '#FFFFFF')
        if stretched.mode == 'RGBA':
            ticket_img.paste(stretched, (0, 0), stretched)
        else:
            ticket_img.paste(stretched, (0, 0))
        
        temp = os.path.join(os.path.dirname(output), "_temp.png")
        ticket_img.save(temp, "PNG")
        img_reader = ImageReader(temp)
        
        title_rgb = self.hex_to_rgb(self.title_color)
        name_rgb = self.hex_to_rgb(self.name_color)
        
        c = canvas.Canvas(output, pagesize=(page_w, page_h))
        
        def draw_ticket(x, y, first, last):
            """Helper to draw a single ticket at position x, y"""
            c.drawImage(img_reader, x, y, width=ticket_w, height=ticket_h, mask='auto')
            
            cx, cy = x + ticket_w/2, y + ticket_h/2
            size_factor = min(ticket_w / (3*inch), ticket_h / (1.75*inch))
            
            # Title
            title = self.title_var.get().strip()
            if title:
                title_size = max(6, int(int(self.title_font_size_var.get()) * size_factor * 1.8))
                font_name = "Helvetica-Bold" if self.title_bold_var.get() else "Helvetica"
                c.setFont(font_name, title_size)
                
                # PDF Y is bottom-up, title_y_pos negative = above center = higher Y in PDF
                title_y = cy - (self.title_y_pos * ticket_h) - title_size * 0.35
                
                if self.title_outline_var.get():
                    c.setFillColorRGB(1, 1, 1)
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if dx or dy:
                                c.drawCentredString(cx + dx, title_y + dy, title.upper())
                
                c.setFillColorRGB(title_rgb[0]/255, title_rgb[1]/255, title_rgb[2]/255)
                c.drawCentredString(cx, title_y, title.upper())
                
                # Title underline
                if self.title_underline_var.get():
                    title_width = c.stringWidth(title.upper(), font_name, title_size)
                    c.setStrokeColorRGB(title_rgb[0]/255, title_rgb[1]/255, title_rgb[2]/255)
                    c.setLineWidth(1)
                    c.line(cx - title_width/2, title_y - 2, cx + title_width/2, title_y - 2)
            
            # Name - First above Last
            name_size = max(6, int(int(self.name_font_size_var.get()) * size_factor * 1.8))
            font_name = "Helvetica-Bold" if self.name_bold_var.get() else "Helvetica"
            c.setFont(font_name, name_size)
            
            # PDF Y is bottom-up, name_y_pos positive = below center in preview = lower Y in PDF
            name_y_center = cy - (self.name_y_pos * ticket_h)
            
            # Small gap between lines
            line_gap = name_size * 0.15
            
            # Position: first name above, last name below - tighter spacing
            first_y = name_y_center + line_gap / 2 + name_size * 0.15
            last_y = name_y_center - line_gap / 2 - name_size * 0.65
            
            if self.name_outline_var.get():
                c.setFillColorRGB(1, 1, 1)
                for dx in [-1, 0, 1]:
                    for dy in [-1, 0, 1]:
                        if dx or dy:
                            c.drawCentredString(cx + dx, first_y + dy, first)
                            c.drawCentredString(cx + dx, last_y + dy, last)
            
            c.setFillColorRGB(name_rgb[0]/255, name_rgb[1]/255, name_rgb[2]/255)
            c.drawCentredString(cx, first_y, first)
            c.drawCentredString(cx, last_y, last)
            
            # Name underlines
            if self.name_underline_var.get():
                c.setStrokeColorRGB(name_rgb[0]/255, name_rgb[1]/255, name_rgb[2]/255)
                c.setLineWidth(1)
                first_width = c.stringWidth(first, font_name, name_size)
                last_width = c.stringWidth(last, font_name, name_size)
                c.line(cx - first_width/2, first_y - 2, cx + first_width/2, first_y - 2)
                c.line(cx - last_width/2, last_y - 2, cx + last_width/2, last_y - 2)
        
        def draw_cutting_guides():
            """Draw dotted cutting lines between tickets"""
            if not self.cutting_guides_var.get():
                return
            
            c.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray color
            c.setLineWidth(0.5)
            c.setDash(3, 3)  # Dotted line pattern
            
            # Vertical lines between columns
            for col in range(cols + 1):
                x = ox + col * ticket_w
                y_start = page_h - oy - rows * ticket_h
                y_end = page_h - oy
                c.line(x, y_start, x, y_end)
            
            # Horizontal lines between rows
            for row in range(rows + 1):
                y = page_h - oy - row * ticket_h
                x_start = ox
                x_end = ox + cols * ticket_w
                c.line(x_start, y, x_end, y)
            
            c.setDash()  # Reset to solid line
        
        if self.batch_mode_var.get():
            # Batch mode ON: group tickets by attendee
            idx = 0
            while idx < len(self.attendees):
                for page_att in range(att_per_page):
                    if idx >= len(self.attendees):
                        break
                    
                    first, last = self.parse_name(self.attendees[idx])
                    start_row = page_att * rows_per_att
                    count = 0
                    
                    for row_off in range(rows_per_att):
                        if count >= tpa:
                            break
                        for col in range(cols):
                            if count >= tpa:
                                break
                            
                            row = start_row + row_off
                            x = ox + col * ticket_w
                            y = page_h - oy - (row + 1) * ticket_h
                            
                            draw_ticket(x, y, first, last)
                            count += 1
                    
                    idx += 1
                
                draw_cutting_guides()
                if idx < len(self.attendees):
                    c.showPage()
        else:
            # Batch mode OFF: fill pages sequentially
            idx = 0  # Current attendee index
            ticket_for_attendee = 0  # How many tickets placed for current attendee
            
            while idx < len(self.attendees):
                # Fill one page
                tickets_on_page = 0
                max_tickets_on_page = att_per_page * tpa
                
                for row in range(rows):
                    for col in range(cols):
                        if tickets_on_page >= max_tickets_on_page or idx >= len(self.attendees):
                            break
                        
                        first, last = self.parse_name(self.attendees[idx])
                        x = ox + col * ticket_w
                        y = page_h - oy - (row + 1) * ticket_h
                        
                        draw_ticket(x, y, first, last)
                        
                        ticket_for_attendee += 1
                        tickets_on_page += 1
                        
                        if ticket_for_attendee >= tpa:
                            idx += 1
                            ticket_for_attendee = 0
                    
                    if tickets_on_page >= max_tickets_on_page or idx >= len(self.attendees):
                        break
                
                draw_cutting_guides()
                if idx < len(self.attendees):
                    c.showPage()
        
        c.save()
        try:
            os.remove(temp)
        except:
            pass
    
    def create_blanks_pdf(self, output):
        """Generate PDF with blank tickets (no names, just extra text if provided)"""
        page_w, page_h = self.get_page_dimensions()
        ticket_w, ticket_h = self.get_ticket_dimensions()
        cols, rows, _, _ = self.calculate_grid()
        pages = int(self.blank_pages_var.get())
        
        gw, gh = cols * ticket_w, rows * ticket_h
        
        if self.align_top_left_var.get():
            ox, oy = 0, 0
        else:
            ox, oy = (page_w - gw) / 2, (page_h - gh) / 2
        
        # Prepare image
        dpi = 3
        iw, ih = int(ticket_w * dpi), int(ticket_h * dpi)
        processed_img = self.get_processed_image()
        stretched = processed_img.resize((iw, ih), Image.LANCZOS)
        final_img = Image.new('RGB', (iw, ih), '#FFFFFF')
        if stretched.mode == 'RGBA':
            final_img.paste(stretched, (0, 0), stretched)
        else:
            final_img.paste(stretched, (0, 0))
        
        temp = output + ".tmp.png"
        final_img.save(temp, dpi=(72*dpi, 72*dpi))
        img_reader = ImageReader(temp)
        
        c = canvas.Canvas(output, pagesize=(page_w, page_h))
        
        title_rgb = self.hex_to_rgb(self.title_color)
        extra_rgb = self.hex_to_rgb(self.name_color)  # Extra text uses "name" color settings
        
        extra_text = self.extra_text_var.get().strip()
        
        def draw_blank_ticket(x, y):
            """Draw a single blank ticket at position x, y"""
            c.drawImage(img_reader, x, y, width=ticket_w, height=ticket_h, mask='auto')
            
            cx, cy = x + ticket_w/2, y + ticket_h/2
            size_factor = min(ticket_w / (3*inch), ticket_h / (1.75*inch))
            
            # Title
            title = self.title_var.get().strip()
            if title:
                title_size = max(6, int(int(self.title_font_size_var.get()) * size_factor * 1.8))
                font_name = "Helvetica-Bold" if self.title_bold_var.get() else "Helvetica"
                c.setFont(font_name, title_size)
                
                title_y = cy - (self.title_y_pos * ticket_h) - title_size * 0.35
                
                if self.title_outline_var.get():
                    c.setFillColorRGB(1, 1, 1)
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if dx or dy:
                                c.drawCentredString(cx + dx, title_y + dy, title.upper())
                
                c.setFillColorRGB(title_rgb[0]/255, title_rgb[1]/255, title_rgb[2]/255)
                c.drawCentredString(cx, title_y, title.upper())
                
                if self.title_underline_var.get():
                    title_width = c.stringWidth(title.upper(), font_name, title_size)
                    c.setStrokeColorRGB(title_rgb[0]/255, title_rgb[1]/255, title_rgb[2]/255)
                    c.setLineWidth(1)
                    c.line(cx - title_width/2, title_y - 2, cx + title_width/2, title_y - 2)
            
            # Extra text (single line, uses "name/extra" settings)
            if extra_text:
                extra_size = max(6, int(int(self.name_font_size_var.get()) * size_factor * 1.8))
                font_name = "Helvetica-Bold" if self.name_bold_var.get() else "Helvetica"
                c.setFont(font_name, extra_size)
                
                extra_y = cy - (self.name_y_pos * ticket_h) - extra_size * 0.35
                
                if self.name_outline_var.get():
                    c.setFillColorRGB(1, 1, 1)
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if dx or dy:
                                c.drawCentredString(cx + dx, extra_y + dy, extra_text)
                
                c.setFillColorRGB(extra_rgb[0]/255, extra_rgb[1]/255, extra_rgb[2]/255)
                c.drawCentredString(cx, extra_y, extra_text)
                
                if self.name_underline_var.get():
                    extra_width = c.stringWidth(extra_text, font_name, extra_size)
                    c.setStrokeColorRGB(extra_rgb[0]/255, extra_rgb[1]/255, extra_rgb[2]/255)
                    c.setLineWidth(1)
                    c.line(cx - extra_width/2, extra_y - 2, cx + extra_width/2, extra_y - 2)
        
        def draw_cutting_guides():
            """Draw dotted cutting lines between tickets"""
            if not self.cutting_guides_var.get():
                return
            
            c.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray color
            c.setLineWidth(0.5)
            c.setDash(3, 3)  # Dotted line pattern
            
            # Vertical lines between columns
            for col in range(cols + 1):
                x = ox + col * ticket_w
                y_start = page_h - oy - rows * ticket_h
                y_end = page_h - oy
                c.line(x, y_start, x, y_end)
            
            # Horizontal lines between rows
            for row in range(rows + 1):
                y = page_h - oy - row * ticket_h
                x_start = ox
                x_end = ox + cols * ticket_w
                c.line(x_start, y, x_end, y)
            
            c.setDash()  # Reset to solid line
        
        # Generate all pages
        for page in range(pages):
            if page > 0:
                c.showPage()
            
            for row in range(rows):
                for col in range(cols):
                    x = ox + col * ticket_w
                    y = page_h - oy - (row + 1) * ticket_h
                    draw_blank_ticket(x, y)
            
            draw_cutting_guides()
        
        c.save()
        try:
            os.remove(temp)
        except:
            pass


def main():
    root = ttk.Window(themename="flatly")
    app = VoucherGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
