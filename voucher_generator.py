#!/usr/bin/env python3
"""
Drink Voucher Generator
GUI app with live preview for creating personalized drink vouchers.
Repository: https://github.com/riconanci/TicketGen
"""

import csv
import os
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import traceback


class VoucherGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drink Voucher Generator")
        self.root.geometry("720x860")
        self.root.resizable(False, False)
        
        # Variables
        self.csv_path = None
        self.image_path = None
        self.attendees = []
        self.voucher_image = None
        self.image_aspect_ratio = 1.71  # Default ratio
        
        # Title settings
        self.title_var = tk.StringVar(value="DRINK VOUCHER")
        self.title_font_size_var = tk.StringVar(value="10")
        self.title_bold_var = tk.IntVar(value=1)
        self.title_color = "#33261A"
        self.title_outline_var = tk.IntVar(value=0)
        
        # Name settings
        self.name_font_size_var = tk.StringVar(value="14")
        self.name_bold_var = tk.IntVar(value=1)
        self.name_color = "#33261A"
        self.name_outline_var = tk.IntVar(value=0)
        self.swap_names_var = tk.IntVar(value=0)  # Swap first/last name order
        
        # Layout variables
        self.orientation_var = tk.StringVar(value="Portrait")
        self.ticket_width_var = tk.StringVar(value="3")
        self.ticket_height_var = tk.StringVar(value="1.75")
        self.tickets_per_attendee_var = tk.StringVar(value="5")
        self.align_top_left_var = tk.IntVar(value=0)
        
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
        
        # === MENU BAR ===
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        
        # === FILE SELECTION ===
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Select Files", padding="6")
        file_frame.pack(fill=tk.X, pady=(0, 4))
        
        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X)
        
        csv_btn = tk.Button(file_row, text="Select CSV", command=self.select_csv, bg="#4CAF50", fg="white", padx=8)
        csv_btn.pack(side=tk.LEFT)
        self.csv_label = ttk.Label(file_row, text="No file", foreground="gray")
        self.csv_label.pack(side=tk.LEFT, padx=(5, 15))
        
        img_btn = tk.Button(file_row, text="Select Image", command=self.select_image, bg="#2196F3", fg="white", padx=8)
        img_btn.pack(side=tk.LEFT)
        self.img_label = ttk.Label(file_row, text="No file", foreground="gray")
        self.img_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # === PREVIEW ===
        preview_frame = ttk.LabelFrame(main_frame, text="Preview (drag title/name to reposition)", padding="4")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        
        toggle_row = ttk.Frame(preview_frame)
        toggle_row.pack(pady=(0, 2))
        
        self.ticket_btn = tk.Button(toggle_row, text="Single Ticket", command=lambda: self.set_preview_mode("ticket"),
                                     bg="#2196F3", fg="white", padx=12, pady=2)
        self.ticket_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.layout_btn = tk.Button(toggle_row, text="Page Layout", command=lambda: self.set_preview_mode("layout"),
                                     bg="#cccccc", fg="black", padx=12, pady=2)
        self.layout_btn.pack(side=tk.LEFT)
        
        self.preview_canvas = tk.Canvas(preview_frame, width=500, height=240, bg="white", relief="sunken", bd=2)
        self.preview_canvas.pack(pady=2)
        self.preview_canvas.create_text(250, 120, text="Select an image to see preview", fill="gray")
        
        self.preview_canvas.bind("<Button-1>", self.on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        self.layout_info_label = ttk.Label(preview_frame, text="", foreground="blue")
        self.layout_info_label.pack()
        
        # === TEXT SETTINGS ===
        text_frame = ttk.LabelFrame(main_frame, text="Step 2: Text Settings", padding="6")
        text_frame.pack(fill=tk.X, pady=(0, 4))
        
        # Header row
        header_row = ttk.Frame(text_frame)
        header_row.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(header_row, text="", width=6).pack(side=tk.LEFT)
        ttk.Label(header_row, text="", width=18).pack(side=tk.LEFT)  # Spacer for entry/label
        ttk.Label(header_row, text="Size", width=6).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(header_row, text="Bold", width=5).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(header_row, text="Color", width=5).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Label(header_row, text="Outline", width=8).pack(side=tk.LEFT, padx=(8, 0))
        
        # Title row
        title_row = ttk.Frame(text_frame)
        title_row.pack(fill=tk.X, pady=2)
        
        ttk.Label(title_row, text="Title:", width=6).pack(side=tk.LEFT)
        title_entry = ttk.Entry(title_row, textvariable=self.title_var, width=18)
        title_entry.pack(side=tk.LEFT)
        title_entry.bind('<KeyRelease>', lambda e: self.on_step2_interact())
        title_entry.bind('<FocusIn>', lambda e: self.on_step2_interact())
        
        title_size = ttk.Combobox(title_row, textvariable=self.title_font_size_var, 
                                   values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24"], width=4, state="readonly")
        title_size.pack(side=tk.LEFT, padx=(8, 0))
        title_size.bind('<<ComboboxSelected>>', lambda e: self.on_step2_interact())
        title_size.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.title_bold_check = tk.Checkbutton(title_row, text="", variable=self.title_bold_var, 
                                                command=self.on_step2_interact)
        self.title_bold_check.pack(side=tk.LEFT, padx=(12, 0))
        self.title_bold_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.title_color_btn = tk.Button(title_row, text="  ", command=self.pick_title_color, 
                                          bg=self.title_color, width=3)
        self.title_color_btn.pack(side=tk.LEFT, padx=(12, 0))
        
        self.title_outline_check = tk.Checkbutton(title_row, text="", variable=self.title_outline_var, 
                                                   command=self.on_step2_interact)
        self.title_outline_check.pack(side=tk.LEFT, padx=(16, 0))
        self.title_outline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Name row
        name_row = ttk.Frame(text_frame)
        name_row.pack(fill=tk.X, pady=2)
        
        ttk.Label(name_row, text="Name:", width=6).pack(side=tk.LEFT)
        ttk.Label(name_row, text="(from CSV)", foreground="gray", width=18, anchor='w').pack(side=tk.LEFT)
        
        name_size = ttk.Combobox(name_row, textvariable=self.name_font_size_var, 
                                  values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24"], width=4, state="readonly")
        name_size.pack(side=tk.LEFT, padx=(8, 0))
        name_size.bind('<<ComboboxSelected>>', lambda e: self.on_step2_interact())
        name_size.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.name_bold_check = tk.Checkbutton(name_row, text="", variable=self.name_bold_var, 
                                               command=self.on_step2_interact)
        self.name_bold_check.pack(side=tk.LEFT, padx=(12, 0))
        self.name_bold_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        self.name_color_btn = tk.Button(name_row, text="  ", command=self.pick_name_color, 
                                         bg=self.name_color, width=3)
        self.name_color_btn.pack(side=tk.LEFT, padx=(12, 0))
        
        self.name_outline_check = tk.Checkbutton(name_row, text="", variable=self.name_outline_var, 
                                                  command=self.on_step2_interact)
        self.name_outline_check.pack(side=tk.LEFT, padx=(16, 0))
        self.name_outline_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # Swap names row
        swap_row = ttk.Frame(text_frame)
        swap_row.pack(fill=tk.X, pady=2)
        ttk.Label(swap_row, text="", width=6).pack(side=tk.LEFT)
        self.swap_names_check = tk.Checkbutton(swap_row, text="Swap First/Last name order", 
                                                variable=self.swap_names_var, command=self.on_step2_interact)
        self.swap_names_check.pack(side=tk.LEFT)
        self.swap_names_check.bind('<Button-1>', lambda e: self.set_preview_mode("ticket"))
        
        # === LAYOUT SETTINGS ===
        layout_frame = ttk.LabelFrame(main_frame, text="Step 3: Ticket Size & Layout", padding="6")
        layout_frame.pack(fill=tk.X, pady=(0, 4))
        
        # Orientation row
        orient_row = ttk.Frame(layout_frame)
        orient_row.pack(fill=tk.X, pady=2)
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
        ttk.Label(tpa_row, text="Tickets:", width=10).pack(side=tk.LEFT)
        self.tpa_combo = ttk.Combobox(tpa_row, textvariable=self.tickets_per_attendee_var,
                                       values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], 
                                       width=5, state="readonly")
        self.tpa_combo.pack(side=tk.LEFT)
        self.tpa_combo.bind('<<ComboboxSelected>>', lambda e: self.on_step3_interact())
        self.tpa_combo.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        self.attendee_info_label = ttk.Label(tpa_row, text="", foreground="blue")
        self.attendee_info_label.pack(side=tk.LEFT, padx=(15, 0))
        
        # Align top-left checkbox
        self.align_check = tk.Checkbutton(tpa_row, text="Align Top-Left (easier cutting)", 
                                           variable=self.align_top_left_var, command=self.on_align_change)
        self.align_check.pack(side=tk.LEFT, padx=(20, 0))
        self.align_check.bind('<Button-1>', lambda e: self.set_preview_mode("layout"))
        
        # Size row (MOVED AFTER tickets)
        size_row = ttk.Frame(layout_frame)
        size_row.pack(fill=tk.X, pady=2)
        
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
        
        self.grid_info_label = ttk.Label(size_row, text="", foreground="blue")
        self.grid_info_label.pack(side=tk.LEFT)
        
        # === GENERATE ===
        generate_frame = ttk.Frame(main_frame)
        generate_frame.pack(fill=tk.X, pady=6)
        
        self.generate_btn = tk.Button(generate_frame, text="GENERATE PDF", command=self.generate_pdf,
                                       bg="#FF5722", fg="white", font=("Arial", 14, "bold"),
                                       padx=30, pady=10, state=tk.DISABLED)
        self.generate_btn.pack()
        
        self.status_label = ttk.Label(generate_frame, text="Select CSV and image to get started", foreground="gray")
        self.status_label.pack(pady=(4, 0))
    
    def on_step2_interact(self):
        self.set_preview_mode("ticket")
        self.update_preview()
    
    def show_about(self):
        """Show About dialog"""
        about_window = tk.Toplevel(self.root)
        about_window.title("About")
        about_window.geometry("280x150")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Center on parent
        about_window.geometry(f"+{self.root.winfo_x() + 200}+{self.root.winfo_y() + 200}")
        
        frame = ttk.Frame(about_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Ticket Generator", font=("Arial", 16, "bold")).pack(pady=(0, 5))
        ttk.Label(frame, text="Version 1.0", font=("Arial", 11)).pack(pady=(0, 15))
        ttk.Label(frame, text="made for Alyssa with love - R", font=("Arial", 8), foreground="gray").pack(pady=(0, 15))
        
        tk.Button(frame, text="Close", command=about_window.destroy, padx=15).pack()
    
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
    
    def pick_title_color(self):
        self.set_preview_mode("ticket")
        color = colorchooser.askcolor(color=self.title_color, title="Choose Title Color")
        if color[1]:
            self.title_color = color[1]
            self.title_color_btn.config(bg=self.title_color)
            self.update_preview()
    
    def pick_name_color(self):
        self.set_preview_mode("ticket")
        color = colorchooser.askcolor(color=self.name_color, title="Choose Name Color")
        if color[1]:
            self.name_color = color[1]
            self.name_color_btn.config(bg=self.name_color)
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
            self.ticket_btn.config(bg="#2196F3", fg="white")
            self.layout_btn.config(bg="#cccccc", fg="black")
        else:
            self.ticket_btn.config(bg="#cccccc", fg="black")
            self.layout_btn.config(bg="#2196F3", fg="white")
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
            self.grid_info_label.config(text=f"({cols}×{rows} grid)")
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
        rows_per_att = math.ceil(tpa / cols)
        att_per_page = max(1, rows // rows_per_att)
        
        return cols, rows, att_per_page, rows_per_att
    
    def calculate_total_pages(self):
        if not self.attendees:
            return 0
        cols, rows, att_per_page, rows_per_att = self.calculate_grid()
        return math.ceil(len(self.attendees) / att_per_page)
    
    def update_calc_display(self):
        cols, rows, att_per_page, rows_per_att = self.calculate_grid()
        tpa = int(self.tickets_per_attendee_var.get())
        total_pages = self.calculate_total_pages()
        
        tw, th = self.ticket_width_var.get(), self.ticket_height_var.get()
        self.layout_info_label.config(text=f"Ticket: {tw}\" × {th}\"  |  {tpa} per person")
        
        if self.attendees:
            self.attendee_info_label.config(text=f"({att_per_page}/page, {total_pages} pages)")
        else:
            self.attendee_info_label.config(text=f"({att_per_page} attendee(s) per page)")
        
    def select_csv(self):
        path = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV", "*.csv"), ("All", "*.*")])
        if path:
            self.csv_path = path
            self.attendees = self.read_attendees(path)
            self.csv_label.config(text=f"{os.path.basename(path)[:15]} ({len(self.attendees)} attendees)", foreground="black")
            self.check_ready()
            self.update_preview()
            
    def select_image(self):
        path = filedialog.askopenfilename(title="Select Image", filetypes=[("Images", "*.jpg *.jpeg *.png *.gif *.bmp"), ("All", "*.*")])
        if path:
            self.image_path = path
            self.img_label.config(text=os.path.basename(path)[:20], foreground="black")
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
        if self.csv_path and self.image_path and self.attendees:
            self.generate_btn.config(state=tk.NORMAL)
            tpa = int(self.tickets_per_attendee_var.get())
            total_pages = self.calculate_total_pages()
            self.status_label.config(
                text=f"✓ Ready! {len(self.attendees)} attendees × {tpa} = {len(self.attendees)*tpa} tickets ({total_pages} pages)", 
                foreground="green"
            )
        else:
            self.generate_btn.config(state=tk.DISABLED)
            
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
    
    def draw_text_with_outline(self, draw, pos, text, font, fill_color, outline=False):
        x, y = pos
        if outline:
            for dx in [-2, -1, 0, 1, 2]:
                for dy in [-2, -1, 0, 1, 2]:
                    if dx != 0 or dy != 0:
                        draw.text((x + dx, y + dy), text, font=font, fill="white")
        draw.text((x, y), text, font=font, fill=fill_color)
    
    def update_ticket_preview(self):
        if not self.voucher_image:
            return
        
        try:
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
            stretched = self.voucher_image.resize((pw, ph), Image.LANCZOS)
            ticket = Image.new('RGB', (pw, ph), '#FFFFFF')
            if stretched.mode == 'RGBA':
                ticket.paste(stretched, (0, 0), stretched)  # Use alpha as mask
            else:
                ticket.paste(stretched, (0, 0))
            
            draw = ImageDraw.Draw(ticket)
            
            scale = min(pw / 216, ph / 126)
            
            # Title font
            title_size = max(8, int(int(self.title_font_size_var.get()) * scale * 1.8))
            try:
                title_font = ImageFont.truetype("arialbd.ttf" if self.title_bold_var.get() else "arial.ttf", title_size)
            except:
                try:
                    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if self.title_bold_var.get() else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
                    title_font = ImageFont.truetype(font_path, title_size)
                except:
                    title_font = ImageFont.load_default()
            
            # Name font
            name_size = max(10, int(int(self.name_font_size_var.get()) * scale * 1.8))
            try:
                name_font = ImageFont.truetype("arialbd.ttf" if self.name_bold_var.get() else "arial.ttf", name_size)
            except:
                try:
                    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if self.name_bold_var.get() else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
                    name_font = ImageFont.truetype(font_path, name_size)
                except:
                    name_font = ImageFont.load_default()
            
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
                                           self.title_color, self.title_outline_var.get())
                # Get actual bbox after positioning
                actual_bbox = draw.textbbox((tx, ty), title.upper(), font=title_font)
                draw.rectangle([actual_bbox[0] - 2, actual_bbox[1] - 2, 
                               actual_bbox[2] + 2, actual_bbox[3] + 2], outline="#2196F3", width=2)
                title_bottom = actual_bbox[3]
            
            # Draw name - First name on top, Last name below
            name_y_center = cy + int(self.name_y_pos * ph)
            
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
                                        self.name_color, self.name_outline_var.get())
            first_actual = draw.textbbox((first_x, first_y), first, font=name_font)
            
            # Position last name
            last_x = cx - last_w // 2 - last_bbox_0[0]
            last_y = first_actual[3] + line_gap - last_bbox_0[1]
            self.draw_text_with_outline(draw, (last_x, last_y), last, name_font, 
                                        self.name_color, self.name_outline_var.get())
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
            self.preview_canvas.create_text(74, canvas_h-15, anchor=tk.W, text="Name", fill="#333", font=("Arial", 8))
            
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
                stretched = self.voucher_image.resize((tw, th), Image.LANCZOS)
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
                        label = f"#{curr_att}"
                        bbox = draw.textbbox((0,0), label, font=font)
                        lw, lh = bbox[2]-bbox[0], bbox[3]-bbox[1]
                        draw.rectangle([x+(tw-lw)//2-2, y+(th-lh)//2-2, x+(tw+lw)//2+2, y+(th+lh)//2+2], fill='white', outline='#666')
                        draw.text((x+(tw-lw)//2, y+(th-lh)//2), label, fill='#333', font=font)
                    else:
                        draw.rectangle([x, y, x+tw-1, y+th-1], fill='#e0e0e0', outline='#ccc')
            
            draw.rectangle([0, 0, pw-1, ph-1], outline='#333', width=2)
            
            self.preview_photo = ImageTk.PhotoImage(page)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image((canvas_w-pw)//2, (canvas_h-ph)//2, anchor=tk.NW, image=self.preview_photo)
            
        except Exception as e:
            print(f"Layout error: {e}")
            traceback.print_exc()
    
    def generate_pdf(self):
        if not self.csv_path or not self.image_path or not self.attendees:
            messagebox.showwarning("Missing", "Select CSV and image first.")
            return
        
        output = filedialog.asksaveasfilename(title="Save PDF", defaultextension=".pdf", 
                                               filetypes=[("PDF", "*.pdf")], initialfile="vouchers.pdf")
        if not output:
            return
        
        self.status_label.config(text="Generating PDF...", foreground="blue")
        self.root.update()
        
        try:
            self.create_pdf(output)
            tpa = int(self.tickets_per_attendee_var.get())
            total_pages = self.calculate_total_pages()
            self.status_label.config(text=f"✓ Created {len(self.attendees)*tpa} vouchers on {total_pages} pages!", foreground="green")
            messagebox.showinfo("Success", f"Created {len(self.attendees)*tpa} vouchers!\n{total_pages} pages\n\nSaved to:\n{output}")
        except Exception as e:
            self.status_label.config(text="Error creating PDF", foreground="red")
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
        stretched = self.voucher_image.resize((iw, ih), Image.LANCZOS)
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
                        
                        c.drawImage(img_reader, x, y, width=ticket_w, height=ticket_h, mask='auto')
                        
                        cx, cy = x + ticket_w/2, y + ticket_h/2
                        size_factor = min(ticket_w / (3*inch), ticket_h / (1.75*inch))
                        
                        # Title
                        title = self.title_var.get().strip()
                        if title:
                            title_size = max(6, int(int(self.title_font_size_var.get()) * size_factor * 1.8))
                            font_name = "Helvetica-Bold" if self.title_bold_var.get() else "Helvetica"
                            c.setFont(font_name, title_size)
                            
                            # In PDF Y goes UP. title_y_pos negative = above center = higher Y
                            # Adjust for baseline (text draws from baseline, add half of font size to center)
                            title_y = cy - (self.title_y_pos * ticket_h) - title_size * 0.35
                            
                            if self.title_outline_var.get():
                                c.setFillColorRGB(1, 1, 1)
                                for dx in [-1, 0, 1]:
                                    for dy in [-1, 0, 1]:
                                        if dx or dy:
                                            c.drawCentredString(cx + dx, title_y + dy, title.upper())
                            
                            c.setFillColorRGB(title_rgb[0]/255, title_rgb[1]/255, title_rgb[2]/255)
                            c.drawCentredString(cx, title_y, title.upper())
                        
                        # Name - First above Last
                        name_size = max(6, int(int(self.name_font_size_var.get()) * size_factor * 1.8))
                        font_name = "Helvetica-Bold" if self.name_bold_var.get() else "Helvetica"
                        c.setFont(font_name, name_size)
                        
                        # Calculate position - name_y_pos positive = below center = lower Y in PDF
                        name_y_center = cy - (self.name_y_pos * ticket_h)
                        
                        # Line gap - small gap between first and last name (matches preview)
                        line_gap = name_size * 0.15
                        
                        # First name above, last name below (in PDF: higher Y = higher on page)
                        first_y = name_y_center + name_size * 0.5 + line_gap
                        last_y = name_y_center - name_size * 0.5
                        
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
                        
                        count += 1
                
                idx += 1
            
            if idx < len(self.attendees):
                c.showPage()
        
        c.save()
        try:
            os.remove(temp)
        except:
            pass


def main():
    root = tk.Tk()
    app = VoucherGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
