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
from tkinter import ttk, filedialog, messagebox
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
        self.root.geometry("700x920")
        self.root.resizable(False, False)
        
        # Variables
        self.csv_path = None
        self.image_path = None
        self.attendees = []
        self.voucher_image = None
        
        # Settings variables
        self.title_var = tk.StringVar(value="DRINK VOUCHER")
        self.font_size_var = tk.StringVar(value="12")
        self.bold_var = tk.IntVar(value=1)
        
        # Layout variables
        self.orientation_var = tk.StringVar(value="Portrait")
        self.ticket_width_var = tk.StringVar(value="3")
        self.ticket_height_var = tk.StringVar(value="1.75")
        self.tickets_per_attendee_var = tk.StringVar(value="5")
        
        # Preview mode: "ticket" or "layout"
        self.preview_mode = tk.StringVar(value="ticket")
        
        # Draggable text positions (as percentage of ticket height, 0 = center)
        # Negative = above center, Positive = below center
        self.title_y_pos = -0.25  # Title starts 25% above center
        self.name_y_pos = 0.10   # Name starts 10% below center
        
        # Dragging state
        self.dragging = None  # None, "title", or "name"
        self.drag_start_y = 0
        self.drag_start_pos = 0
        
        # Preview dimensions (updated when preview is drawn)
        self.preview_ticket_height = 0
        self.preview_offset_y = 0
        
        self.setup_ui()
        self.update_valid_sizes()
        
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === FILE SELECTION SECTION ===
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Select Files", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # CSV button and label
        csv_row = ttk.Frame(file_frame)
        csv_row.pack(fill=tk.X, pady=2)
        csv_btn = tk.Button(csv_row, text="Select CSV File", command=self.select_csv, bg="#4CAF50", fg="white", padx=10)
        csv_btn.pack(side=tk.LEFT)
        self.csv_label = ttk.Label(csv_row, text="No file selected", foreground="gray")
        self.csv_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Image button and label
        img_row = ttk.Frame(file_frame)
        img_row.pack(fill=tk.X, pady=2)
        img_btn = tk.Button(img_row, text="Select Voucher Image", command=self.select_image, bg="#2196F3", fg="white", padx=10)
        img_btn.pack(side=tk.LEFT)
        self.img_label = ttk.Label(img_row, text="No file selected", foreground="gray")
        self.img_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # === PREVIEW SECTION ===
        preview_frame = ttk.LabelFrame(main_frame, text="Preview (drag title/name to reposition)", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Preview toggle buttons
        toggle_row = ttk.Frame(preview_frame)
        toggle_row.pack(pady=(0, 10))
        
        self.ticket_btn = tk.Button(toggle_row, text="Single Ticket", command=lambda: self.set_preview_mode("ticket"),
                                     bg="#2196F3", fg="white", padx=15, pady=5)
        self.ticket_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.layout_btn = tk.Button(toggle_row, text="Page Layout", command=lambda: self.set_preview_mode("layout"),
                                     bg="#cccccc", fg="black", padx=15, pady=5)
        self.layout_btn.pack(side=tk.LEFT)
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(preview_frame, width=500, height=350, bg="#e0e0e0", relief="sunken", bd=2)
        self.preview_canvas.pack(pady=5)
        self.preview_canvas.create_text(250, 175, text="Select an image to see preview", fill="gray")
        
        # Bind mouse events for dragging
        self.preview_canvas.bind("<Button-1>", self.on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        # Layout info label
        self.layout_info_label = ttk.Label(preview_frame, text="", foreground="blue")
        self.layout_info_label.pack()
        
        # Drag instruction
        self.drag_hint_label = ttk.Label(preview_frame, text="↕ Click and drag title or name to reposition", foreground="gray")
        self.drag_hint_label.pack()
        
        # === SETTINGS SECTION ===
        settings_frame = ttk.LabelFrame(main_frame, text="Step 2: Customize Text", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Title field
        title_row = ttk.Frame(settings_frame)
        title_row.pack(fill=tk.X, pady=5)
        ttk.Label(title_row, text="Title:", width=15).pack(side=tk.LEFT)
        title_entry = ttk.Entry(title_row, textvariable=self.title_var, width=25)
        title_entry.pack(side=tk.LEFT, padx=(5, 0))
        title_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        ttk.Label(title_row, text="(leave empty for no title)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        
        # Font size dropdown
        font_row = ttk.Frame(settings_frame)
        font_row.pack(fill=tk.X, pady=5)
        ttk.Label(font_row, text="Font Size:", width=15).pack(side=tk.LEFT)
        self.font_combo = ttk.Combobox(font_row, textvariable=self.font_size_var, 
                                        values=["8", "9", "10", "11", "12", "13", "14", "16", "18"], 
                                        width=5, state="readonly")
        self.font_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.font_combo.set("12")
        self.font_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Bold checkbox
        bold_row = ttk.Frame(settings_frame)
        bold_row.pack(fill=tk.X, pady=5)
        ttk.Label(bold_row, text="Bold Names:", width=15).pack(side=tk.LEFT)
        self.bold_check = tk.Checkbutton(bold_row, variable=self.bold_var, command=self.update_preview)
        self.bold_check.pack(side=tk.LEFT, padx=(5, 0))
        self.bold_check.select()
        
        # Reset positions button
        reset_row = ttk.Frame(settings_frame)
        reset_row.pack(fill=tk.X, pady=5)
        ttk.Label(reset_row, text="Text Position:", width=15).pack(side=tk.LEFT)
        reset_btn = tk.Button(reset_row, text="Reset to Center", command=self.reset_text_positions, 
                              bg="#888", fg="white", padx=10)
        reset_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # === LAYOUT SECTION ===
        layout_frame = ttk.LabelFrame(main_frame, text="Step 3: Ticket Size & Layout", padding="10")
        layout_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Page Orientation
        orient_row = ttk.Frame(layout_frame)
        orient_row.pack(fill=tk.X, pady=5)
        ttk.Label(orient_row, text="Page:", width=15).pack(side=tk.LEFT)
        orient_frame = ttk.Frame(orient_row)
        orient_frame.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Radiobutton(orient_frame, text="Portrait (8.5\" × 11\")", 
                        variable=self.orientation_var, value="Portrait",
                        command=self.on_orientation_change).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(orient_frame, text="Landscape (11\" × 8.5\")", 
                        variable=self.orientation_var, value="Landscape",
                        command=self.on_orientation_change).pack(side=tk.LEFT)
        
        # Ticket Width
        width_row = ttk.Frame(layout_frame)
        width_row.pack(fill=tk.X, pady=5)
        ttk.Label(width_row, text="Ticket Width:", width=15).pack(side=tk.LEFT)
        self.width_combo = ttk.Combobox(width_row, textvariable=self.ticket_width_var,
                                    values=["1.5", "2", "2.5", "3", "3.5", "4", "4.25"], 
                                    width=8, state="readonly")
        self.width_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.width_combo.bind('<<ComboboxSelected>>', lambda e: self.on_size_change())
        ttk.Label(width_row, text="inches", foreground="gray").pack(side=tk.LEFT, padx=(5, 0))
        self.width_info_label = ttk.Label(width_row, text="", foreground="blue")
        self.width_info_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Ticket Height
        height_row = ttk.Frame(layout_frame)
        height_row.pack(fill=tk.X, pady=5)
        ttk.Label(height_row, text="Ticket Height:", width=15).pack(side=tk.LEFT)
        self.height_combo = ttk.Combobox(height_row, textvariable=self.ticket_height_var,
                                     values=["1", "1.25", "1.5", "1.75", "2", "2.5", "3", "3.5", "4"], 
                                     width=8, state="readonly")
        self.height_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.height_combo.bind('<<ComboboxSelected>>', lambda e: self.on_size_change())
        ttk.Label(height_row, text="inches", foreground="gray").pack(side=tk.LEFT, padx=(5, 0))
        self.height_info_label = ttk.Label(height_row, text="", foreground="blue")
        self.height_info_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Tickets per attendee
        tpa_row = ttk.Frame(layout_frame)
        tpa_row.pack(fill=tk.X, pady=5)
        ttk.Label(tpa_row, text="Tickets/Attendee:", width=15).pack(side=tk.LEFT)
        tpa_combo = ttk.Combobox(tpa_row, textvariable=self.tickets_per_attendee_var,
                                  values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], 
                                  width=5, state="readonly")
        tpa_combo.pack(side=tk.LEFT, padx=(5, 0))
        tpa_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Calculated info display
        calc_frame = ttk.Frame(layout_frame)
        calc_frame.pack(fill=tk.X, pady=(10, 5))
        
        self.calc_label = ttk.Label(calc_frame, text="", foreground="#333", font=("Arial", 9))
        self.calc_label.pack(anchor=tk.W)
        
        self.warning_label = ttk.Label(calc_frame, text="", foreground="orange")
        self.warning_label.pack(anchor=tk.W)
        
        # === GENERATE SECTION ===
        generate_frame = ttk.Frame(main_frame)
        generate_frame.pack(fill=tk.X, pady=10)
        
        self.generate_btn = tk.Button(generate_frame, text="GENERATE PDF", command=self.generate_pdf,
                                       bg="#FF5722", fg="white", font=("Arial", 14, "bold"),
                                       padx=30, pady=12, state=tk.DISABLED)
        self.generate_btn.pack(pady=10)
        
        # Status label
        self.status_label = ttk.Label(generate_frame, text="Select a CSV file and image to get started", foreground="gray")
        self.status_label.pack()
    
    def reset_text_positions(self):
        """Reset title and name to default positions"""
        self.title_y_pos = -0.25
        self.name_y_pos = 0.10
        self.update_preview()
    
    def on_canvas_click(self, event):
        """Handle mouse click on canvas - start dragging if on text"""
        if self.preview_mode.get() != "ticket" or not self.voucher_image:
            return
        
        # Check if click is within the ticket area
        canvas_height = 350
        ticket_top = self.preview_offset_y
        ticket_bottom = ticket_top + self.preview_ticket_height
        
        if event.y < ticket_top or event.y > ticket_bottom:
            return
        
        # Calculate relative Y position within ticket (0 = center)
        relative_y = (event.y - ticket_top) / self.preview_ticket_height - 0.5
        
        # Check if near title or name position (within 15% tolerance)
        title_pos = self.title_y_pos
        name_pos = self.name_y_pos
        
        if abs(relative_y - title_pos) < 0.15:
            self.dragging = "title"
            self.drag_start_y = event.y
            self.drag_start_pos = self.title_y_pos
            self.preview_canvas.config(cursor="sb_v_double_arrow")
        elif abs(relative_y - name_pos) < 0.15:
            self.dragging = "name"
            self.drag_start_y = event.y
            self.drag_start_pos = self.name_y_pos
            self.preview_canvas.config(cursor="sb_v_double_arrow")
    
    def on_canvas_drag(self, event):
        """Handle mouse drag on canvas"""
        if not self.dragging or self.preview_ticket_height == 0:
            return
        
        # Calculate position change as percentage of ticket height
        delta_y = (event.y - self.drag_start_y) / self.preview_ticket_height
        new_pos = self.drag_start_pos + delta_y
        
        # Clamp to reasonable bounds (-0.45 to 0.45)
        new_pos = max(-0.45, min(0.45, new_pos))
        
        if self.dragging == "title":
            self.title_y_pos = new_pos
        else:
            self.name_y_pos = new_pos
        
        self.update_preview()
    
    def on_canvas_release(self, event):
        """Handle mouse release"""
        self.dragging = None
        self.preview_canvas.config(cursor="")
        
    def set_preview_mode(self, mode):
        """Toggle between ticket and layout preview modes"""
        self.preview_mode.set(mode)
        
        if mode == "ticket":
            self.ticket_btn.config(bg="#2196F3", fg="white")
            self.layout_btn.config(bg="#cccccc", fg="black")
            self.drag_hint_label.config(text="↕ Click and drag title or name to reposition")
        else:
            self.ticket_btn.config(bg="#cccccc", fg="black")
            self.layout_btn.config(bg="#2196F3", fg="white")
            self.drag_hint_label.config(text="")
        
        self.update_preview()
    
    def on_orientation_change(self):
        """Handle orientation change"""
        self.update_valid_sizes()
        self.update_preview()
    
    def on_size_change(self):
        """Handle ticket size change"""
        self.update_valid_sizes()
        self.update_preview()
    
    def update_valid_sizes(self):
        """Update width/height options based on page orientation"""
        page_w, page_h = self.get_page_dimensions()
        page_w_in = page_w / inch
        page_h_in = page_h / inch
        
        all_widths = ["1.5", "2", "2.5", "3", "3.5", "4", "4.25", "5.5"]
        all_heights = ["1", "1.25", "1.5", "1.75", "2", "2.5", "3", "3.5", "4", "5.5"]
        
        valid_widths = [w for w in all_widths if float(w) <= page_w_in]
        valid_heights = [h for h in all_heights if float(h) <= page_h_in]
        
        current_width = self.ticket_width_var.get()
        current_height = self.ticket_height_var.get()
        
        self.width_combo['values'] = valid_widths
        self.height_combo['values'] = valid_heights
        
        if current_width not in valid_widths and valid_widths:
            self.ticket_width_var.set("3" if "3" in valid_widths else valid_widths[-1])
        if current_height not in valid_heights and valid_heights:
            self.ticket_height_var.set("1.75" if "1.75" in valid_heights else valid_heights[-1])
        
        try:
            cols = int(page_w_in / float(self.ticket_width_var.get()))
            rows = int(page_h_in / float(self.ticket_height_var.get()))
            self.width_info_label.config(text=f"({cols} across)")
            self.height_info_label.config(text=f"({rows} down)")
        except:
            pass
        
    def get_page_dimensions(self):
        """Get page dimensions in points"""
        if self.orientation_var.get() == "Landscape":
            return landscape(letter)
        return letter
    
    def get_ticket_dimensions(self):
        """Get ticket dimensions in points"""
        width = float(self.ticket_width_var.get()) * inch
        height = float(self.ticket_height_var.get()) * inch
        return width, height
    
    def calculate_grid(self):
        """Calculate how many tickets fit on a page (no mixing attendees per row)"""
        page_w, page_h = self.get_page_dimensions()
        ticket_w, ticket_h = self.get_ticket_dimensions()
        
        cols = max(1, int(page_w // ticket_w))
        rows = max(1, int(page_h // ticket_h))
        
        tickets_per_attendee = int(self.tickets_per_attendee_var.get())
        rows_per_attendee = math.ceil(tickets_per_attendee / cols)
        attendees_per_page = max(1, rows // rows_per_attendee)
        
        return cols, rows, attendees_per_page, rows_per_attendee
    
    def update_calc_display(self):
        """Update the calculated layout info"""
        cols, rows, attendees_per_page, rows_per_attendee = self.calculate_grid()
        tickets_per_attendee = int(self.tickets_per_attendee_var.get())
        
        info = f"Layout: {cols}×{rows} grid  |  {attendees_per_page} attendee(s)/page  |  {rows_per_attendee} row(s)/attendee"
        self.calc_label.config(text=info)
        
        ticket_w = float(self.ticket_width_var.get())
        ticket_h = float(self.ticket_height_var.get())
        self.layout_info_label.config(text=f"Ticket: {ticket_w}\" × {ticket_h}\"  |  {tickets_per_attendee} per person")
        
        wasted = (cols * rows_per_attendee) - tickets_per_attendee
        if wasted > 0:
            self.warning_label.config(text=f"Note: {wasted} empty slot(s) per attendee to keep rows separate")
        else:
            self.warning_label.config(text="")
        
    def select_csv(self):
        try:
            path = filedialog.askopenfilename(
                title="Select CSV file with attendee names",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if path:
                self.csv_path = path
                self.csv_label.config(text=os.path.basename(path), foreground="black")
                self.attendees = self.read_attendees(path)
                self.status_label.config(text=f"✓ Loaded {len(self.attendees)} attendees", foreground="green")
                self.check_ready()
                self.update_preview()
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting CSV:\n{e}")
            
    def select_image(self):
        try:
            path = filedialog.askopenfilename(
                title="Select voucher image",
                filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"), ("All files", "*.*")]
            )
            if path:
                self.image_path = path
                self.img_label.config(text=os.path.basename(path), foreground="black")
                self.voucher_image = Image.open(path)
                self.check_ready()
                self.update_preview()
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting image:\n{e}")
            
    def check_ready(self):
        """Enable generate button when both files are selected"""
        if self.csv_path and self.image_path and len(self.attendees) > 0:
            self.generate_btn.config(state=tk.NORMAL)
            tpa = int(self.tickets_per_attendee_var.get())
            total = len(self.attendees) * tpa
            self.status_label.config(text=f"✓ Ready! {len(self.attendees)} attendees × {tpa} = {total} tickets", foreground="green")
        else:
            self.generate_btn.config(state=tk.DISABLED)
            
    def read_attendees(self, csv_path):
        """Read attendee names from CSV file"""
        attendees = []
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and row[0].strip():
                        if len(row) >= 2 and row[1].strip():
                            full_name = f"{row[0].strip()}, {row[1].strip()}"
                        else:
                            full_name = row[0].strip()
                        attendees.append(full_name)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV:\n{e}")
        return attendees
    
    def parse_name(self, full_name):
        """Parse name from 'Last, First' format"""
        if ',' in full_name:
            parts = full_name.split(',', 1)
            return parts[1].strip() if len(parts) > 1 else "", parts[0].strip()
        return full_name.strip(), ""
            
    def update_preview(self):
        """Update the preview canvas"""
        self.update_calc_display()
        self.check_ready()
        
        if not self.voucher_image:
            return
        
        if self.preview_mode.get() == "ticket":
            self.update_ticket_preview()
        else:
            self.update_layout_preview()
    
    def resize_image_to_fit(self, img, target_w, target_h):
        """Resize image to fit inside dimensions (no cropping)"""
        img_ratio = img.width / img.height
        target_ratio = target_w / target_h
        
        if img_ratio > target_ratio:
            new_w = target_w
            new_h = int(target_w / img_ratio)
        else:
            new_h = target_h
            new_w = int(target_h * img_ratio)
        
        return img.resize((new_w, new_h), Image.LANCZOS)
    
    def update_ticket_preview(self):
        """Update preview showing single ticket with draggable text"""
        if not self.voucher_image:
            return
        
        try:
            first_name, last_name = self.parse_name(self.attendees[0]) if self.attendees else ("First", "Last")
            
            ticket_w_in = float(self.ticket_width_var.get())
            ticket_h_in = float(self.ticket_height_var.get())
            aspect_ratio = ticket_w_in / ticket_h_in
            
            canvas_width = 500
            canvas_height = 350
            max_w, max_h = 460, 320
            
            if aspect_ratio > max_w / max_h:
                preview_w = max_w
                preview_h = int(max_w / aspect_ratio)
            else:
                preview_h = max_h
                preview_w = int(max_h * aspect_ratio)
            
            # Store for drag calculations
            self.preview_ticket_height = preview_h
            self.preview_offset_y = (canvas_height - preview_h) // 2
            
            # Create ticket image
            ticket = Image.new('RGB', (preview_w, preview_h), '#FFFFFF')
            
            if self.voucher_image:
                fitted = self.resize_image_to_fit(self.voucher_image, preview_w, preview_h)
                x = (preview_w - fitted.width) // 2
                y = (preview_h - fitted.height) // 2
                ticket.paste(fitted, (x, y))
            
            draw = ImageDraw.Draw(ticket)
            
            title = self.title_var.get()
            font_size = int(self.font_size_var.get())
            bold = self.bold_var.get() == 1
            
            scale = min(preview_w / 216, preview_h / 126)
            scaled_size = max(10, int(font_size * scale * 1.5))
            
            try:
                if bold:
                    name_font = ImageFont.truetype("arialbd.ttf", scaled_size)
                else:
                    name_font = ImageFont.truetype("arial.ttf", scaled_size)
                title_font = ImageFont.truetype("arialbd.ttf", int(scaled_size * 0.8))
            except:
                try:
                    if bold:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", scaled_size)
                    else:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", scaled_size)
                    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", int(scaled_size * 0.8))
                except:
                    name_font = ImageFont.load_default()
                    title_font = ImageFont.load_default()
            
            center_x = preview_w // 2
            center_y = preview_h // 2
            text_color = (51, 38, 26)
            
            # Draw title at draggable position
            if title.strip():
                title_y = center_y + int(self.title_y_pos * preview_h)
                bbox = draw.textbbox((0, 0), title.upper(), font=title_font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                draw.text((center_x - tw // 2, title_y - th // 2), title.upper(), fill=text_color, font=title_font)
                
                # Draw drag indicator
                draw.rectangle([center_x - tw//2 - 5, title_y - th//2 - 3, 
                               center_x + tw//2 + 5, title_y + th//2 + 3], outline="#2196F3", width=1)
            
            # Draw name at draggable position
            name_y = center_y + int(self.name_y_pos * preview_h)
            full_name = f"{first_name} {last_name}".strip()
            bbox = draw.textbbox((0, 0), full_name, font=name_font)
            nw = bbox[2] - bbox[0]
            nh = bbox[3] - bbox[1]
            draw.text((center_x - nw // 2, name_y - nh // 2), full_name, fill=text_color, font=name_font)
            
            # Draw drag indicator for name
            draw.rectangle([center_x - nw//2 - 5, name_y - nh//2 - 3, 
                           center_x + nw//2 + 5, name_y + nh//2 + 3], outline="#4CAF50", width=1)
            
            self.preview_photo = ImageTk.PhotoImage(ticket)
            
            self.preview_canvas.delete("all")
            x = (canvas_width - preview_w) // 2
            y = (canvas_height - preview_h) // 2
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_photo)
            
            # Draw legend
            self.preview_canvas.create_text(10, canvas_height - 30, anchor=tk.W, 
                                           text="■ Title", fill="#2196F3", font=("Arial", 9))
            self.preview_canvas.create_text(60, canvas_height - 30, anchor=tk.W, 
                                           text="■ Name", fill="#4CAF50", font=("Arial", 9))
            
        except Exception as e:
            print(f"Preview error: {e}")
            traceback.print_exc()
    
    def update_layout_preview(self):
        """Update preview showing full page layout"""
        try:
            page_w, page_h = self.get_page_dimensions()
            cols, rows, attendees_per_page, rows_per_attendee = self.calculate_grid()
            tickets_per_attendee = int(self.tickets_per_attendee_var.get())
            ticket_w, ticket_h = self.get_ticket_dimensions()
            
            canvas_width, canvas_height = 500, 350
            margin = 20
            
            scale = min((canvas_width - 2*margin) / page_w, (canvas_height - 2*margin) / page_h)
            
            pw = int(page_w * scale)
            ph = int(page_h * scale)
            tw = int(ticket_w * scale)
            th = int(ticket_h * scale)
            
            page_img = Image.new('RGB', (pw, ph), 'white')
            draw = ImageDraw.Draw(page_img)
            
            try:
                font = ImageFont.truetype("arial.ttf", 9)
            except:
                try:
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 9)
                except:
                    font = ImageFont.load_default()
            
            mini = None
            if self.voucher_image and tw > 10 and th > 10:
                mini = Image.new('RGB', (tw, th), '#FFFFFF')
                fitted = self.resize_image_to_fit(self.voucher_image, tw, th)
                mini.paste(fitted, ((tw - fitted.width)//2, (th - fitted.height)//2))
            
            gw = cols * tw
            gh = rows * th
            ox = (pw - gw) // 2
            oy = (ph - gh) // 2
            
            for row in range(rows):
                attendee_row = row // rows_per_attendee
                row_in_attendee = row % rows_per_attendee
                current_attendee = attendee_row + 1
                
                if current_attendee > attendees_per_page:
                    break
                
                for col in range(cols):
                    x = ox + col * tw
                    y = oy + row * th
                    ticket_idx = row_in_attendee * cols + col
                    
                    if ticket_idx < tickets_per_attendee:
                        if mini:
                            page_img.paste(mini, (x, y))
                        draw.rectangle([x, y, x+tw-1, y+th-1], outline='#999')
                        
                        label = f"#{current_attendee}"
                        bbox = draw.textbbox((0,0), label, font=font)
                        lw, lh = bbox[2]-bbox[0], bbox[3]-bbox[1]
                        lx, ly = x + (tw-lw)//2, y + (th-lh)//2
                        draw.rectangle([lx-2, ly-2, lx+lw+2, ly+lh+2], fill='white', outline='#666')
                        draw.text((lx, ly), label, fill='#333', font=font)
                    else:
                        draw.rectangle([x, y, x+tw-1, y+th-1], fill='#e0e0e0', outline='#ccc')
            
            draw.rectangle([0, 0, pw-1, ph-1], outline='#333', width=2)
            
            self.preview_photo = ImageTk.PhotoImage(page_img)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image((canvas_width-pw)//2, (canvas_height-ph)//2, anchor=tk.NW, image=self.preview_photo)
            
        except Exception as e:
            print(f"Layout error: {e}")
            traceback.print_exc()
    
    def generate_pdf(self):
        """Generate the PDF file"""
        if not self.csv_path or not self.image_path or not self.attendees:
            messagebox.showwarning("Missing Data", "Please select CSV and image files first.")
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                title="Save PDF as...",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile="drink_vouchers.pdf"
            )
            
            if not output_path:
                return
            
            self.status_label.config(text="Generating PDF...", foreground="blue")
            self.root.update()
            
            self.create_vouchers_pdf(output_path)
            
            tpa = int(self.tickets_per_attendee_var.get())
            total = len(self.attendees) * tpa
            self.status_label.config(text=f"✓ Created {total} vouchers!", foreground="green")
            messagebox.showinfo("Success", f"Created {total} vouchers!\n\nSaved to:\n{output_path}")
            
        except Exception as e:
            self.status_label.config(text="Error creating PDF", foreground="red")
            messagebox.showerror("Error", f"Could not create PDF:\n{e}")
            traceback.print_exc()
            
    def create_vouchers_pdf(self, output_path):
        """Generate PDF with all vouchers"""
        title_text = self.title_var.get().strip()
        font_size = int(self.font_size_var.get())
        bold = self.bold_var.get() == 1
        
        page_w, page_h = self.get_page_dimensions()
        ticket_w, ticket_h = self.get_ticket_dimensions()
        cols, rows, attendees_per_page, rows_per_attendee = self.calculate_grid()
        tickets_per_attendee = int(self.tickets_per_attendee_var.get())
        
        gw = cols * ticket_w
        gh = rows * ticket_h
        ox = (page_w - gw) / 2
        oy = (page_h - gh) / 2
        
        # Prepare ticket image
        dpi_scale = 3
        img_w = int(ticket_w * dpi_scale)
        img_h = int(ticket_h * dpi_scale)
        
        fitted = self.resize_image_to_fit(self.voucher_image, img_w, img_h)
        ticket_img = Image.new('RGB', (img_w, img_h), '#FFFFFF')
        ticket_img.paste(fitted, ((img_w - fitted.width)//2, (img_h - fitted.height)//2))
        
        temp_path = os.path.join(os.path.dirname(output_path), "_temp_voucher.png")
        ticket_img.save(temp_path, "PNG")
        img_reader = ImageReader(temp_path)
        
        c = canvas.Canvas(output_path, pagesize=(page_w, page_h))
        
        attendee_idx = 0
        while attendee_idx < len(self.attendees):
            for page_att in range(attendees_per_page):
                if attendee_idx >= len(self.attendees):
                    break
                
                first_name, last_name = self.parse_name(self.attendees[attendee_idx])
                start_row = page_att * rows_per_attendee
                ticket_count = 0
                
                for row_off in range(rows_per_attendee):
                    if ticket_count >= tickets_per_attendee:
                        break
                    for col in range(cols):
                        if ticket_count >= tickets_per_attendee:
                            break
                        
                        row = start_row + row_off
                        x = ox + col * ticket_w
                        y = page_h - oy - (row + 1) * ticket_h
                        
                        c.drawImage(img_reader, x, y, width=ticket_w, height=ticket_h, mask='auto')
                        
                        center_x = x + ticket_w / 2
                        center_y = y + ticket_h / 2
                        
                        c.setFillColorRGB(0.2, 0.15, 0.1)
                        name_font = "Helvetica-Bold" if bold else "Helvetica"
                        
                        size_factor = min(ticket_w / (3*inch), ticket_h / (1.75*inch))
                        scaled = max(6, int(font_size * size_factor))
                        
                        # Apply draggable positions
                        if title_text:
                            title_y = center_y + self.title_y_pos * ticket_h
                            c.setFont("Helvetica-Bold", int(scaled * 0.8))
                            c.drawCentredString(center_x, title_y, title_text.upper())
                        
                        name_y = center_y + self.name_y_pos * ticket_h
                        full_name = f"{first_name} {last_name}".strip()
                        c.setFont(name_font, scaled)
                        c.drawCentredString(center_x, name_y, full_name)
                        
                        ticket_count += 1
                
                attendee_idx += 1
            
            if attendee_idx < len(self.attendees):
                c.showPage()
        
        c.save()
        
        try:
            os.remove(temp_path)
        except:
            pass


def main():
    root = tk.Tk()
    app = VoucherGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
