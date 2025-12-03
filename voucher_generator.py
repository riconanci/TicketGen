#!/usr/bin/env python3
"""
Drink Voucher Generator
GUI app with live preview for creating personalized drink vouchers.
Repository: https://github.com/riconanci/TicketGen
"""

import csv
import os
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
        self.root.geometry("650x850")
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
        self.columns_var = tk.StringVar(value="2")
        self.rows_var = tk.StringVar(value="6")
        self.tickets_per_attendee_var = tk.StringVar(value="5")
        
        # Preview mode: "ticket" or "layout"
        self.preview_mode = tk.StringVar(value="ticket")
        
        self.setup_ui()
        
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
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
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
        self.preview_canvas = tk.Canvas(preview_frame, width=450, height=320, bg="#f0f0f0", relief="sunken", bd=2)
        self.preview_canvas.pack(pady=5)
        self.preview_canvas.create_text(225, 160, text="Select an image to see preview", fill="gray")
        
        # Ticket size info label
        self.size_info_label = ttk.Label(preview_frame, text="", foreground="gray")
        self.size_info_label.pack()
        
        # Warning label for size issues
        self.warning_label = ttk.Label(preview_frame, text="", foreground="orange")
        self.warning_label.pack()
        
        # === SETTINGS SECTION ===
        settings_frame = ttk.LabelFrame(main_frame, text="Step 2: Customize", padding="10")
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
        
        # === LAYOUT SECTION ===
        layout_frame = ttk.LabelFrame(main_frame, text="Step 3: Page Layout", padding="10")
        layout_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Orientation
        orient_row = ttk.Frame(layout_frame)
        orient_row.pack(fill=tk.X, pady=5)
        ttk.Label(orient_row, text="Orientation:", width=15).pack(side=tk.LEFT)
        orient_combo = ttk.Combobox(orient_row, textvariable=self.orientation_var,
                                     values=["Portrait", "Landscape"], width=10, state="readonly")
        orient_combo.pack(side=tk.LEFT, padx=(5, 0))
        orient_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Columns
        cols_row = ttk.Frame(layout_frame)
        cols_row.pack(fill=tk.X, pady=5)
        ttk.Label(cols_row, text="Columns:", width=15).pack(side=tk.LEFT)
        cols_combo = ttk.Combobox(cols_row, textvariable=self.columns_var,
                                   values=["1", "2", "3", "4"], width=5, state="readonly")
        cols_combo.pack(side=tk.LEFT, padx=(5, 0))
        cols_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        ttk.Label(cols_row, text="(tickets across)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        
        # Rows
        rows_row = ttk.Frame(layout_frame)
        rows_row.pack(fill=tk.X, pady=5)
        ttk.Label(rows_row, text="Rows:", width=15).pack(side=tk.LEFT)
        rows_combo = ttk.Combobox(rows_row, textvariable=self.rows_var,
                                   values=["1", "2", "3", "4", "5", "6"], width=5, state="readonly")
        rows_combo.pack(side=tk.LEFT, padx=(5, 0))
        rows_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        ttk.Label(rows_row, text="(tickets down)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        
        # Tickets per attendee
        tpa_row = ttk.Frame(layout_frame)
        tpa_row.pack(fill=tk.X, pady=5)
        ttk.Label(tpa_row, text="Tickets/Attendee:", width=15).pack(side=tk.LEFT)
        tpa_combo = ttk.Combobox(tpa_row, textvariable=self.tickets_per_attendee_var,
                                  values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], width=5, state="readonly")
        tpa_combo.pack(side=tk.LEFT, padx=(5, 0))
        tpa_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Calculated info
        calc_row = ttk.Frame(layout_frame)
        calc_row.pack(fill=tk.X, pady=5)
        ttk.Label(calc_row, text="", width=15).pack(side=tk.LEFT)
        self.calc_label = ttk.Label(calc_row, text="Attendees per page: 2", foreground="blue")
        self.calc_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # === GENERATE SECTION ===
        generate_frame = ttk.Frame(main_frame)
        generate_frame.pack(fill=tk.X, pady=10)
        
        self.generate_btn = tk.Button(generate_frame, text="GENERATE PDF", command=self.generate_pdf,
                                       bg="#FF5722", fg="white", font=("Arial", 12, "bold"),
                                       padx=20, pady=10, state=tk.DISABLED)
        self.generate_btn.pack(pady=10)
        
        # Status label
        self.status_label = ttk.Label(generate_frame, text="Select a CSV file and image to get started", foreground="gray")
        self.status_label.pack()
        
    def set_preview_mode(self, mode):
        """Toggle between ticket and layout preview modes"""
        self.preview_mode.set(mode)
        
        if mode == "ticket":
            self.ticket_btn.config(bg="#2196F3", fg="white")
            self.layout_btn.config(bg="#cccccc", fg="black")
        else:
            self.ticket_btn.config(bg="#cccccc", fg="black")
            self.layout_btn.config(bg="#2196F3", fg="white")
        
        self.update_preview()
        
    def get_page_dimensions(self):
        """Get page dimensions based on orientation"""
        if self.orientation_var.get() == "Landscape":
            return landscape(letter)
        return letter
    
    def get_ticket_dimensions(self):
        """Calculate ticket dimensions based on layout settings"""
        page_width, page_height = self.get_page_dimensions()
        cols = int(self.columns_var.get())
        rows = int(self.rows_var.get())
        
        ticket_width = page_width / cols
        ticket_height = page_height / rows
        
        return ticket_width, ticket_height
    
    def get_attendees_per_page(self):
        """Calculate how many attendees fit on one page"""
        cols = int(self.columns_var.get())
        rows = int(self.rows_var.get())
        tickets_per_attendee = int(self.tickets_per_attendee_var.get())
        
        total_tickets = cols * rows
        attendees = total_tickets // tickets_per_attendee
        return max(1, attendees)
    
    def check_ticket_size(self):
        """Check if ticket size is reasonable and return warnings"""
        ticket_width, ticket_height = self.get_ticket_dimensions()
        
        # Convert to inches for display
        width_in = ticket_width / inch
        height_in = ticket_height / inch
        
        warnings = []
        
        # Minimum size check (1.5" x 1")
        if width_in < 1.5:
            warnings.append(f"Width ({width_in:.2f}\") may be too narrow for text")
        if height_in < 1.0:
            warnings.append(f"Height ({height_in:.2f}\") may be too short for text")
            
        # Maximum size check (4.25" x 3.5")
        if width_in > 4.25:
            warnings.append(f"Width ({width_in:.2f}\") is quite large")
        if height_in > 3.5:
            warnings.append(f"Height ({height_in:.2f}\") is quite large")
        
        return width_in, height_in, warnings
        
    def select_csv(self):
        try:
            path = filedialog.askopenfilename(
                title="Select your CSV file with attendee names",
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
                title="Select your voucher image",
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
            tickets_per_attendee = int(self.tickets_per_attendee_var.get())
            total_tickets = len(self.attendees) * tickets_per_attendee
            self.status_label.config(text=f"✓ Ready! {len(self.attendees)} attendees × {tickets_per_attendee} = {total_tickets} tickets", foreground="green")
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
            messagebox.showerror("Error", f"Could not read CSV file:\n{e}")
        return attendees
    
    def parse_name(self, full_name):
        """Parse name from 'Last, First' format"""
        full_name = full_name.strip()
        if ',' in full_name:
            parts = full_name.split(',', 1)
            last_name = parts[0].strip()
            first_name = parts[1].strip() if len(parts) > 1 else ""
            return first_name, last_name
        else:
            return full_name, ""
            
    def update_preview(self):
        """Update the preview canvas based on current mode"""
        # Update calculated values
        attendees_per_page = self.get_attendees_per_page()
        self.calc_label.config(text=f"Attendees per page: {attendees_per_page}")
        
        # Update size info
        width_in, height_in, warnings = self.check_ticket_size()
        self.size_info_label.config(text=f"Ticket size: {width_in:.2f}\" × {height_in:.2f}\"")
        
        if warnings:
            self.warning_label.config(text="⚠ " + "; ".join(warnings))
        else:
            self.warning_label.config(text="")
        
        # Update status if ready
        self.check_ready()
        
        if not self.voucher_image:
            return
        
        if self.preview_mode.get() == "ticket":
            self.update_ticket_preview()
        else:
            self.update_layout_preview()
    
    def update_ticket_preview(self):
        """Update preview showing single ticket"""
        if not self.voucher_image:
            return
        
        try:
            title = self.title_var.get()
            font_size = int(self.font_size_var.get())
            bold = self.bold_var.get() == 1
            
            if self.attendees:
                first_name, last_name = self.parse_name(self.attendees[0])
            else:
                first_name, last_name = "First", "Last"
            
            # Get ticket aspect ratio
            ticket_width, ticket_height = self.get_ticket_dimensions()
            aspect_ratio = ticket_width / ticket_height
            
            # Create preview at appropriate size maintaining aspect ratio
            max_preview_width = 420
            max_preview_height = 300
            
            if aspect_ratio > max_preview_width / max_preview_height:
                preview_width = max_preview_width
                preview_height = int(max_preview_width / aspect_ratio)
            else:
                preview_height = max_preview_height
                preview_width = int(max_preview_height * aspect_ratio)
            
            # Resize and crop image to fit ticket dimensions
            preview_img = self.resize_image_to_fill(self.voucher_image, preview_width, preview_height)
            
            draw = ImageDraw.Draw(preview_img)
            
            # Scale font size based on preview size
            scale = preview_height / 175  # Base scale
            scaled_font_size = int(font_size * scale)
            
            # Try to load fonts
            try:
                if bold:
                    name_font = ImageFont.truetype("arialbd.ttf", scaled_font_size * 2)
                else:
                    name_font = ImageFont.truetype("arial.ttf", scaled_font_size * 2)
                title_font = ImageFont.truetype("arialbd.ttf", int(scaled_font_size * 1.5))
            except:
                try:
                    if bold:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", scaled_font_size * 2)
                    else:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", scaled_font_size * 2)
                    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", int(scaled_font_size * 1.5))
                except:
                    name_font = ImageFont.load_default()
                    title_font = ImageFont.load_default()
            
            img_width, img_height = preview_img.size
            center_x = img_width // 2
            center_y = img_height // 2
            
            text_color = (51, 38, 26)
            
            if title.strip():
                title_bbox = draw.textbbox((0, 0), title.upper(), font=title_font)
                title_width = title_bbox[2] - title_bbox[0]
                draw.text((center_x - title_width // 2, center_y - 45), title.upper(), fill=text_color, font=title_font)
                name_y_offset = -5
            else:
                name_y_offset = -20
            
            fn_bbox = draw.textbbox((0, 0), first_name, font=name_font)
            fn_width = fn_bbox[2] - fn_bbox[0]
            draw.text((center_x - fn_width // 2, center_y + name_y_offset), first_name, fill=text_color, font=name_font)
            
            ln_bbox = draw.textbbox((0, 0), last_name, font=name_font)
            ln_width = ln_bbox[2] - ln_bbox[0]
            draw.text((center_x - ln_width // 2, center_y + name_y_offset + scaled_font_size * 2 + 5), last_name, fill=text_color, font=name_font)
            
            self.preview_photo = ImageTk.PhotoImage(preview_img)
            
            self.preview_canvas.delete("all")
            canvas_width = 450
            canvas_height = 320
            x = (canvas_width - preview_img.width) // 2
            y = (canvas_height - preview_img.height) // 2
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_photo)
            
        except Exception as e:
            print(f"Preview error: {e}")
            traceback.print_exc()
    
    def update_layout_preview(self):
        """Update preview showing full page layout"""
        try:
            page_width, page_height = self.get_page_dimensions()
            cols = int(self.columns_var.get())
            rows = int(self.rows_var.get())
            tickets_per_attendee = int(self.tickets_per_attendee_var.get())
            
            # Calculate preview scale to fit in canvas
            canvas_width = 450
            canvas_height = 320
            margin = 20
            
            available_width = canvas_width - 2 * margin
            available_height = canvas_height - 2 * margin
            
            scale_x = available_width / page_width
            scale_y = available_height / page_height
            scale = min(scale_x, scale_y)
            
            preview_page_width = int(page_width * scale)
            preview_page_height = int(page_height * scale)
            
            # Create page image
            page_img = Image.new('RGB', (preview_page_width, preview_page_height), 'white')
            draw = ImageDraw.Draw(page_img)
            
            # Calculate ticket size in preview
            ticket_w = preview_page_width // cols
            ticket_h = preview_page_height // rows
            
            # Load small font for labels
            try:
                label_font = ImageFont.truetype("arial.ttf", 10)
            except:
                try:
                    label_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 10)
                except:
                    label_font = ImageFont.load_default()
            
            # Create a small version of the voucher image
            if self.voucher_image:
                ticket_img = self.resize_image_to_fill(self.voucher_image, ticket_w, ticket_h)
            else:
                ticket_img = None
            
            # Draw tickets
            attendee_num = 1
            ticket_count = 0
            
            for row in range(rows):
                for col in range(cols):
                    x = col * ticket_w
                    y = row * ticket_h
                    
                    if ticket_img:
                        page_img.paste(ticket_img, (x, y))
                    else:
                        # Draw placeholder rectangle
                        draw.rectangle([x, y, x + ticket_w - 1, y + ticket_h - 1], 
                                       fill='#f5f5dc', outline='#ccc')
                    
                    # Draw grid lines
                    draw.rectangle([x, y, x + ticket_w - 1, y + ticket_h - 1], outline='#999')
                    
                    # Draw attendee label
                    label = f"Attendee {attendee_num}"
                    label_bbox = draw.textbbox((0, 0), label, font=label_font)
                    label_width = label_bbox[2] - label_bbox[0]
                    label_x = x + (ticket_w - label_width) // 2
                    label_y = y + ticket_h // 2 - 5
                    
                    # Draw label with background for readability
                    padding = 2
                    draw.rectangle([label_x - padding, label_y - padding, 
                                   label_x + label_width + padding, label_y + 12 + padding],
                                  fill='white', outline='#999')
                    draw.text((label_x, label_y), label, fill='#333', font=label_font)
                    
                    ticket_count += 1
                    if ticket_count >= tickets_per_attendee:
                        ticket_count = 0
                        attendee_num += 1
            
            # Draw page border
            draw.rectangle([0, 0, preview_page_width - 1, preview_page_height - 1], outline='#333', width=2)
            
            self.preview_photo = ImageTk.PhotoImage(page_img)
            
            self.preview_canvas.delete("all")
            x = (canvas_width - preview_page_width) // 2
            y = (canvas_height - preview_page_height) // 2
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_photo)
            
        except Exception as e:
            print(f"Layout preview error: {e}")
            traceback.print_exc()
    
    def resize_image_to_fill(self, img, target_width, target_height):
        """Resize and crop image to fill exact dimensions (cover mode)"""
        img_ratio = img.width / img.height
        target_ratio = target_width / target_height
        
        if img_ratio > target_ratio:
            # Image is wider - scale by height and crop width
            new_height = target_height
            new_width = int(target_height * img_ratio)
        else:
            # Image is taller - scale by width and crop height
            new_width = target_width
            new_height = int(target_width / img_ratio)
        
        # Resize
        resized = img.resize((new_width, new_height), Image.LANCZOS)
        
        # Crop to center
        left = (new_width - target_width) // 2
        top = (new_height - target_height) // 2
        right = left + target_width
        bottom = top + target_height
        
        return resized.crop((left, top, right, bottom))
    
    def generate_pdf(self):
        """Handle generate button click"""
        print("Generate PDF clicked!")
        
        if not self.csv_path or not self.image_path:
            messagebox.showwarning("Missing Files", "Please select both a CSV file and an image first.")
            return
            
        if not self.attendees:
            messagebox.showwarning("No Attendees", "No attendees found in the CSV file.")
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                title="Save PDF as...",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialfile="drink_vouchers.pdf"
            )
            
            print(f"Save path: {output_path}")
            
            if not output_path:
                print("No output path selected")
                return
            
            self.status_label.config(text="Generating PDF...", foreground="blue")
            self.root.update()
            
            self.create_vouchers_pdf(output_path)
            
            tickets_per_attendee = int(self.tickets_per_attendee_var.get())
            total_tickets = len(self.attendees) * tickets_per_attendee
            self.status_label.config(text=f"✓ Success! Created {total_tickets} vouchers", foreground="green")
            messagebox.showinfo("Success", f"Created {total_tickets} vouchers!\n\nSaved to:\n{output_path}")
            
        except Exception as e:
            error_msg = traceback.format_exc()
            print(f"Error: {error_msg}")
            self.status_label.config(text="Error creating PDF", foreground="red")
            messagebox.showerror("Error", f"Could not create PDF:\n{e}\n\nDetails:\n{error_msg}")
            
    def create_vouchers_pdf(self, output_path):
        """Generate the PDF with all vouchers - all tickets touching"""
        
        title_text = self.title_var.get().strip()
        font_size = int(self.font_size_var.get())
        bold = self.bold_var.get() == 1
        
        cols = int(self.columns_var.get())
        rows = int(self.rows_var.get())
        tickets_per_attendee = int(self.tickets_per_attendee_var.get())
        
        page_width, page_height = self.get_page_dimensions()
        pagesize = (page_width, page_height)
        
        ticket_width = page_width / cols
        ticket_height = page_height / rows
        
        tickets_per_page = cols * rows
        attendees_per_page = tickets_per_page // tickets_per_attendee
        
        print(f"Creating PDF: {cols}x{rows} grid, {tickets_per_attendee} tickets/attendee")
        print(f"Ticket size: {ticket_width/inch:.2f}\" x {ticket_height/inch:.2f}\"")
        
        # Prepare image - resize to match ticket aspect ratio
        img_ratio = ticket_width / ticket_height
        temp_img = self.resize_image_to_fill(self.voucher_image, 
                                              int(ticket_width * 2), 
                                              int(ticket_height * 2))
        temp_path = os.path.join(os.path.dirname(output_path), "_temp_voucher.png")
        temp_img.save(temp_path, "PNG")
        img_reader = ImageReader(temp_path)
        
        c = canvas.Canvas(output_path, pagesize=pagesize)
        
        # Track position on page
        current_ticket = 0
        
        for attendee in self.attendees:
            first_name, last_name = self.parse_name(attendee)
            
            for ticket_num in range(tickets_per_attendee):
                # Check if we need a new page
                if current_ticket >= tickets_per_page:
                    c.showPage()
                    current_ticket = 0
                
                # Calculate position
                col = current_ticket % cols
                row = current_ticket // cols
                
                x = col * ticket_width
                y = page_height - (row + 1) * ticket_height
                
                # Draw ticket image
                c.drawImage(img_reader, x, y, width=ticket_width, height=ticket_height,
                           preserveAspectRatio=False, mask='auto')
                
                # Draw text
                center_x = x + (ticket_width / 2)
                center_y = y + (ticket_height / 2)
                
                c.setFillColorRGB(0.2, 0.15, 0.1)
                
                if bold:
                    name_font = "Helvetica-Bold"
                else:
                    name_font = "Helvetica"
                
                # Scale font based on ticket size
                size_factor = min(ticket_width / (3 * inch), ticket_height / (1.75 * inch))
                scaled_font_size = int(font_size * size_factor)
                title_font_size = int(scaled_font_size * 0.8)
                lastname_font_size = int(scaled_font_size * 0.9)
                
                if title_text:
                    c.setFont("Helvetica-Bold", title_font_size)
                    c.drawCentredString(center_x, center_y + 22 * size_factor, title_text.upper())
                    
                    c.setFont(name_font, scaled_font_size)
                    c.drawCentredString(center_x, center_y + 2 * size_factor, first_name)
                    
                    c.setFont(name_font, lastname_font_size)
                    c.drawCentredString(center_x, center_y - 14 * size_factor, last_name)
                else:
                    c.setFont(name_font, scaled_font_size)
                    c.drawCentredString(center_x, center_y + 8 * size_factor, first_name)
                    
                    c.setFont(name_font, lastname_font_size)
                    c.drawCentredString(center_x, center_y - 10 * size_factor, last_name)
                
                current_ticket += 1
        
        c.save()
        
        # Clean up temp file
        try:
            os.remove(temp_path)
        except:
            pass
        
        print(f"PDF saved to: {output_path}")


def main():
    root = tk.Tk()
    app = VoucherGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
