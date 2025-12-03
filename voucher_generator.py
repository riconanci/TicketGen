#!/usr/bin/env python3
"""
Drink Voucher Generator
GUI app with live preview for creating personalized drink vouchers.
"""

import csv
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import traceback


class VoucherGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drink Voucher Generator")
        self.root.geometry("600x750")
        self.root.resizable(False, False)
        
        # Variables
        self.csv_path = None
        self.image_path = None
        self.attendees = []
        self.voucher_image = None
        
        # Settings variables
        self.title_var = tk.StringVar(value="DRINK VOUCHER")
        self.font_size_var = tk.StringVar(value="12")
        self.bold_var = tk.IntVar(value=1)  # Use IntVar for checkbox (1=checked, 0=unchecked)
        
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
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(preview_frame, width=400, height=250, bg="#f0f0f0", relief="sunken", bd=2)
        self.preview_canvas.pack(pady=10)
        self.preview_canvas.create_text(200, 125, text="Select an image to see preview", fill="gray")
        
        # === SETTINGS SECTION ===
        settings_frame = ttk.LabelFrame(main_frame, text="Step 2: Customize", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Title field
        title_row = ttk.Frame(settings_frame)
        title_row.pack(fill=tk.X, pady=5)
        ttk.Label(title_row, text="Title:", width=12).pack(side=tk.LEFT)
        title_entry = ttk.Entry(title_row, textvariable=self.title_var, width=30)
        title_entry.pack(side=tk.LEFT, padx=(5, 0))
        title_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        ttk.Label(title_row, text="(leave empty for no title)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        
        # Font size dropdown
        font_row = ttk.Frame(settings_frame)
        font_row.pack(fill=tk.X, pady=5)
        ttk.Label(font_row, text="Font Size:", width=12).pack(side=tk.LEFT)
        self.font_combo = ttk.Combobox(font_row, textvariable=self.font_size_var, 
                                        values=["8", "9", "10", "11", "12", "13", "14", "16", "18"], 
                                        width=5, state="readonly")
        self.font_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.font_combo.set("12")
        self.font_combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Bold checkbox
        bold_row = ttk.Frame(settings_frame)
        bold_row.pack(fill=tk.X, pady=5)
        ttk.Label(bold_row, text="Bold Names:", width=12).pack(side=tk.LEFT)
        self.bold_check = tk.Checkbutton(bold_row, variable=self.bold_var, command=self.update_preview)
        self.bold_check.pack(side=tk.LEFT, padx=(5, 0))
        self.bold_check.select()  # Start checked
        
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
            self.status_label.config(text=f"✓ Ready! {len(self.attendees)} attendees loaded", foreground="green")
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
        """Update the preview canvas with current settings"""
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
            
            preview_img = self.voucher_image.copy()
            preview_img.thumbnail((380, 230), Image.LANCZOS)
            
            draw = ImageDraw.Draw(preview_img)
            
            # Try to load fonts
            try:
                if bold:
                    name_font = ImageFont.truetype("arialbd.ttf", font_size * 2)
                else:
                    name_font = ImageFont.truetype("arial.ttf", font_size * 2)
                title_font = ImageFont.truetype("arialbd.ttf", int(font_size * 1.5))
            except:
                try:
                    if bold:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size * 2)
                    else:
                        name_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size * 2)
                    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", int(font_size * 1.5))
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
            draw.text((center_x - ln_width // 2, center_y + name_y_offset + font_size * 2 + 5), last_name, fill=text_color, font=name_font)
            
            self.preview_photo = ImageTk.PhotoImage(preview_img)
            
            self.preview_canvas.delete("all")
            canvas_width = 400
            canvas_height = 250
            x = (canvas_width - preview_img.width) // 2
            y = (canvas_height - preview_img.height) // 2
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_photo)
        except Exception as e:
            print(f"Preview error: {e}")
    
    def generate_pdf(self):
        """Handle generate button click"""
        print("Generate PDF clicked!")  # Debug
        
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
            
            print(f"Save path: {output_path}")  # Debug
            
            if not output_path:
                print("No output path selected")  # Debug
                return
            
            self.status_label.config(text="Generating PDF...", foreground="blue")
            self.root.update()
            
            self.create_vouchers_pdf(output_path)
            
            self.status_label.config(text=f"✓ Success! Created {len(self.attendees) * 5} vouchers", foreground="green")
            messagebox.showinfo("Success", f"Created {len(self.attendees) * 5} vouchers!\n\nSaved to:\n{output_path}")
            
        except Exception as e:
            error_msg = traceback.format_exc()
            print(f"Error: {error_msg}")  # Debug
            self.status_label.config(text="Error creating PDF", foreground="red")
            messagebox.showerror("Error", f"Could not create PDF:\n{e}\n\nDetails:\n{error_msg}")
            
    def create_vouchers_pdf(self, output_path):
        """Generate the PDF with all vouchers - ALL tickets touching"""
        
        title_text = self.title_var.get().strip()
        font_size = int(self.font_size_var.get())
        bold = self.bold_var.get() == 1
        
        print(f"Creating PDF: title='{title_text}', font_size={font_size}, bold={bold}")  # Debug
        
        # Ticket dimensions (3" x 1.75")
        ticket_width = 3 * inch
        ticket_height = 1.75 * inch
        
        page_width, page_height = letter
        
        # Center 2 columns - tickets touching horizontally
        total_width = 2 * ticket_width
        margin_x = (page_width - total_width) / 2
        col1_x = margin_x
        col2_x = margin_x + ticket_width  # No gap - touching
        
        # Start from very top
        margin_top = 0.25 * inch
        
        img_reader = ImageReader(self.voucher_image)
        c = canvas.Canvas(output_path, pagesize=letter)
        
        for attendee_idx, attendee in enumerate(self.attendees):
            first_name, last_name = self.parse_name(attendee)
            
            position_on_page = attendee_idx % 2
            
            # Start new page for every 2 attendees (after the first page)
            if attendee_idx > 0 and position_on_page == 0:
                c.showPage()
            
            # Which rows this attendee uses
            start_row = position_on_page * 3
            
            # All tickets touching - no gaps
            ticket_positions = [
                (col1_x, page_height - margin_top - (start_row + 1) * ticket_height),
                (col2_x, page_height - margin_top - (start_row + 1) * ticket_height),
                (col1_x, page_height - margin_top - (start_row + 2) * ticket_height),
                (col2_x, page_height - margin_top - (start_row + 2) * ticket_height),
                (col1_x, page_height - margin_top - (start_row + 3) * ticket_height),
            ]
            
            for x, y in ticket_positions:
                c.drawImage(img_reader, x, y, width=ticket_width, height=ticket_height,
                           preserveAspectRatio=True, mask='auto')
                
                center_x = x + (ticket_width / 2)
                center_y = y + (ticket_height / 2)
                
                c.setFillColorRGB(0.2, 0.15, 0.1)
                
                # Set font based on bold setting
                if bold:
                    name_font = "Helvetica-Bold"
                else:
                    name_font = "Helvetica"
                
                title_font_size = int(font_size * 0.8)
                name_font_size = font_size
                lastname_font_size = int(font_size * 0.9)
                
                if title_text:
                    c.setFont("Helvetica-Bold", title_font_size)
                    c.drawCentredString(center_x, center_y + 22, title_text.upper())
                    
                    c.setFont(name_font, name_font_size)
                    c.drawCentredString(center_x, center_y + 2, first_name)
                    
                    c.setFont(name_font, lastname_font_size)
                    c.drawCentredString(center_x, center_y - 14, last_name)
                else:
                    c.setFont(name_font, name_font_size)
                    c.drawCentredString(center_x, center_y + 8, first_name)
                    
                    c.setFont(name_font, lastname_font_size)
                    c.drawCentredString(center_x, center_y - 10, last_name)
        
        c.save()
        print(f"PDF saved to: {output_path}")  # Debug


def main():
    root = tk.Tk()
    app = VoucherGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
