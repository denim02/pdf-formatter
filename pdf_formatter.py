import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch, mm
import io
import fitz  # PyMuPDF
import os

class ExamReformatter:
    def __init__(self, input_pdf_path, output_pdf_path):
        self.input_pdf = input_pdf_path
        self.output_pdf = output_pdf_path
        self.exams = {}  # Dictionary to store exam pages
        # Use landscape A4
        self.page_width = landscape(A4)[0]
        self.page_height = landscape(A4)[1]
        self.grid_cols = 3
        self.grid_rows = 2
        self.margin = 10 * mm  # Smaller margins for landscape A4
        
    def parse_pdf(self):
        """Parse the PDF and group pages by exam number."""
        with open(self.input_pdf, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            current_exam = None
            
            # Iterate through all pages
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                
                # Look for "TEZA" followed by a number
                import re
                teza_match = re.search(r'TEZA\s*(\d+)', text, re.IGNORECASE)
                
                if teza_match:
                    # Found a new exam
                    current_exam = int(teza_match.group(1))
                    if current_exam not in self.exams:
                        self.exams[current_exam] = []
                    self.exams[current_exam].append(page_num)
                elif current_exam is not None:
                    # This page belongs to the current exam
                    self.exams[current_exam].append(page_num)
    
    def create_blank_page(self):
        """Create a blank PDF page."""
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=landscape(A4))
        c.drawString(100, 100, " ")
        c.showPage()
        c.save()
        packet.seek(0)
        return packet
    
    def create_grid_page(self, exam_numbers, page_indices, title=""):
        """Create a page with exams in a grid layout."""
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=landscape(A4))
        
        # Calculate cell dimensions
        cell_width = (self.page_width - 2 * self.margin) / self.grid_cols
        cell_height = (self.page_height - 2 * self.margin) / self.grid_rows
        
        # Add title if provided (also make this smaller)
        if title:
            c.setFont("Helvetica", 6)
            c.setFillColorRGB(0.7, 0.7, 0.7)  # Light gray
            c.drawString(self.margin, self.page_height - 10, title)
        
        pdf_document = fitz.open(self.input_pdf)
        has_content = False
        
        for idx, (exam_num, page_idx) in enumerate(zip(exam_numbers, page_indices)):
            if exam_num is None or page_idx is None:
                continue
                
            row = idx // self.grid_cols
            col = idx % self.grid_cols
            
            # Calculate position
            x = self.margin + col * cell_width
            y = self.page_height - self.margin - (row + 1) * cell_height
            
            # Draw cell border
            c.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray border
            c.setLineWidth(0.5)
            c.rect(x, y, cell_width, cell_height, stroke=1, fill=0)
            
            # Add exam number label - MUCH SMALLER
            c.setFont("Helvetica", 5)  # Tiny font
            c.setFillColorRGB(0.7, 0.7, 0.7)  # Light gray text
            label_height = 8  # Much smaller label space
            c.drawString(x + 2, y + cell_height - label_height + 2, f"E{exam_num}-P{page_idx - self.exams[exam_num][0] + 1}")
            
            # Convert PDF page to image
            has_content = True
            page = pdf_document[page_idx]
            
            # Use higher resolution for better quality
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            img_data = pix.tobytes("png")
            
            # Save temporary image
            temp_img = f"temp_exam_{exam_num}_page_{page_idx}.png"
            with open(temp_img, "wb") as img_file:
                img_file.write(img_data)
            
            try:
                # Draw image filling the entire cell (minus tiny label space)
                # No padding - use full cell space
                img_x = x + 1  # 1 point border
                img_y = y + 1
                img_width = cell_width - 2
                img_height = cell_height - label_height - 2
                
                c.drawImage(temp_img, img_x, img_y, 
                           width=img_width, 
                           height=img_height,
                           preserveAspectRatio=True, mask='auto')
            finally:
                try:
                    os.remove(temp_img)
                except:
                    pass
        
        c.showPage()
        c.save()
        packet.seek(0)
        pdf_document.close()
        return packet, has_content
    
    def reformat_pdf(self):
        """Main method to reformat the entire PDF."""
        self.parse_pdf()
        
        print(f"\nFound {len(self.exams)} exams in the PDF")
        
        # Categorize exams by number of pages
        one_page_exams = []
        two_page_exams = []
        three_plus_page_exams = []
        
        for exam_num, pages in sorted(self.exams.items()):
            if len(pages) == 1:
                one_page_exams.append(exam_num)
            elif len(pages) == 2:
                two_page_exams.append(exam_num)
            else:
                three_plus_page_exams.append(exam_num)
        
        print(f"  1-page exams: {len(one_page_exams)}")
        print(f"  2-page exams: {len(two_page_exams)}")
        print(f"  3+ page exams: {len(three_plus_page_exams)} - {three_plus_page_exams}")
        
        # Sort exam numbers
        exam_numbers = sorted(self.exams.keys())
        
        # Create output PDF
        output_writer = PyPDF2.PdfWriter()
        
        # Process exams in batches of 6
        batch_count = 0
        for i in range(0, len(exam_numbers), 6):
            batch = exam_numbers[i:i+6]
            batch_count += 1
            
            # Pad with None if less than 6
            while len(batch) < 6:
                batch.append(None)
            
            print(f"\nProcessing batch {batch_count}: Exams {[b for b in batch if b is not None]}")
            
            # SHEET 1: First pages (front side)
            page_indices = []
            for exam_num in batch:
                if exam_num and exam_num in self.exams:
                    page_indices.append(self.exams[exam_num][0])  # First page
                else:
                    page_indices.append(None)
            
            front_packet, _ = self.create_grid_page(batch, page_indices, f"Batch {batch_count} - First Pages")
            
            # SHEET 1: Second pages (back side) - reversed for double-sided printing
            back_order = [
                batch[2], batch[1], batch[0],  # First row reversed
                batch[5], batch[4], batch[3]   # Second row reversed
            ]
            page_indices = []
            for exam_num in back_order:
                if exam_num and exam_num in self.exams and len(self.exams[exam_num]) > 1:
                    page_indices.append(self.exams[exam_num][1])  # Second page
                else:
                    page_indices.append(None)
            
            back_packet, has_second_pages = self.create_grid_page(back_order, page_indices, f"Batch {batch_count} - Second Pages")
            
            # Add first sheet (pages 1-2)
            try:
                front_reader = PyPDF2.PdfReader(front_packet)
                output_writer.add_page(front_reader.pages[0])
                
                if has_second_pages:
                    back_reader = PyPDF2.PdfReader(back_packet)
                    output_writer.add_page(back_reader.pages[0])
                else:
                    # Add blank back page
                    blank_packet = self.create_blank_page()
                    blank_reader = PyPDF2.PdfReader(blank_packet)
                    output_writer.add_page(blank_reader.pages[0])
                    
            except Exception as e:
                print(f"Error processing batch: {e}")
                continue
        
        # ADDITIONAL SHEETS: Third pages for exams that have them
        # Process third pages in groups of 6
        third_page_exams = [exam for exam in exam_numbers if len(self.exams[exam]) >= 3]
        
        if third_page_exams:
            print(f"\nProcessing third pages for {len(third_page_exams)} exams...")
            
            for i in range(0, len(third_page_exams), 6):
                batch = third_page_exams[i:i+6]
                
                # Pad with None
                while len(batch) < 6:
                    batch.append(None)
                
                # Create page with third pages
                page_indices = []
                for exam_num in batch:
                    if exam_num and len(self.exams[exam_num]) >= 3:
                        page_indices.append(self.exams[exam_num][2])  # Third page
                    else:
                        page_indices.append(None)
                
                third_packet, _ = self.create_grid_page(batch, page_indices, f"Third Pages - Exams {[b for b in batch if b is not None]}")
                
                # Add blank back page for single-sided printing of third pages
                blank_packet = self.create_blank_page()
                
                try:
                    third_reader = PyPDF2.PdfReader(third_packet)
                    output_writer.add_page(third_reader.pages[0])
                    
                    blank_reader = PyPDF2.PdfReader(blank_packet)
                    output_writer.add_page(blank_reader.pages[0])
                except Exception as e:
                    print(f"Error processing third pages: {e}")
        
        # Write output PDF
        with open(self.output_pdf, 'wb') as output_file:
            output_writer.write(output_file)
        
        print(f"\nReformatted PDF saved to: {self.output_pdf}")
        print(f"Total pages in output: {len(output_writer.pages)}")
        print("\nPrinting instructions:")
        print("1. First set of sheets: Contains pages 1 (front) and 2 (back) for all exams")
        print("   - Print double-sided (flip on short edge)")
        print("2. Additional sheets at the end: Contains page 3 for exams that have them")
        print("   - Can be printed single-sided or double-sided")
        print("- Use A4 paper in landscape orientation")
        print("- Each sheet shows 6 exam pages in a 3x2 grid")

# Update with your PDF filename
reformatter = ExamReformatter('Tezat.pdf', 'output_grid_exams_A4_landscape.pdf')
reformatter.reformat_pdf()