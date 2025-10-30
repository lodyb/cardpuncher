#!/usr/bin/env python3
"""
Trading Card Print Layout Tool

Generates print-ready PDF sheets of trading cards arranged in a 3x3 grid 
on A4 pages, optimized for machine cutting with color-coded alignment markers.
"""

import argparse
import os
import sys
import glob
import re
from pathlib import Path
from typing import List, Optional, Tuple
from dataclasses import dataclass

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.pdfdoc import PDFDictionary
    from PIL import Image
except ImportError as e:
    print(f"Error: Missing required dependency: {e}")
    print("Please install required packages: pip install reportlab Pillow")
    sys.exit(1)


# =============================================================================
# CONFIGURATION
# =============================================================================

@dataclass
class Config:
    card_width_mm: float = 63.0
    card_height_mm: float = 88.0
    bleed_mm: float = 3.175
    spacing_mm: float = 0.0
    margin_mm: float = 0.0
    dpi: int = 1200
    
    grid_cols: int = 3
    grid_rows: int = 3
    
    marker_distance_mm: float = 10.0
    marker_size_mm: float = 2.0
    marker_line_width_mm: float = 0.3
    
    guide_line_width_mm: float = 0.1
    
    left_marker_color: tuple = (1, 0, 1)  # Magenta
    right_marker_color: tuple = (0, 1, 0)  # Green
    guide_line_color: tuple = (0, 1, 1)   # Cyan
    
    supported_extensions: tuple = ('.png', '.jpg', '.jpeg', '.tiff', '.webp')
    inch_to_mm: float = 25.4
    
    @classmethod
    def from_directory_name(cls, folder_path: str, **overrides) -> 'Config':
        config = cls(**overrides)
        folder_name = os.path.basename(os.path.normpath(folder_path))
        
        match = re.match(r'^b(\d+(?:\.\d+)?)_m(\d+(?:\.\d+)?)_(.+)$', folder_name)
        if not match:
            return config
            
        bleed_val, margin_val = float(match.group(1)), float(match.group(2))
        
        if bleed_val < 1.0:
            config.bleed_mm = bleed_val * config.inch_to_mm
            config.margin_mm = margin_val * config.inch_to_mm
            print(f"Auto-detected: Bleed {bleed_val}\" ({config.bleed_mm:.3f}mm), Margin {margin_val}\" ({config.margin_mm:.3f}mm)")
        else:
            config.bleed_mm = bleed_val
            config.margin_mm = margin_val
            print(f"Auto-detected: Bleed {bleed_val}mm, Margin {margin_val}mm")
            
        return config


@dataclass
class Layout:
    start_x: float
    start_y: float
    card_width: float
    card_height: float
    placed_width: float
    placed_height: float


# =============================================================================
# CORE FUNCTIONALITY
# =============================================================================

class PDFGenerator:
    def __init__(self, config: Config):
        self.config = config
        self.page_width, self.page_height = A4
    
    def mm_to_points(self, mm_val: float) -> float:
        return mm_val * mm
        
    def find_images(self, folder_path: str) -> List[str]:
        files = []
        for ext in self.config.supported_extensions:
            files.extend(glob.glob(os.path.join(folder_path, f"*{ext}"), recursive=False))
            files.extend(glob.glob(os.path.join(folder_path, f"*{ext.upper()}"), recursive=False))
        return sorted(list(set(files)))
    
    def process_image(self, image_path: str) -> Optional[Tuple[Image.Image, Image.Image]]:
        """Process image with mirrored edge bleed.
        
        Creates a bleed area by mirroring the edges of the card image outward.
        This creates a natural-looking bleed that extends the card design.
        
        Returns:
            Tuple of (bleed_image, card_image) or None if error
        """
        card_w = int(self.config.card_width_mm * self.config.dpi / self.config.inch_to_mm)
        card_h = int(self.config.card_height_mm * self.config.dpi / self.config.inch_to_mm)
        
        bleed_px = int(self.config.bleed_mm * self.config.dpi / self.config.inch_to_mm)
        bleed_w = card_w + 2 * bleed_px
        bleed_h = card_h + 2 * bleed_px
        
        try:
            with Image.open(image_path) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize to exact card dimensions
                card_img = img.resize((card_w, card_h), Image.Resampling.LANCZOS).copy()
                
                # Create bleed image with mirrored edges
                bleed_img = Image.new('RGB', (bleed_w, bleed_h))
                
                # Place main card image in center
                bleed_img.paste(card_img, (bleed_px, bleed_px))
                
                # Mirror edges outward
                # Left edge
                left_strip = card_img.crop((0, 0, bleed_px, card_h))
                bleed_img.paste(left_strip.transpose(Image.FLIP_LEFT_RIGHT), (0, bleed_px))
                
                # Right edge
                right_strip = card_img.crop((card_w - bleed_px, 0, card_w, card_h))
                bleed_img.paste(right_strip.transpose(Image.FLIP_LEFT_RIGHT), (bleed_px + card_w, bleed_px))
                
                # Top edge
                top_strip = card_img.crop((0, 0, card_w, bleed_px))
                bleed_img.paste(top_strip.transpose(Image.FLIP_TOP_BOTTOM), (bleed_px, 0))
                
                # Bottom edge
                bottom_strip = card_img.crop((0, card_h - bleed_px, card_w, card_h))
                bleed_img.paste(bottom_strip.transpose(Image.FLIP_TOP_BOTTOM), (bleed_px, bleed_px + card_h))
                
                # Corners (mirrored in both directions)
                # Top-left corner
                tl_corner = card_img.crop((0, 0, bleed_px, bleed_px))
                bleed_img.paste(tl_corner.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM), (0, 0))
                
                # Top-right corner
                tr_corner = card_img.crop((card_w - bleed_px, 0, card_w, bleed_px))
                bleed_img.paste(tr_corner.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM), (bleed_px + card_w, 0))
                
                # Bottom-left corner
                bl_corner = card_img.crop((0, card_h - bleed_px, bleed_px, card_h))
                bleed_img.paste(bl_corner.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM), (0, bleed_px + card_h))
                
                # Bottom-right corner
                br_corner = card_img.crop((card_w - bleed_px, card_h - bleed_px, card_w, card_h))
                bleed_img.paste(br_corner.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM), (bleed_px + card_w, bleed_px + card_h))
                
                return (bleed_img, card_img)
        except (IOError, OSError):
            return None
    
    def draw_card_with_bleed(self, canvas_obj: canvas.Canvas, bleed_img: Image.Image,
                            card_img: Image.Image, x: float, y: float, layout: Layout):
        """Draw scaled image for bleed area, then exact card image centered on top."""
        bleed = self.mm_to_points(self.config.bleed_mm)
        
        # Draw scaled image to fill entire bleed area (background)
        canvas_obj.drawImage(ImageReader(bleed_img), x, y,
                           width=layout.placed_width, height=layout.placed_height)
        
        # Draw exact card image centered on top (offset by bleed on all sides)
        canvas_obj.drawImage(ImageReader(card_img), 
                           x + bleed, y + bleed,
                           width=layout.card_width, 
                           height=layout.card_height)
    
    def calculate_layout(self) -> Layout:
        placed_w = self.config.card_width_mm + 2 * self.config.bleed_mm
        placed_h = self.config.card_height_mm + 2 * self.config.bleed_mm
        
        grid_w = self.config.grid_cols * placed_w + (self.config.grid_cols - 1) * self.config.spacing_mm
        grid_h = self.config.grid_rows * placed_h + (self.config.grid_rows - 1) * self.config.spacing_mm
        
        page_w_mm, page_h_mm = self.page_width / mm, self.page_height / mm
        center_x = (page_w_mm - grid_w) / 2
        center_y = (page_h_mm - grid_h) / 2
        
        return Layout(
            start_x=self.mm_to_points(center_x),
            start_y=self.mm_to_points(center_y),
            card_width=self.mm_to_points(self.config.card_width_mm),
            card_height=self.mm_to_points(self.config.card_height_mm),
            placed_width=self.mm_to_points(placed_w),
            placed_height=self.mm_to_points(placed_h)
        )
    
    def draw_marker(self, canvas_obj: canvas.Canvas, x: float, y: float, size: float):
        canvas_obj.line(x - size/2, y, x + size/2, y)
        canvas_obj.line(x, y - size/2, x, y + size/2)
    
    def draw_guides(self, canvas_obj: canvas.Canvas, layout: Layout):
        spacing = self.mm_to_points(self.config.spacing_mm)
        
        # Cutting guide lines
        canvas_obj.setStrokeColorRGB(*self.config.guide_line_color)
        canvas_obj.setLineWidth(self.mm_to_points(self.config.guide_line_width_mm))
        
        for col in range(1, self.config.grid_cols):
            x = layout.start_x + col * (layout.placed_width + spacing)
            canvas_obj.line(x, 0, x, self.page_height)
            
        for row in range(1, self.config.grid_rows):
            y = layout.start_y + row * (layout.placed_height + spacing)
            canvas_obj.line(0, y, self.page_width, y)
        
        # Alignment markers
        canvas_obj.setLineWidth(self.mm_to_points(self.config.marker_line_width_mm))
        marker_dist = self.mm_to_points(self.config.marker_distance_mm)
        marker_size = self.mm_to_points(self.config.marker_size_mm)
        
        for row in range(self.config.grid_rows):
            for col in range(self.config.grid_cols):
                card_x = layout.start_x + col * (layout.placed_width + spacing)
                card_y = layout.start_y + row * (layout.placed_height + spacing)
                
                left_x = card_x - marker_dist
                right_x = card_x + layout.placed_width + marker_dist
                top_y = card_y + layout.placed_height
                bottom_y = card_y
                
                if left_x > 0:
                    canvas_obj.setStrokeColorRGB(*self.config.left_marker_color)
                    self.draw_marker(canvas_obj, left_x, top_y, marker_size)
                    self.draw_marker(canvas_obj, left_x, bottom_y, marker_size)
                
                if right_x < self.page_width:
                    canvas_obj.setStrokeColorRGB(*self.config.right_marker_color)
                    self.draw_marker(canvas_obj, right_x, top_y, marker_size)
                    self.draw_marker(canvas_obj, right_x, bottom_y, marker_size)
    
    def embed_icc_profile(self, canvas_obj: canvas.Canvas, icc_path: str) -> bool:
        try:
            with open(icc_path, 'rb') as f:
                icc_data = f.read()
            
            output_intent = PDFDictionary({
                'Type': '/OutputIntent',
                'S': '/GTS_PDFX',
                'OutputConditionIdentifier': 'sRGB',
                'Info': 'sRGB color space',
                'OutputCondition': 'sRGB'
            })
            
            canvas_obj._doc.Reference(output_intent)
            return True
        except Exception:
            return False
    
    def generate_pdf(self, folder_path: str, output_path: str, icc_path: Optional[str] = None) -> str:
        images = self.find_images(folder_path)
        if not images:
            raise FileNotFoundError(f"No supported images found in {folder_path}")
        
        layout = self.calculate_layout()
        canvas_obj = canvas.Canvas(output_path, pagesize=A4)
        
        if icc_path and os.path.exists(icc_path):
            self.embed_icc_profile(canvas_obj, icc_path)
        
        cards_per_page = self.config.grid_cols * self.config.grid_rows
        pages = (len(images) + cards_per_page - 1) // cards_per_page
        spacing = self.mm_to_points(self.config.spacing_mm)
        
        for page in range(pages):
            start_idx = page * cards_per_page
            page_images = images[start_idx:start_idx + cards_per_page]
            
            for i, img_path in enumerate(page_images):
                row, col = divmod(i, self.config.grid_cols)
                x = layout.start_x + col * (layout.placed_width + spacing)
                y = layout.start_y + (self.config.grid_rows - 1 - row) * (layout.placed_height + spacing)
                
                processed_imgs = self.process_image(img_path)
                if processed_imgs:
                    bleed_img, card_img = processed_imgs
                    self.draw_card_with_bleed(canvas_obj, bleed_img, card_img, x, y, layout)
            
            self.draw_guides(canvas_obj, layout)
            
            if page < pages - 1:
                canvas_obj.showPage()
        
        canvas_obj.save()
        return output_path


# =============================================================================
# CLI INTERFACE
# =============================================================================

def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate print-ready PDF sheets of trading cards in 3x3 grid layout",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pdf.py "my_deck"
  python pdf.py "b0.125_m0_my_deck"  # Auto-detect 1/8" bleed, 0 margin
  python pdf.py "b2_m1_my_deck"      # Auto-detect 2mm bleed, 1mm margin

Directory naming: b{bleed}_m{margin}_{name}
  Values < 1.0 = inches, >= 1.0 = millimeters
        """
    )
    
    parser.add_argument('folder_path', help='Path to folder containing card images')
    parser.add_argument('--card-w-mm', type=float, default=63.0, help='Card width (mm)')
    parser.add_argument('--card-h-mm', type=float, default=88.0, help='Card height (mm)')
    parser.add_argument('--bleed-mm', type=float, help='Bleed area (mm)')
    parser.add_argument('--margin-mm', type=float, help='Page margin (mm)')
    parser.add_argument('--dpi', type=int, default=1200, help='Image DPI')
    parser.add_argument('--icc', type=str, default="sRGB.icc", help='ICC profile path')
    parser.add_argument('--card-spacing-mm', type=float, default=0.0, help='Card spacing (mm)')
    parser.add_argument('--no-auto-detect', action='store_true', help='Disable auto-detection')
    
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    
    # Validate inputs
    for val, name in [(args.card_w_mm, "card width"), (args.card_h_mm, "card height"), (args.dpi, "DPI")]:
        if val <= 0:
            print(f"Error: {name} must be positive")
            sys.exit(1)
    
    for val, name in [(args.bleed_mm, "bleed"), (args.margin_mm, "margin"), (args.card_spacing_mm, "spacing")]:
        if val is not None and val < 0:
            print(f"Error: {name} cannot be negative")
            sys.exit(1)
    
    # Create configuration
    overrides = {
        'card_width_mm': args.card_w_mm,
        'card_height_mm': args.card_h_mm,
        'spacing_mm': args.card_spacing_mm,
        'dpi': args.dpi
    }
    
    if args.bleed_mm is not None:
        overrides['bleed_mm'] = args.bleed_mm
    if args.margin_mm is not None:
        overrides['margin_mm'] = args.margin_mm
    
    if args.no_auto_detect:
        config = Config(**overrides)
        if args.bleed_mm is None:
            config.bleed_mm = 3.175
        if args.margin_mm is None:
            config.margin_mm = 0.0
    else:
        config = Config.from_directory_name(args.folder_path, **overrides)
    
    # Generate PDF
    output_path = f"{os.path.basename(os.path.normpath(args.folder_path))}.pdf"
    icc_path = args.icc if os.path.exists(args.icc) else None
    
    try:
        generator = PDFGenerator(config)
        generator.generate_pdf(args.folder_path, output_path, icc_path)
        
        print(f"\nSuccess! Generated {output_path}")
        print(f"Card size: {config.card_width_mm}x{config.card_height_mm}mm")
        print(f"With bleed: {config.card_width_mm + 2*config.bleed_mm}x{config.card_height_mm + 2*config.bleed_mm}mm")
        print(f"Spacing: {config.spacing_mm}mm (perfect grid)")
        print(f"Markers: Left (magenta) and right (green) at {config.marker_distance_mm}mm offset")
        print(f"Guides: Cyan cutting lines for strip slicing")
        
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()