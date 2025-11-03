#!/usr/bin/env python3

import argparse
import glob
import os
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    from PIL import Image
    import yaml
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Install: pip install reportlab Pillow PyYAML")
    sys.exit(1)


@dataclass
class Config:
    card_width_mm: float = 63.0
    card_height_mm: float = 88.0
    bleed_mm: float = 1.0
    grid_cols: int = 3
    grid_rows: int = 3
    dpi: int = 600
    corner_bevel_mm: float = 2.0
    corner_line_width_mm: float = 1
    separator_width_mm: float = 0.2
    spacing_mm: float = 0.3
    
    @classmethod
    def from_yaml(cls, yaml_path: str):
        if os.path.exists(yaml_path):
            with open(yaml_path) as f:
                data = yaml.safe_load(f)
                return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
        return cls()


class CardPuncher:
    def __init__(self, config: Config):
        self.config = config
        self.page_width, self.page_height = A4
        self.cache: Dict[str, Tuple[ImageReader, ImageReader]] = {}
    
    def mm_to_pt(self, mm_val: float) -> float:
        return mm_val * mm
        
    def find_images(self, folder: str) -> List[str]:
        patterns = ['*.png', '*.jpg', '*.jpeg']
        files = []
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(folder, pattern)))
            files.extend(glob.glob(os.path.join(folder, pattern.upper())))
        return sorted(set(files))
    
    def create_mirrored_bleed(self, img: Image.Image) -> Image.Image:
        c = self.config
        card_w = int(c.card_width_mm * c.dpi / 25.4)
        card_h = int(c.card_height_mm * c.dpi / 25.4)
        bleed_px = int(c.bleed_mm * c.dpi / 25.4)
        
        card_img = img.resize((card_w, card_h), Image.Resampling.LANCZOS)
        bleed_img = Image.new('RGB', (card_w + 2 * bleed_px, card_h + 2 * bleed_px))
        bleed_img.paste(card_img, (bleed_px, bleed_px))
        
        edges = [
            (card_img.crop((0, 0, bleed_px, card_h)).transpose(Image.FLIP_LEFT_RIGHT), (0, bleed_px)),
            (card_img.crop((card_w - bleed_px, 0, card_w, card_h)).transpose(Image.FLIP_LEFT_RIGHT), (bleed_px + card_w, bleed_px)),
            (card_img.crop((0, 0, card_w, bleed_px)).transpose(Image.FLIP_TOP_BOTTOM), (bleed_px, 0)),
            (card_img.crop((0, card_h - bleed_px, card_w, card_h)).transpose(Image.FLIP_TOP_BOTTOM), (bleed_px, bleed_px + card_h)),
        ]
        
        corners = [
            (0, 0, 0, 0), (card_w - bleed_px, 0, bleed_px + card_w, 0),
            (0, card_h - bleed_px, 0, bleed_px + card_h), (card_w - bleed_px, card_h - bleed_px, bleed_px + card_w, bleed_px + card_h)
        ]
        
        for edge, pos in edges:
            bleed_img.paste(edge, pos)
        
        for x1, y1, x2, y2 in corners:
            corner = card_img.crop((x1, y1, x1 + bleed_px, y1 + bleed_px))
            corner = corner.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM)
            bleed_img.paste(corner, (x2, y2))
        
        return bleed_img
    
    def process_image(self, path: str) -> Optional[Tuple[ImageReader, ImageReader]]:
        if path in self.cache:
            return self.cache[path]
        
        try:
            with Image.open(path) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                c = self.config
                card_w = int(c.card_width_mm * c.dpi / 25.4)
                card_h = int(c.card_height_mm * c.dpi / 25.4)
                
                card_img = img.resize((card_w, card_h), Image.Resampling.LANCZOS)
                bleed_img = self.create_mirrored_bleed(img)
                
                result = (ImageReader(bleed_img), ImageReader(card_img))
                self.cache[path] = result
                return result
        except Exception:
            return None
    
    def draw_corner_guides(self, c: canvas.Canvas, x: float, y: float, w: float, h: float):
        bevel = self.mm_to_pt(self.config.corner_bevel_mm) / 2
        line_w = self.mm_to_pt(self.config.corner_line_width_mm)
        offset = line_w / 2
        
        c.setLineWidth(line_w)
        c.setDash([1, 1])
        
        c.setStrokeColorRGB(0, 1, 1)
        c.line(x, y - offset, x + bevel, y - offset)
        c.line(x - offset, y, x - offset, y + bevel)
        
        c.setStrokeColorRGB(1, 1, 0)
        c.setDash([1, 1], 1)
        c.line(x, y - offset, x + bevel, y - offset)
        c.line(x - offset, y, x - offset, y + bevel)
        
        c.setDash([1, 1], 0)
        c.setStrokeColorRGB(0, 1, 1)
        c.line(x + w - bevel, y - offset, x + w, y - offset)
        c.line(x + w + offset, y, x + w + offset, y + bevel)
        
        c.setStrokeColorRGB(1, 1, 0)
        c.setDash([1, 1], 1)
        c.line(x + w - bevel, y - offset, x + w, y - offset)
        c.line(x + w + offset, y, x + w + offset, y + bevel)
        
        c.setDash([1, 1], 0)
        c.setStrokeColorRGB(0, 1, 1)
        c.line(x, y + h + offset, x + bevel, y + h + offset)
        c.line(x - offset, y + h - bevel, x - offset, y + h)
        
        c.setStrokeColorRGB(1, 1, 0)
        c.setDash([1, 1], 1)
        c.line(x, y + h + offset, x + bevel, y + h + offset)
        c.line(x - offset, y + h - bevel, x - offset, y + h)
        
        c.setDash([1, 1], 0)
        c.setStrokeColorRGB(0, 1, 1)
        c.line(x + w - bevel, y + h + offset, x + w, y + h + offset)
        c.line(x + w + offset, y + h - bevel, x + w + offset, y + h)
        
        c.setStrokeColorRGB(1, 1, 0)
        c.setDash([1, 1], 1)
        c.line(x + w - bevel, y + h + offset, x + w, y + h + offset)
        c.line(x + w + offset, y + h - bevel, x + w + offset, y + h)
        
        c.setDash()
    
    def draw_separators(self, c: canvas.Canvas, layout: dict):
        c.setStrokeColorRGB(0, 1, 1)
        c.setLineWidth(self.mm_to_pt(self.config.separator_width_mm))
        
        spacing = self.mm_to_pt(self.config.spacing_mm)
        placed_w = self.mm_to_pt(self.config.card_width_mm + 2 * self.config.bleed_mm)
        placed_h = self.mm_to_pt(self.config.card_height_mm + 2 * self.config.bleed_mm)
        
        for col in range(1, self.config.grid_cols):
            x = layout['start_x'] + col * (placed_w + spacing)
            c.line(x, 0, x, self.page_height)
            
        for row in range(1, self.config.grid_rows):
            y = layout['start_y'] + row * (placed_h + spacing)
            c.line(0, y, self.page_width, y)
    
    def draw_header(self, c: canvas.Canvas, timestamp: str, total_cards: int, page: int, pages: int):
        c.setFillColorRGB(0, 0, 1)
        c.setFont("Helvetica", 7)
        
        info = (f"Page {page}/{pages} | {timestamp} | Cards: {total_cards} | "
                f"Size: {self.config.card_width_mm}x{self.config.card_height_mm}mm | "
                f"Grid: {self.config.grid_cols}x{self.config.grid_rows} | "
                f"DPI: {self.config.dpi} | Bleed: {self.config.bleed_mm}mm | "
                f"Corner bevel: {self.config.corner_bevel_mm}mm")
        
        x_offset = 10 + self.mm_to_pt(10)
        y_offset = self.page_height - 6 - self.mm_to_pt(10)
        c.drawString(x_offset, y_offset, info)
    
    def calculate_layout(self) -> dict:
        placed_w = self.mm_to_pt(self.config.card_width_mm + 2 * self.config.bleed_mm)
        placed_h = self.mm_to_pt(self.config.card_height_mm + 2 * self.config.bleed_mm)
        spacing = self.mm_to_pt(self.config.spacing_mm)
        
        grid_w = self.config.grid_cols * placed_w + (self.config.grid_cols - 1) * spacing
        grid_h = self.config.grid_rows * placed_h + (self.config.grid_rows - 1) * spacing
        
        return {
            'start_x': (self.page_width - grid_w) / 2,
            'start_y': (self.page_height - grid_h) / 2,
            'card_w': self.mm_to_pt(self.config.card_width_mm),
            'card_h': self.mm_to_pt(self.config.card_height_mm),
            'placed_w': placed_w,
            'placed_h': placed_h,
            'spacing': spacing
        }
    
    def generate(self, input_folder: str, output_path: str):
        images = self.find_images(input_folder)
        if not images:
            raise FileNotFoundError(f"No images found in {input_folder}")
        
        cardback_path = os.path.join(input_folder, 'cardback.png')
        if not os.path.exists(cardback_path):
            cardback_path = os.path.join(input_folder, 'CARDBACK.PNG')
        has_cardback = os.path.exists(cardback_path)
        
        images = [img for img in images if 'cardback' not in img.lower()]
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        layout = self.calculate_layout()
        c = canvas.Canvas(output_path, pagesize=A4)
        
        cards_per_page = self.config.grid_cols * self.config.grid_rows
        total_pages = (len(images) + cards_per_page - 1) // cards_per_page
        bleed_pt = self.mm_to_pt(self.config.bleed_mm)
        
        for page in range(total_pages):
            self.draw_header(c, timestamp, len(images), page + 1, total_pages)
            
            page_images = images[page * cards_per_page:(page + 1) * cards_per_page]
            
            for i, img_path in enumerate(page_images):
                row, col = divmod(i, self.config.grid_cols)
                x = layout['start_x'] + col * (layout['placed_w'] + layout['spacing'])
                y = layout['start_y'] + (self.config.grid_rows - 1 - row) * (layout['placed_h'] + layout['spacing'])
                
                result = self.process_image(img_path)
                if result:
                    bleed_reader, card_reader = result
                    
                    c.drawImage(bleed_reader, x, y, width=layout['placed_w'], height=layout['placed_h'])
                    c.drawImage(card_reader, x + bleed_pt, y + bleed_pt, 
                              width=layout['card_w'], height=layout['card_h'])
                    
                    card_x = x + bleed_pt
                    card_y = y + bleed_pt
                    self.draw_corner_guides(c, card_x, card_y, layout['card_w'], layout['card_h'])
            
            self.draw_separators(c, layout)
            
            if has_cardback:
                c.showPage()
                self.draw_header(c, timestamp, len(images), page + 1, total_pages)
                
                cardback_result = self.process_image(cardback_path)
                if cardback_result:
                    bleed_reader, card_reader = cardback_result
                    
                    for i in range(len(page_images)):
                        row, col = divmod(i, self.config.grid_cols)
                        mirrored_col = self.config.grid_cols - 1 - col
                        x = layout['start_x'] + mirrored_col * (layout['placed_w'] + layout['spacing'])
                        y = layout['start_y'] + (self.config.grid_rows - 1 - row) * (layout['placed_h'] + layout['spacing'])
                        
                        c.drawImage(bleed_reader, x, y, width=layout['placed_w'], height=layout['placed_h'])
                        c.drawImage(card_reader, x + bleed_pt, y + bleed_pt, 
                                  width=layout['card_w'], height=layout['card_h'])
                        
                        card_x = x + bleed_pt
                        card_y = y + bleed_pt
                        self.draw_corner_guides(c, card_x, card_y, layout['card_w'], layout['card_h'])
                
                self.draw_separators(c, layout)
            
            if page < total_pages - 1:
                c.showPage()
        
        c.save()
        return output_path


def main():
    parser = argparse.ArgumentParser(description='CardPuncher - Print-ready card layouts')
    parser.add_argument('folder', help='Folder containing card images')
    parser.add_argument('--config', help='YAML config file', default='cardpuncher.yaml')
    parser.add_argument('--dpi', type=int, help='Image resolution')
    parser.add_argument('--card-width-mm', type=float, help='Card width (mm)')
    parser.add_argument('--card-height-mm', type=float, help='Card height (mm)')
    
    args = parser.parse_args()
    
    config = Config.from_yaml(args.config)
    if args.dpi:
        config.dpi = args.dpi
    if args.card_width_mm:
        config.card_width_mm = args.card_width_mm
    if args.card_height_mm:
        config.card_height_mm = args.card_height_mm
    
    script_dir = Path(__file__).parent
    output_dir = script_dir / 'output'
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = Path(args.folder).name
    output_path = output_dir / f"{folder_name}_{timestamp}.pdf"
    
    try:
        puncher = CardPuncher(config)
        puncher.generate(args.folder, str(output_path))
        
        fronts = [img for img in puncher.find_images(args.folder) if 'cardback' not in img.lower()]
        cardback_path = os.path.join(args.folder, 'cardback.png')
        has_back = os.path.exists(cardback_path) or os.path.exists(os.path.join(args.folder, 'CARDBACK.PNG'))
        
        print(f"\nSuccess! {output_path}")
        print(f"Cards: {len(fronts)} | "
              f"Size: {config.card_width_mm}x{config.card_height_mm}mm | "
              f"Grid: {config.grid_cols}x{config.grid_rows} | "
              f"DPI: {config.dpi}")
        if has_back:
            print(f"Double-sided: Yes (interleaved backs)")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
