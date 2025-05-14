from PIL import Image, ImageDraw
import os

def create_checkmark_icon():
    # Create images of different sizes for the icon
    sizes = [(16,16), (32,32), (48,48), (64,64), (128,128)]
    images = []
    
    for size in sizes:
        # Create a new image with transparency
        image = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # Calculate checkmark points based on image size
        width, height = size
        padding = width // 4
        
        # Points for checkmark (✓)
        points = [
            (padding, height//2),  # Start point
            (width//2.5, height-padding),  # Bottom point
            (width-padding, padding)  # Top right point
        ]
        
        # Draw checkmark with light gray color (200, 200, 200)
        draw.line(points, fill=(200, 200, 200, 255), width=max(2, width//8))
        
        images.append(image)
    
    # Save as ICO file
    icon_path = 'taskmaster.ico'
    # Save all sizes in the ICO file
    images[0].save(icon_path, format='ICO', append_images=images[1:], sizes=[(x,x) for x in [16, 32, 48, 64, 128]])
    return icon_path

if __name__ == '__main__':
    create_checkmark_icon() 