"""Generate a feather application icon."""
from PIL import Image, ImageDraw

def create_feather_icon(filename: str = "app_icon.png", size: int = 256) -> None:
    """Create a simple feather app icon."""
    # Create image with white background
    img = Image.new("RGB", (size, size), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # Draw feather shape - simplified
    # Feather shaft (main line)
    center_x = size // 2
    center_y = size // 2
    shaft_len = size // 3
    
    # Draw shaft (brown/tan)
    draw.line(
        [(center_x, center_y - shaft_len), (center_x, center_y + shaft_len // 2)],
        fill=(139, 69, 19),  # Saddle brown
        width=max(2, size // 32)
    )
    
    # Draw feather vanes (left and right curves) - blue
    feather_color = (70, 130, 200)  # Steel blue
    
    # Left vane - multiple curves
    for i in range(5):
        y = center_y - shaft_len + (i * shaft_len // 4)
        x_curve = center_x - (shaft_len // 3 - i * shaft_len // 20)
        draw.arc(
            [x_curve - size // 16, y - size // 32, x_curve + size // 16, y + size // 32],
            0, 360,
            fill=feather_color,
            width=max(1, size // 64)
        )
    
    # Right vane - mirror
    for i in range(5):
        y = center_y - shaft_len + (i * shaft_len // 4)
        x_curve = center_x + (shaft_len // 3 - i * shaft_len // 20)
        draw.arc(
            [x_curve - size // 16, y - size // 32, x_curve + size // 16, y + size // 32],
            0, 360,
            fill=feather_color,
            width=max(1, size // 64)
        )
    
    img.save(filename)
    print(f"Feather icon created: {filename}")

if __name__ == "__main__":
    create_feather_icon()
