




from PIL import Image, ImageDraw, ImageFont
# Open the image
img = Image.open('template/certificate_template.png')  # Ensure this path exists
name = 'KE SOTHIN'
# Get image size (fixed: no parentheses)
width, height = img.size
# Create a drawing context
draw = ImageDraw.Draw(img)
# Load a font (optional; adjust path as needed)
try:
    font = ImageFont.truetype('arialbd.ttf', 100)  # Example: Bold Arial, size 100
except OSError:
    font = ImageFont.load_default()  # Fallback to default font
# Draw the text (centered horizontally, adjust y-position as needed)
draw.text((width // 2, (height//2)+60), name, fill='navy', font=font, anchor='mm')
# Show the image (for preview)
img.show()