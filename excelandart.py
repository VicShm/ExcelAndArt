import xlsxwriter
from PIL import Image
import time

start_time = time.time()

img_file = './cat/cat.png'	# Image file.
exl_file = './cat/cat.xlsx'	# Excel file.

# Open image file
img = Image.open(img_file)
width, height = img.size
obj = img.load()

exl_book = xlsxwriter.Workbook(exl_file)	# Create Excel book
exl_sheet = exl_book.add_worksheet()		# Create Excel sheet

cell_format_list = []	# Cells format list.
pixel_list = []		# List of all pixel colors

# List of deduplicated colors.
for x in range(0, width):
    for y in range(0, height):
        rgb = (obj[x, y][0], obj[x, y][1], obj[x, y][2])  # pillow gets pixel color in RGB
        pixel_list.append(f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}')  # In Excel  must be HEX cells color
unique_colors = list(set(pixel_list))

# Creating Excel Cell formats list.
for color in unique_colors:
    cell_format = exl_book.add_format()
    cell_format.set_pattern(1)
    cell_format.set_bg_color(color.strip())
    cell_format_list.append(cell_format)

# Consistently read the color of each pixel and specify the color of the corresponding cells in Excel.
for x in range(0, width):
    for y in range(0, height):
        rgb = (obj[x, y][0], obj[x, y][1], obj[x, y][2])
        hex_pixel_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
        items = [item for item in cell_format_list if item.bg_color == hex_pixel_color]
        exl_sheet.write(y, x, '', items[0])
    print(width-x)

# Specify the required cell size and page scale.
print("Few seconds more...")
exl_sheet.set_default_row(5)
exl_sheet.set_column(0, width, 0.5)
exl_sheet.zoom = 20
exl_sheet
exl_book.close()
print("Done!")
print(time.time() - start_time)
