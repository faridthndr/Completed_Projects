from pdf2image import convert_from_path

pdf_path = 'D:\\Ashraf\\1247371.pdf'

images = convert_from_path(pdf_path)

for i, img in enumerate(images):
    img.save(f'page_{i+1}.png', 'PNG')

print("Conversion successful!")
