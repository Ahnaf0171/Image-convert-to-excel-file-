from PIL import Image
import pytesseract
import pandas as pd

# Path to the Tesseract executable (Update this path if necessary)
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# Path to the image file
image_path = "table1.png"

# Load the image and apply OCR
img = Image.open(image_path)
data = pytesseract.image_to_string(img, config='--psm 6')

# Debug: Print raw OCR output
print("Raw OCR Output:")
print(data)

# Split the data into rows
rows = data.split("\n")
data_list = []

# Process each row
for row in rows:
    # Split row into columns by whitespace
    row_data = row.split()
    
    # Append valid rows with at least 5 columns
    if len(row_data) >= 5:
        data_list.append(row_data[:5])  # Use the first 5 columns only

# Create a DataFrame
columns = ["Month", "Product", "Product color", "Country", "Sales Revenue"]
df = pd.DataFrame(data_list, columns=columns)

# Save the DataFrame to an Excel file
excel_path = "output.xlsx"
df.to_excel(excel_path, index=False, engine='openpyxl')

print(f"Data successfully saved to {excel_path}")
