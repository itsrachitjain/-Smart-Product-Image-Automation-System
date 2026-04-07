🚀 Smart Product Image Automation System

End-to-end Excel Automation & Image Processing System using Python to eliminate manual effort in inserting and formatting product images in large-scale datasets.

📌 Project Overview

This project automates:

📊 Bulk Excel Processing (1000+ rows)
🖼️ Product Image Mapping & Embedding
📁 Multi-format Image Handling (PNG, JPG, etc.)

The goal is to automatically match product images with Excel records using UPC codes and insert them into the sheet with proper alignment and formatting, significantly reducing manual workload.

🎯 Objective

Automate product image insertion into Excel sheets
Match images dynamically using UPC identifiers
Maintain consistent formatting and alignment
Handle missing or mismatched image scenarios
Improve operational efficiency in catalog management

📂 Input Requirements

Excel File (.xlsx)

Required Columns:

UPC → Unique product identifier
ARTICLE IMAGE → Target column for image insertion
Image Folder
Images must be named using UPC values

Example:

12345.png  
12345.jpg 

⚙️ Data Processing & Automation

🔹 File Handling

Dynamic path input system
Auto folder creation for output
Timestamp-based output file generation
🔹 Image Processing

Detects all supported image formats using PIL
Converts images to optimized PNG format
Resizes images to fit within Excel cells
Maintains image clarity and aspect ratio
🔹 Excel Automation

Dynamically detects column positions
Inserts images into correct cells
Centers images using anchor positioning
Adjusts row height and column width automatically
Removes unwanted cell formatting

📊 Core Logic

Image Matching
Matches images using UPC-based naming convention
Iterates through supported extensions
Inserts first matching image found
Alignment System
Uses pixel-to-EMU conversion for precision
Applies vertical centering logic
Adds padding for clean visual layout

📈 Key Features

✔ Automated bulk image insertion
✔ Dynamic column detection (no hardcoding)
✔ Multi-format image support
✔ Smart centering & resizing
✔ Missing image detection
✔ Clean Excel formatting
✔ Timestamp-based output tracking

📁 Project Structure
├── Smart Product Image Automation System.py                 
├── IMAGE/               
├── OUTPUT/                
└── README.md               

🛠 Tech Stack

Python
openpyxl
Pillow (PIL)
OS & Datetime modules

🚀 How to Run

pip install openpyxl pillow
python main.py
Inputs Required:
Excel file path
Image folder path (default available)
Output folder path (default available)

📤 Output

Final Excel file with embedded images

Format:

Image Inserted File YYYY-MM-DD_HH-MM-SS.xlsx

⚠️ Error Handling

Skips rows with missing UPC
Detects and reports missing images
Handles invalid or corrupted image files
Prevents crashes using try-except blocks

📈 Performance Insights

Efficient for hundreds to thousands of records
Reduces manual effort by 80–90%
Scalable for large catalog operations

💼 Business Impact

✔ Eliminated manual image insertion effort
✔ Improved product catalog preparation speed
✔ Reduced human error in image mapping
✔ Standardized formatting across datasets

👤 Author

Rachit Jain
Data Analyst | Automation | Business Intelligence

🔗 Connect With Me

LinkedIn: https://www.linkedin.com/in/itsrachitjain/
GitHub: https://github.com/itsrachitjain
