from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import tkinter as tk
from tkinter import filedialog

# Function to browse for a Word document
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(tk.END, file_path)

# Function to apply changes to the selected document
def apply_changes():
    file_path = entry_file_path.get()
    if file_path:
        doc = Document(file_path)

        # Define the color mappings
        color_mappings = {
            'Improbable': 'CCCCCC',  # Gray
            'Peu probable': 'FFFF00',  # Yellow
            'Probable': 'FFA500',  # Orange
            'Très probable': 'FF0000',  # Red
            'Négligeable': 'CCCCCC',  # Gray
            'Limité': 'FFFF00',  # Yellow
            'Important': 'FFA500',  # Orange
            'Critique': 'FF0000',  # Red
            'Simple': 'FFFF00',  # Yellow
            'Raisonnable': 'FFA500',  # Orange
            'Complexe': 'FF0000'  # Red
        }

        # Iterate through all tables in the document
        for table in doc.tables:
            # Iterate through each row in the table
            for row in table.rows:
                # Iterate through each cell in the row
                for cell in row.cells:
                    # Check the text content of the cell and set the color accordingly
                    text = cell.text.strip()
                    if text in color_mappings:
                        color_code = color_mappings[text]
                        # Create a new shading element with the specified fill color
                        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_code}"/>'
                        shading = parse_xml(shading_xml)
                        # Apply the shading to the cell
                        cell._element.tcPr.append(shading)

        # Save the modified document with the same filename to overwrite the original document
        doc.save(file_path)
        lbl_status.config(text="Changes applied successfully.", fg="green")
    else:
        lbl_status.config(text="Please select a file.", fg="red")

# Create the main window
window = tk.Tk()
window.title("Word Document Colorizer")

# Create the file path entry field
entry_file_path = tk.Entry(window, width=50)
entry_file_path.pack(pady=10)

# Create the Browse button
btn_browse = tk.Button(window, text="Browse", command=browse_file)
btn_browse.pack(pady=5)

# Create the Apply button
btn_apply = tk.Button(window, text="APPLY", command=apply_changes)
btn_apply.pack(pady=10)

# Create the status label
lbl_status = tk.Label(window, text="")
lbl_status.pack()

# Start the main GUI loop
window.mainloop()
