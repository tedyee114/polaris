import os
import boto3
import fitz
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk  # Import ttk module for progress bar

# Define global variables
inputfile = ""
inputfolder = ""
outputfolder = ""
scoresonoff = "off"
file_label = None  # Initialize to None
folder_label = None  # Initialize to None
result_label = None  # Initialize to None
progress_bar = None  # Initialize to None
progress_label = None  # Initialize to None
result_log = ""  # Initialize an empty string to store log messages

# Function to update result_label with a new message
def update_result_label(new_message):
    global result_log
    result_log += new_message + "\n"  # Append the new message to the log with a newline character
    result_label.config(text=result_log)  # Update the result_label with the updated log

def browse_file():
    global inputfile
    filetypes = [
        ("Image files", "*.jpg;*.jpeg;*.png;*.tiff;*.tif;*.pdf"),
        ("All files", "*.*")
    ]
    inputfile = filedialog.askopenfilename(title="Select Input File", filetypes=filetypes, initialdir=os.getcwd())
    # Update the file label
    file_label.config(text="Selected File: " + os.path.basename(inputfile))

def browse_folder():
    global inputfolder
    inputfolder = filedialog.askdirectory(title="Select Input Folder (must contain only images)", initialdir=os.getcwd())
    # Update the folder label
    folder_label.config(text="Selected Folder: " + inputfolder)

def out_folder():
    global outputfolder
    outputfolder = filedialog.askdirectory(title="Select Output Folder", initialdir=os.getcwd())
    # Update the output folder label
    update_result_label(f"Selected Output Folder: {outputfolder}")

def same_out_folder():
    global inputfolder
    global outputfolder
    outputfolder = inputfolder
    update_result_label(f"Selected Output Folder: {outputfolder}")

def scores():
    global scoresonoff
    global outputfolder
    outputfolder = inputfolder

def update_progress(progress):
    progress_bar["value"] = progress
    progress_label.config(text="Progress: {}%".format(progress))

def get_rows_columns_map(table_result, blocks_map):
    rows = {}
    scores = []
    for relationship in table_result['Relationships']:
        if relationship['Type'] == 'CHILD':
            for child_id in relationship['Ids']:
                cell = blocks_map[child_id]
                if cell['BlockType'] == 'CELL':
                    row_index = cell['RowIndex']
                    col_index = cell['ColumnIndex']
                    if row_index not in rows:
                        # create new row
                        rows[row_index] = {}
                    
                    # get confidence score
                    scores.append(str(cell['Confidence']))
                        
                    # get the text value
                    rows[row_index][col_index] = get_text(cell, blocks_map)
    return rows, scores

def get_text(result, blocks_map):
    text = ''
    if 'Relationships' in result:
        for relationship in result['Relationships']:
            if relationship['Type'] == 'CHILD':
                for child_id in relationship['Ids']:
                    word = blocks_map[child_id]
                    if word['BlockType'] == 'WORD':
                        if "," in word['Text'] and word['Text'].replace(",", "").isnumeric():
                            text += '"' + word['Text'] + '"' + ' '
                        else:
                            text += word['Text'] + ' '
                    if word['BlockType'] == 'SELECTION_ELEMENT':
                        if word['SelectionStatus'] =='SELECTED':
                            text +=  'X '
    return text

def get_table_csv_results(file_name):

    with open(file_name, 'rb') as file:
        img_test = file.read()
        bytes_test = bytearray(img_test)
        print('Image loaded', file_name)

    # process using image bytes
    # get the results
    

    session = boto3.session.Session(profile_name = 'default')
    client = session.client('textract', region_name='us-east-2')
    response = client.analyze_document(Document={'Bytes': bytes_test}, FeatureTypes=['TABLES'])

    # Get the text blocks
    blocks=response['Blocks']
    # pprint(blocks)

    blocks_map = {}
    table_blocks = []
    for block in blocks:
        blocks_map[block['Id']] = block
        if block['BlockType'] == "TABLE":
            table_blocks.append(block)

    if len(table_blocks) <= 0:
        return "<b> NO Table FOUND </b>"

    csv = ''
    for index, table in enumerate(table_blocks):
        csv += generate_table_csv(table, blocks_map, index +1)
        csv += '\n\n'

    return csv

def generate_table_csv(table_result, blocks_map, table_index):
    rows, scores = get_rows_columns_map(table_result, blocks_map)

    table_id = 'Table_' + str(table_index)
    
    # get cells.
    csv = 'Table: {0}\n\n'.format(table_id)

    for row_index, cols in rows.items():
        for col_index, text in cols.items():
            col_indices = len(cols.items())
            csv += '{}'.format(text) + ","
        csv += '\n'
        
    csv += '\n\n Confidence Scores % (Table Cell) \n'
    cols_count = 0
    for score in scores:
        cols_count += 1
        csv += score + ","
        if cols_count == col_indices:
            csv += '\n'
            cols_count = 0

    csv += '\n\n\n'
    return csv

def toggle_switch():
    global scoresonoff
    if scoresonoff == "off":
        switch_button.config(text="Print Confidence Scores ON (click to change)", bg="green")
        scoresonoff = "on"
    else:
        switch_button.config(text="Print Confidence Scores OFF (click to change)", bg="red")
        scoresonoff = "off"

def main(input_file, output_file):
    # If the input file is a PDF, convert it to PNGs
    if input_file.lower().endswith('.pdf'):
        doc = fitz.open(input_file)
        for page_number in range(doc.page_count):
            page = doc.load_page(page_number)
            image = page.get_pixmap()
            image_file = f"{input_file[:-4]}_page_{page_number + 1}.png"
            image.save(image_file)
            update_result_label(f"Page {page_number + 1} converted to PNG: {image_file}")
        doc.close()
        input_file = f"{input_file[:-4]}_page_1.png"  # Use the first page for processing
    table_csv = get_table_csv_results(input_file)

    # replace content
    with open(output_file, "wt") as fout:
        fout.write(table_csv)

    # show the results
    print('CSV OUTPUT FILE: ', output_file)

def processimages():
    global inputfile
    global inputfolder
    global outputfolder
    global scoresonoff

    # Initialize progress variables
    total_stages = 2  # Total number of stages in the conversion process
    current_stage = 0

    if inputfile:
        current_stage += 2
        update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for loading images
        update_result_label("File Loaded")
        
        output_file = inputfile + '_output.csv'
        main(inputfile, output_file)
        
        
    elif inputfolder:
        completed_files = 0
        for fileeach in os.listdir(inputfolder):
            current_stage += 2
            update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for loading images

            filename = os.path.join(inputfolder, fileeach)
            output_file = os.path.join(outputfolder, fileeach + '_output.csv')
            main(filename, output_file)
            completed_files += 1
            update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for analyzing documents

    update_result_label(f"Conversion Complete, Result located at {outputfolder}")

def gui():
    global inputfile
    global inputfolder
    global outputfolder
    global switch_button
    global scoresonoff
    global file_label  # Declare as global
    global folder_label  # Declare as global
    global result_label  # Declare as global
    global progress_bar  # Declare as global
    global progress_label  # Declare as global

    inputfile = ""
    inputfolder = ""
    outputfolder = ""
    scoresonoff = ""
    scoresonoff = "off"

    # Create main window
    root = tk.Tk()
    root.title("Image to Excel Converter")
    # root.geometry("700x400")
    root.configure(bg="#00A3E0")

    # File Button - Top Left
    file_button = tk.Button(root, text="Select Input File", command=browse_file, bg="#FAA12A")
    file_button['font'] = ('System', 12, 'bold')
    file_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

    # Folder Button - Top Right
    folder_button = tk.Button(root, text="Select Input Folder (must contain only images)", command=browse_folder, bg="#FAA12A")
    folder_button['font'] = ('System', 12, 'bold')
    folder_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    # Output Folder Button - Middle Left
    output_folder_button = tk.Button(root, text="Select Output Folder", command=out_folder, bg="#FAA12A")
    output_folder_button['font'] = ('System', 12, 'bold')
    output_folder_button.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    # Same Output Folder Button - Middle Right
    same_output_button = tk.Button(root, text="Same Output Folder as Input Folder", command=same_out_folder, bg="#FAA12A")
    same_output_button['font'] = ('System', 12, 'bold')
    same_output_button.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

    # radio button
    switch_button = tk.Button(root, text="Print Confidence Scores OFF (click to change)", bg="red", width=8, command=toggle_switch)
    switch_button['font'] = ('System', 12, 'bold')
    switch_button.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

    # Submit Button - Bottom Center
    submit_button = tk.Button(root, text="Submit", command=processimages, bg="#FAA12A")
    submit_button['font'] = ('System', 12, 'bold')
    submit_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    # Labels to display selected file, folder, and result
    file_label = tk.Label(root, text="", bg="#00A3E0")
    file_label['font'] = ('System', 12, 'bold')
    file_label.grid(row=5, column=0, columnspan=2, padx=10, pady=0, sticky="ew")

    folder_label = tk.Label(root, text="", bg="#00A3E0")
    folder_label['font'] = ('System', 12, 'bold')
    folder_label.grid(row=6, column=0, columnspan=2, padx=10, pady=0, sticky="ew")

    result_label = tk.Label(root, text="", bg="#00A3E0")
    result_label['font'] = ('System', 12, 'bold')
    result_label.grid(row=7, column=0, columnspan=2, padx=10, pady=0, sticky="ew")

    # Progress bar and label
    progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')
    progress_bar.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    progress_label = tk.Label(root, text="Progress: 0%", bg="#00A3E0")
    progress_label['font'] = ('System', 12, 'bold')
    progress_label.grid(row=9, column=0, columnspan=2, padx=10, pady=0, sticky="ew")

    root.mainloop()

if __name__ == "__main__":
    gui()
