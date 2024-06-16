import os                                           #used to access files
import boto3                                        #the library to access AWS services 
import botocore                                     #another AWS library
import fitz                                         #used to convert pdf's to images
import tkinter as tk                                #the library for simple GUIs
from tkinter import messagebox, filedialog, ttk     #used for error messages, accessing files, and the progress bar, respectively


# Define global variables
inputfile       = ""    #(textstring) full filepath of user-chosen file to be operated on
inputfolder     = ""    #(textstring) full filepath of user-chosen folder whose entire contents will be operated on, one at a time
outputfolder    = ""    #(textstring) full filepath where output(s) will be created to
scoresonoff     = "off" #(textstring) either "off" or "on", tells program to print confidence scores at the end of the csv or leave them off
file_label      = ""    #(textstring) full filepath of user-chosen file to display on the GUI for visual confirmation
folder_label    = ""    #(textstring) full filepath of user-chosen folder to display on the GUI for visual confirmation
result_label    = ""    #(textstring) full output filepath and/or progress updates to display on GUI as operations happen
progress_bar    = None  #(idk what type) the progress bar on the GUI
progress_label  = ""    #(textstring) label showing % complete on the GUI
result_log      = ""    #(textstring) progress updates to display on GUI as operations happen

#update progress bar in GUI
def update_progress(progress):
    progress_bar["value"] = progress
    progress_label.config(text="Progress: {}%".format(progress))
    
#update result_label with a new message in GUI
def update_result_label(new_message):
    global result_log
    if new_message:  # Only append non-empty messages
        result_log += new_message + "\n"  # Append the new message to the log with a newline character
        result_label.config(text=result_log)  # Update the result_label with the updated log

#gets words from Textract output, called by get_rows_columns_map()
def get_text(result, blocks_map):
    print("Now passing through get_text()")
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

#arranges Textract output into a folder, called by generate_table_csv()
def get_rows_columns_map(table_result, blocks_map):
    print("Now passing through get_rows_columns_map()")
    global scoresonoff
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
                    # get confidence scores
                    scores.append(str(cell['Confidence']))
                        
                    # get the text value
                    rows[row_index][col_index] = get_text(cell, blocks_map)
    return rows, scores

#turns the extracted words (and scores if selected) into a table format, called by get_table_csv_results()
def generate_table_csv(table_result, blocks_map, table_index):
    print("Now passing through generate_table_csv()")
    rows, scores = get_rows_columns_map(table_result, blocks_map)

    table_id = 'Table_' + str(table_index)
    
    # get cells.
    csv = 'Table: {0}\n\n'.format(table_id)

    for row_index, cols in rows.items():
        for col_index, text in cols.items():
            col_indices = len(cols.items())
            csv += '{}'.format(text) + ","
        csv += '\n'
    
    if scoresonoff == "on":                                                       #print confidence scores if user selected it
        csv += '\n\n Confidence Scores % (Table Cell) \n'
        cols_count = 0
        for score in scores:
            print("Now Adding Confidence Scores to Table")
            cols_count += 1
            csv += score + ","
            if cols_count == col_indices:
                csv += '\n'
                cols_count = 0

    csv += '\n\n\n'
    return csv

#accesses Textract to actually use the AI for output, called by main()
def get_table_csv_results(input_file):
    print("Now opening client to AWS and passing through get_table_csv_results()")
    with open(input_file, 'rb') as file:
        img_test = file.read()
        bytes_test = bytearray(img_test)
        print('Image loaded: ', input_file)

    access_key = app.access_key_value                       #get the input values from inside the LoginPage class
    secret_key = app.secret_key_value                       #get the input values from inside the LoginPage class

    client = boto3.client(
        'textract',
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name='us-east-2'
    )
    response = client.analyze_document(Document={'Bytes': bytes_test}, FeatureTypes=['TABLES'])

    #gets output in JSON format. pprint(blocks) allows user to see that in the terminal
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

#turns PDFs into PNGs and then turns table-formatted text into a .csv file, called by processimages()
def main(input_file, output_file):
    print("Now converting to PNGs (if needed) and passing through main()")

    # If the input file is a PDF, convert it to PNGs
    if input_file.lower().endswith('.pdf'):
        doc = fitz.open(input_file)
        for page_number in range(doc.page_count):
            page = doc.load_page(page_number)
            image = page.get_pixmap()
            image_file = f"{input_file[:-4]}_page_{page_number + 1}.png"
            image.save(image_file)
            update_result_label(f"Page {page_number + 1} converted to PNG: {image_file}")
            input_file = image_file            #Use the first page for processing
            table_csv = get_table_csv_results(input_file)               #calls the next function
            os.remove(image_file)
            with open(output_file, "at") as fout:                       #create the actual .csv file
                fout.write(table_csv)
        doc.close()
    else:
        table_csv = get_table_csv_results(input_file)               #calls the next function
        with open(output_file, "at") as fout:                       #create the actual .csv file
            fout.write(table_csv)

    print('CSV OUTPUT FILE: ', output_file)

#starts the whole process and updates the GUI progress updates and bar, called by gui()
def processimages():
    global inputfile
    global inputfolder
    global outputfolder
    global scoresonoff

    print("Step 2 Complete: User has submitted inputs, process_images() will now begin with the following: inputfile=",inputfile," inputfolder=",inputfolder," outpufolder=",outputfolder," outpufolder=",outputfolder," scoresonoff=",scoresonoff)


    total_stages = 1  # Total number of stages in the conversion process
    current_stage = 0

    if inputfile:
        current_stage += 1
        update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for loading images
        update_result_label("File Loaded")
        
        filename = inputfile.split("/")[-1]
        output_file = outputfolder +"/"+ filename + '_output.csv'
        main(inputfile, output_file)
        
        
    elif inputfolder:
        completed_files = 0
        for fileeach in os.listdir(inputfolder):
            current_stage += 1
            update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for loading images

            filenamewithpath = os.path.join(inputfolder, fileeach)
            output_file = os.path.join(outputfolder +"/"+ fileeach + '_output.csv')
            main(filenamewithpath, output_file)
            completed_files += 1
            update_progress(int((current_stage / total_stages) * 100))  # Update progress bar for analyzing documents
    else:
        output_file = "ERROR: NO FILE CREATED"
    update_result_label(f"Conversion Complete, Result located at {output_file}")
    print("All Steps Complete###################################################")

#creates the only the visual components of the GUI, started when login is successful
def gui():
    global inputfile                                            #declare that these variables are all the global ones listed at the top
    global inputfolder
    global outputfolder
    global switch_button
    global scoresonoff
    global file_label
    global folder_label
    global result_label
    global progress_bar
    global progress_label

    #allow user to select one or multiple input files
    def browse_file():
        global inputfile
        filetypes = [
            ("Image files", "*.jpg;*.jpeg;*.png;*.tiff;*.tif;*.pdf"),
            ("All files", "*.*")]
        inputfile = filedialog.askopenfilename(title="Select Input File", filetypes=filetypes, initialdir=os.getcwd())
        if inputfile:
            # Update the file label
            file_label.config(text="Selected File: " + os.path.basename(inputfile))
            inputfile = ""                                          #either a folder or a file can be used, so it clears the folder
            folder_label.config(text="")                              #also removes the file from the GUI
            #make the same output as input invisible
            or2.config(fg="#1F252F")                                  
            same_output_button.config(text="This Button is Only Available if Folder Selected", command=same_out_folder, bg="#1F252F", fg="white")

    #allow user to select an input folder
    def browse_folder():
        global inputfile
        global inputfolder
        inputfolder = filedialog.askdirectory(title="Select Input Folder (must contain only images)", initialdir=os.getcwd())
        if inputfolder:
            # Update the folder label
            folder_label.config(text="Selected Folder: " + inputfolder)
            inputfile = ""                                          #either a folder or a file can be used, so it clears the file
            file_label.config(text="")                              #also removes the file from the GUI
            #make the same output as input visible
            or2.config(fg="white")                                  
            same_output_button.config(text="Same Output Folder as Input Folder", bg="#3C3C3C")


    #allow user to select an output location
    def out_folder():
        global outputfolder
        outputfolder = filedialog.askdirectory(title="Select Output Folder", initialdir=os.getcwd())
        # Update the output folder label
        out_folder_label.config(text="Selected Output Location: " + outputfolder)

    #operate the toggle button about printing scores
    def toggle_switch():
        global scoresonoff
        if scoresonoff == "off":
            switch_button.config(text="Show Confidence Scores in Document ON (click to change)", bg="green")
            scoresonoff = "on"
        else:
            switch_button.config(text="Show Confidence Scores in Document OFF (click to change)", bg="red")
            scoresonoff = "off"

    #allow user to set the output location to be the same as the input folder
    def same_out_folder():
        global inputfolder
        global outputfolder
        if inputfolder:                                         #button only works if an input folder exists , otherwise both receieve blank filepaths
            outputfolder = inputfolder
            out_folder_label.config(text="Selected Output Location: " + outputfolder)
            
    #submit button kicks off the rest if the program
    def submit_function():
        if (inputfile and outputfolder) or (inputfolder and outputfolder):
            processimages()
        else:
            messagebox.showerror("Error", "Please Select an output folder and either an input file or folder containing only images and PDFs")


    ############################################################### Visual Elements of the GUI are created here#######
    # Create main window
    root = tk.Tk()                                              #create popup window
    root.title("Image to Excel Converter")                      #popup window title
    # root.geometry("700x400")                                  #window size (removed to allow automatic resizing due to large messages)
    root.configure(bg="#1F252F")                                #popup window background color

    # Input File Selector Button - Top Left
    file_button = tk.Button(root, text="Select Input File", command=browse_file, bg="#3C3C3C", fg="white")
    file_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")    #location, horizontal padding, vertical padding, make the cell fit the text in an East-West format (i.e., it autosizes horiztonally)

    # Text "or" - Top Middle
    or1 = tk.Label(root, text="OR", bg="#1F252F", fg="white")
    or1.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    # Folder Button - Top Right
    folder_button = tk.Button(root, text="Select Input Folder (must contain only images and PDFs)", command=browse_folder, bg="#3C3C3C", fg="white")
    folder_button.grid(row=0, column=2, padx=10, pady=10, sticky="ew")

    # Output Folder Button - Second Row Left
    output_folder_button = tk.Button(root, text="Select Output Folder", command=out_folder, bg="#3C3C3C", fg="white")
    output_folder_button.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    # Text "or" - Second Row Middle
    or2 = tk.Label(root, text="OR", bg="#1F252F", fg="#1F252F") #this begins hidden and only appears when an input folder is selected
    or2.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

    # Same Output Folder Button - Second Row Right
    same_output_button = tk.Button(root, text="This Button is Only Available if a Folder is Selected", command=same_out_folder, bg="#1F252F", fg="white") #this button appears correctly once an input folder is selected
    same_output_button.grid(row=1, column=2, padx=10, pady=10, sticky="ew")

    # Confidence Scores Radio Button - Third Row Middle and Right
    switch_button = tk.Button(root, text="Print Confidence Scores OFF (click to change)", bg="red", width=8, command=toggle_switch)
    switch_button.grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

    # Submit Button - Fourth Row, Left and Middle
    submit_button = tk.Button(root, text="Submit", command=submit_function, bg="#FAA12A") ########this is the button that starts the next process!!!!!!!!!!!!
    submit_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    # Labels to display selected file, folder, and result
    file_label = tk.Label(root, text="", bg="#1F252F", fg="white")
    file_label.grid(row=5, column=0, columnspan=3, padx=10, pady=0, sticky="ew")

    folder_label = tk.Label(root, text="", bg="#1F252F", fg="white")
    folder_label.grid(row=6, column=0, columnspan=3, padx=10, pady=0, sticky="ew")
    
    out_folder_label = tk.Label(root, text="", bg="#1F252F", fg="white")
    out_folder_label.grid(row=7, column=0, columnspan=3, padx=10, pady=0, sticky="ew")

    result_label = tk.Label(root, text="", bg="#1F252F", fg="white")
    result_label.grid(row=8, column=0, columnspan=3, padx=10, pady=0, sticky="ew")

    # Progress bar and label
    progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')
    progress_bar.grid(row=9, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    progress_label = tk.Label(root, text="Progress: 0%", bg="#1F252F", fg="white")
    progress_label.grid(row=10, column=0, columnspan=3, padx=10, pady=0, sticky="ew")
    
    # Text to show Copyright info
    copyright = tk.Label(root, text="\n\n\n© Ted Yee 2024\n", bg="#1F252F", fg="white")
    copyright.grid(row=11, column=0, columnspan=3, padx=10, pady=0, sticky="ew")

    root.mainloop()

#handles the entire login page, started when program starts
class LoginPage:
    #creates only the visual components of the login page
    def __init__(self, master):
        self.master = master
        self.master.title("AWS Login To Access Converter")      #the title of the popup
        self.master.configure(bg="#1F252F")                     #background color of the popup

        self.username_label = tk.Label(master, text="\n\n\n\n\n\n\nAWS Access Key ID:", bg="#1F252F", fg="white")   #a bunch of newlines to move the title down, then title text, background color, and text color
        self.username_label.pack()                              #place the object (the label in this case) as closely above the previous object as possible

        self.access_key = tk.Entry(master, width=40)            #text entry box for the access_key
        self.access_key.pack()                                  #place the object (the entrybox in this case) as closely above the previous object as possible

        self.password_label = tk.Label(master, text="\nAWS Secret Access Key:", bg="#1F252F", fg="white")   #newlines to move the title down, then title text, background color, and text color
        self.password_label.pack()                              #place the object (the label in this case) as closely above the previous object as possible

        self.secret_key = tk.Entry(master, width=40, show="*")  #text entry box for the secret_key, showing asterisks instead of characters
        self.secret_key.pack()                                  #place the object (the entrybox in this case) as closely above the previous object as possible

        self.login_button = tk.Button(master, text="Login", bg="#FAA12A", command=self.validate_login) #the login button, background color, and tells it to start the validate_login() function below when clicked
        self.login_button.pack()                                #place the object (the button in this case) as closely above the previous object as possible
        
        self.copyright = tk.Label(master, text="\n\n\n\n\n\nData will be validated with AWS via an internet connection.\n© Ted Yee 2024", bg="#1F252F", fg="white")  #a bunch of newlines to move it further down, then the Copyright text
        self.copyright.pack()                                   #place the object (the text in this case) as closely above the previous object as possible

        self.aws_credentials_valid = False                      #variable starts saying that login is unsuccesful
        

    #sends input values to AWS in the form of an S3 request about which files exist under the account. If request returns an error, the values aren't a real AWS token, else update access_key_value and secret_key_valu variables
    def validate_login(self):
        self.access_key_value = self.access_key.get()           #gets the input value from the popup window
        self.secret_key_value = self.secret_key.get()           #gets the input value from the popup window
        print("User has entered the following credentials: Access Key=", self.access_key_value, ", Secret Key=", self.secret_key_value)
        if self.access_key_value and self.secret_key_value:     #if both boxes have input
            try:
                self.s3 = boto3.client('s3', aws_access_key_id = self.access_key_value, aws_secret_access_key= self.secret_key_value)
                buckets = self.s3.list_buckets()['Buckets']     #This is a random request to AWS just to see if login is valid (if isn't, returns an error). It actually retrieves a list of buckets owned by the account
                self.aws_credentials_valid = True               #variable says that login is successful now
                self.master.destroy()                           #Destroy the login window after successful login
                
            except botocore.exceptions.ClientError as e:        #if the request to AWS returns any kind of error
                if e.response['Error']['Code'] == 'InvalidAccessKeyId': #if this error is returned, show the below message in the error message popup
                    messagebox.showerror("Error", "AWS was not able to validate that Access Key ID or Secret Access Key. Please provide a valid set of AWS credentials or see here for info: https://docs.aws.amazon.com/cli/v1/userguide/cli-configure-files.html")
                else:                                           #if any other error is returned, show the below message in the error message popup
                    messagebox.showerror("Error", f"An error occurred: {e}")
        else:                                                   #if both boxes do not have input, show the below message in the error message popup
            messagebox.showerror("Error", "Please enter both AWS Access Key ID and Secret Access Key.")

#what to do when program is started
if __name__ == "__main__":
    #create the popup window for the login page
    root = tk.Tk()                                              #create the popup page object
    root.geometry("700x400")                                    #size of the popup window
    app = LoginPage(root)                                       #a way to reference things inside the class
    root.mainloop()                                             #tells it to start the popup
    
    #if login successful, open next page
    if app.aws_credentials_valid:                               #checks that the aws_credentials_valid variable inside the LoginPage class is True
        access_key = app.access_key_value                       #get the input values from inside the LoginPage class
        secret_key = app.secret_key_value                       #get the input values from inside the LoginPage class
        print("Step 1 Complete: The Following Credentials are Valid: ", access_key, secret_key)
        gui()                                                   #start the gui, which in turn will kick off the whole process
    # gui()                                                       #if login page needs to be skipped for dev testing, comment out everything in if __name__ == "__main__": and uncomment this line, also note that the access and secret keys need to be inputted into get_table_csv_results()