from tkinter import Tk, Button, filedialog
from docx import Document

def upload_file():
    # Open file dialog to select a file
    file_path = filedialog.askopenfilename()
    
    # Print the selected file path (you can modify this according to your requirements)
    print("Selected file:", file_path)
    window.destroy()  # Close the window

    with open(file_path, 'rb') as file:
        # Process the file content (print it in this example)
        process_file(file_path)

def process_file(file_path):
    # Print the file path and content
    double_slash_path = file_path.replace('\\', '\\\\')
    
    def modify_document():
        # Load the existing document
        doc = Document(double_slash_path)

        # Modify the document as needed (e.g., add comments to table cells)
        # ...
        # ******
        #***********
        # *******

        # Save the modified document
        doc.save('C:\\Users\\yiren\\OneDrive\\Desktop\\new(1).docx')

        # Return the path to the modified document
        return 'C:\\Users\\yiren\\OneDrive\\Desktop\\new(1).docx'


    # Call the function to modify the document
    modified_document_path = modify_document()

# Create the main Tkinter window
window = Tk()
window.title("Select File")

# Create the upload button
upload_button = Button(window, text="Upload", command=upload_file)
upload_button.pack()

# Display the button on the window
upload_button.pack()

# Start the Tkinter event loop
window.mainloop()
