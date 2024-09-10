# Custom Avery Labels Generator

This VBA script is designed to generate custom Avery labels from a table in a Word document. Users can specify the Avery label template code, font size, and colors for specific keywords.

## Features

- Supports multiple Avery label templates (6560, 5160, 6576).
- Prompts the user to input font size.
- Allows users to specify colors for the keywords **Active**, **Reagent**, and **Corrosive**.
- Generates labels in a new Word document.

## How to Use

1. **Open the Word document** containing the table you want to use for generating labels.
2. Press `Alt + F11` to open the VBA editor.
3. **Insert a new module**:
   - Right-click on any existing module (or the document name) in the Project Explorer.
   - Choose `Insert > Module`.
4. **Copy and paste the VBA code** into the new module.
5. **Run the macro** by pressing `F5` or by going to `Run > Run Sub/UserForm`.

## Script Explanation

### Main Subroutine

`CreateCustomLabels`: This is the main subroutine that generates the labels.

- **Input Prompts**:
  - Avery label template code.
  - Colors for the keywords **Active**, **Reagent**, and **Corrosive**.
  - Font size for the labels.
- **Document References**:
  - References the active document and checks for the existence of at least one table.
- **Label Document Creation**:
  - Creates a new document for the labels and sets up the label template based on the user's input.
- **Table Iteration**:
  - Iterates through each row in the table (skipping the header) and extracts the necessary data.
  - Concatenates the data into the required label format.
  - Populates the labels in the new document.
- **Color Application**:
  - Applies the user-specified colors to the keywords in the labels.

### Helper Subroutines and Functions

- `SetupLabelTemplate`: Configures the label template in the new document based on the specified dimensions and margins.
- `CleanCellText`: Cleans up cell text by removing end-of-cell markers.
- `IsValidColor`: Validates if the provided color is a valid color name.
- `ApplyColorsToLabels`: Applies colors to the text in the labels based on user input.
- `ApplyColorToWord`: Applies a specific color to a word in a range.

## Error Handling

The script includes error handling to display a message if any error occurs during the execution of the macro.

## Example Usage

1. **Run the Macro**: Execute the macro by pressing `F5` in the VBA editor.
2. **Enter Avery Label Template Code**: When prompted, enter a supported Avery label template code (e.g., 6560, 5160, 6576).
3. **Enter Colors**: Input the desired colors for **Active**, **Reagent**, and **Corrosive** when prompted.
4. **Enter Font Size**: Input the desired font size for the labels.
5. **Generate Labels**: The macro will generate a new Word document with the custom labels, applying the specified colors and font size.

## Notes

- Ensure the table in the Word document has the correct structure and data before running the macro.
- The script supports a few Avery label templates; you can add more templates by modifying the `Select Case` block in the `CreateCustomLabels` subroutine.

# Save the content to a file
file_path = "/mnt/data/README.md"
with open(file_path, "w") as file:
    file.write(readme_content)

file_path
