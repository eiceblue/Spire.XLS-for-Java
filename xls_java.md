# Create Excel with Five Sheets
## Create an Excel workbook containing five worksheets with specific names
```java
//Create a workbook
Workbook workbook = new Workbook();
workbook.createEmptySheets(5);
for (int i = 0; i < 5; i++)
{
    Worksheet sheet = workbook.getWorksheets().get(i);
    sheet.setName("Sheet" + i);
}
```

---

# Create Excel with One Sheet
## Creates a workbook with a single empty sheet
```java
//Create a workbook
Workbook workbook = new Workbook();
workbook.createEmptySheets(1);
Worksheet sheet = workbook.getWorksheets().get(0);
```

---

# Create Multiple Excel Files
## Generate 50 Excel workbooks with multiple worksheets and data
```java
public class createFiftyExcelFiles {
    public static void main(String[] args) {
        for (int n = 0; n < 50; n++)
        {
            Workbook workbook = new Workbook();
            workbook.createEmptySheets(5);
            for (int i = 0; i < 5; i++)
            {
                Worksheet sheet = workbook.getWorksheets().get(i);
                sheet.setName("Sheet" + i);
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 50; col++)
                    {
                        sheet.get(row, col).setText("row" + row + " col" + col);
                    }
                }
            }

            workbook.saveToFile("output/workbook"+n+".xlsx", ExcelVersion.Version2010);
        }
    }
}
```

---

# spire.xls hello world example
## Create a simple Excel file with Hello World text
```java
//Create a workbook
Workbook workbook = new Workbook();
//Get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);
sheet.get("A1").setText("Hello World");

sheet.get("A1").autoFitColumns();

//Save to file
workbook.saveToFile("output/helloWorld.xlsx", ExcelVersion.Version2010);
```

---

# Open and Modify Excel File
## This code demonstrates how to open an existing Excel file and modify it by adding a new sheet and setting cell values.
```java
//Create a workbook
Workbook workbook = new Workbook();

workbook.loadFromFile("data/templateAz2.xlsx");

//Add a new sheet, named MySheet
Worksheet sheet = workbook.getWorksheets().add("MySheet");

//Get the reference of "A1" cell from the cells collection of a worksheet
sheet.get("A1").setText("Hello World");
```

---

# Excel Scroll Bar Control
## Add a scroll bar control to an Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Set a value for range B10
sheet.getCellRange("B10").setValue("1");
sheet.getCellRange("B10").getStyle().getFont().isBold(true);
//Add scroll bar control
IScrollBarShape scrollBar = sheet.getScrollBarShapes().addScrollBar(10, 3, 150, 20);
scrollBar.setLinkedCell(sheet.getCellRange("B10"));
scrollBar.setMin(1);
scrollBar.setMax(150);
scrollBar.setIncrementalChange(1);
scrollBar.setDisplay3DShading(true);
```

---

# Excel Spinner Control
## Add a spinner control to an Excel worksheet
```java
//Set text for range C11
sheet.getCellRange("C11").setText("Value:");
sheet.getCellRange("C11").getStyle().getFont().isBold(true);

//Set value for range B10
sheet.getCellRange("C12").setNumberValue(0);

//Add spinner control
ISpinnerShape spinner = sheet.getSpinnerShapes().addSpinner(12, 4, 20, 20);
spinner.setLinkedCell(sheet.getCellRange("C12"));
spinner.setMin(0);
spinner.setMax(100);
spinner.setIncrementalChange(5);
spinner.setDisplay3DShading(true);
```

---

# Excel Table with Filter
## Add a table with filter to Excel worksheet
```java
//Create a List Object named in Table.
sheet.getListObjects().create("Table", sheet.get(1, 1, sheet.getLastRow(), sheet.getLastColumn()));

//Set the BuiltInTableStyle for List object.
sheet.getListObjects().get(0).setBuiltInTableStyle(TableBuiltInStyles.TableStyleLight9);
```

---

# Add Total Row to Excel Table
## Core functionality for adding a total row to a table in Excel using Spire.XLS for Java
```java
//Create a table with the data from the specific cell range.
IListObject table = sheet.getListObjects().create("Table", sheet.get("A1:D4"));

//Display total row.
table.setDisplayTotalRow(true);

//Add a total row.
table.getColumns().get(0).setTotalsRowLabel("Total");
table.getColumns().get(1).setTotalsCalculation(ExcelTotalsCalculation.Sum);
table.getColumns().get(2).setTotalsCalculation(ExcelTotalsCalculation.Sum);
table.getColumns().get(3).setTotalsCalculation(ExcelTotalsCalculation.Sum);
```

---

# Apply Subscript and Superscript in Excel
## This code demonstrates how to apply subscript and superscript formatting to text in Excel cells using Spire.XLS for Java.
```java
//Set the rtf value of "B3" to "R100-0.06".
IXLSRange range = sheet.get("B3");
range.getRichText().setText("R100-0.06");

//Create a font. Set the IsSubscript property of the font to "true".
ExcelFont font = workbook.createFont();
font.isSubscript(true);
font.setColor(Color.green);

//Set font for specified range of the text in "B3".
range.getRichText().setFont(4, 8, font);

//Set the rtf value of "D3" to "a2 + b2 = c2".
range = sheet.get("D3");
range.getRichText().setText("a2 + b2 = c2");

//Create a font. Set the IsSuperscript property of the font to "true".
font = workbook.createFont();
font.isSuperscript(true);

//Set font for specified range of the text in "D3".
range.getRichText().setFont(1, 1, font);
range.getRichText().setFont(6, 6, font);
range.getRichText().setFont(11, 11, font);

sheet.getAllocatedRange().autoFitColumns();
```

---

# Excel Font Style Cloning
## Clone Excel font styles and modify properties
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Set A1 cell range's CellStyle.
CellStyle style = workbook.getStyles().addStyle("style");
style.getFont().setFontName("Calibri");
style.getFont().setColor(Color.red);
style.getFont().setSize(12);
style.getFont().isBold(true);
style.getFont().isItalic(true);

//Clone the same style for B2 cell range.
CellStyle csO1 = style.clone();

//Clone the same style for C3 cell range and then reset the font color for the text.
CellStyle csGreen = style.clone();
csGreen.getFont().setColor(Color.green);
```

---

# Spire.XLS Cell Range Copying
## Copy a range of cells to another location in Excel
```java
//Get the first worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Specify a destination range
CellRange cells = sheet1.getRange().get("G1:H19");

//Copy the selected range to destination range
sheet1.getRange().get("B1:C19").copy(cells);
```

---

# Excel Data Copy with Style
## Copy cell range data with formatting styles from source to destination
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the default first worksheet
Worksheet worksheet = workbook.getWorksheets().get(0);

//Get a source range (A1:D3).
CellRange srcRange = worksheet.getRange().get("A1:D3");

//Create a style object.
CellStyle style = workbook.getStyles().addStyle("style");

//Specify the font attribute.
style.getFont().setFontName("Calibri");

//Specify the shading color.
style.getFont().setColor(Color.red);

//Specify the border attributes.
style.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
style.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setColor(Color.blue);
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setColor(Color.blue);
style.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
style.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setColor(Color.blue);

srcRange.setCellStyleName(style.getName());

//Set the destination range
CellRange destRange = worksheet.getRange().get("A12:D14");

//Copy the range data with style
srcRange.copy(destRange,true,true);
```

---

# Excel Nested Group Creation
## Create nested row groups in Excel worksheet
```java
//Set the summary rows appear above detail rows.
sheet.getPageSetup().isSummaryRowBelow(false);

//Group the rows that you want to group.
sheet.groupByRows(2, 9, false);
sheet.groupByRows(4, 5, false);
sheet.groupByRows(8, 9, false);
```

---

# Create Excel Table
## Create a table in Excel worksheet and apply built-in style
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Add a new List Object to the worksheet
sheet.getListObjects().create("table", sheet.getCellRange(1, 1, 19, 5));

//Add Default Style to the table
sheet.getListObjects().get(0).setBuiltInTableStyle(TableBuiltInStyles.TableStyleLight9);
```

---

# Excel Custom Sorting
## Implement custom sorting in Excel using Spire.XLS for Java
```java
// Set whether header participates in sorting
wb.getDataSorter().isIncludeTitle(false);

// Custom sort
wb.getDataSorter().getSortColumns().add(0, new String[]
        {"DD","CC", "BB", "AA", "HH","GG","FF","EE"}
);
wb.getDataSorter().sort(wb.getWorksheets().get(0).getRange().get("A1:A8"));
```

---

# Excel Data Export
## Export worksheet data to DataTable
```java
//Open xls document
Workbook workbook = new Workbook();
workbook.loadFromFile("data/DataExport.xlsx");

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Export to dataTable
DataTable dataTable = sheet.exportDataTable();
```

---

# Excel Data Import and Styling
## Import data from a DataTable into an Excel worksheet and apply styling
```java
//Create a workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.getWorksheets().get(0);

//Insert data table
sheet.insertDataTable(datatable, true, 1, 1, -1, -1);

//Style the imported data
CellStyle oddStyle = workbook.getStyles().addStyle("oddStyle");
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
oddStyle.setKnownColor(ExcelColors.LightGreen1);
CellStyle evenStyle = workbook.getStyles().addStyle("evenStyle");
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
evenStyle.setKnownColor(ExcelColors.LightTurquoise);
for (CellRange range : sheet.getAllocatedRange().getRows()) {
    if (range.getRow() % 2 == 0)
        range.setCellStyleName(evenStyle.getName());
    else
        range.setCellStyleName(oddStyle.getName());
}

//Style the header
CellStyle styleHeader = sheet.getAllocatedRange().getRows()[0].getCellStyle();
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
styleHeader.setVerticalAlignment(VerticalAlignType.Center);
styleHeader.setKnownColor(ExcelColors.Green);
styleHeader.getExcelFont().setKnownColor(ExcelColors.White);
styleHeader.getExcelFont().isBold(true);
sheet.getAllocatedRange().autoFitColumns();
sheet.getAllocatedRange().autoFitRows();
sheet.getAllocatedRange().getRows()[0].setRowHeight(20);
```

---

# Excel Data Sorting
## Sort data in Excel worksheet by multiple columns in ascending order
```java
Worksheet worksheet = workbook.getWorksheets().get(0);

workbook.getDataSorter().getSortColumns().add(2, OrderBy.Ascending);
workbook.getDataSorter().getSortColumns().add(3, OrderBy.Ascending);

workbook.getDataSorter().sort(worksheet.getCellRange("A1:E19"));
```

---

# Excel Group Management
## Expand and collapse grouped rows in Excel
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Expand the grouped rows with ExpandCollapseFlags set to expand parent
sheet.getCellRange("A16:G19").expandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent);

//Collapse the grouped rows
sheet.getCellRange("A10:G12").collapseGroup(GroupByType.ByRows);
```

---

# Export Excel Data to DataTable
## Export data from Excel worksheet to DataTable while controlling data format preservation
```java
//Create a workbook
Workbook workbook=new Workbook();
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
//Export to datatable with data format options
ExportTableOptions options = new ExportTableOptions();
options.setKeepDataFormat(false);
options.setRenameStrategy(RenameStrategy.Digit);
DataTable table = sheet.exportDataTable(1, 1, sheet.getLastDataRow(), sheet.getLastDataColumn(), options);
```

---

# Find and Replace Data in Excel
## This code demonstrates how to find specific text in an Excel worksheet and replace it with new text, then highlight the modified cells.
```java
// Find the "Area" string
CellRange[] ranges = worksheet.findAllString("Area", false, false);

// Traverse the found ranges
for (CellRange range : ranges) {
    // Replace it with "Area Code"
    range.setText("Area Code");
    // Highlight the color
    range.getStyle().setColor(Color.yellow);
}
```

---

# Find Text by Regular Expression in Excel
## This code demonstrates how to find text in an Excel worksheet using regular expressions
```java
// Find cell ranges by Regex
CellRange[] ranges = worksheet.findAllString(".*North.", false, false, true);
String information = "";

// Get the information of every cell range
for (int i = 0; i < ranges.length; i++) {
    information += "RangeAddressLocal:" + ranges[i].getRangeAddressLocal() + "\r\n";
    information += "Text:" + ranges[i].getText() + "\r\n";
}
```

---

# Find Text in Cell Range
## Core functionality to find specific text within a specified range of cells in an Excel worksheet
```java
// Get the cell range from A16 to B20
CellRange range = sheet.getRange().get("A16:B20");

// Find all occurrences of "e-iceblue1" in the specified range and type
CellRange[] resultRange = range.findAll("e-iceblue1", FindType.Text, ExcelFindOptions.None);

// Check if any occurrences were found
if (resultRange.length != 0)
{
    // Iterate through each found occurrence
    for(CellRange r:resultRange)
    {
        // Get the address of the found cell
        String address = r.getRangeAddress();
    }
}
```

---

# Excel Table Formatting
## Format table style and set total row calculations
```java
//Add Default Style to the table
sheet.getListObjects().get(0).setBuiltInTableStyle(TableBuiltInStyles.TableStyleMedium9);

//Show total
sheet.getListObjects().get(0).setDisplayTotalRow(true);

//Set calculation type
sheet.getListObjects().get(0).getColumns().get(0).setTotalsRowLabel("Total");
sheet.getListObjects().get(0).getColumns().get(1).setTotalsCalculation(ExcelTotalsCalculation.None);
sheet.getListObjects().get(0).getColumns().get(2).setTotalsCalculation(ExcelTotalsCalculation.None);
sheet.getListObjects().get(0).getColumns().get(3).setTotalsCalculation(ExcelTotalsCalculation.Sum);
sheet.getListObjects().get(0).getColumns().get(4).setTotalsCalculation(ExcelTotalsCalculation.Sum);
sheet.getListObjects().get(0).setShowTableStyleRowStripes(true);
sheet.getListObjects().get(0).setShowTableStyleColumnStripes(true);
```

---

# Excel Goal Seek Implementation
## Perform goal seek calculation to find input value that achieves desired result
```java
//target cell
CellRange targetCell = worksheet.getCellRange("A2");
targetCell.setFormula("=SUM(A1+B1)");

//variable cell
CellRange gussCell = worksheet.getCellRange("B1");

//trial solution
com.spire.xls.GoalSeek goalSeek = new com.spire.xls.GoalSeek();
GoalSeekResult result = goalSeek.TryCalculate(targetCell, 500, gussCell);

//determine the solution
result.Determine();
```

---

# Import Data from ArrayList to Excel
## Core functionality for importing data from an ArrayList into an Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.createEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Create an ArrayList object
ArrayList list = new ArrayList();

//Add strings in list
list.add("Spire.Doc for Java");
list.add("Spire.XLS for Java");
list.add("Spire.PDF for Java");
list.add("Spire.Presentation for Java");

//Insert array list in worksheet
sheet.insertArrayList(list, 1, 1, true);

sheet.getAllocatedRange().autoFitColumns();
```

---

# Import Data from Data Column
## This code demonstrates how to import specific columns from a DataTable to a worksheet

```java
//Create a DataTable object
DataTable dataTable = new DataTable();
dataTable.getColumns().add("No", Integer.class);
dataTable.getColumns().add("Name", String.class);
dataTable.getColumns().add("City", String.class);

//Import the two columns of the data table to worksheet
DataColumn[] columns={dataTable.getColumns().get(1),dataTable.getColumns().get(2)};
sheet.insertDataColumns(columns, true, 1, 1);
```

---

# Import Data from DataTable to Excel
## This code demonstrates how to import data from a DataTable object into an Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.createEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Create a DataTable object
DataTable dataTable = new DataTable();
dataTable.getColumns().add("No", Integer.class);
dataTable.getColumns().add("Name", String.class);
dataTable.getColumns().add("City", String.class);

//Insert data from DataTable to worksheet
sheet.insertDataTable(dataTable, true, 1, 1);
```

---

# Excel Controls Insertion
## Insert various controls (textbox, checkbox, radio button, combobox) into Excel worksheet
```java
//Add a text box
ITextBoxShape textbox = ws.getTextBoxes().addTextBox(9, 2, 25, 100);
textbox.setText("Hello World");
//Add a checkbox
ICheckBox cb = ws.getCheckBoxes().addCheckBox(11, 2, 15, 100);
cb.setCheckState(CheckState.Checked);
cb.setText("Check Box 1");
//Add a Radio Button
IRadioButton rb = ws.getRadioButtons().add(13, 2, 15, 100);
rb.setText("Option 1");

//Add a combox
IComboBoxShape cbx = (IComboBoxShape)ws.getComboBoxes().addComboBox(15, 2, 15, 100);
cbx.setListFillRange(ws.getRange().get("A41:A47"));
```

---

# Excel HTML String Insertion
## Insert HTML content into Excel cells
```java
//create a workbook
Workbook workbook = new Workbook();
//get the first worksheet
Worksheet worksheet = workbook.getWorksheets().get(0);
//get the A1 range
CellRange range = worksheet.getCellRange("A1");
//insert html string
range.setHtmlString(htmlCode);
```

---

# Excel Text Replacement and Highlighting
## Find and replace text in Excel cells and highlight them with color
```java
CellRange[] ranges = sheet.findAllString("Total", true, true);

for (CellRange range : ranges)
{
    //reset the text, in other words, replace the text
    range.setText("Sum");

    //set the color
    range.getStyle().setColor(Color.yellow);
}
```

---

# Excel Partial Text Replacement
## Replace specific text within a cell's content
```java
//create a new workbook
Workbook workbook = new Workbook();

//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//set text value
sheet.getRange().get("A1").setText("Hello World");
sheet.getRange().get("A1").autoFitColumns();

//replace partial Text
sheet.getCellList().get(0).textPartReplace("World","Spire");
```

---

# Excel Data Retrieval and Extraction
## Extract specific rows from one Excel worksheet to another based on cell value
```java
// Create workbooks and get worksheets
Workbook newBook = new Workbook();
Worksheet newSheet = newBook.getWorksheets().get(0);
Workbook workbook = new Workbook();
Worksheet sheet = workbook.getWorksheets().get(0);

// Retrieve data and extract it to the first worksheet of the new excel workbook.
int i = 1;
int columnCount =  sheet.getColumns().length;
for(CellRange range :sheet.getColumns()[0].getCellList()){
    if (range.getText().equals("teacher")) {
        CellRange sourceRange = sheet.getRange().get(range.getRow(), 1, range.getRow(), columnCount);
        CellRange destRange = newSheet.getRange().get(i, 1, i, columnCount);
        sheet.copy(sourceRange, destRange,true);
        i++;
    }
}
```

---

# Set Array of Values into Excel Range
## Insert a 2D array of values into a worksheet range using Spire.XLS for Java
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Create an empty worksheet.
workbook.createEmptySheets(1);

//Get the worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Set the value of max row and column.
int maxRow = 10000;
int maxCol = 200;

//Output an array of data to a range of worksheet.
Object[][] myarray = new Object[maxRow + 1][maxCol + 1];
Boolean[][] isred = new Boolean[maxRow + 1][maxCol + 1];
for (int i = 0; i <= maxRow; i++) {
    for (int j = 0; j <= maxCol; j++) {
        myarray[i][j] = i + j;
        if ((int) myarray[i][j] > 8)
            isred[i][j] = true;
    }
}
sheet.insertArray(myarray, 1, 1);
```

---

# Split Excel Data Into Multiple Columns
## This code demonstrates how to split data in a single column into multiple columns based on space delimiter
```java
//Split data into separate columns by the delimited characters of space.
String[] splitText = null;
String text = null;
for (int i = 1; i < sheet.getLastRow(); i++)
{
    text = sheet.getRange().get(i + 1, 1).getText();
    splitText = text.split(" ");
    for (int j = 0; j < splitText.length; j++)
    {
        sheet.getRange().get(i + 1, 1 + j + 1).setText(splitText[j]);
    }
}
```

---

# Spire.XLS Table Data Sorting
## Sort table data in Excel using Spire.XLS library
```java
//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//add a new List Object to the worksheet
IListObject listObject = sheet.getListObjects().create("table", sheet.getCellRange(1,1,19,5));

//add default Style to the table
listObject.setBuiltInTableStyle(TableBuiltInStyles.TableStyleLight9);

//sorting
listObject.getAutoFilters().getSorter().getSortColumns().add(2, OrderBy.Ascending);
listObject.getAutoFilters().getSorter().sort(sheet.getCellRange(1,1,19,5));
```

---

# Excel Rich Text Formatting
## Apply different font styles to specific ranges of text in Excel cells
```java
// Create a workbook
Workbook workbook = new Workbook();
Worksheet sheet = workbook.getWorksheets().get(0);

// Create font
ExcelFont fontBold = workbook.createFont();
fontBold.isBold(true);

ExcelFont fontUnderline = workbook.createFont();
fontUnderline.setUnderline(FontUnderlineType.Single);

ExcelFont fontItalic = workbook.createFont();
fontItalic.isItalic(true);

ExcelFont fontColor = workbook.createFont();
fontColor.setKnownColor(ExcelColors.Green);

// Set rich text format
RichText richText = sheet.getCellRange("B11").getRichText();
richText.setText("Bold and underlined and italic and colored text.");
richText.setFont(0,3,fontBold);
richText.setFont(9,18,fontUnderline);
richText.setFont(24, 29, fontItalic);
richText.setFont(35,41,fontColor);
```

---

# Accessing Excel Cells
## Different methods to access cells in an Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Access cell by its name
CellRange range1 = sheet.getRange().get("A1");

//Access cell by index of row and column
CellRange range2 = sheet.getRange().get(2,1);

//Access cell in cell collection
CellRange range3 = sheet.getCells()[2];
```

---

# Apply Multiple Fonts in Single Cell
## Demonstrates how to apply different fonts to different parts of text in a single Excel cell
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Create a font object in workbook, setting the font color, size and type.
ExcelFont font1 = workbook.createFont();
font1.setKnownColor(ExcelColors.LightBlue);
font1.isBold(true);
font1.setSize(10);

//Create another font object specifying its properties.
ExcelFont font2 = workbook.createFont();
font2.setKnownColor(ExcelColors.Red);
font2.isBold(true);
font2.isItalic(true);
font2.setFontName("Times New Roman");
font2.setSize(11);

//Write a RichText string to the cell 'A1', and set the font for it.
RichText richText = sheet.getRange().get("B5").getRichText();
richText.setText("This document was created with Spire.XLS for Java.");
richText.setFont(0, 29, font1);
richText.setFont(31, 48, font2);
```

---

# Spire.XLS Cell Style Application
## Apply style to used cells in Excel worksheets
```java
//Create a cell style and set the parameter
CellStyle cellStyle = workbook.getStyles().addStyle("Mystyle");
cellStyle.setColor(Color.white);
cellStyle.getBorders().setKnownColor(ExcelColors.Black);
cellStyle.getBorders().setLineStyle(LineStyleType.Thin);
cellStyle.getBorders().getByBordersLineType(BordersLineType.DiagonalDown).setLineStyle(LineStyleType.None);
cellStyle.getBorders().getByBordersLineType(BordersLineType.DiagonalUp).setLineStyle(LineStyleType.None);

//Apply style for used cell
for (int i = 0; i < workbook.getWorksheets().size(); i++) {
    //false--false means only apply style to the used cells
    workbook.getWorksheets().get(i).applyStyle(cellStyle, false, false);
}
```

---

# Spire.XLS AutoFit Cell Size
## Auto-fit column width and row height based on cell value
```java
//Set value for B8
CellRange cell = worksheet.getRange().get("B8");
cell.setText("Welcome to Spire.XLS!");

//Set the cell style
CellStyle style = cell.getCellStyle();
style.getFont().setSize(16);
style.getFont().isBold(true);

//Auto fit column width and row height based on cell value
cell.autoFitColumns();
cell.autoFitRows();
```

---

# Excel Text to Number Conversion
## Convert text string format to number format in Excel cells
```java
//Convert text string format to number format
worksheet.getRange().get("D2:D8").convertToNumber();
```

---

# Spire.XLS Cell Format Copying
## Copy cell format from one column to another
```java
//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Copy the cell format from column 2 and apply to cells of column 5.
int count = sheet.getRows().length;
for (int i = 1; i < count + 1; i++) {
    sheet.getRange().get("E"+i).setStyle(sheet.getRange().get("B"+i).getStyle());
}
```

---

# Count Number of Cells in Excel Worksheet
## This code demonstrates how to count the total number of cells in a worksheet
```java
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the number of cells
int cellCount = sheet.getCells().length;
```

---

# Excel Cell Cutting Operation
## Cut cells from one position to another in Excel spreadsheet
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

CellRange Ori = sheet.getRange().get("A1:C5");
CellRange Dest = sheet.getRange().get("A26:C30");

//Copy the range to other position
sheet.copy(Ori, Dest, true, true, true);

//Remove all content in original cells
for(CellRange cr : Ori.getCellList()) {
    cr.clearAll();
}
```

---

# Detect and Unmerge Cells
## This code demonstrates how to detect merged cells in a worksheet and unmerge them.
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Get the merged cell ranges in the first worksheet and put them into a CellRange array.
CellRange[] range = sheet.getMergedCells();

//Traverse through the array and unmerge the merged cells.
for(CellRange cell : range){
    cell.unMerge();
}
```

---

# Duplicate Cell Range
## Copy data from source range to destination range and maintain the format
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);
//Copy data from source range to destination range and maintain the format.
sheet.copy(sheet.getRange().get("A6:F6"), sheet.getRange().get("A16:F16"), true);
```

---

# Spire.XLS Empty Cell Operations
## Clear content from Excel cells using different methods
```java
//Set the value as null to remove the original content from the Excel Cell.
sheet.getRange().get("C6").setValue("");

//Clear the content to remove the original content from the Excel Cell.
sheet.getRange().get("B6").clearContents();

//Remove the contents with format from the Excel cell.
sheet.getRange().get("D6").clearAll();
```

---

# Filter cells by cell color
## Apply auto-filter to cells based on their fill color
```java
//Create an auto filter in the sheet and specify the range to be filterd
sheet.getAutoFilters().setRange(sheet.getRange().get("G1:G19"));

//Get the column to be filter
FilterColumn filtercolumn = sheet.getAutoFilters().get(0);

//Add a color filter to filter the column based on cell color
sheet.getAutoFilters().addFillColorFilter(filtercolumn, Color.red);

//Filter the data.
sheet.getAutoFilters().filter();
```

---

# Find Cells with Style Name
## Find cells that have the same style as a reference cell and mark them
```java
//Get the cell style name
String styleName = sheet.getRange().get("A1").getCellStyleName();

for (int i = 1; i <= sheet.getLastRow(); i ++)
{
  for (int j =1; j <= sheet.getLastColumn(); j ++)
  {
    CellRange cr =  sheet.getCellRange(i,j);
    if (cr.getCellStyleName().equals(styleName))
    {
        cr.setValue("Same style");
    }
  }
}
```

---

# Find Formula Cells in Excel
## Code to find cells containing specific formulas in an Excel worksheet
```java
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Find the cells that contain formula "=SUM(A11,A12)"
CellRange[] ranges = sheet.findAll("=SUM(A11,A12)", EnumSet.of(FindType.Formula), EnumSet.of(ExcelFindOptions.None));
//Create a string builder
StringBuilder builder = new StringBuilder();

//Append the address of found cells to builder
String address;
if (ranges.length != 0) {
    for(CellRange range : ranges) {
        address = range.getRangeAddress();
        builder.append("The address of found cell is: " + address+"\n");
    }
} else {
    builder.append("No cell contain the formula"+"\n");
}
```

---

# Excel Cell Address Retrieval
## Get cell range addresses and properties in Excel
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Get a cell range
CellRange range = sheet.getRange().get("A1:B5");

//Get address of range
String address = range.getRangeAddressLocal();

//Get the cell count of range
int count = range.getCellsCount();

//Get the address of the entire column of range
String entireColAddress = range.getEntireColumn().getRangeAddressLocal();

//Get the address of the entire row of range
String entireRowAddress = range.getEntireRow().getRangeAddressLocal();
```

---

# Get Cell Data Type
## Retrieve and display cell data types in an Excel spreadsheet
```java
//Get the cell types of the cells in range H2:H7
for(CellRange range : sheet.getRange().get("H2:H7").getCellList()){
    Object cellType = sheet.getCellType(range.getRow(), range.getColumn(), false);
    sheet.getRange().get(range.getRow(), range.getColumn() + 1).setText(cellType.toString());
    sheet.getRange().get(range.getRow(), range.getColumn() + 1).getStyle().getFont().setColor(Color.red);
    sheet.getRange().get(range.getRow(), range.getColumn() + 1).getStyle().getFont().isBold(true);
}
```

---

# Get Cell Displayed Text
## Retrieve the formatted text displayed in an Excel cell
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get first worksheet of the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Set value for B8
CellRange cell = sheet.getRange().get("B8");
cell.setNumberValue(0.012345);

//Set the cell style
CellStyle style = cell.getCellStyle();
style.setNumberFormat("0.00");

//Get the cell value
String cellValue = cell.getValue();

//Get the displayed text of the cell
String displayedText = cell.getDisplayedText();
```

---

# Get Cell Value by Cell Name
## Demonstrates how to get the value of a specific cell using its name in an Excel worksheet
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Specify a cell by its name.
CellRange cell = sheet.getRange().get("A2");

//Get value of cell "A2".
cell.getValue()
```

---

# Excel Range Intersection
## Get intersection of two cell ranges in Excel
```java
//Get the two ranges.
CellRange range = sheet.getRange().get("A2:D7").intersect(sheet.getRange().get("B2:E8"));

StringBuilder content = new StringBuilder();
content.append("The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:"+"\n");

//Get the intersection of the two ranges.
for(CellRange r : range.getCellList())
{
    content.append(r.getValue()+"\n");
}
```

---

# Hide Cell Content in Excel
## Hide cell content by setting number format
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Hide the area by setting the number format as ";;;".
sheet.getRange().get("C5:D6").setNumberFormat(";;;");
```

---

# Spire.XLS Cell Merging
## Merge cells in Excel using column index and range reference
```java
//Merge the seventh column in Excel file.
workbook.getWorksheets().get(0).getColumns()[6].merge();

//Merge the particular range in Excel file.
workbook.getWorksheets().get(0).getRange().get("A14:D14").merge();
```

---

# Excel Formula Value Copying
## Copy only formula values from one cell range to another in Excel
```java
CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;
CellRange sourceRange = sheet.getCellRange("A6:E6");
sheet.copy(sourceRange, sheet.getCellRange("A8:E8"), EnumSet.of(copyOptions));
sourceRange.copy(sheet.getCellRange("A10:E10"), EnumSet.of(copyOptions));
```

---

# Excel Cell Fill Pattern
## Set cell color and fill pattern in Excel
```java
//Set cell color
worksheet.getRange().get("B7:F7").getStyle().setColor(Color.yellow);
//Set cell fill pattern
worksheet.getRange().get("B8:F8").getStyle().setFillPattern(ExcelPatternType.Percent125Gray);
```

---

# Excel DB Number Formatting
## Set DB number format for cells in Excel
```java
//Get the cell range
CellRange range = sheet.getRange().get("A1:A3");

//Set the DB num format
range.setNumberFormat("[DBNum2][$-804]General");

//Auto fit columns
range.autoFitColumns();
```

---

# Excel Cell Text Shrinking
## Shrink text to fit in a cell using Spire.XLS for Java
```java
// The cell range to shrink text
CellRange cell = sheet.getRange().get("B13:C13");

// Enable ShrinkToFit
CellStyle style = cell.getCellStyle();
style.setShrinkToFit(true);
```

---

# Traverse Cell Values in Excel
## Iterate through all cells in a worksheet and get their values
```java
//Get first worksheet of the workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

//Get the cell range collection
CellRange[] cellRangeCollection = worksheet.getCells();

//Traverse cells value
for(CellRange cellRange : cellRangeCollection)
{
    //Get cell address and value
    String cellAddress = cellRange.getRangeAddress();
    Object cellValue = cellRange.getValue();
}
```

---

# Excel Cell Ungrouping
## Ungroup specific rows in an Excel worksheet
```java
//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Ungroup the row 10 to 12.
sheet.ungroupByRows(10, 12);

//Ungroup the row 16 to 19.
sheet.ungroupByRows(16, 19);
```

---

# Excel Cell Unmerging
## Unmerge specific cells in an Excel worksheet
```java
//Unmerge the cells.
sheet.getRange().get("F2").unMerge();

//Unmerge the cells.
sheet.getRange().get("F7").unMerge();
```

---

# Spire.XLS for Java - Explicit Line Breaks
## Using explicit line breaks in Excel cells
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first default worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Specify a cell range
CellRange c5 = sheet1.getRange().get("C5");

//Set the cell width for specified range
sheet1.setColumnWidth(c5.getColumn(), 70);

//Put the string value with explicit line breaks
c5.setValue("Spire.XLS for Java is a professional Excel Java API\n that can be used to create, read, \nwrite, convert and print Excel files in Java application \nSpire.XLS for Java offers object model\n Excel API for speeding up Excel programming in Java platform -\n create new Excel documents from template, edit existing \nExcel documents and \nconvert Excel files.");

//Set Text wrap
c5.isWrapText(true);
```

---

# Excel Cell Text Wrapping
## Wrap or unwrap text in Excel cells
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Wrap the excel text;
sheet.getRange().get("C1").setText("e-iceblue is in facebook and welcome to like us");
sheet.getRange().get("C1").getStyle().setWrapText(true);
sheet.getRange().get("D1").setText("e-iceblue is in twitter and welcome to follow us");
sheet.getRange().get("D1").getStyle().setWrapText(true);

//Unwrap the excel text;
sheet.getRange().get("C2").setText("http://www.facebook.com/pages/e-iceblue/139657096082266");
sheet.getRange().get("C2").getStyle().setWrapText(false);
sheet.getRange().get("D2").setText("https://twitter.com/eiceblue");
sheet.getRange().get("D2").getStyle().setWrapText(false);
```

---

# Excel Column Auto-Fit
## Auto-fit columns within a specified range in an Excel worksheet
```java
//Auto fit the Column of the worksheet
sheet.autoFitColumn(2, 2, 5);
```

---

# Excel Row AutoFit
## Auto fit a specific row within a range in Excel worksheet
```java
// Auto fit the second row of the worksheet
sheet.autoFitRow(2, 1, 2, false);
```

---

# Spire.XLS Column Copying
## Copy columns within and between worksheets
```java
//Copy the first column to the third column in the same sheet
sheet1.copy(sheet1.getColumns()[0],sheet1.getColumns()[2],true,true,true);

//Copy the first column to the second column in the different sheet
sheet1.copy(sheet1.getColumns()[0],sheet2.getColumns()[1],true,true,true);
```

---

# Spire.XLS Copy Rows
## Copy rows within the same sheet and between different sheets
```java
Workbook workbook = new Workbook();
Worksheet sheet1 = workbook.getWorksheets().get(0);
Worksheet sheet2 = workbook.getWorksheets().get(1);

//Copy the first row to the third row in the same sheet
sheet1.copy(sheet1.getRows()[0], sheet1.getRows()[2], true, true, true);

//Copy the first row to the second row in the different sheet
sheet1.copy(sheet1.getRows()[0], sheet2.getRows()[1], true, true, true);
```

---

# Copy Single Column and Row in Excel
## Demonstrates how to copy a single column and row in Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Specify a destination range to copy one column
CellRange columnCells = sheet1.getRange().get("G1:G19");

//Copy the second column to destination range
sheet1.getColumns()[1].copy(columnCells);

//Specify a destination range to copy one row
CellRange rowCells = sheet1.getRange().get("A21:E21");

//Copy the first row to destination range
sheet1.getRows()[0].copy(rowCells);
```

---

# Spire.XLS Copy Range with Options
## Copy a range of cells from one worksheet to another with style preservation and reference updating
```java
//Get the first worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Add a new worksheet as destination sheet
Worksheet destinationSheet = workbook.getWorksheets().add("DestSheet");

//Specify a copy range of original sheet
CellRange cellRange = sheet1.getRange().get("B2:D4");

//Copy the specified range to added worksheet and keep original styles and update reference
workbook.getWorksheets().get(0).copy(cellRange, destinationSheet, 2, 1, true, true);
```

---

# Delete Blank Rows and Columns in Excel
## Core functionality to remove empty rows and columns from an Excel worksheet
```java
//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Delete blank rows from the worksheet.
for (int i = sheet.getRows().length - 1; i >= 0; i--) {
    if (sheet.getRows()[i].isBlank()) {
        sheet.deleteRow(i + 1);
    }
}

//Delete blank columns from the worksheet.
for (int j = sheet.getColumns().length - 1; j >= 0; j--) {
    if (sheet.getColumns()[j].isBlank()) {
        sheet.deleteColumn(j + 1);
    }
}
```

---

# Excel Row and Column Deletion
## Delete multiple rows and columns from an Excel worksheet
```java
//Delete 4 rows from the fifth row.
worksheet.deleteRow(5, 4);

//Delete 2 columns from the second column.
worksheet.deleteColumn(2, 2);
```

---

# Get Default Row and Column Count
## Retrieve the default number of rows and columns in a worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();
//Clear all worksheets
workbook.getWorksheets().clear();

//Create a new worksheet
Worksheet sheet = workbook.createEmptySheet();

//Get row and column count
int rowCount = sheet.getRows().length;
int columnCount = sheet.getColumns().length;
```

---

# Excel Row and Column Grouping
## Group rows and columns in an Excel worksheet
```java
//Group rows.
sheet.groupByRows(1,5,false);

//Group columns.
sheet.groupByColumns(1,3,false);
```

---

# Hide or Show Excel Row and Column Headers
## Demonstrates how to hide or show row and column headers in an Excel worksheet
```java
//Hide the headers of rows and columns
worksheet.setRowColumnHeadersVisible(false);

//Show the headers of rows and columns
//worksheet.setRowColumnHeadersVisible(true);
```

---

# Hide Excel Rows and Columns
## Demonstrate how to hide specific rows and columns in an Excel worksheet
```java
// Get the first worksheet.
Worksheet worksheet = workbook.getWorksheets().get(0);

// Hide the column of the worksheet.
worksheet.hideColumn(2);

// Hide the row of the worksheet.
worksheet.hideRow(4);
```

---

# Excel Row and Column Insertion
## Insert rows and columns in an Excel worksheet
```java
//Get the first worksheet.
Worksheet worksheet = workbook.getWorksheets().get(0);

//Insert a row into the worksheet.
worksheet.insertRow(2);

//Insert a column into the worksheet.
worksheet.insertColumn(2);

//Insert multiple rows into the worksheet.
worksheet.insertRow(5, 2);

//Insert multiple columns into the worksheet.
worksheet.insertColumn(5, 2);
```

---

# remove row based on keyword
## Remove a row from Excel worksheet that contains a specific keyword
```java
//Find the string
CellRange cr = sheet.findString("Address", false, false);

//Delete the row which includes the string
sheet.deleteRow(cr.getRow());
```

---

# Set Column Width in Pixels
## This code demonstrates how to set the width of a column in pixels in an Excel worksheet.
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet worksheet = workbook.getWorksheets().get(0);

//Set the width of the third column to 400 pixels
worksheet.setColumnWidthInPixels(3, 400);
```

---

# Setting Default Row and Column Styles
## Demonstrates how to set default styles for rows and columns in Excel
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Create a cell style and set the color
CellStyle style = workbook.getStyles().addStyle("MyStyle");
style.setColor(Color.yellow);

//Set the default style for the first row and column
sheet.setDefaultRowStyle(1, style);
sheet.setDefaultColumnStyle(1, style);
```

---

# Excel Default Row Height Setting
## Set default row height for Excel worksheet
```java
//Get the first worksheet.
Worksheet worksheet = workbook.getWorksheets().get(0);

//Set default row height
worksheet.setDefaultRowHeight(30);
```

---

# Excel Row Height and Column Width Setting
## Set the height and width of cells in an Excel worksheet
```java
//Get the first worksheet.
Worksheet worksheet = workbook.getWorksheets().get(0);

//Set the column width to 30.
worksheet.setColumnWidth(4, 30);

//Set the row height to 30.
worksheet.setRowHeight(4,30);
```

---

# Spire.XLS Summary Column Direction
## Set the direction of summary columns in Excel
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Group columns
sheet.groupByColumns(1, 3, true);

//Set summary columns to right of details
sheet.getPageSetup().isSummaryColumnRight(false);
```

---

# Spire.XLS Summary Row Direction
## Set the position of summary rows relative to grouped data
```java
//Get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//Group rows
sheet.groupByRows(1, 3, true);

//Set summary rows below details.
sheet.getPageSetup().isSummaryRowBelow(false);
```

---

# Excel Row and Column Operations
## Unhide hidden rows and columns in Excel worksheet
```java
//Unhide the row
sheet.showRow(15);

//Unhide the column
sheet.showColumn(4);
```

---

# Picture Alignment in Excel Cell
## Aligning a picture within a cell using Spire.XLS for Java
```java
//Insert an image to the specific cell
ExcelPicture picture = sheet.getPictures().add(1, 1, "data/SpireXls.png");

//Adjust the column width and row height so that the cell can contain the picture
sheet.setColumnWidth(1,40);
sheet.setRowHeight(1,200);

//Vertically and horizontally align the image
picture.setLeftColumnOffset(100);
picture.setTopRowOffset(25);
```

---

# Excel Picture Copying
## Copy a picture from one worksheet to another in an Excel file
```java
//Get the first worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Add a new worksheet as destination sheet
Worksheet destinationSheet = workbook.getWorksheets().add("DestSheet");

//Get the first picture from the first worksheet
ExcelPicture sourcePicture = sheet1.getPictures().get(0);

//Get the image
BufferedImage image = sourcePicture.getPicture();

//Add the image into the added worksheet
destinationSheet.getPictures().add(2,2,image);
```

---

# Get Cropped Position of Excel Picture
## Extract the position and dimensions of a cropped picture in an Excel worksheet
```java
// Get the first picture from the worksheet
ExcelPicture picture = sheet.getPictures().get(0);

// Get the cropped position
int left = picture.getLeft();
int top = picture.getTop();
int width = picture.getWidth();
int height = picture.getHeight();
```

---

# Excel Image Deletion
## Delete all images from an Excel worksheet
```java
//Get pictures from the first worksheet
PicturesCollection Pictures = sheet1.getPictures();

//Delete all images in the worksheet.
for (int i = Pictures.getCount() - 1; i >= 0; i--)
{
    Pictures.get(i).remove();
}
```

---

# Extract Embedded Images from Excel
## Retrieve and save images embedded in an Excel worksheet
```java
// Create a new Workbook object
Workbook wb = new Workbook();

// Load the Excel file "EmbedImageViaWps.xlsx" from the "data" directory
wb.loadFromFile("data/EmbedImageViaWps.xlsx");

// Access the first worksheet in the workbook
Worksheet sheet = wb.getWorksheets().get(0);

// Get an array of ExcelPicture objects representing cell images in the worksheet
ExcelPicture[] pc = sheet.getCellImages();

// Iterate through each ExcelPicture object in the array
for (int i = 0; i < pc.length; i++) {
    ExcelPicture ep = pc[i];

    // Get the BufferedImage of the ExcelPicture
    BufferedImage image = ep.getPicture();

    // Write the image to a PNG file in the output directory
    ImageIO.write(image, "PNG", new File(outputFile + String.format("result_%d.png", i)));
}
```

---

# Excel Background Image Insertion
## Code to insert a background image into an Excel worksheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//Set the image to be background image of the worksheet.
BufferedImage bufferedImage = ImageIO.read(new File("data/Background.png"));
sheet1.getPageSetup().setBackgoundImage(bufferedImage);
```

---

# Insert Image in WPS Cell
## Demonstrates how to insert an image into a specific cell in a worksheet
```java
//Create a new instance of Workbook
Workbook workbook = new Workbook();

//Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Embed an image into cell B1
sheet.getCellRange("B1").insertOrUpdateCellImage(inputImage,true);
```

---

# Insert Web Image into Excel
## This code demonstrates how to insert an image from a web URL into an Excel worksheet
```java
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the image
URL url = new URL("https://www.e-iceblue.com/downloads/demo/SpireXls.png");
BufferedImage bufferedImage = ImageIO.read(url);

sheet.getPictures().add(3, 2, bufferedImage);
```

---

# Excel Image Positioning
## Locate and reposition images in Excel worksheet
```java
//Get the first image in the sheet
ExcelPicture pic = sheet.getPictures().get(0);
//Set the position
pic.setLeftColumnOffset(300);
pic.setTopRowOffset(300);
```

---

# Excel Picture Offset Management
## Set left and top offsets for an image in Excel worksheet
```java
//Add an image to the specific cell
ExcelPicture pic = sheet.getPictures().add(2, 2,"data/SpireXls.png");

//Set left offset and top offset from the current range
pic.setLeftColumnOffset(200);
pic.setTopRowOffset(100);
```

---

# Excel Picture Reference Range
## Set reference range for a picture in Excel worksheet
```java
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Get the first picture in worksheet
ExcelPicture picture = sheet.getPictures().get(0);

//Set the referenced range of the picture to A1:B3
picture.setRefRange("A1:B3");
```

---

# Read Images from Excel
## Extract images from Excel worksheet
```java
//Create a Workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.loadFromFile("data/ReadImages.xlsx");

//Get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Get the first image
ExcelPicture pic = sheet.getPictures().get(0);
BufferedImage loImage = pic.getPicture();
ImageIO.write(loImage,"png",new File("output/ReadImages.png"));
```

---

# Reset Image Size and Position in Excel
## This code demonstrates how to reset the size and position of an image in an Excel worksheet
```java
//Add an image to the specific cell
ExcelPicture picture = sheet.getPictures().add(2, 2,"data/SpireXls.png");

//Set the size for the picture.
picture.setWidth(200);
picture.setHeight(200);

//Set the position for the picture.
picture.setLeft(200);
picture.setTop(200);
```

---

# Setting Image Offset for Excel Chart
## This code demonstrates how to set the image offset for a chart background in Excel using Spire.XLS for Java.
```java
//Add chart to worksheet
Chart chart1 = sheet1.getCharts().add(ExcelChartType.ColumnClustered);

//Chart Position
chart1.setLeftColumn(1);
chart1.setTopRow(11);
chart1.setRightColumn(8);
chart1.setBottomRow(33);

//Add picture as background
chart1.getChartArea().getFill().customPicture(bufferedImage,"None");

//Set the image offset
chart1.getChartArea().getFill().getPicStretch().setLeft(20);
chart1.getChartArea().getFill().getPicStretch().setTop(20);
chart1.getChartArea().getFill().getPicStretch().setRight(5);
chart1.getChartArea().getFill().getPicStretch().setBottom(5);
```

---

# Spire.XLS Image Writing
## Add image to Excel worksheet
```java
//Add an image to the specific cell
sheet.getPictures().add(14, 5,"data/SpireXls.png");
```

---

# Add Comment with Author in Excel
## This code demonstrates how to add a comment with author information to an Excel cell using Spire.XLS for Java
```java
// Open xls document
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Specify the cell range for regular comment
CellRange range = sheet.getCellRange("C1");

// Set the author and comment content
String author = "E-iceblue";
String text = "This is demo to show how to add a comment with editable Author property.";

// Add comment to the range and set properties
ExcelComment comment = range.addComment();
comment.setWidth(200);
comment.setVisible(true);
comment.setText(author + "\r" + text);

// Set the font of the author
ExcelFont font = workbook.createFont();
font.setFontName("Tahoma");
font.setKnownColor(ExcelColors.Black);
font.isBold(true);
comment.getRichText().setFont(0, author.length(), font);
```

---

# Excel Comment with Picture
## Add a comment with a background picture to an Excel cell
```java
// Create workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the cell range for comment
CellRange range = sheet.getCellRange("C6");

// Add comment to the cell
ExcelComment comment = range.addComment();

// Load the image file
BufferedImage bufferedImage = ImageIO.read(new File("data/Logo.png"));

// Fill the comment with a customized background picture
comment.getFill().customPicture(bufferedImage, "logo.png");

// Set the height and width of comment
comment.setHeight(bufferedImage.getHeight());
comment.setWidth(bufferedImage.getWidth());
comment.setVisible(true);
```

---

# Edit Excel Comment
## How to edit an existing comment in an Excel worksheet
```java
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Get the first comment.
ExcelComment comment = sheet.getComments().get(0);

//Edit the comment.
comment.setText("This comment has been edited by Spire.XLS.");
```

---

# Excel Name Manager Comment Extraction
## Extract comments from Name Manager in Excel workbook
```java
// Access the NameRanges property of the workbook
INameRanges nameManager = workbook.getNameRanges();

// Create a StringBuilder to store the result
StringBuilder sb = new StringBuilder();

// Iterate through each name in the NameRanges collection
for (int i = 0; i < nameManager.getCount(); i++)
{
    // Get the XlsName object at index i
    XlsName name = (XlsName)nameManager.get(i);

    // Append the name and comment value to the StringBuilder
    sb.append("Name: " + name.getName() + ", Comment: " + name.getCommentValue() + "\r\n");
}
```

---

# Excel Comment Visibility Control
## Hide or show Excel comments using Spire.XLS for Java
```java
//Hide the second comment
sheet.getComments().get(1).isVisible(false);

//Show the third comment
sheet.getComments().get(2).isVisible(true);
```

---

# Read Excel Cell Comments
## Demonstrates how to read comments from Excel cells
```java
//Open workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Read plain text comment from cell A1
String commentA1 = sheet.getCellRange("A1").getComment().getText();

//Read rich text comment from cell A2
String commentA2 = sheet.getCellRange("A2").getComment().getRichText().getRtfText();
```

---

# Excel Comment Management
## Edit and remove comments from an Excel worksheet
```java
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Get all comments in the sheet
CommentsCollection comments = sheet.getComments();

//Edit the content of first comment.
comments.get(0).setText("This comment has been edited by Spire.XLS.");
//Remove the second comment.
comments.get(1).remove();
```

---

# Excel Comment Text Rotation
## Set text rotation for cell comments in Excel
```java
//Create font
ExcelFont font = workbook.createFont();
font.setFontName("Arial");
font.setSize(11);
font.setKnownColor(ExcelColors.Orange);

CellRange range = sheet.getCellRange("E1");
range.getComment().setText("This is a comment");
range.getComment().getRichText().setFont(0, (range.getComment().getText().length() - 1), font);

// Set its vertical and horizontal alignment
range.getComment().setVAlignment(CommentVAlignType.Center);
range.getComment().setHAlignment(CommentHAlignType.Left);

//Set the comment text rotation
range.getComment().setTextRotation(TextRotationType.LeftToRight);
```

---

# Excel Comment Position and Alignment
## Set position and alignment for Excel comments
```java
//Create font
ExcelFont font1 = workbook.createFont();
font1.setFontName("Calibri");
font1.setSize(12);
font1.setColor(Color.orange);
font1.isBold(true);

ExcelFont font2 = workbook.createFont();
font2.setFontName("Calibri");
font2.setSize(12);
font2.setColor(Color.blue);
font2.isBold(true);

//Add comment 1 and set its size, text, position and alignment
sheet.getCellRange("G5").setText("Spire.XLS");
ExcelComment Comment1 = sheet.getCellRange("G5").getComment();
Comment1.isVisible(true);
Comment1.setHeight(150);
Comment1.setWidth(300);
Comment1.getRichText().setText("Spire.XLS for Java:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ");
Comment1.getRichText().setFont(0, 19, font1);
Comment1.setTextRotation(TextRotationType.LeftToRight);
//Set the position of Comment
Comment1.setTop(20);
Comment1.setLeft(40);
//Set the alignment of text in Comment
Comment1.setVAlignment(CommentVAlignType.Center);
Comment1.setHAlignment(CommentHAlignType.Justified);

//Add comment2 and set its size, text, position and alignment for comparison
sheet.getCellRange("D14").setText("E-iceblue");
ExcelComment Comment2 = sheet.getCellRange("D14").getComment();
Comment2.isVisible(true);
Comment2.setHeight(150);
Comment2.setWidth(300);
Comment2.getRichText().setText("About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.");
Comment2.getRichText().setFont(0, 16, font2);
Comment2.setTextRotation(TextRotationType.LeftToRight);
//Set the position of Comment
Comment2.setTop(170);
Comment2.setLeft(450);
//Set the alignment of text in Comment
Comment2.setVAlignment(CommentVAlignType.Top);
Comment2.setHAlignment(CommentHAlignType.Justified);
```

---

# Excel Comment Writing
## Write regular and rich text comments to Excel cells
```java
//Create font
ExcelFont font = workbook.createFont();
font.setFontName("Arial");
font.setSize(11);
font.setKnownColor(ExcelColors.Orange);
ExcelFont fontBlue = workbook.createFont();
fontBlue.setKnownColor(ExcelColors.LightBlue);
ExcelFont fontGreen = workbook.createFont();
fontGreen.setKnownColor(ExcelColors.LightGreen);

//Specify the cell range for regular comment
CellRange range = sheet.getCellRange("B11");
range.setText("Regular comment");
range.getComment().setText("Regular comment");
range.autoFitColumns();

//Specify the cell range for rich text comment
range = sheet.getCellRange("B12");
range.setText("Rich text comment");
range.getRichText().setFont(0, 16, font);
range.autoFitColumns();

//Set font color for rich text comment
range.getComment().getRichText().setText("Rich text comment");
range.getComment().getRichText().setFont(0, 4, fontGreen);
range.getComment().getRichText().setFont(5, 9, fontBlue);
```

---

# Chart Sheet to SVG Conversion
## Convert Excel chart sheet to SVG format
```java
//Get the chartsheet by name
ChartSheet sheet = workbook.getChartSheetByName("Chart1");
FileOutputStream stream = new FileOutputStream("output/chartSheetToSVG.svg");

sheet.toSVGStream(stream);
stream.flush();
stream.close();
```

---

# CSV to DataTable Conversion
## Convert CSV file to DataTable using Spire.XLS
```java
//Create a workbook and load a csv file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile, ",", 1, 1);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
//Export to dataTable
DataTable dataTable = sheet.exportDataTable();
```

---

# CSV to Excel Conversion
## Convert CSV file to Excel format using Spire.XLS library
```java
//Create a workbook and load a csv file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile, ",", 1, 1);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
sheet.getCellRange("D2:E19").setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));
sheet.getAllocatedRange().autoFitColumns();

//Save the Excel file
workbook.saveToFile(outputFile, ExcelVersion.Version2013);
```

---

# CSV to PDF Conversion
## Convert CSV file to PDF format using Spire.XLS library
```java
//Create a workbook and load a csv file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile, ",", 1, 1);

//Set the setSheetFitToPage property as true
workbook.getConverterSetting().setSheetFitToPage(true);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Autofit a column if the characters in the column exceed column width
for (int i = 1; i < sheet.getColumns().length; i++)
{
    sheet.autoFitColumn(i);
}

//Save to PDF document
workbook.saveToFile(output,  FileFormat.PDF);
```

---

# Excel Worksheet to PDF Conversion
## Convert each worksheet in an Excel workbook to a separate PDF file
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load a file
workbook.loadFromFile(inputFile);

for(int i = 0; i < workbook.getWorksheets().getCount(); i ++)
{
   Worksheet worksheet = workbook.getWorksheets().get(i);
   String result = "output/sheet-" + i + "-result.pdf";
   worksheet.saveToPdf(result);
}
```

---

# ET to XLS Conversion
## Convert ET file format to XLS format using Spire.XLS
```java
Workbook workbook = new Workbook();
workbook.loadFromFile("data/Sample.et");
workbook.saveToFile("output/ETtoXls.xls", FileFormat.Version97to2003);
```

---

# ETT to XLS Conversion
## Convert ETT file format to XLS format using Spire.XLS
```java
Workbook workbook = new Workbook();
workbook.loadFromFile("data/Sample.ett");
workbook.saveToFile("output/ETTtoXls.xls", FileFormat.Version97to2003);
```

---

# Excel to HTML Conversion
## Convert Excel workbook to HTML format
```java
Workbook workbook = new Workbook();
workbook.loadFromFile("data/WorkbookToHTML.xlsx");

//Convert to html
workbook.saveToFile("output/excelToHtml_result.html",FileFormat.HTML);
```

---

# Excel to PDF Conversion with Width Fitting
## Configure Excel worksheets to fit width when converting to PDF
```java
for(int i = 0; i < workbook.getWorksheets().getCount(); i ++)
{
   Worksheet worksheet = workbook.getWorksheets().get(i);
    //Auto fit page height
    worksheet.getPageSetup().setFitToPagesTall(0);
    //Fit one page width
    worksheet.getPageSetup().setFitToPagesWide(1);
}
```

---

# Spire.XLS SVG Conversion
## Get dimensions of converted SVG from Excel worksheet
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Convert the worksheet to an SVG stream and get the dimensions of the resulting SVG
Dimension dimension = sheet.toSVGStream(stream, sheet.getFirstRow(), sheet.getFirstColumn(), sheet.getLastRow(), sheet.getLastColumn());

// Get the dimensions of the SVG
double height = dimension.getHeight();
double width = dimension.getWidth();
```

---

# HTML to Excel Conversion
## Convert HTML files to Excel format using Spire.XLS library
```java
// Create a workbook
Workbook workbook = new Workbook();
// Load HTML file
workbook.loadFromHtml(inputFile);

// Save workbook to Excel file
workbook.saveToFile(outputFile, ExcelVersion.Version2013);
```

---

# Office Open XML to Excel Conversion
## Convert XML file to Excel format using Spire.XLS library
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load from xml file
FileInputStream fileStream = new FileInputStream(inputFile);
try {
    workbook.loadFromXml(fileStream);
}finally {
    if(fileStream != null)
        fileStream.close();
}

//Save to Excel file
workbook.saveToFile(outputFile, ExcelVersion.Version2010);
```

---

# Excel Range to PDF Conversion
## Convert a selected range of cells from Excel to PDF format
```java
// Create a workbook
Workbook workbook = new Workbook();

// Add a new sheet to workbook
workbook.getWorksheets().add("newsheet");

// Copy your area to new sheet
workbook.getWorksheets().get(0).getCellRange("A9:E15").copy(workbook.getWorksheets().get(1).getCellRange("A9:E15"), true);

// Auto fit column width
workbook.getWorksheets().get(1).getCellRange("A9:E15").autoFitColumns();

// Save the worksheet to PDF
Worksheet worksheet = workbook.getWorksheets().get(1);
worksheet.saveToPdf("output/selectedRangeToPDF.pdf");
```

---

# Excel Sheet to Image Conversion
## Convert an Excel worksheet to an image file
```java
//Create a workbook and load a file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Save the sheet to image
BufferedImage bufferedImage = sheet.toImage(sheet.getFirstRow(), sheet.getFirstColumn(), sheet.getLastRow(), sheet.getLastColumn());
ImageIO.write(bufferedImage,"PNG",new File(outputFile));
```

---

# Convert specific Excel cells to image
## This code demonstrates how to convert a specific range of cells in an Excel worksheet to an image file
```java
// Create a workbook and load a file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile);

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Convert specific cell ranges to image
BufferedImage bufferedImage = sheet.toImage(1, 1, 7, 5);

// Save the image to a file
ImageIO.write(bufferedImage, "PNG", new File(outputFile));
```

---

# Spire.XLS Font Directory Specification
## Specify custom font directory for Excel to PDF conversion
```java
//Create a workbook
Workbook workbook = new Workbook();

//Specifies the font file directory
workbook.setCustomFontFileDirectory(new String[]{"./data/Font"});

//Get custom font analysis results
Hashtable hashtable = workbook.getCustomFontParsedResult();
```

---

# Excel to CSV Conversion
## Convert Excel worksheet to CSV format
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Convert to CSV file
sheet.saveToFile(outputFile, ",", Charset.forName("UTF-8"));
```

---

# Excel to CSV Conversion with Double Quotes
## Convert Excel file to CSV format with double quotes using Spire.XLS library
```java
// Create a workbook
Workbook workbook = new Workbook();

// Convert to CSV file with double quotes
// When the last parameter is set to true, there are double quotes. The default parameter is false
workbook.saveToFile("ToCSVAddQuotation.csv", ",", true);
```

---

# Excel to HTML Conversion
## Convert Excel worksheet to HTML format with embedded images
```java
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Set embedded image
HTMLOptions options = new HTMLOptions();
options.setImageEmbedded(true);

// Save to HTML file
sheet.saveToHtml(outputFile, options);
```

---

# Excel to HTML Stream Conversion
## Convert Excel worksheet to HTML format using stream output
```java
// Get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Set embedded image
HTMLOptions options = new HTMLOptions();
options.setImageEmbedded(true);

// Save to HTML stream
sheet.saveToHtml(stream, options);
```

---

# Excel to HTML Conversion with Hidden Worksheets
## Convert Excel file to HTML while preserving hidden worksheets
```java
//Create a workbook
Workbook book = new Workbook();

//Load the Excel document
book.loadFromFile("data/ToHtmlWithHiddenWorksheets.xlsx");

//Save to HTML with hidden worksheets (false = include hidden worksheets)
String result = "output/ToHtml_result.html";
book.saveToHtml(result, false);
```

---

# Excel to Image Conversion with Comments
## Convert Excel worksheet to image while preserving comments
```java
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.getWorksheets().get(0);
worksheet.getPageSetup().setPrintComments(PrintCommentType.InPlace);
int firstRow = worksheet.getFirstRow();
int firstColumn = worksheet.getFirstColumn();
int lastRow = worksheet.getLastRow();
int lastColumn = worksheet.getLastColumn();
BufferedImage bufferedImage = worksheet.toImage(firstRow, firstColumn, lastRow, lastColumn);
```

---

# Excel to High Resolution Image Conversion
## Convert Excel worksheets to images with high resolution settings
```java
//Create a Workbook
Workbook workbook = new Workbook();

//Iterate through all the worksheets
for (int i = 0; i < workbook.getWorksheets().size(); i++) {
    Worksheet sheet = workbook.getWorksheets().get(i);
    //Set the image resolution
    workbook.getConverterSetting().setXDpi(300);
    workbook.getConverterSetting().setYDpi(300);
    //Save the image
    sheet.saveToImage("output/imageResolution-" + i + ".png");
}
```

---

# Excel to Image Conversion Without White Space
## Convert Excel worksheet to image without surrounding white space by setting margins to zero
```java
//Create a workbook and load a file
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile);

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Set the margin as 0 to remove the white space around the image
sheet.getPageSetup().setLeftMargin(0);
sheet.getPageSetup().setTopMargin(0);
sheet.getPageSetup().setRightMargin(0);
sheet.getPageSetup().setBottomMargin(0);

//Save the sheet to image
BufferedImage bufferedImage = sheet.toImage(sheet.getFirstRow(), sheet.getFirstColumn(), sheet.getLastRow(), sheet.getLastColumn());
```

---

# Excel to ODS Conversion
## Convert Excel files to ODS format using Spire.XLS library
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load an excel document
workbook.loadFromFile(inputFile);

//Convert to ODS file
workbook.saveToFile(outputFile, FileFormat.ODS);
```

---

# Excel to Office Open XML Conversion
## Convert Excel workbook to Office Open XML format
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Set text for cell
sheet.getCellRange("A1").setText("Hello World");
//Set color for cell
sheet.getCellRange("B1").getCellStyle().setKnownColor(ExcelColors.Gray25Percent);
sheet.getCellRange("C1").getCellStyle().setKnownColor(ExcelColors.Gold);

//Save to xml file
workbook.saveAsXml(outputFile);
```

---

# Excel to PDF conversion
## Convert Excel file to PDF format using Spire.XLS
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load a file
workbook.loadFromFile(inputFile);
//Fit to page
workbook.getConverterSetting().setSheetFitToPage(true);
//Save to PDF file
workbook.saveToFile(outputFile, FileFormat.PDF);

workbook.dispose();
```

---

# Excel to PDF Conversion
## Simple conversion of Excel files to PDF format
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load a file
workbook.loadFromFile(inputFile);

//Save to PDF file
workbook.saveToFile(outputFile, FileFormat.PDF);
```

---

# Excel to PDF Conversion with Page Size Change
## Convert Excel file to PDF after changing the page size to A3 for all worksheets
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load a file
workbook.loadFromFile(inputFile);

for(int i = 0; i < workbook.getWorksheets().getCount(); i ++)
{
   Worksheet worksheet = workbook.getWorksheets().get(i);
   //Change the page size
   worksheet.getPageSetup().setPaperSize(PaperSizeType.PaperA3);
}
workbook.saveToFile(outputPath, FileFormat.PDF);
```

---

# Excel to PDF conversion with custom paper size
## Convert Excel file to PDF with custom paper size settings
```java
// Create a Workbook
Workbook workbook = new Workbook();

// Get the first worksheet and set the custom paper size
workbook.getWorksheets().get(0).getPageSetup().setCustomPaperSize(100,100);
```

---

# Excel to PostScript Conversion
## Convert Excel files to PostScript format using Spire.XLS for Java
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load an excel document
workbook.loadFromFile("data/ToPostScript.xlsx");

String result = "output/ToPostScript.ps";
//Convert to PS file
workbook.saveToFile(result, FileFormat.PostScript);
```

---

# Excel to HTML Conversion
## Convert Excel file to standalone HTML format
```java
// Create a workbook
Workbook wb = new Workbook();

// Set HTMLOptions for standalone HTML
HTMLOptions.Default.isStandAloneHtmlFile(true);

// Save workbook to HTML format
wb.saveToStream(fileStream, FileFormat.HTML);
```

---

# Excel to SVG Conversion
## Convert Excel worksheets to SVG format
```java
//Open xls document
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile);
//Traverse worksheets
for (int i = 0; i < workbook.getWorksheets().size(); i++)
{
    FileOutputStream stream = new FileOutputStream("output/sheet"+i+".svg");
    //Convert worksheet to svg file
    Worksheet sheet = workbook.getWorksheets().get(i);
    sheet.toSVGStream(stream, sheet.getFirstRow(), sheet.getFirstColumn(), sheet.getLastRow(), sheet.getLastColumn());
    stream.flush();
    stream.close();
}
```

---

# Excel to Text Conversion
## Convert Excel worksheet to text format using Spire.XLS
```java
// Open xls document
Workbook workbook = new Workbook();

Worksheet worksheet = workbook.getWorksheets().get(0);
// Convert to text
Charset charset = Charset.forName("utf8");
worksheet.saveToFile(outputFile, " ", charset);
```

---

# Excel to TIFF Conversion
## Convert Excel worksheet to TIFF image format
```java
//Create a Workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Save the first worksheet to tiff
String outputName="output/toTiffResult.tif";
sheet.saveToTiff(outputName);
```

---

# Excel to XPS Conversion
## Convert Excel file to XPS format using Spire.XLS library
```java
String  inputFile = "data/CreateTable.xlsx";
String  outputFile = "output/ToXPS.xps";
//Open xls document
Workbook workbook = new Workbook();
workbook.loadFromFile(inputFile);
//Convert to XPS
workbook.saveToFile(outputFile, FileFormat.XPS);
```

---

# UOS to Excel Conversion
## Convert UOS format file to Excel format
```java
//Create a workbook
Workbook workbook=new Workbook();
//Load the UOS from disk
workbook.loadFromFile("data/input.uos",ExcelVersion.UOS);
//Convert to Excel
workbook.saveToFile("output/output.xlsx",ExcelVersion.Version2013);
```

---

# XLSB Data Conversion and Styling
## Convert XLSB data to DataTable and apply styling
```java
//Export the first sheet data to dataTable
DataTable datatable=ToDataTable(inputFile,0);
//create a workbook
Workbook workbook = new Workbook();
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
//Insert datatable
sheet.insertDataTable(datatable,true,1,1,-1,-1);
//Set body style
CellStyle oddStyle = workbook.getStyles().addStyle("oddStyle");
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
oddStyle.setKnownColor(ExcelColors.LightGreen1);
CellStyle evenStyle = workbook.getStyles().addStyle("evenStyle");
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
evenStyle.setKnownColor(ExcelColors.LightTurquoise);
for (CellRange range : sheet.getAllocatedRange().getRows())
{
    if (range.getRow() % 2 == 0)
        range.setCellStyleName(evenStyle.getName());
    else
        range.setCellStyleName(oddStyle.getName());
}
//Set header style
CellStyle styleHeader = sheet.getRows()[0].getCellStyle();
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
styleHeader.setVerticalAlignment(VerticalAlignType.Center);
styleHeader.setKnownColor(ExcelColors.Green);
styleHeader.getExcelFont().setKnownColor(ExcelColors.White);
styleHeader.getExcelFont().isBold(true);
sheet.getAllocatedRange().autoFitColumns();
sheet.getAllocatedRange().autoFitRows();
sheet.getRows()[0].setRowHeight(20);

private static DataTable ToDataTable (String inputFile, int worksheet) {
    //Open xls document
    Workbook workbook = new Workbook();
    workbook.loadFromFile(inputFile);
    //Export the first sheet data to dataTable
    Worksheet sheet = workbook.getWorksheets().get(worksheet);
    return sheet.exportDataTable();
}
```

---

# Excel to ET Conversion
## Convert XLS files to ET format using Spire.XLS
```java
Workbook workbook = new Workbook();
workbook.saveToFile("output/XlsToET.et", FileFormat.ET);
```

---

# XLS to ETT Conversion
## Convert Excel XLS files to ETT format using Spire.XLS library
```java
Workbook workbook = new Workbook();
workbook.loadFromFile("data/Sample.xls");
workbook.saveToFile("output/XlsToETT.ett", FileFormat.ETT);
```

---

# Excel file conversion
## Convert XLS to XLSM format
```java
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.loadFromFile("data/MacroSample.xls",ExcelVersion.Version97to2003);

//Save the workbook as a new XLSM file
String output = "output/XLSToXLSM.xlsm";
workbook.saveToFile(output);
```

---

# Excel AutoFilter for Blank Cells
## Apply auto-filter to match blank cells in Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Match the blank data in the first column
sheet.getAutoFilters().matchBlanks(0);

// Apply the filter
sheet.getAutoFilters().filter();
```

---

# Excel AutoFilter Non-Blank Cells
## Apply auto-filter to show only non-blank cells in Excel sheet
```java
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Match the non blank data
sheet.getAutoFilters().matchNonBlanks(0);

//Filter
sheet.getAutoFilters().filter();
```

---

# Excel AutoFilter Creation
## Create auto filter for a specific cell range in Excel worksheet
```java
//Create filter
sheet.getAutoFilters().setRange(sheet.getCellRange("A1:J1"));
```

---

# Excel Data Validation
## Implement different types of data validation in Excel cells
```java
//Decimal DataValidation
sheet.getCellRange("B11").setText("Input Number(3-6):");
CellRange rangeNumber = sheet.getCellRange("B12");
//Set the operator for the data validation.
rangeNumber.getDataValidation().setCompareOperator(ValidationComparisonOperator.Between);
//Set the value or expression associated with the data validation.
rangeNumber.getDataValidation().setFormula1("3");
//The value or expression associated with the second part of the data validation.
rangeNumber.getDataValidation().setFormula2("6");
//Set the data validation type.
rangeNumber.getDataValidation().setAllowType(CellDataType.Decimal);
//Set the data validation error message.
rangeNumber.getDataValidation().setErrorMessage("Please input correct number!");
//Enable the error.
rangeNumber.getDataValidation().setShowError(true);

//Date DataValidation
sheet.getCellRange("B14").setText("Input Date:");
CellRange rangeDate = sheet.getCellRange("B15");
rangeDate.getDataValidation().setAllowType(CellDataType.Date);
rangeDate.getDataValidation().setCompareOperator(ValidationComparisonOperator.Between);
rangeDate.getDataValidation().setFormula1("1/1/1970");
rangeDate.getDataValidation().setFormula2("12/31/1970");
rangeDate.getDataValidation().setErrorMessage("Please input correct date!");
rangeDate.getDataValidation().setShowError(true);
rangeDate.getDataValidation().setAlertStyle(AlertStyleType.Warning);

//TextLength DataValidation
sheet.getCellRange("B17").setText("Input Text:");
CellRange rangeTextLength = sheet.getCellRange("B18");
rangeTextLength.getDataValidation().setAllowType(CellDataType.TextLength);
rangeTextLength.getDataValidation().setCompareOperator(ValidationComparisonOperator.LessOrEqual);
rangeTextLength.getDataValidation().setFormula1("5");
rangeTextLength.getDataValidation().setErrorMessage("Enter a Valid String!");
rangeTextLength.getDataValidation().setShowError(true);
rangeTextLength.getDataValidation().setAlertStyle(AlertStyleType.Stop);
```

---

# Filter Cells By String
## Filter cells in Excel that start with a specific string value
```java
// Filter the cell which starts with "South".
sheet.getAutoFilters().setRange(sheet.get("D1:D19"));
FilterColumn filtercolumn = (FilterColumn) sheet.getAutoFilters().get(0);
sheet.getAutoFilters().customFilter(filtercolumn, FilterOperatorType.Equal, "South*");
sheet.getAutoFilters().filter();
```

---

# Excel Data Validation Settings
## Get settings of data validation from a cell
```java
//Cell B4 has the Decimal Validation
CellRange cell = sheet.getCellRange("B4");

//Get the validation of this cell
Validation validation = cell.getDataValidation();

//Get the settings
String allowType = validation.getAllowType().toString();
String data = validation.getCompareOperator().toString();
String minimum = validation.getFormula1().toString();
String maximum = validation.getFormula2().toString();
boolean ignoreBlank = validation.getIgnoreBlank();
```

---

# Excel List Data Validation
## Implement list data validation for Excel cells
```java
//Set data validation for cell
CellRange range = sheet.getCellRange("D10");
range.getDataValidation().setShowError(true);
range.getDataValidation().setAlertStyle(AlertStyleType.Stop);
range.getDataValidation().setErrorTitle("Error");
range.getDataValidation().setErrorMessage("Please select a city from the list");
range.getDataValidation().setDataRange(sheet.getCellRange("A7:A10"));
```

---

# Remove Auto Filters from Excel Worksheet
## Remove auto filters using Spire.XLS for Java
```java
//Remove the auto filters
sheet.getAutoFilters().clear();
```

---

# Excel Data Validation Removal
## Remove data validation from specified ranges in Excel worksheet
```java
//Create a workbook.
Workbook workbook = new Workbook();

//Create an array of rectangles, which is used to locate the ranges in worksheet.
Rectangle[] rectangles = new Rectangle[1];

//Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles[0] = new Rectangle(0, 0, 1, 2);

//Remove validations in the ranges represented by rectangles.
workbook.getWorksheets().get(0).getDVTable().remove(rectangles);
```

---

# Excel Data Validation Across Sheets
## Set data validation on a cell that references a range on a separate sheet
```java
//get the first sheet
Worksheet sheet1 = workbook.getWorksheets().get(0);

//get a cellRange
sheet1.getCellRange("B10").setText("Here is a dataValidation example.");

//get the second sheet
Worksheet sheet2 = workbook.getWorksheets().get(1);

//enable the data can be from different sheet.
sheet2.getParentWorkbook().setAllow3DRangesInDataValidation(true);
sheet1.getCellRange("B11").getDataValidation().setDataRange(sheet2.getCellRange("A1:A7"));
```

---

# Excel Time Data Validation
## Set time validation rules for a cell in Excel
```java
//Create a workbook
Workbook workbook = new Workbook();
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

sheet.getCellRange("C12").setText("Please enter time between 09:00 and 18:00:");
sheet.getCellRange("C12").autoFitColumns();

//Set Time data validation for cell "D12"
CellRange range = sheet.getCellRange("D12");
range.getDataValidation().setAllowType(CellDataType.Time);
range.getDataValidation().setCompareOperator(ValidationComparisonOperator.Between);

range.getDataValidation().setFormula1("09:00");
range.getDataValidation().setFormula2("18:00");

range.getDataValidation().setAlertStyle(AlertStyleType.Info);
range.getDataValidation().setShowError(true);
range.getDataValidation().setErrorTitle("Time Error");
range.getDataValidation().setErrorMessage("Please enter a valid time");
range.getDataValidation().setInputMessage("Time Validation Type");
range.getDataValidation().setIgnoreBlank(true);
range.getDataValidation().setShowInput(true);
```

---

# Excel Data Validation Verification
## Verify cell values against Excel data validation rules
```java
//create a workbook
Workbook workbook = new Workbook();

//get first worksheet of the workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

//cell B4 has the Decimal Validation
CellRange cell = worksheet.getRange().get("B4");

//get the validation of this cell
Validation validation = cell.getDataValidation();

//get the specified data range
Double minimum = Double.parseDouble(validation.getFormula1());
Double maximum = Double.parseDouble(validation.getFormula2());

//verify if a value is within the validation range
if (cell.getNumberValue() < minimum || cell.getNumberValue() > maximum)
{
    //value is not valid
}
else
{
    //value is valid
}
```

---

# Excel Whole Number Data Validation
## Implement whole number data validation in Excel cells using Spire.XLS for Java
```java
//get cellRange
sheet.getCellRange("C12").setText("Please enter number between 10 and 100:");
sheet.getCellRange("C12").autoFitColumns();

//set Whole Number data validation for cell "D12"
CellRange range = sheet.getCellRange("D12");
range.getDataValidation().setAllowType(CellDataType.Integer);
range.getDataValidation().setCompareOperator(ValidationComparisonOperator.Between);
range.getDataValidation().setFormula1("10");
range.getDataValidation().setFormula1("100");
range.getDataValidation().setAlertStyle(AlertStyleType.Info);
range.getDataValidation().setShowError(true);
range.getDataValidation().setErrorTitle("Error");
range.getDataValidation().setErrorMessage("Please enter a valid number");
range.getDataValidation().setInputMessage("Whole Number Validation Type");
range.getDataValidation().setIgnoreBlank(true);
range.getDataValidation().setShowInput(true);
```

---

# Add Data Table to Chart
## This code demonstrates how to add a data table to a chart in an Excel file using Spire.XLS for Java
```java
//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the first chart
Chart chart = sheet.getCharts().get(0);

//add data table
chart.hasDataTable(true);
```

---

# Adding Picture to Excel Chart
## This code demonstrates how to add a picture to a chart in an Excel worksheet
```java
// Get the first sheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the chart from the worksheet
Chart chart = sheet.getCharts().get(0);

// Add the picture to the chart
chart.getShapes().addPicture("data/SpireXls.png");
```

---

# Add TextBox to Excel Chart
## This code demonstrates how to add a textbox to a chart in an Excel worksheet.
```java
//get the first chart (assuming workbook and sheet exist)
Chart chart = sheet.getCharts().get(0);

//add a textbox
ITextBoxLinkShape textbox = chart.getShapes().addTextBox();
textbox.setWidth(1200);
textbox.setHeight(320);
textbox.setLeft(1000);
textbox.setTop(480);
textbox.setText("This is a textbox");
```

---

# Add Trendline to Excel Charts
## Demonstrates how to add different types of trendlines to Excel charts
```java
Worksheet sheet = workbook.getWorksheets().get(0);
//select chart and set logarithmic trendline
Chart chart = sheet.getCharts().get(0);
chart.setChartTitle("Logarithmic Trendline");
chart.getSeries().get(0).getTrendLines().add(TrendLineType.Logarithmic);
//select chart and set moving_average trendline
Chart chart1 = sheet.getCharts().get(1);
chart1.setChartTitle("Moving Average Trendline");
chart1.getSeries().get(0).getTrendLines().add(TrendLineType.Moving_Average);
//select chart and set linear trendline
Chart chart2 = sheet.getCharts().get(2);
chart2.setChartTitle("Linear Trendline");
chart2.getSeries().get(0).getTrendLines().add(TrendLineType.Linear);
//select chart and set exponential trendline
Chart chart3 = sheet.getCharts().get(3);
chart3.setChartTitle("Exponential Trendline");
chart3.getSeries().get(0).getTrendLines().add(TrendLineType.Exponential);
```

---

# Adjust Bar Space in Excel Chart
## Adjust the gap width and overlap between bars in an Excel chart
```java
//get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.getWorksheets().get(0);
Chart chart = ws.getCharts().get(0);

//adjust the space between bars
for(int i = 0;i<chart.getSeries().getCount(); i++)
{
    ChartSerie cs=  chart.getSeries().get(i);
    cs.getFormat().getOptions().setGapWidth(200);
    cs.getFormat().getOptions().setOverlap(0);
}
```

---

# Apply Soft Edges Effect to Excel Chart
## This code demonstrates how to apply a soft edges effect to a chart in an Excel worksheet using Spire.XLS for Java.
```java
//get the chart
Chart chart = sheet.getCharts().get(0);

//specify the size of the soft edge. Value can be set from 0 to 100
chart.getChartArea().getShadow().setSoftEdge(25);
```

---

# Change Chart Size and Position in Excel
## This code demonstrates how to modify the size and position of a chart in an Excel worksheet
```java
//get the chart
Chart chart = sheet.getCharts().get(0);

//change chart size
chart.setWidth( 600);
chart.setHeight( 500);

//change chart position
chart.setLeftColumn(3);
chart.setTopRow(7);
```

---

# Spire.XLS Chart Data Label Modification
## Change data label in Excel chart
```java
//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the chart
Chart chart = sheet.getCharts().get(0);

//change data label of the first data point of the first series
chart.getSeries().get(0).getDataPoints().get(0).getDataLabels().setText("changed data label");
```

---

# Change Chart Data Range
## This code demonstrates how to change the data range of an existing chart in an Excel worksheet
```java
//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get chart
Chart chart = sheet.getCharts().get(0);

//change data range
chart.setDataRange(sheet.getCellRange("A1:C4"));
```

---

# Excel Chart Gridlines Modification
## Change the color of major gridlines in an Excel chart
```java
//create a workbook
Workbook workbook = new Workbook();

//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the chart
Chart chart = sheet.getCharts().get(0);

//change the color of major gridlines
chart.getPrimaryValueAxis().getMajorGridLines().getLineProperties().setColor(Color.RED);
```

---

# Change Excel Chart Series Color
## This code demonstrates how to change the color of a chart series in an Excel file using Spire.XLS for Java
```java
//get the first chart
Chart chart = sheet.getCharts().get(0);

//get the second series
ChartSerie cs = chart.getSeries().get(1);

//set the fill type
cs.getFormat().getFill().setFillType(ShapeFillType.SolidColor);

//change the fill color
cs.getFormat().getFill().setForeColor(Color.orange);
```

---

# Chart Axis Title Configuration
## Set titles and font size for chart axes
```java
//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the chart
Chart chart = sheet.getCharts().get(0);

//set axis title
chart.getPrimaryCategoryAxis().setTitle("Category Axis");
chart.getPrimaryValueAxis().setTitle("Value axis");

//set font size
chart.getPrimaryCategoryAxis().getFont().setSize(12);
chart.getPrimaryValueAxis().getFont().setSize(12);
```

---

# Chart to Image Conversion
## Convert Excel chart to image format
```java
//Load Excel file
Workbook workbook = new Workbook();
workbook.loadFromFile(input);

//Save chart as image
BufferedImage image = workbook.saveChartAsImage(workbook.getWorksheets().get(0), 0);
ImageIO.write(image, "png", new File(output));
```

---

# Spire.XLS Doughnut Chart Creation
## Create a doughnut chart with percentage labels and legend positioning
```java
//add a new chart, set chart type as doughnut
Chart chart = sheet.getCharts().add();
chart.setChartType( ExcelChartType.Doughnut);
chart.setDataRange(sheet.getCellRange("A1:B5"));
chart.setSeriesDataFromRange(false);

//set position of chart
chart.setLeftColumn(4);
chart.setTopRow(2);
chart.setRightColumn(12);
chart.setBottomRow(22);

//chart title
chart.setChartTitle("Market share by country");
chart.getChartTitleArea().isBold(true );
chart.getChartTitleArea().setSize(12);

for( int i =0; i<chart.getSeries().getCount();i++)
{
    ChartSerie cs = chart.getSeries().get(i);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasPercentage(true);
}
chart.getLegend().setPosition( LegendPositionType.Top);
```

---

# Create Multi-Level Chart in Excel
## This code demonstrates how to create a multi-level bar chart in Excel using Spire.XLS library
```java
//add a clustered bar chart to worksheet
Chart chart = sheet.getCharts().add(ExcelChartType.BarClustered);
chart.setChartTitle("Value");
chart.getPlotArea().getFill().setFillType(ShapeFillType.NoFill);
chart.getLegend().delete();
chart.setLeftColumn(5);
chart.setTopRow(1);
chart.setRightColumn(14);

//set the data source of series data
chart.setDataRange(sheet.getCellRange("C2:C9"));
chart.setSeriesDataFromRange(false);

//set the data source of category labels
ChartSerie serie = chart.getSeries().get(0);
serie.setCategoryLabels(sheet.getCellRange("A2:B9"));

//show multi-level category labels
chart.getPrimaryCategoryAxis().setMultiLevelLable(true);
```

---

# Create Pivot Chart in Excel
## Create a clustered column chart based on a pivot table
```java
//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
//get the first pivot table in the worksheet
IPivotTable pivotTable = sheet.getPivotTables().get(0);

//create a clustered column chart based on the pivot table
Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered, pivotTable);
//set chart position
chart.setTopRow(12);
chart.setLeftColumn(1);
chart.setRightColumn(8);
chart.setBottomRow(30);
//set chart title
chart.setChartTitle("Pivot Chart");
```

---

# Create Excel Radar Chart
## Core code for creating a radar chart in Excel using Spire.XLS library
```java
//Add a new chart worksheet to workbook
Chart chart = sheet.getCharts().add();

//Set position of chart
chart.setLeftColumn(1);
chart.setTopRow(6);
chart.setRightColumn(11);
chart.setBottomRow(29);

//Set region of chart data
chart.setDataRange(sheet.getCellRange("A1:C5"));
chart.setSeriesDataFromRange(false);
chart.setChartType(ExcelChartType.Radar);

//Chart title
chart.setChartTitle("Sale market by region");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(12);

chart.getPlotArea().getFill().setVisible(false);

chart.getLegend().setPosition(LegendPositionType.Corner);
```

---

# Customize Chart Data Markers
## Create a scatter chart with customized data markers including colors, size, style and transparency
```java
//Create a Scatter-Markers chart based on the sample data
Chart chart = sheet.getCharts().add(ExcelChartType.ScatterMarkers);

//Set region of chart data
chart.setDataRange(sheet.getCellRange("A1:B7"));
chart.setSeriesDataFromRange(false);
chart.getPlotArea().setVisible(false);

//Set position of chart
chart.setLeftColumn(4);
chart.setTopRow(5);
chart.setRightColumn(11);
chart.setBottomRow(22);

chart.setChartTitle("Chart with Markers");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(10);

//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
ChartSerie cs1 = chart.getSeries().get(0);
cs1.getDataFormat().setMarkerBackgroundColor(Color.blue);
cs1.getDataFormat().setMarkerForegroundColor(Color.orange);
cs1.getDataFormat().setMarkerSize(7);
cs1.getDataFormat().setMarkerStyle(ChartMarkerType.PlusSign);
cs1.getDataFormat().setMarkerTransparencyValue(0.8);

ChartSerie cs2 = chart.getSeries().get(1);
cs2.getDataFormat().setMarkerBackgroundColor(Color.pink);
cs2.getDataFormat().setMarkerSize(9);
cs2.getDataFormat().setMarkerStyle(ChartMarkerType.Triangle);
cs2.getDataFormat().setMarkerTransparencyValue(0.9);
```

---

# Excel Chart Data Callout Configuration
## Configure data callouts for chart series
```java
//get the first chart
Chart chart = sheet.getCharts().get(0);
ChartSeries series = chart.getSeries();
for (int i = 0; i < series.size(); i++) {
    ChartSerie cs = series.get(i);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasWedgeCallout(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasCategoryName(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasSeriesName(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasLegendKey(true);
}
```

---

# spire.xls chart legend manipulation
## delete specific legend entries from a chart
```java
//get the chart
Chart chart = sheet.getCharts().get(0);

//delete the first and the second legend entries from the chart
chart.getLegend().getLegendEntries().get(0).delete();
chart.getLegend().getLegendEntries().get(1).delete();
```

---

# Excel Chart with Discontinuous Data
## Create a column chart using non-continuous data ranges
```java
//add a chart
Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
chart.setSeriesDataFromRange(false);

//set the position of chart
chart.setLeftColumn(1);
chart.setTopRow(10);
chart.setRightColumn(10);
chart.setBottomRow(24);

//add a series
ChartSerie cs1 = (ChartSerie)chart.getSeries().add();

//set the name of the cs1
cs1.setName(sheet.getCellRange("B1").getValue());

//set discontinuous values for cs1
cs1.setCategoryLabels(sheet.getCellRange("A2:A3").addCombinedRange(sheet.getCellRange("A5:A6"))
        .addCombinedRange(sheet.getCellRange("A8:A9")));
cs1.setValues(sheet.getCellRange("B2:B3").addCombinedRange(sheet.getCellRange("B5:B6"))
        .addCombinedRange(sheet.getCellRange("B8:B9")));

//set the chart type
cs1.setSerieType(ExcelChartType.ColumnClustered);

//add a series
ChartSerie cs2 = (ChartSerie)chart.getSeries().add();
cs2.setName(sheet.getCellRange("C1").getValue());
cs2.setCategoryLabels(sheet.getCellRange("A2:A3").addCombinedRange(sheet.getCellRange("A5:A6"))
        .addCombinedRange(sheet.getCellRange("A8:A9")));
cs2.setValues(sheet.getCellRange("C2:C3").addCombinedRange(sheet.getCellRange("C5:C6"))
        .addCombinedRange(sheet.getCellRange("C8:C9")));
cs2.setSerieType(ExcelChartType.ColumnClustered);

chart.setChartTitle("Chart");
chart.getChartTitleArea().getFont().setSize(20);
chart.getChartTitleArea().setColor(Color.black);

chart.getPrimaryValueAxis().hasMajorGridLines(false);
```

---

# Spire.XLS Line Chart Editing
## Add a new series to an existing line chart
```java
//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the line chart
Chart chart = sheet.getCharts().get(0);

//add a new series
ChartSerie cs = chart.getSeries().add("Added");

//set the values for the series
cs.setValues(sheet.getCellRange("I1:L1"));
```

---

# Embed Non-Installed Fonts in Excel Chart
## Apply custom fonts to chart elements that are not installed on the system
```java
//Load the font file from disk
workbook.setCustomFontFilePaths(new String[]{ "data/PT_Serif-Caption-Web-Regular.ttf"});
Hashtable result  = workbook.getCustomFontParsedResult();
ArrayList valueList = new ArrayList(result.values());

//Apply the font for PrimaryValueAxis of chart
chart.getPrimaryValueAxis().getFont().setFontName(valueList.get(0).toString());

//Apply the font for PrimaryCategoryAxis of chart
chart.getPrimaryCategoryAxis().getFont().setFontName(valueList.get(0).toString());

//Apply the font for the first chartSerie of chart
ChartSerie chartSerie1 = chart.getSeries().get(0);
chartSerie1.getDataPoints().getDefaultDataPoint().getDataLabels().setFontName(valueList.get(0).toString());
```

---

# Exploded Doughnut Chart Creation
## Creating an exploded doughnut chart in Excel using Spire.XLS for Java
```java
//add a chart
Chart chart = sheet.getCharts().add();
chart.setChartType(ExcelChartType.DoughnutExploded);

//set position of chart
chart.setLeftColumn(1);
chart.setTopRow(6);
chart.setRightColumn(11);
chart.setBottomRow(29);

//set region of chart data
chart.setDataRange(sheet.getCellRange("A1:B5"));
chart.setSeriesDataFromRange(false);

//chart title
chart.setChartTitle("Sales market by country");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(12);

ChartSeries series = chart.getSeries();
for (int i = 0;i < series.size();i++ )
{
    ChartSerie cs = series.get(i);
    cs.getFormat().getOptions().isVaryColor(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
}
chart.getPlotArea().getFill().setVisible(false);
chart.getLegend().setPosition(LegendPositionType.Top);
```

---

# Extract Chart Trendline Formula
## Extract the equation of a trendline from an Excel chart
```java
//get the chart from the first worksheet
Chart chart = workbook.getWorksheets().get(0).getCharts().get(0);

//get the trendline of the chart and then extract the equation of the trendline
IChartTrendLine trendLine = chart.getSeries().get(1).getTrendLines().get(0);
String formula = trendLine.getFormula();
```

---

# Fill Chart Elements with Picture
## Demonstrates how to fill chart elements with images in Excel using Spire.XLS for Java
```java
// Get the first chart
Chart chart = ws.getCharts().get(0);

// Fill chart area with image
BufferedImage image = ImageIO.read(new File("image.png"));
chart.getChartArea().getFill().customPicture(image,"None");
chart.getPlotArea().getFill().setTransparency(0.9);

// Fill plot area with image
//chart.getPlotArea().getFill().customPicture(image,"None");
```

---

# Chart Marker Filling
## Demonstrates three different ways to fill chart markers in Excel charts
```java
//Fill chart marker with custom picture
chart.getSeries().get(0).getFormat().getLineProperties().setColor(Color.yellow);
chart.getSeries().get(0).getFormat().setMarkerStyle(ChartMarkerType.Picture);
IShapeFill markerFill = chart.getSeries().get(0).getDataFormat().getMarkerFill();
markerFill.customPicture("path/to/image.png");

//Fill chart marker with texture
IShapeFill markerFill2 = chart.getSeries().get(1).getDataFormat().getMarkerFill();
chart.getSeries().get(1).getFormat().getLineProperties().setColor(Color.red);
markerFill2.setTexture(GradientTextureType.Granite);

//Fill chart marker with pattern
chart.getSeries().get(2).getFormat().getLineProperties().setColor(Color.BLUE); //Line color of the series
IShapeFill markerFill3 = chart.getSeries().get(2).getDataFormat().getMarkerFill();
markerFill3.setPattern(GradientPatternType.Pat10Percent);
markerFill3.setForeColor(Color.lightGray);
markerFill3.setBackColor(Color.ORANGE);
```

---

# Excel Chart Axis Formatting
## Format axis properties and data points in Excel chart
```java
//add a chart
Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
chart.setDataRange(sheet.getCellRange("B1:B9"));
chart.setSeriesDataFromRange(false);
chart.getPlotArea().setVisible(false);
chart.setTopRow(10);
chart.setBottomRow(28);
chart.setLeftColumn(2);
chart.setRightColumn(10);
chart.setChartTitle("Chart with Customized Axis");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(12);
ChartSerie cs1 = chart.getSeries().get(0);
cs1.setCategoryLabels(sheet.getCellRange("A2:A9"));

//format axis
chart.getPrimaryValueAxis().setMajorUnit(8);
chart.getPrimaryValueAxis().setMinorUnit(2);
chart.getPrimaryValueAxis().setMaxValue(50);
chart.getPrimaryValueAxis().setMinValue(0);
chart.getPrimaryValueAxis().isReverseOrder(false);
chart.getPrimaryValueAxis().setMajorTickMark(TickMarkType.TickMarkOutside);
chart.getPrimaryValueAxis().setMinorTickMark(TickMarkType.TickMarkInside);
chart.getPrimaryValueAxis().setTickLabelPosition(TickLabelPositionType.TickLabelPositionNextToAxis);
chart.getPrimaryValueAxis().setCrossesAt(0);

//set number format
chart.getPrimaryValueAxis().setNumberFormat("$#,##0");
chart.getPrimaryValueAxis().isSourceLinked(false);
ChartSerie serie = chart.getSeries().get(0);
ChartDataPointsCollection dataPoints = serie.getDataPoints();
Iterator<ChartDataPoint> it = dataPoints.iterator();
while (it.hasNext()) {
    ChartDataPoint dataPoint = it.next();
    //format series
    dataPoint.getDataFormat().getFill().setFillType(ShapeFillType.SolidColor);
    dataPoint.getDataFormat().getFill().setForeColor(Color.lightGray);

    //set transparency
    dataPoint.getDataFormat().getFill().setTransparency(.3);
}
```

---

# Excel Gauge Chart Creation
## Creating a gauge chart using Spire.XLS for Java
```java
//Create a Workbook
Workbook workbook = new Workbook();

//Get the first sheet and set its name
Worksheet sheet = workbook.getWorksheets().get(0);
sheet.setName("Gauge Chart");

//Add a Doughnut chart
Chart chart = sheet.getCharts().add(ExcelChartType.Doughnut);
chart.setDataRange(sheet.getCellRange("A1:A5"));
chart.setSeriesDataFromRange(false);
chart.hasLegend(true);

//Set the position of chart
chart.setLeftColumn(2);
chart.setTopRow(7);
chart.setRightColumn(9);
chart.setBottomRow(25);

//Get the series 1
ChartSerie cs1 = (ChartSerie)chart.getSeries().get("Value");
cs1.getFormat().getOptions().setDoughnutHoleSize(60);
cs1.getDataFormat().getOptions().setFirstSliceAngle(270);

//Set the fill color
cs1.getDataPoints().get(0).getDataFormat().getFill().setForeColor(Color.yellow);
cs1.getDataPoints().get(1).getDataFormat().getFill().setForeColor(Color.pink);
cs1.getDataPoints().get(2).getDataFormat().getFill().setForeColor(Color.orange);
cs1.getDataPoints().get(3).getDataFormat().getFill().setVisible(false);

//Add a series with pie chart
ChartSerie cs2 = (ChartSerie)chart.getSeries().add("Pointer", ExcelChartType.Pie);

//Set the value
cs2.setValues(sheet.getCellRange("D2:D4"));
cs2.setUsePrimaryAxis(false);
cs2.getDataPoints().get(0).getDataLabels().hasValue(true);
cs2.getDataFormat().getOptions().setFirstSliceAngle(270);
cs2.getDataPoints().get(0).getDataFormat().getFill().setVisible(false);
cs2.getDataPoints().get(1).getDataLabels().hasValue(true);
cs2.getDataPoints().get(1).getDataFormat().getFill().setFillType(ShapeFillType.SolidColor);
cs2.getDataPoints().get(1).getDataFormat().getFill().setForeColor(Color.black);
cs2.getDataPoints().get(2).getDataFormat().getFill().setVisible(false);
```

---

# Get Chart Category Labels
## Extract category labels from a chart in an Excel file
```java
// create a workbook
Workbook workbook = new Workbook();

// get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

// get the chart
Chart chart = sheet.getCharts().get(0);

// get the cell range of the category labels
CellRange cr = chart.getPrimaryCategoryAxis().getCategoryLabels();

StringBuilder sb = new StringBuilder();
for(int i = 0; i < cr.getCount(); i++)
{
    CellRange cell = cr.getCellList().get(i);
    sb.append(cell.getValue() + "\r\n");
}
```

---

# Get Excel Chart Data Point Values
## Retrieve values from data points in an Excel chart
```java
//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the chart
Chart chart = sheet.getCharts().get(0);

//get the first series of the chart
ChartSerie cs = chart.getSeries().get(0);

for(int i = 0; i < cs.getValues().getCount(); i++)
{
    CellRange cell = cs.getValues().getCellList().get(i);
    //get the data point range address
    String rangeAddress = cell.getRangeAddress();
    //get the data point value
    Object value = cell.getValue();
}
```

---

# get worksheet of chart
## retrieve the worksheet that contains a specific chart in Excel
```java
//create a workbook
Workbook workbook = new Workbook();

//access first worksheet of the workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

//access the first chart inside this worksheet
Chart chart = worksheet.getCharts().get(0);

//get its worksheet
Worksheet wSheet = (Worksheet)chart.getSheet();
```

---

# Spire.XLS Chart Category Label Management
## Hide category labels in an Excel chart
```java
//Create a new instance of Workbook
Workbook workbook = new Workbook();

//Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Get the first chart
Chart chart = sheet.getCharts().get(0);

//Get all category labels
String[] labels = chart.getCategoryLabels();

//Hide the first category label
chart.hideCategoryLabels(new String[]{labels[0]});
```

---

# Hide Chart Major Gridlines
## This code demonstrates how to hide major gridlines in a chart using Spire.XLS for Java.
```java
//get the first sheet
Worksheet sheet = workbook.getWorksheets().get(0);

//get the chart
Chart chart = sheet.getCharts().get(0);

//hide major gridlines
chart.getPrimaryValueAxis().hasMajorGridLines(false);
```

---

# Spire.XLS Line Chart Creation
## Create and configure a line chart in Excel using Spire.XLS library
```java
//Add a chart
Chart chart = sheet.getCharts().add();
chart.setChartType(ExcelChartType.Line);

//Set region of chart data
chart.setDataRange(sheet.getCellRange("A1:E5"));

//Set position of chart
chart.setLeftColumn(1);
chart.setTopRow(6);
chart.setRightColumn(11);
chart.setBottomRow(29);

//Set chart title
chart.setChartTitle("Sales market by country");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(12);

chart.getPrimaryCategoryAxis().setTitle("Month");
chart.getPrimaryCategoryAxis().getFont().isBold(true);
chart.getPrimaryCategoryAxis().getTitleArea().isBold(true);

chart.getPrimaryValueAxis().setTitle("Sales(in Dollars)");
chart.getPrimaryValueAxis().hasMajorGridLines(false);
chart.getPrimaryValueAxis().getTitleArea().setTextRotationAngle(90);
chart.getPrimaryValueAxis().setMinValue(1000);
chart.getPrimaryValueAxis().getTitleArea().isBold(true);

for(ChartSerie cs : (Iterable<ChartSerie>) chart.getSeries())
{
    cs.getFormat().getOptions().isVaryColor(true);
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
    cs.getDataFormat().setMarkerStyle(ChartMarkerType.Circle);
}

chart.getPlotArea().getFill().setVisible(false);

chart.getLegend().setPosition(LegendPositionType.Top);
```

---

# Excel Pie Chart Creation
## Create regular and 3D pie charts using Spire.XLS for Java
```java
public static void pie(){
    executePie(false,"output/Pie.xlsx");
}
public static void pie3D(){
    executePie(true,"output/Pie3D.xlsx");
}
private static void executePie(boolean is3D,String output)
{
    //create a Workbook
    Workbook workbook = new Workbook();

    //get the first sheet and set its name
    Worksheet sheet = workbook.getWorksheets().get(0);
    sheet.setName("Pie Chart");

    //add a chart
    Chart chart = null;
    if (is3D)
    {
        chart = sheet.getCharts().add(ExcelChartType.Pie3D);
    }
    else
    {
        chart = sheet.getCharts().add(ExcelChartType.Pie);
    }
    //set chart data
    createChartData(sheet);

    //set region of chart data
    chart.setDataRange(sheet.getCellRange("B2:B5"));
    chart.setSeriesDataFromRange(false);

    //set position of chart
    chart.setLeftColumn(1);
    chart.setTopRow(6);
    chart.setRightColumn(9);
    chart.setBottomRow(25);

    //chart title
    chart.setChartTitle("Sales by year");
    chart.getChartTitleArea().isBold(true);
    chart.getChartTitleArea().setSize(12);

    ChartSerie cs = chart.getSeries().get(0);
    cs.setCategoryLabels(sheet.getCellRange("A2:A5"));
    cs.setValues(sheet.getCellRange("B2:B5"));
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);

    chart.getPlotArea().getFill().setVisible(false);
}
```

---

# Spire.XLS Pyramid Column Chart
## Create pyramid column charts (2D and 3D) with custom styling
```java
public static void pyramidColumn(){
    executePyramidColumn(false,"output/pyramidColumn.xlsx");
}

public static void pyramidColumn3D(){
    executePyramidColumn(true,"output/pyramidColumn3D.xlsx");
}

private static void executePyramidColumn(boolean is3D,String output) {
    //create a Workbook
    Workbook workbook = new Workbook();

    //get the first sheet and set its name
    Worksheet sheet = workbook.getWorksheets().get(0);
    sheet.setName("Chart");

    //set chart data
    createChartData(sheet);

    //add a chart
    Chart chart = sheet.getCharts().add();

    //set region of chart data
    chart.setDataRange(sheet.getCellRange("B2:B5"));
    chart.setSeriesDataFromRange(false);

    //set position of chart
    chart.setLeftColumn(1);
    chart.setTopRow(6);
    chart.setRightColumn(11);
    chart.setBottomRow(29);

    if (is3D)
    {
        chart.setChartType(ExcelChartType.Pyramid3DClustered);
    }
    else
    {
        chart.setChartType(ExcelChartType.PyramidClustered);
    }

    //chart title
    chart.setChartTitle("Sales by year");
    chart.getChartTitleArea().isBold(true);
    chart.getChartTitleArea().setSize(12);

    chart.getPrimaryCategoryAxis().setTitle("Year");
    chart.getPrimaryCategoryAxis().getFont().isBold(true);
    chart.getPrimaryCategoryAxis().getTitleArea().isBold(true);

    chart.getPrimaryValueAxis().setTitle("Sales(in Dollars)");
    chart.getPrimaryValueAxis().hasMajorGridLines(false);
    chart.getPrimaryValueAxis().setMinValue(1000);
    chart.getPrimaryValueAxis().getTitleArea().isBold(true);
    chart.getPrimaryValueAxis().getTitleArea().setTextRotationAngle(90);

    ChartSerie cs = chart.getSeries().get(0);
    cs.setCategoryLabels(sheet.getCellRange("A2:A5"));
    cs.getFormat().getOptions().isVaryColor(true);

    chart.getLegend().setPosition(LegendPositionType.Top);
}
```

---

# spire.xls chart removal
## remove chart from Excel worksheet
```java
//get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//get the first chart from the first worksheet
IChartShape chart = sheet.getCharts().get(0);

//remove the chart
chart.remove();
```

---

# Spire.XLS Rich Text for Data Labels
## Set rich text formatting for chart data labels
```java
//Get the first chart inside this worksheet
Chart chart = sheet.getCharts().get(0);

//Get the first datalabel of the first series
ChartDataLabels datalabel = chart.getSeries().get(0).getDataPoints().get(0).getDataLabels();

//Set the text
datalabel.setText("Rich Text Label");

//Show the value
chart.getSeries().get(0).getDataPoints().get(0).getDataLabels().hasValue(true);

//Set styles for the text
chart.getSeries().get(0).getDataPoints().get(0).getDataLabels().setColor(Color.RED);
chart.getSeries().get(0).getDataPoints().get(0).getDataLabels().isBold(true);
```

---

# Spire.XLS 3D Chart Rotation
## Rotate a 3D chart by setting X and Y rotation angles
```java
//get the chart from the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);
Chart chart = sheet.getCharts().get(0);

//X rotation:
chart.setRotation(30);
//Y rotation:
chart.setElevation(20);
```

---

# Excel Chart Data Labels Formatting
## Set and format data labels in an Excel chart
```java
// Create a line chart with markers
Chart chart = sheet.getCharts().add(ExcelChartType.LineMarkers);
chart.setDataRange(sheet.getCellRange("B1:B7"));
chart.getPlotArea().setVisible(false);
chart.setSeriesDataFromRange(false);
chart.setTopRow(5);
chart.setBottomRow(26);
chart.setLeftColumn(2);
chart.setRightColumn(11);
chart.setChartTitle("Data Labels Demo");
chart.getChartTitleArea().isBold(true);
chart.getChartTitleArea().setSize(12);

// Set up series and category labels
ChartSerie cs1 = chart.getSeries().get(0);
cs1.setCategoryLabels(sheet.getCellRange("A2:A7"));

// Configure data labels properties
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().hasLegendKey(false);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().hasPercentage(false);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().hasSeriesName(true);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().hasCategoryName(true);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().setDelimiter(". ");

// Format data labels appearance
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().setSize(9);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().setColor(Color.RED);
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().setFontName("Calibri");
cs1.getDataPoints().getDefaultDataPoint().getDataLabels().setPosition(DataLabelPositionType.Center);
```

---

# Excel Chart Border Customization
## Set border color and style for Excel chart
```java
//create a workbook
Workbook workbook = new Workbook();

//get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.getWorksheets().get(0);
Chart chart = ws.getCharts().get(0);

//set CustomLineWeight property for Series line
(chart.getSeries().get(0).getDataPoints().get(0).getDataFormat().getLineProperties()).setCustomLineWeight(2.5f);

//set color property for Series line
(chart.getSeries().get(0).getDataPoints().get(0).getDataFormat().getLineProperties()).setColor(Color.RED);
```

---

# Spire.XLS Chart Category Text Direction
## Set vertical text direction for chart category axis
```java
//get the first chart
Chart chart = sheet.getCharts().get(0);

//set Category text direction
chart.getPrimaryCategoryAxis().setTextDirection(TextVerticalValue.Vertical);
```

---

# Chart Background Color Setting
## Set the background color of a chart in Excel
```java
//get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.getWorksheets().get(0);
Chart chart = ws.getCharts().get(0);

//set background color
chart.getChartArea().setForeGroundColor(Color.YELLOW);
```

---

# spire.xls chart font
## Set font for chart data labels
```java
//create a font
ExcelFont font = workbook.createFont();
font.setSize(15.0);
font.setColor(Color.lightGray);
for (ChartSerie cs : (Iterable<ChartSerie>)chart.getSeries())
{
    //set font
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().getTextArea().setFont(font);
}
```

---

# Set Font for Chart Legend and Data Table
## This code demonstrates how to set custom font properties for chart legend and data labels in an Excel chart
```java
//create a font with specified size and color
ExcelFont font = workbook.createFont();
font.setSize(14);
font.setColor(Color.RED);

//apply the font to chart Legend
chart.getLegend().getTextArea().setFont(font);

//apply the font to chart DataLabel
for (ChartSerie cs : (Iterable<ChartSerie>)chart.getSeries())
{
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().getTextArea().setFont(font);
}
```

---

# spire.xls font formatting
## set font for chart title and axis
```java
//Set font for chart title and chart axis
Worksheet worksheet = workbook.getWorksheets().get(0);
Chart chart = worksheet.getCharts().get(0);

//Format the font for the chart title
chart.getChartTitleArea().setColor(Color.blue);
chart.getChartTitleArea().setSize(20.0);
chart.getChartTitleArea().setFontName("Arial");

//Format the font for the chart Axis
chart.getPrimaryValueAxis().getFont().setColor(Color.orange);
chart.getPrimaryValueAxis().getFont().setSize(10.0);

chart.getPrimaryCategoryAxis().getFont().setFontName("Arial");
chart.getPrimaryCategoryAxis().getFont().setColor(Color.red);
chart.getPrimaryCategoryAxis().getFont().setSize(20.0);
```

---

# Set Chart Legend Background Color
## Change the background color of a chart legend in Excel using Spire.XLS for Java
```java
Worksheet ws = workbook.getWorksheets().get(0);
Chart chart = ws.getCharts().get(0);

XlsChartFrameFormat x = (XlsChartFrameFormat)chart.getLegend().getFrameFormat();
x.getFill().setFillType(ShapeFillType.SolidColor);
x.setForeGroundColor(Color.LIGHT_GRAY);
```

---

# Set Marker Color for Sparkline
## This code demonstrates how to set the marker color for sparklines in an Excel worksheet.
```java
public class setMarkerColorForSparkline {
    public static void main(String[] args) {
        Workbook book = new Workbook();
        Worksheet worksheet = book.getWorksheets().get(0);
        SparklineGroupCollection sparklineGroup = worksheet.getSparklineGroups();
        sparklineGroup.get(1).setMarkersColor(Color.RED);
    }
}
```

---

# Excel Chart Leader Lines
## Show leader lines for data labels in a stacked bar chart
```java
// Create a stacked bar chart
Chart chart = sheet.getCharts().add(ExcelChartType.BarStacked);
chart.setDataRange(sheet.getCellRange("A1:C3"));
chart.setTopRow(4);
chart.setLeftColumn(2);
chart.setWidth(450);
chart.setHeight(300);

// Configure leader lines for data labels
for (ChartSerie cs : (Iterable<ChartSerie>) chart.getSeries()) {
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
    // Show the leader lines
    cs.getDataPoints().getDefaultDataPoint().getDataLabels().setShowLeaderLines(true);
}
```

---

# Excel Sparkline Creation
## Add line sparklines to worksheet cells
```java
//add sparkline
SparklineGroup sparklineGroup = sheet.getSparklineGroups().addGroup(SparklineType.Line);
SparklineCollection sparklines = sparklineGroup.add();
sparklines.add(sheet.getCellRange("A2:D2"), sheet.getCellRange("E2"));
sparklines.add(sheet.getCellRange("A3:D3"), sheet.getCellRange("E3"));
sparklines.add(sheet.getCellRange("A4:D4"), sheet.getCellRange("E4"));
sparklines.add(sheet.getCellRange("A5:D5"), sheet.getCellRange("E5"));
sparklines.add(sheet.getCellRange("A6:D6"), sheet.getCellRange("E6"));
sparklines.add(sheet.getCellRange("A7:D7"), sheet.getCellRange("E7"));
sparklines.add(sheet.getCellRange("A8:D8"), sheet.getCellRange("E8"));
sparklines.add(sheet.getCellRange("A9:D9"), sheet.getCellRange("E9"));
sparklines.add(sheet.getCellRange("A10:D10"), sheet.getCellRange("E10"));
sparklines.add(sheet.getCellRange("A11:D11"), sheet.getCellRange("E11"));
```

---

# Adding Arrow Lines to Excel
## Demonstrates how to add various types of arrow lines to an Excel worksheet using Spire.XLS
```java
//create a workbook.
Workbook workbook = new Workbook();

//get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//add a Double Arrow and fill the line with solid color.
XlsLineShape line = (XlsLineShape)sheet.getTypedLines().addLine();
line.setTop(10);
line.setLeft( 20);
line.setWidth(100);
line.setHeight(0);
line.setColor( Color.BLUE);
line.setBeginArrowHeadStyle( ShapeArrowStyleType.LineArrow);
line.setEndArrowHeadStyle(ShapeArrowStyleType.LineArrow);

//add an Arrow and fill the line with solid color.
XlsLineShape line_1 = (XlsLineShape)sheet.getTypedLines().addLine();
line_1.setTop(50);
line_1.setLeft(30);
line_1.setWidth(100);
line_1.setHeight(100);
line_1.setColor( Color.RED);
line_1.setBeginArrowHeadStyle(ShapeArrowStyleType.LineNoArrow);
line_1.setEndArrowHeadStyle(ShapeArrowStyleType.LineArrow);

//add an Elbow Arrow Connector.
XlsLineShape line3 = (XlsLineShape)sheet.getTypedLines().addLine();
line3.setLineShapeType( LineShapeType.ElbowLine);
line3.setWidth(30);
line3.setHeight(50);
line3.setEndArrowHeadStyle( ShapeArrowStyleType.LineArrow);
line3.setTop(100);
line3.setLeft(50);

//add an Elbow Double-Arrow Connector.
XlsLineShape line2 = (XlsLineShape)sheet.getTypedLines().addLine();
line2.setLineShapeType(LineShapeType.ElbowLine);
line2.setWidth(50);
line2.setHeight(50);
line2.setEndArrowHeadStyle( ShapeArrowStyleType.LineArrow);
line2.setBeginArrowHeadStyle( ShapeArrowStyleType.LineArrow);
line2.setLeft(120);
line2.setTop( 100);

//add a Curved Arrow Connector.
line3 = (XlsLineShape)sheet.getTypedLines().addLine();
line3.setLineShapeType(LineShapeType.CurveLine);
line3.setWidth(30);
line3.setHeight(50);
line3.setEndArrowHeadStyle( ShapeArrowStyleType.LineArrowOpen);
line3.setTop(100);
line3.setLeft(200);

//add a Curved Double-Arrow Connector.
line2 = (XlsLineShape)sheet.getTypedLines().addLine();
line2.setLineShapeType(LineShapeType.CurveLine);
line2.setWidth(30);
line2.setHeight(50);
line2.setEndArrowHeadStyle( ShapeArrowStyleType.LineArrowOpen);
line2.setBeginArrowHeadStyle(ShapeArrowStyleType.LineArrowOpen);
line2.setLeft( 250);
line2.setTop(100);
```

---

# Adding Line Shapes to Excel Worksheet
## This code demonstrates how to add different types of line shapes with various styles to an Excel worksheet
```java
//Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//Add shape line1
ILineShape line1 = sheet.getLines().addLine(10, 2, 200, 1, LineShapeType.Line);
//Set dash style type
line1.setDashStyle(ShapeDashLineStyleType.Solid);
//Set color
line1.setColor(Color.BLUE);
//Set weight
line1.setWeight(2);
//Set end arrow style type
line1.setEndArrowHeadStyle(ShapeArrowStyleType.LineArrow);

//Add shape line2
ILineShape line2 = sheet.getLines().addLine(12, 2, 200, 1, LineShapeType.CurveLine);
line2.setDashStyle(ShapeDashLineStyleType.Dotted);
line2.setColor(Color.ORANGE);
line2.setWeight(2);

//Add shape line3
ILineShape line3 = sheet.getLines().addLine(14, 2, 200, 1, LineShapeType.ElbowLine);
line3.setDashStyle(ShapeDashLineStyleType.DashDotDot);
line3.setColor(Color.PINK);
line3.setWeight(2);

//Add shape line4
ILineShape line4 = sheet.getLines().addLine(16, 2, 200, 1, LineShapeType.LineInv);
line4.setDashStyle(ShapeDashLineStyleType.Dashed);
line4.setColor(Color.green);
line4.setWeight(2);
line4.setBeginArrowHeadStyle(ShapeArrowStyleType.LineArrow);
line4.setEndArrowHeadStyle(ShapeArrowStyleType.LineArrow);
```

---

# Excel Oval Shape Creation
## Add oval shapes with different fill styles to Excel worksheet
```java
//Add oval shape1
IOvalShape ovalShape1 = sheet.getOvalShapes().addOval(11, 2, 100, 100);
ovalShape1.getLine().setWeight(0);
//Fill shape with solid color
ovalShape1.getFill().setFillType(ShapeFillType.SolidColor);
ovalShape1.getFill().setForeColor(Color.orange);

//Add oval shape2
IOvalShape ovalShape2 = sheet.getOvalShapes().addOval(11, 5, 100, 100);
ovalShape2.getLine().setDashStyle(ShapeDashLineStyleType.Solid);
ovalShape2.getLine().setWeight(1);
//Fill shape with picture
ovalShape2.getFill().customPicture("data/E-iceblueLogo.png");
```

---

# Copy shapes between worksheets
## Demonstrates how to create and copy various shapes between Excel worksheets
```java
//create a workbook
Workbook workbook = new Workbook();

//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//create line shape
ILineShape line = sheet.getTypedLines().addLine();
line.setTop(50);
line.setLeft(30);
line.setWidth(30);
line.setHeight(50);
line.setBeginArrowHeadStyle(ShapeArrowStyleType.LineArrowDiamond);
line.setEndArrowHeadStyle(ShapeArrowStyleType.LineArrow);

//get the second worksheet
Worksheet CopyShapes = workbook.getWorksheets().get(1);

//copy the line into the second worksheet
CopyShapes.getTypedLines().addCopy(line);

//create a button and then copy into other sheet
IRadioButton button = sheet.getTypedRadioButtons().add(5, 5, 20, 20);
CopyShapes.getTypedRadioButtons().addCopy(button);

//create a textbox and then copy into other sheet
ITextBoxLinkShape textbox = sheet.getTypedTextBoxes().addTextBox(5, 7, 50, 100);
CopyShapes.getTypedTextBoxes().addCopy(textbox);

//create a checkbox and then copy into other sheet
ICheckBox checkbox = sheet.getTypedCheckBoxes().addCheckBox(10, 1, 20, 20);
CopyShapes.getTypedCheckBoxes().addCopy(checkbox);

//create a comboboxes and then copy into other sheet
IComboBoxShape ComboBoxes = sheet.getTypedComboBoxes().addComboBox(10, 5, 30, 30);
ComboBoxes.setListFillRange(sheet.getCellRange("A14:A15"));
CopyShapes.getTypedComboBoxes().addCopy(ComboBoxes);
```

---

# Delete All Shapes in Excel Worksheet
## This code demonstrates how to delete all shapes from a worksheet in an Excel file
```java
//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//delete all shapes in the worksheet
for (int i = sheet.getPrstGeomShapes().getCount()-1; i >= 0; i--)
{
    sheet.getPrstGeomShapes().get(i).remove();
}
```

---

# Delete Particular Shape in Excel
## Remove a specific shape from worksheet
```java
//create a workbook
Workbook workbook = new Workbook();

//get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

//delete the first shape in the worksheet
sheet.getPrstGeomShapes().get(0).remove();
```

---

# Excel Shape Text and Image Extraction
## Extract text and images from shapes in Excel worksheets
```java
//get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//extract text from the first shape and save to a txt file.
IPrstGeomShape shape1 = sheet.getPrstGeomShapes().get(2);
String s = shape1.getText();
StringBuilder sb = new StringBuilder();
sb.append("The text in the third shape is: " + s);

//extract image from the second shape and save to a local folder.
IPrstGeomShape shape2 = sheet.getPrstGeomShapes().get(1);
Image image = shape2.getFill().getPicture();
```

---

# Excel Shape Grouping
## Group shapes in Excel worksheet
```java
//add shapes
IPrstGeomShape shape1 = worksheet.getPrstGeomShapes().addPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect);
IPrstGeomShape shape2 = worksheet.getPrstGeomShapes().addPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle);

//group
GroupShapeCollection groupShapeCollection = worksheet.getGroupShapes();
groupShapeCollection.group(new com.spire.xls.core.IShape[]{shape1,shape2});
```

---

# Spire.XLS Shape Visibility Control
## Hide or unhide shapes in Excel worksheet
```java
//create a workbook.
Workbook workbook = new Workbook();

//get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//hide the second shape in the worksheet
sheet.getPrstGeomShapes().get(1).setVisible(false);

//show the second shape in the worksheet
//sheet.getPrstGeomShapes().get(1).setVisible(true);
```

---

# Insert Shapes to Excel Sheet
## Add various geometric shapes with different fill styles to an Excel worksheet
```java
//create a workbook.
Workbook workbook = new Workbook();

//get the first worksheet.
Worksheet sheet = workbook.getWorksheets().get(0);

//add a triangle shape.
IPrstGeomShape triangle = sheet.getPrstGeomShapes().addPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle);

//fill the triangle with solid color.
triangle.getFill().setForeColor( Color.YELLOW);
triangle.getFill().setFillType( ShapeFillType.SolidColor);

//add a heart shape.
IPrstGeomShape heart = sheet.getPrstGeomShapes().addPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart);

//fill the heart with gradient color.
heart.getFill().setForeColor(Color.RED);
heart.getFill().setFillType(ShapeFillType.Gradient);

//add an arrow shape with default color.
IPrstGeomShape arrow = sheet.getPrstGeomShapes().addPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow);

//add a cloud shape.
IPrstGeomShape cloud = sheet.getPrstGeomShapes().addPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud);

//fill the cloud with custom picture
cloud.getFill().customPicture(image, "SpireXls.png");
cloud.getFill().setFillType( ShapeFillType.Picture);
```

---

# Shape Text Centering
## Center text within a shape in Excel worksheet
```java
//Add a rectangle shape
IPrstGeomShape rect = sheet.getPrstGeomShapes().addPrstGeomShape(11,2,300,300, PrstGeomShapeType.Rect);

rect.setText("E-iceblue");
//Middle centered the text of IPrstGeomShape
rect.setTextVerticalAlignment(ExcelVerticalAlignment.MiddleCentered);
```

---

# Spire.XLS Shape Shadow Style Modification
## Modify shadow style properties for Excel shapes
```java
//get the third shape from the worksheet.
IPrstGeomShape shape = sheet.getPrstGeomShapes().get(2);

//set the shadow style for the shape.
shape.getShadow().setAngle(90);
shape.getShadow().setTransparency(30);
shape.getShadow().setDistance(10);
shape.getShadow().setSize(130);
shape.getShadow().setColor(Color.YELLOW);
shape.getShadow().setBlur(30);
shape.getShadow().hasCustomStyle(true);
```

---

# Set Shadow Style for Shape
## Configure shadow properties for Excel shapes using Spire.XLS for Java
```java
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Add an ellipse shape
IPrstGeomShape ellipse = sheet.getPrstGeomShapes().addPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse);

// Set the shadow style for the ellipse
ellipse.getShadow().setAngle(90);
ellipse.getShadow().setDistance(10);
ellipse.getShadow().setSize(150);
ellipse.getShadow().setColor(Color.GRAY);
ellipse.getShadow().setBlur(30);
ellipse.getShadow().setTransparency(1);
ellipse.getShadow().hasCustomStyle(true);
```

---

# Excel Shape to Image Conversion
## Convert shapes from Excel worksheet to image files
```java
// Convert shapes to images
SaveShapeTypeOption shapelist = new SaveShapeTypeOption();

// Save all shapes in the worksheet to images
shapelist.setSaveAll(true);

// Save the shapes in the worksheet to images and get a HashMap of shapes and their corresponding images
HashMap<IShape, BufferedImage> images = sheet.saveAndGetShapesToImage(shapelist);

// Iterate through each shape in the HashMap
for (IShape shape : images.keySet()) {
    // Get the image corresponding to the current shape
    BufferedImage image = images.get(shape);

    // Generate a filename based on the shape's properties
    String fileName = shape.getName() + "_" + shape.getHeight() + "_" + shape.getWidth() + "_" + shape.getShapeType() + ".png";

    // Save the image to a file
    ImageIO.write(image, "PNG", new File("testImage/" + fileName));
}
```

---

# Excel Shape Texture Fill
## Tiling picture as texture in Excel shape
```java
//Get the first shape
IPrstGeomShape shape = sheet.getPrstGeomShapes().get(0);

//Fill shape with texture
shape.getFill().setFillType(ShapeFillType.Texture);

//Custom texture with picture
shape.getFill().customTexture("data/logo.png");

//Tile picture as texture
shape.getFill().setTile(true);
```

---

# Apply Built-in Styles to Excel
## Apply built-in style to a cell range in Excel worksheet
```java
//Set the built-in style "Title" for the range A1:J1 in the worksheet
sheet.getCellRange("A1:J1").setBuiltInStyle(BuiltInStyles.Title);
```

---

# Excel Color Scale Conditional Formatting
## Apply color scales to data range in Excel using conditional formatting
```java
//Create a new instance of Workbook
Workbook workbook = new Workbook();

//Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Add conditional formatting to the worksheet
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(sheet.getAllocatedRange());
IConditionalFormat format = xcfs.addCondition();
format.setFormatType(ConditionalFormatType.ColorScale);
```

---

# Excel Conditional Formatting
## Apply conditional formatting rules to Excel cells based on cell values
```java
//Add the first conditional formatting rule to the worksheet based on cell values
XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
xcfs1.addRange(sheet.getAllocatedRange());
IConditionalFormat format1 = xcfs1.addCondition();
format1.setFormatType(ConditionalFormatType.CellValue);
format1.setFirstFormula("800");
format1.setOperator(ComparisonOperatorType.Greater);
format1.setFontColor(Color.RED);
format1.setBackColor(Color.LIGHT_GRAY);

//Add the second conditional formatting rule to the worksheet based on cell values
XlsConditionalFormats xcfs2 = sheet.getConditionalFormats().add();
xcfs2.addRange(sheet.getAllocatedRange());
IConditionalFormat format2 = xcfs1.addCondition();
format2.setFormatType(ConditionalFormatType.CellValue);
format2.setFirstFormula("300");
format2.setOperator(ComparisonOperatorType.Less);
format2.setFontColor(Color.GREEN);
format2.setBackColor(Color.BLUE);
```

---

# Excel Data Bars Conditional Formatting
## Apply data bars to cell range in Excel using conditional formatting
```java
//Add conditional formatting (data bars) to the worksheet
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(sheet.getAllocatedRange());
IConditionalFormat format = xcfs.addCondition();
format.setFormatType(ConditionalFormatType.DataBar);
format.getDataBar().setBarColor(Color.BLUE);
```

---

# Spire.XLS Gradient Fill Effects
## Apply gradient fill effects to Excel cells
```java
//Get the CellRange object for cell B5
CellRange range =sheet.getCellRange("B5");

//Set the row height of the range to 50
range.setRowHeight(50);
//Set the column width of the range to 30
range.setColumnWidth(30);
//Set the text in the range to "Hello"
range.setText( "Hello");

//Set the horizontal alignment of the range to center
range.getStyle().setHorizontalAlignment( HorizontalAlignType.Center);

//Set the fill pattern of the range to Gradient
range.getStyle().getInterior().setFillPattern( ExcelPatternType.Gradient);
//Set the fore color of the gradient
range.getStyle().getInterior().getGradient().setForeColor(Color.CYAN);
//Set the back color of the gradient
range.getStyle().getInterior().getGradient().setBackColor( Color.BLUE);
//Apply a two-color horizontal gradient shading effect to the gradient
range.getStyle().getInterior().getGradient().twoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1);
```

---

# Excel icon sets formatting
## Apply conditional formatting with icon sets to cell range
```java
//Add conditional formatting to the worksheet
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(sheet.getAllocatedRange());
//Add a condition for the conditional format
IConditionalFormat format = xcfs.addCondition();
format.setFormatType(ConditionalFormatType.IconSet);
format.getIconSet().setIconSetType(IconSetType.ThreeTrafficLights1);
```

---

# Excel Colors and Palette Manipulation
## Working with color palettes and cell formatting in Excel
```java
//Create a new workbook
Workbook workbook = new Workbook();

//Change the palette color to Orchid at index 60
workbook.changePaletteColor(Color.YELLOW, 60);

//Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);
//Get the cell range B2
CellRange cell = sheet.getCellRange("B2");
//Set the text in the cell
cell.setText("Welcome to use Spire.XLS");

//Set the font color, size, and autofit the columns and rows of the cell
cell.getStyle().getFont().setColor(Color.YELLOW);
cell.getStyle().getFont().setSize(20);
cell.autoFitColumns();
cell.autoFitRows();
```

---

# Excel Conditional Formatting
## Add conditional formatting rules to Excel cells at runtime
```java
private static void addComparisonRule1(Worksheet sheet)
{
    //Add conditional formats to the worksheet for range A1:D1
    XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
    xcfs1.addRange(sheet.getCellRange("A1:D1"));
    //Add a condition for the conditional format
    IConditionalFormat cf1 = xcfs1.addCondition();
    cf1.setFormatType( ConditionalFormatType.CellValue);
    cf1.setFirstFormula("150");
    cf1.setOperator(ComparisonOperatorType.Greater);
    cf1.setFontColor(Color.RED);
    cf1.setBackColor( Color.BLUE);
}
private static void addComparisonRule2(Worksheet sheet)
{
    //Add conditional formats to the worksheet for range A2:D2
    XlsConditionalFormats xcfs2 = sheet.getConditionalFormats().add();
    xcfs2.addRange(sheet.getCellRange("A2:D2"));
    //Add a condition for the conditional format
    IConditionalFormat cf2 = xcfs2.addCondition();
    cf2.setFormatType( ConditionalFormatType.CellValue);
    cf2.setFirstFormula( "500");
    cf2.setOperator( ComparisonOperatorType.Less);
    //Set border color
    cf2.setLeftBorderColor(Color.BLUE);
    cf2.setRightBorderColor( Color.BLUE);
    cf2.setTopBorderColor( Color.GREEN);
    cf2.setBottomBorderColor( Color.GREEN);
    cf2.setLeftBorderStyle( LineStyleType.Medium);
    cf2.setRightBorderStyle( LineStyleType.Thick);
    cf2.setTopBorderStyle(LineStyleType.Double);
    cf2.setBottomBorderStyle(LineStyleType.Double);
}

private static void addComparisonRule3(Worksheet sheet)
{
    //Add conditional formats to the worksheet for range A3:D3
    XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
    xcfs1.addRange(sheet.getCellRange("A3:D3"));
    //Add a condition for the conditional format
    IConditionalFormat cf1 = xcfs1.addCondition();
    cf1.setFormatType( ConditionalFormatType.CellValue);
    cf1.setFirstFormula("300");
    cf1.setSecondFormula( "500");
    cf1.setOperator(ComparisonOperatorType.Between);
    cf1.setBackColor( Color.YELLOW);
}

private static void addComparisonRule4(Worksheet sheet)
{
    //Add conditional formats to the worksheet for range A4:D4
    XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
    xcfs1.addRange(sheet.getCellRange("A4:D4"));
    //Add a condition for the conditional format
    IConditionalFormat cf1 = xcfs1.addCondition();
    cf1.setFormatType( ConditionalFormatType.CellValue);
    cf1.setFirstFormula( "100");
    cf1.setSecondFormula( "200");
    cf1.setOperator(ComparisonOperatorType.NotBetween);
    cf1.setFillPattern( ExcelPatternType.ReverseDiagonalStripe);
    cf1.setColor( Color.LIGHT_GRAY);
    cf1.setBackColor( Color.BLACK);
}
```

---

# Spire.XLS Conditional Date Formatting
## Format dates conditionally based on time period
```java
//Add conditional formats to the worksheet for the allocated range
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(sheet.getAllocatedRange());
//Add a time period condition (last 7 days) to the conditional format
IConditionalFormat conditionalFormat = xcfs.addTimePeriodCondition(TimePeriodType.Last7Days);
conditionalFormat.setBackColor(Color.orange);
```

---

# Excel Formula Conditional Formatting
## Create formula-based conditional formatting for cells in Excel
```java
// Get a range of cells in the first column of the worksheet
CellRange range = sheet.getColumns()[0];

// Add conditional formatting to the worksheet
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(range);

// Add a condition to the conditional formatting
IConditionalFormat conditional = xcfs.addCondition();
conditional.setFormatType(ConditionalFormatType.Formula);
conditional.setFirstFormula("=($A1<$B1)");
conditional.setBackKnownColor(ExcelColors.Yellow);
```

---

# Excel Font Styling
## Apply various font styles to Excel cells
```java
//Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Set the font name for cell B1 to "Comic Sans MS"
sheet.getCellRange("B1").getCellStyle().getExcelFont().setFontName("Comic Sans MS");
//Set the font name for cells B2 to D2 to "Corbel"
sheet.getCellRange("B2:D2").getCellStyle().getExcelFont().setFontName("Corbel");
//Set the font name for cells B3 to D7 to "Aleo"
sheet.getCellRange("B3:D7").getCellStyle().getExcelFont().setFontName("Aleo");

//Set the font size for cell B1 to 45
sheet.getCellRange("B1").getCellStyle().getExcelFont().setSize(45);
//Set the font size for cells B2 to D3 to 25
sheet.getCellRange("B2:D3").getCellStyle().getExcelFont().setSize(25);
//Set the font size for cells B3 to D7 to 12
sheet.getCellRange("B3:D7").getCellStyle().getExcelFont().setSize(12);

//Set the font style of cells B2 to D2 to bold
sheet.getCellRange("B2:D2").getCellStyle().getExcelFont().isBold(true);

//Set the font style of cells B3 to B7 to underline
sheet.getCellRange("B3:B7").getCellStyle().getExcelFont().setUnderline(FontUnderlineType.Single);

//Set the font color of cell B1 to blue
sheet.getCellRange("B1").getCellStyle().getExcelFont().setColor(Color.blue);
//Set the font color of cells B2 to D2 to pink
sheet.getCellRange("B2:D2").getCellStyle().getExcelFont().setColor(Color.pink);
//Set the font color of cells B3 to D7 to darkGray
sheet.getCellRange("B3:D7").getCellStyle().getExcelFont().setColor(Color.darkGray);

//Set the font style of cells B3 to D7 to italic
sheet.getCellRange("B3:D7").getCellStyle().getExcelFont().isItalic(true);
```

---

# Excel Cell Style Formatting
## Set foreground and background colors and patterns for Excel cells
```java
//Add a new cell style named "newStyle1"
CellStyle style = workbook.getStyles().addStyle("newStyle1");

//Set the fill pattern of the interior to vertical stripes
style.getInterior().setFillPattern(ExcelPatternType.VerticalStripe);

//Set the background color of the gradient to green
style.getInterior().getGradient().setBackKnownColor(ExcelColors.Green);

//Set the foreground color of the gradient to yellow
style.getInterior().getGradient().setForeKnownColor(ExcelColors.Yellow);

//Apply the "newStyle1" cell style to cell B2
sheet.getCellRange("B2").setCellStyleName(style.getName());

//Set the text of cell B2 to "Test"
sheet.getCellRange("B2").setText("Test");
//Set the row height of cell B2 to 30
sheet.getCellRange("B2").setRowHeight(30);
//Set the column width of cell B2 to 50
sheet.getCellRange("B2").setColumnWidth(50);


//Add a new cell style named "newStyle2"
style = workbook.getStyles().addStyle("newStyle2");

//Set the fill pattern of the interior to thin horizontal stripes
style.getInterior().setFillPattern(ExcelPatternType.ThinHorizontalStripe);
//Set the foreground color of the gradient to red
style.getInterior().getGradient().setForeKnownColor(ExcelColors.Red);

//Apply the "newStyle2" cell style to cell B4
sheet.getCellRange("B4").setCellStyleName(style.getName());
//Set the row height of cell B4 to 30
sheet.getCellRange("B4").setRowHeight(30);
//Set the column width of cell B4 to 60
sheet.getCellRange("B4").setColumnWidth(60);
```

---

# Excel Column Formatting
## Format a column with custom style including alignment, font color, and borders
```java
//Create a new workbook
Workbook workbook = new Workbook();

//Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Add a new cell style named "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");

//Set the vertical alignment of the style to center
style.setVerticalAlignment(VerticalAlignType.Center);

//Set the horizontal alignment of the style to center
style.setHorizontalAlignment(HorizontalAlignType.Center);

//Set the font color of the style to blue
style.getFont().setColor(Color.BLUE);

//Enable the "shrink to fit" option for the style
style.setShrinkToFit(true);

//Set the bottom border color and style of the style
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setColor(Color.ORANGE);
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Dotted);

//Apply the "newStyle" cell style to all cells in column 0
sheet.getColumns()[0].setCellStyleName(style.getName());
//Set the text of all cells in column 0 to "Test"
sheet.getColumns()[0].setText("Test");
```

---

# Excel Row Formatting
## Format a row in Excel with custom cell style
```java
//Create a new workbook object
Workbook workbook = new Workbook();

//Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

//Add a new cell style named "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");

//Set the vertical alignment of the style to center
style.setVerticalAlignment( VerticalAlignType.Center);

//Set the horizontal alignment of the style to center
style.setHorizontalAlignment(HorizontalAlignType.Center);

//Set the font color of the style to blue
style.getFont().setColor(Color.BLUE);

//Enable the "shrink to fit" option for the style
style.setShrinkToFit(true);

//Set the bottom border color of the style to orange
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setColor(Color.ORANGE);

//Set the line style of the bottom border of the style to dotted
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle( LineStyleType.Dotted);

//Apply the "newStyle" cell style to all cells in row 1
sheet.getRows()[1].setCellStyleName( style.getName());

//Set the text of all cells in row 1 to "Test"
sheet.getRows()[1].setText( "Test");
```

---

# Excel Cell Style Formatting
## Create and apply custom style to Excel cells
```java
//Create a new CellStyle object named "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");

//Set the color of the style to DarkGray
style.setColor(Color.DARK_GRAY);

//Set the font color of the style to White
style.getFont().setColor(Color.WHITE);

//Set the font name of the style to "Times New Roman"
style.getFont().setFontName("Times New Roman");

//Set the font size of the style to 12
style.getFont().setSize(12);

//Make the font bold in the style
style.getFont().isBold(true);

// Set the rotation angle of the style to 45 degrees
style.setRotation(45);

//Set the horizontal alignment of the style to Center
style.setHorizontalAlignment(HorizontalAlignType.Center);
//Set the vertical alignment of the style to Center
style.setVerticalAlignment(VerticalAlignType.Center);

//Apply the style to the range A1:J1 in the first worksheet of the workbook
workbook.getWorksheets().get(0).getCellRange("A1:J1").setCellStyleName(style.getName());
```

---

# Excel Cell Style Manipulation
## Get and set cell style properties in Excel
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the cell range "B4" from the worksheet
CellRange range = sheet.getCellRange("B4");

// Get the cell style of the range
CellStyle style = range.getCellStyle();
// Set the font name to "Calibri"
style.getFont().setFontName("Calibri");
// Make the font bold
style.getFont().isBold(true);
// Set the font size to 15
style.getFont().setSize(15);
// Set the font color to blue
style.getFont().setColor(Color.BLUE);
// Apply the style to the range
range.setStyle(style);
```

---

# Highlight Average Values in Excel
## Apply conditional formatting to highlight cells above and below average values
```java
// Add conditional formats to apply formatting based on conditions

// Create a new XlsConditionalFormats object
XlsConditionalFormats format1 = sheet.getConditionalFormats().add();
// Add the cell range "E2:E10" to the conditional format
format1.addRange(sheet.getCellRange("E2:E10"));
// Create an average condition of type "Below"
IConditionalFormat cf1 = format1.addAverageCondition(AverageType.Below);
// Set the background color to blue for cells that meet the average condition
cf1.setBackColor(Color.BLUE);

// Create another XlsConditionalFormats object
XlsConditionalFormats format2 = sheet.getConditionalFormats().add();
// Add the same cell range "E2:E10" to the second conditional format
format2.addRange(sheet.getCellRange("E2:E10"));
// Create an average condition of type "Above"
IConditionalFormat cf2 = format1.addAverageCondition(AverageType.Above);
// Set the background color to orange for cells that meet the average condition
cf2.setBackColor(Color.ORANGE);
```

---

# Excel Conditional Formatting
## Highlight duplicate and unique values in a cell range
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add conditional formats to apply formatting based on conditions

// Create a new XlsConditionalFormats object
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
// Add the cell range "C2:C10" to the conditional format
xcfs.addRange(sheet.getCellRange("C2:C10"));
// Create a conditional format for duplicate values
IConditionalFormat format1 = xcfs.addCondition();
// Set the format type to DuplicateValues
format1.setFormatType(ConditionalFormatType.DuplicateValues);
// Set the background color to red for cells with duplicate values
format1.setBackColor(Color.RED);

// Create another XlsConditionalFormats object
XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
// Add the same cell range "C2:C10" to the second conditional format
xcfs1.addRange(sheet.getCellRange("C2:C10"));
// Create a conditional format for unique values
IConditionalFormat format2 = xcfs.addCondition();
// Set the format type to UniqueValues
format2.setFormatType(ConditionalFormatType.UniqueValues);
// Set the background color to yellow for cells with unique values
format2.setBackColor(Color.YELLOW);
```

---

# Excel Conditional Formatting for Ranked Values
## Highlight top and bottom ranked values in Excel ranges using conditional formatting
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add conditional formats to the worksheet for range D2:D10
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
xcfs.addRange(sheet.getCellRange("D2:D10"));

// Add a top-bottom conditional format to highlight the top 2 values in range D2:D10
IConditionalFormat format1 = xcfs.addTopBottomCondition(TopBottomType.Top, 2);
format1.setFormatType(ConditionalFormatType.TopBottom);
format1.setBackColor(Color.RED);

// Add conditional formats to the worksheet for range E2:E10
XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
xcfs1.addRange(sheet.getCellRange("E2:E10"));

// Add a top-bottom conditional format to highlight the bottom 2 values in range E2:E10
IConditionalFormat format2 = xcfs1.addTopBottomCondition(TopBottomType.Bottom, 2);
format2.setFormatType(ConditionalFormatType.TopBottom);
format2.setBackColor(Color.GREEN);
```

---

# Excel Cell Indentation
## Set indentation level for Excel cell
```java
//Get the CellRange object representing the cell at B5
CellRange cell = sheet.getCellRange("B5");

//Set the text of the cell to "Hello Spire!"
cell.setText("Hello Spire!");

//Set the indentation level of the cell's style to 2
cell.getStyle().setIndentLevel(2);
```

---

# Excel Cell Interior Formatting
## Apply gradient fill to cell interiors in Excel
```java
// Apply gradient formatting to cells
CellRange range = sheet.getCellRange("E1:K1");
range.getCellStyle().getInterior().setFillPattern(ExcelPatternType.Gradient);
range.getCellStyle().getInterior().getGradient().setBackKnownColor(ExcelColors.LightBlue);
range.getCellStyle().getInterior().getGradient().setForeKnownColor(ExcelColors.White);
range.getCellStyle().getInterior().getGradient().setGradientStyle(GradientStyleType.Vertical);
range.getCellStyle().getInterior().getGradient().setGradientVariant(GradientVariantsType.ShadingVariants1);
```

---

# make cell active in Excel worksheet
## Set active cell and adjust visible area
```java
// Get the second worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(1);

// Activate the worksheet
sheet.activate();

// Set the active cell to cell range "B2"
sheet.setActiveCell(sheet.getCellRange("B2"));

// Set the first visible column to column index 1
sheet.setFirstVisibleColumn(1);

// Set the first visible row to row index 1
sheet.setFirstVisibleRow(1);
```

---

# Excel Number Formatting
## Apply various number formatting styles to Excel cells
```java
// Set text and formatting for cell B10
sheet.getCellRange("B10").setText("NUMBER FORMATTING");
sheet.getCellRange("B10").getCellStyle().getExcelFont().isBold(true);

// Set text, value, and number format for cell B13
sheet.getCellRange("B13").setText("0");
sheet.getCellRange("C13").setNumberValue(1234.5678);
sheet.getCellRange("C13").setNumberFormat("0");

// Set text, value, and number format for cell B14
sheet.getCellRange("B14").setText("0.00");
sheet.getCellRange("C14").setNumberValue(1234.5678);
sheet.getCellRange("C14").setNumberFormat("0.00");

// Set text, value, and number format for cell B15
sheet.getCellRange("B15").setText("#,##0.00");
sheet.getCellRange("C15").setNumberValue(1234.5678);
sheet.getCellRange("C15").setNumberFormat("#,##0.00");

// Set text, value, and number format for cell B16
sheet.getCellRange("B16").setText("$#,##0.00");
sheet.getCellRange("C16").setNumberValue(1234.5678);
sheet.getCellRange("C16").setNumberFormat("$#,##0.00");

// Set text, value, and number format for cell B17
sheet.getCellRange("B17").setText("0;[Red]-0");
sheet.getCellRange("C17").setNumberValue(-1234.5678);
sheet.getCellRange("C17").setNumberFormat("0;[Red]-0");

// Set text, value, and number format for cell B18
sheet.getCellRange("B18").setText("0.00;[Red]-0.00");
sheet.getCellRange("C18").setNumberValue(-1234.5678);
sheet.getCellRange("C18").setNumberFormat("0.00;[Red]-0.00");

// Set text, value, and number format for cell B19
sheet.getCellRange("B19").setText("#,##0;[Red]-#,##0");
sheet.getCellRange("C19").setNumberValue(-1234.5678);
sheet.getCellRange("C19").setNumberFormat("#,##0;[Red]-#,##0");

// Set text, value, and number format for cell B20
sheet.getCellRange("B20").setText("#,##0.00;[Red]-#,##0.000");
sheet.getCellRange("C20").setNumberValue(-1234.5678);
sheet.getCellRange("C20").setNumberFormat("#,##0.00;[Red]-#,##0.00");

// Set text, value, and number format for cell B21
sheet.getCellRange("B21").setText("0.00E+00");
sheet.getCellRange("C21").setNumberValue(1234.5678);
sheet.getCellRange("C21").setNumberFormat("0.00E+00");

// Set text, value, and number format for cell B22
sheet.getCellRange("B22").setText("0.00%");
sheet.getCellRange("C22").setNumberValue(1234.5678);
sheet.getCellRange("C22").setNumberFormat("0.00%");

// Apply known color Gray25Percent to cells B13:B22
sheet.getCellRange("B13:B22").getCellStyle().setKnownColor(ExcelColors.Gray25Percent);

// Auto-fit the width of columns 2 and 3
sheet.autoFitColumn(2);
sheet.autoFitColumn(3);
```

---

# Excel Border Formatting
## Set cell borders in Excel worksheet
```java
//Define a CellRange object that includes all cells in the worksheet
CellRange cr = sheet.getCellRange(sheet.getFirstRow(), sheet.getFirstColumn(), sheet.getLastRow(), sheet.getLastColumn());

//Set the border style of the CellRange to Double line
cr.getBorders().setLineStyle(LineStyleType.Double);
//Set the diagonal down border style of the CellRange to None
cr.getBorders().getByBordersLineType(BordersLineType.DiagonalDown).setLineStyle(LineStyleType.None);
//Set the diagonal up border style of the CellRange to None
cr.getBorders().getByBordersLineType(BordersLineType.DiagonalUp).setLineStyle(LineStyleType.None);
//Set the border color of the CellRange to Blue
cr.getBorders().setColor(Color.BLUE);
```

---

# Excel Conditional Formatting with Formulas
## Demonstrates how to set conditional formatting based on formulas in Excel using Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Create a new XlsConditionalFormats object
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
// Add the cell range "B5" to the conditional format
xcfs.addRange(sheet.getCellRange("B5"));
// Create a conditional format based on cell value
IConditionalFormat format = xcfs.addCondition();
// Set the format type to CellValue
format.setFormatType(ConditionalFormatType.CellValue);
// Set the first formula to "1000"
format.setFirstFormula("1000");
// Set the comparison operator to Greater
format.setOperator(ComparisonOperatorType.Greater);
// Set the background color to orange for cells meeting the condition
format.setBackColor(Color.orange);
```

---

# Excel Conditional Formatting
## Set row colors based on even/odd row numbers
```java
// Get the range of cells with data in the worksheet
CellRange dataRange = sheet.getAllocatedRange();

// Add conditional formats to apply formatting based on conditions

// Create a new XlsConditionalFormats object
XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
// Add the data range to the conditional format
xcfs.addRange(dataRange);
// Create a conditional format for even rows
IConditionalFormat format1 = xcfs.addCondition();
// Set the first formula to check if the row number is even using MOD function
format1.setFirstFormula("=MOD(ROW(),2)=0");
// Set the format type to Formula
format1.setFormatType(ConditionalFormatType.Formula);
// Set the background color to light gray for even rows
format1.setBackColor(Color.lightGray);

// Create another XlsConditionalFormats object
XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
// Add the same data range to the second conditional format
xcfs1.addRange(dataRange);
// Create a conditional format for odd rows
IConditionalFormat format2 = xcfs.addCondition();
// Set the first formula to check if the row number is odd using MOD function
format2.setFirstFormula("=MOD(ROW(),2)=1");
// Set the format type to Formula
format2.setFormatType(ConditionalFormatType.Formula);
// Set the background color to yellow for odd rows
format2.setBackColor(Color.yellow);
```

---

# Spire.XLS Traffic Light Icons
## Set conditional formatting with traffic light icons in Excel cells
```java
// Add conditional formatting to the allocated range of cells
XlsConditionalFormats conditional = sheet.getConditionalFormats().add();
conditional.addRange(sheet.getAllocatedRange());

// Add condition for applying traffic lights icons
IConditionalFormat format = conditional.addCondition();
format.setFormatType(ConditionalFormatType.IconSet);
format.getIconSet().setIconSetType(IconSetType.ThreeTrafficLights1);
```

---

# Excel Conditional Formatting
## Apply various conditional formatting rules to Excel cells
```java
private static void addConditionalFormattingForExistingSheet(Worksheet sheet)
{
    // Set row height and column width for the entire allocated range
    sheet.getAllocatedRange().setRowHeight(15);
    sheet.getAllocatedRange().setColumnWidth(16);

    // Add conditional formatting for range A1:D1
    XlsConditionalFormats xcfs1 = sheet.getConditionalFormats().add();
    xcfs1.addRange(sheet.getCellRange("A1:D1"));
    IConditionalFormat cf1 = xcfs1.addCondition();
    cf1.setFormatType(ConditionalFormatType.CellValue);
    cf1.setFirstFormula("150");
    cf1.setOperator(ComparisonOperatorType.Greater);
    cf1.setFontColor(Color.red);
    cf1.setBackColor(Color.pink);

    // Add conditional formatting for range A2:D2
    XlsConditionalFormats xcfs2 = sheet.getConditionalFormats().add();
    xcfs2.addRange(sheet.getCellRange("A2:D2"));
    IConditionalFormat cf2 = xcfs2.addCondition();
    cf2.setFormatType(ConditionalFormatType.CellValue);
    cf2.setFirstFormula("300");
    cf2.setOperator(ComparisonOperatorType.Less);
    cf2.setLeftBorderColor(Color.pink);
    cf2.setRightBorderColor(Color.pink);
    cf2.setTopBorderColor(Color.blue);
    cf2.setBottomBorderColor(Color.blue);
    cf2.setLeftBorderStyle(LineStyleType.Medium);
    cf2.setRightBorderStyle(LineStyleType.Thick);
    cf2.setTopBorderStyle(LineStyleType.Double);
    cf2.setBottomBorderStyle(LineStyleType.Double);

    // Add conditional formatting for range A3:D3
    XlsConditionalFormats xcfs3 = sheet.getConditionalFormats().add();
    xcfs3.addRange(sheet.getCellRange("A3:D3"));
    IConditionalFormat cf3 = xcfs3.addCondition();
    cf3.setFormatType(ConditionalFormatType.DataBar);
    cf3.getDataBar().setBarColor(Color.yellow);

    // Add conditional formatting for range A4:D4
    XlsConditionalFormats xcfs4 = sheet.getConditionalFormats().add();
    xcfs4.addRange(sheet.getCellRange("A4:D4"));
    IConditionalFormat cf4 = xcfs4.addCondition();
    cf4.setFormatType(ConditionalFormatType.IconSet);
    cf4.getIconSet().setIconSetType(IconSetType.ThreeTrafficLights1);

    // Add conditional formatting for range A5:D5
    XlsConditionalFormats xcfs5 = sheet.getConditionalFormats().add();
    xcfs5.addRange(sheet.getCellRange("A5:D5"));
    IConditionalFormat cf5 = xcfs5.addCondition();
    cf5.setFormatType(ConditionalFormatType.ColorScale);

    // Add conditional formatting for range A6:D6
    XlsConditionalFormats xcfs6 = sheet.getConditionalFormats().add();
    xcfs6.addRange(sheet.getCellRange("A6:D6"));
    IConditionalFormat cf6 = xcfs6.addCondition();
    cf6.setFormatType(ConditionalFormatType.DuplicateValues);
    cf6.setBackColor(Color.orange);
}
```

---

# Excel Cell Text Alignment
## Set vertical and horizontal alignment, rotation angle, and row height for Excel cells
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set vertical alignment of range B1:C1 to Top
sheet.getCellRange("B1:C1").getCellStyle().setVerticalAlignment(VerticalAlignType.Top);

// Set vertical alignment of range B2:C2 to Center
sheet.getCellRange("B2:C2").getCellStyle().setVerticalAlignment(VerticalAlignType.Center);

// Set vertical alignment of range B3:C3 to Bottom
sheet.getCellRange("B3:C3").getCellStyle().setVerticalAlignment(VerticalAlignType.Bottom);

// Set horizontal alignment of range B4:C4 to General
sheet.getCellRange("B4:C4").getCellStyle().setHorizontalAlignment(HorizontalAlignType.General);

// Set horizontal alignment of range B5:C5 to Left
sheet.getCellRange("B5:C5").getCellStyle().setHorizontalAlignment(HorizontalAlignType.Left);

// Set horizontal alignment of range B6:C6 to Center
sheet.getCellRange("B6:C6").getCellStyle().setHorizontalAlignment(HorizontalAlignType.Center);

// Set horizontal alignment of range B7:C7 to Right
sheet.getCellRange("B7:C7").getCellStyle().setHorizontalAlignment(HorizontalAlignType.Right);

// Set rotation angle of range B8:C8 to 45 degrees
sheet.getCellRange("B8:C8").getCellStyle().setRotation(45);
// Set rotation angle of range B9:C9 to 90 degrees
sheet.getCellRange("B9:C9").getCellStyle().setRotation(90);

// Set row height of range B8:C9 to 60
sheet.getCellRange("B8:C9").setRowHeight(60);
```

---

# Spire.XLS Text Direction
## Set text reading order in Excel cells
```java
// Create a workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the cell range
CellRange cell = sheet.getCellRange("B5");

// Set the reading order of the cell to right-to-left
cell.getStyle().setReadingOrder(ReadingOrderType.RightToLeft);
```

---

# Excel Cell Style Application
## Apply predefined styles to Excel cells
```java
// Create a new cell style named "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");
style.getFont().setFontName("Calibri");
style.getFont().isBold(true);
style.getFont().setSize(15);
style.getFont().setColor(Color.blue);

// Get the cell range B5
CellRange range = sheet.getCellRange("B5");
// Set the text of the cell and apply the "newStyle" to it
range.setText("Welcome to use Spire.XLS");
range.setCellStyleName(style.getName());

// Auto-fit the columns in the range
range.autoFitColumns();
```

---

# Excel Style Object Usage
## Create and apply cell styles in Excel using Spire.XLS
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Add a new worksheet with the name "new sheet"
Worksheet sheet = workbook.getWorksheets().add("new sheet");

// Create a new CellStyle object and give it a name "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");

// Set the vertical alignment of the style to Center
style.setVerticalAlignment(VerticalAlignType.Center);

// Set the horizontal alignment of the style to Center
style.setHorizontalAlignment(HorizontalAlignType.Center);

// Set the font color of the style to blue
style.getFont().setColor(Color.blue);

// Enable shrink to fit for the style
style.setShrinkToFit(true);

// Set the bottom border color of the style to yellow
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setColor(Color.yellow);

// Set the bottom border's line style to Medium
style.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Medium);

// Apply the style to cell ranges
sheet.getCellRange("B1").setStyle(style);
sheet.getCellRange("B4").setStyle(style);
sheet.getCellRange("C3").setCellStyleName(style.getName());
sheet.getCellRange("D4").setStyle(style);
```

---

# Excel Conditional Formatting Implementation
## Various types of conditional formatting using Spire.XLS library
```java
public class variousConditionalFormatting {

    public static void main(String[] args) {
        // Create a new Workbook object
        Workbook workbook = new Workbook();
        // Create an empty sheet in the workbook
        workbook.createEmptySheets(1);
        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);
        // Add conditional formatting to the new sheet
        AddConditionalFormattingForNewSheet(sheet);
    }

    private static void AddConditionalFormattingForNewSheet(Worksheet sheet) {
        AddDefaultIconSet(sheet);
        AddIconSet2(sheet);
        AddIconSet3(sheet);
        AddDefaultColorScale(sheet);
        Add3ColorScale(sheet);
        Add2ColorScale(sheet);
        AddEboveEverage(sheet);
        AddTop10_1(sheet);
        AddDataBar1(sheet);
        AddContainsText(sheet);
        AddTimePeriod_1(sheet);
    }

    // This method implements the TimePeriod conditional formatting type with Today attribute.
    private static void AddTimePeriod_1(Worksheet sheet) {
        XlsConditionalFormats conds = sheet.getConditionalFormats().add();
        conds.addRange(sheet.getCellRange("I1:K2"));
        sheet.getCellRange("I1:K2").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("I1:K2").getStyle().setColor(Color.gray);
        IConditionalFormat cf = conds.addTimePeriodCondition(TimePeriodType.Today);
        cf.setFillPattern(ExcelPatternType.Solid);
        cf.setBackColor(Color.pink);
    }

    // This method implements the ContainsText conditional formatting type.
    private static void AddContainsText(Worksheet sheet) {
        XlsConditionalFormats conds = sheet.getConditionalFormats().add();
        conds.addRange(sheet.getCellRange("E5:G6"));
        sheet.getCellRange("E5:G6").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("E5:G6").getStyle().setColor(Color.blue);
        IConditionalFormat cf = conds.addContainsTextCondition("abc");
        cf.setFillPattern(ExcelPatternType.Solid);
        cf.setBackColor(Color.yellow);
    }

    // This method implements the DataBars conditional formatting type.
    private static void AddDataBar1(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("E1:G2"));
        sheet.getCellRange("E1:G2").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("E1:G2").getStyle().setColor(Color.green);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.DataBar);
        cf.getDataBar().setBarColor(Color.blue);
        cf.getDataBar().getMinPoint().setType(ConditionValueType.Percent);
        cf.getDataBar().setShowValue(true);
    }

    // This method implements a Top10 conditional formatting type.
    private static void AddTop10_1(Worksheet sheet) {
        XlsConditionalFormats conds = sheet.getConditionalFormats().add();
        conds.addRange(sheet.getCellRange("A17:C20"));
        sheet.getCellRange("A17:C20").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A17:C20").getStyle().setColor(Color.gray);
        IConditionalFormat cf = conds.addTopBottomCondition(TopBottomType.Top, 10);
        cf.setFillPattern(ExcelPatternType.Solid);
        cf.setBackColor(Color.yellow);
    }

    // This method implements the AboveAverage conditional formatting type.
    private static void AddEboveEverage(Worksheet sheet) {
        XlsConditionalFormats conds = sheet.getConditionalFormats().add();
        conds.addRange(sheet.getCellRange("A11:C12"));
        sheet.getCellRange("A11:C12").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A11:C12").getStyle().setColor(Color.red);
        IConditionalFormat cf = conds.addAverageCondition(AverageType.Above);
        cf.setFillPattern(ExcelPatternType.Solid);
        cf.setBackColor(Color.pink);
    }

    // This method implements the ColorScale conditional formatting type with some color scale attributes.
    private static void Add2ColorScale(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("A9:C10"));
        sheet.getCellRange("A9:C10").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A9:C10").getStyle().setColor(Color.white);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.ColorScale);
        cf.getColorScale().setMinColor(Color.yellow);
        cf.getColorScale().setMaxColor(Color.blue);
    }

    // This method implements the ColorScale conditional formatting type with some color scale attributes.
    private static void Add3ColorScale(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("A7:C8"));
        sheet.getCellRange("A7:C8").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A7:C8").getStyle().setColor(Color.green);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.ColorScale);
        cf.getColorScale().getMinValue().setType(ConditionValueType.Number);
        cf.getColorScale().getMinValue().setValue(9);
        cf.getColorScale().setMinColor(Color.pink);
    }

    // This method implements the ColorScale conditional formatting type.
    private static void AddDefaultColorScale(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("A5:C6"));
        sheet.getCellRange("A5:C6").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A5:C6").getStyle().setColor(Color.pink);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.ColorScale);
    }

    // This method implements the IconSet conditional formatting type with ThreeArrows colored attribute.
    private static void AddIconSet2(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("M1:O2"));
        sheet.getCellRange("M1:O2").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("M1:O2").getStyle().setColor(Color.blue);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.IconSet);
        cf.getIconSet().setIconSetType(IconSetType.ThreeArrows);
    }

    // This method implements the IconSet conditional formatting type with FourArrows colored attribute.
    private static void AddIconSet3(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("M3:O4"));
        sheet.getCellRange("M3:O4").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("M3:O4").getStyle().setColor(Color.white);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.IconSet);
        cf.getIconSet().setIconSetType(IconSetType.FourArrows);
    }

    // This method implements the IconSet conditional formatting type.
    private static void AddDefaultIconSet(Worksheet sheet) {
        XlsConditionalFormats xcfs = sheet.getConditionalFormats().add();
        xcfs.addRange(sheet.getCellRange("A1:C2"));
        sheet.getCellRange("A1:C2").getStyle().setFillPattern(ExcelPatternType.Solid);
        sheet.getCellRange("A1:C2").getStyle().setColor(Color.yellow);
        IConditionalFormat cf = xcfs.addCondition();
        cf.setFormatType(ConditionalFormatType.IconSet);
    }
}
```

---

# Excel Formula Calculation
## Calculate formula values in Excel using Spire.XLS for Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();
// Load the workbook from the input stream
workbook.loadFromStream(inputStream);

// Calculate the value of formula "Sheet1!$B$3"
Object b3 = workbook.calculateFormulaValue("Sheet1!$B$3");
// Calculate the value of formula "Sheet1!$C$3"
Object c3 = workbook.calculateFormulaValue("Sheet1!$C$3");
// Create a formula string "Sheet1!$B$3 + Sheet1!$C$3"
String formula = "Sheet1!$B$3 + Sheet1!$C$3";
// Calculate the value of the formula
Object value = workbook.calculateFormulaValue(formula);
```

---

# Excel Named Range Formula
## Insert formula using named range in Excel
```java
// Create a named range object and add it to the workbook's named ranges collection
INamedRange NamedRange = workbook.getNameRanges().add("NewNamedRange");
// Set the local name of the named range to the formula "=SUM(A1+A2)"
NamedRange.setNameLocal("=SUM(A1+A2)");

// Set the formula of cell C1 to the named range "NewNamedRange"
sheet.getCellRange("C1").setFormula("NewNamedRange");
```

---

# Excel Formula Implementation
## Demonstrate how to implement various Excel formulas in Java using Spire.XLS
```java
// Create a new workbook object
Workbook workbook = new Workbook();
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the number format of the first column to text format
sheet.getColumns()[0].setNumberFormat("@");

// Set various Excel formulas in cells
sheet.getCellRange("B1").setFormula("=CEILING.MATH(-2.78, 5, -1)");
sheet.getCellRange("B2").setFormula("=BITOR(23,10)");
sheet.getCellRange("B3").setFormula("=BITAND(23,10)");
sheet.getCellRange("B4").setFormula("=BITLSHIFT(23,2)");
sheet.getCellRange("B5").setFormula("=BITRSHIFT(23,2)");
sheet.getCellRange("B6").setFormula("=FLOOR.MATH(12.758, 2, -1)");
sheet.getCellRange("B7").setFormula("=ISOWEEKNUM(DATE(2012, 1, 1))");
sheet.getCellRange("B8").setFormula("=CEILING.PRECISE(-4.6, 3)");
sheet.getCellRange("B9").setFormula("=ENCODEURL(\"https://www.e-iceblue.com\")");

// Calculate all the formulas in the workbook
workbook.calculateAllValue();

// Auto-fit the columns in the allocated range of the worksheet
sheet.getAllocatedRange().autoFitColumns();
```

---

# Excel Formula Reader
## Read formulas and their calculated values from Excel cells
```java
// Get the formula and its calculated value from a cell
String formula = sheet.getCellRange("C14").getFormula();
double value = sheet.getCellRange("C14").getFormulaNumberValue();
```

---

# Excel Add-in Function Registration
## Register and use custom add-in functions in Excel workbook
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Add custom add-in functions to the workbook
workbook.getAddInFunctions().add("test.xlam", "TEST_UDF");
workbook.getAddInFunctions().add("test.xlam", "TEST_UDF1");

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the formula of cell A1 to use the "TEST_UDF" function
sheet.getCellRange("A1").setFormula("=TEST_UDF()");
// Set the formula of cell A2 to use the "TEST_UDF1" function
sheet.getCellRange("A2").setFormula("=TEST_UDF1()");
```

---

# Remove Formulas But Keep Values
## Extracts formula values from Excel cells and replaces formulas with their calculated values
```java
// Iterate over each worksheet in the workbook
for (Worksheet sheet : (Iterable<Worksheet>) workbook.getWorksheets())
{
    // Iterate over each cell range in the worksheet
    for (CellRange cell : (Iterable<CellRange>) sheet.getRange())
    {
        // Check if the cell contains a formula
        if (cell.hasFormula())
        {
            // Get the value of the formula
            Object value = cell.getFormulaValue();
            
            // Clear the cell's content
            cell.clear(ExcelClearOptions.ClearContent);
            
            // Set the cell's value to the string representation of the formula value
            cell.setValue(value.toString());
        }
    }
}
```

---

# Excel Array Formula Implementation
## Set array formula in Excel worksheet cells
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set an array formula for a range of cells in the worksheet
sheet.getCellRange("A5:C6").setFormulaArray("=LINEST(A1:A3,B1:C3,TRUE,TRUE)");

// Calculate all the formulas and update their values
workbook.calculateAllValue();
```

---

# Excel R1C1 Formula Implementation
## Using R1C1 notation to set array formulas in Excel cells
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set an array formula using R1C1 notation for a specific cell in the worksheet
sheet.getCellRange("C4").setFormulaArrayR1C1("=SUM(R[-3]C[-2]:R[-1]C)");

// Calculate all the formulas and update their values
workbook.calculateAllValue();
```

---

# Spire.XLS R1C1 Formula
## Using R1C1 notation for Excel formulas
```java
// Set formula for cell C4 using R1C1 notation
sheet.getCellRange("C4").setFormulaR1C1("=SUM(R[-3]C[-2]:R[-1]C)");

// Calculate all values in the workbook
workbook.calculateAllValue();
```

---

# Excel Formula Writing with Spire.XLS
## Demonstrates how to write various types of formulas to Excel cells
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set various types of formulas in cells
// String formula
currentFormula = "=\"hello\"";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Numeric formula
currentFormula = "=300";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Boolean formula
currentFormula = "=false";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Arithmetic formula
currentFormula = "=1+2+3+4+5-6-7+8-9";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Cell reference formula
currentFormula = "=Sheet1!$B$3";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Function formula with range
currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);

// Date/Time formula with formatting
currentFormula = "=NOW()";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow, 2).setFormula(currentFormula);
sheet.getCellRange(currentRow, 2).getCellStyle().setNumberFormat("yyyy-MM-DD");

// Mathematical function formulas
currentFormula = "=SQRT(40)";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow++, 2).setFormula(currentFormula);

// Statistical function formulas
currentFormula = "=MAX(10,30)";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow++, 2).setFormula(currentFormula);

// Conditional formula
currentFormula = "=IF(4,2,2)";
sheet.getCellRange(++currentRow, 1).setText("'"+currentFormula);
sheet.getCellRange(currentRow++, 2).setFormula(currentFormula);
```

---

# Excel Header Footer Image
## Add image to first page header and footer
```java
// Set the first page to have different header/footer
sheet.getPageSetup().setDifferentFirst((byte)1);

// Set the image header
sheet.getPageSetup().setFirstLeftHeaderImage(bufferedImage);
sheet.getPageSetup().setFirstCenterHeaderImage(bufferedImage);
sheet.getPageSetup().setFirstRightHeaderImage(bufferedImage);

// Set the image footer
sheet.getPageSetup().setFirstLeftFooterImage(bufferedImage);
sheet.getPageSetup().setFirstCenterFooterImage(bufferedImage);
sheet.getPageSetup().setFirstRightFooterImage(bufferedImage);
```

---

# Excel Watermark Addition
## Adding text watermark to Excel worksheets using header images
```java
// Define the font and watermark text
Font font = new Font("Arial", Font.PLAIN, 40);
String watermark = "Confidential";

// Loop through each worksheet in the workbook
for (Worksheet sheet : (Iterable<Worksheet>) workbook.getWorksheets()) {
    // Draw the watermark image
    BufferedImage imgWtrmrk = drawText(watermark, font, Color.pink, Color.white, 
                                     sheet.getPageSetup().getPageHeight(), 
                                     sheet.getPageSetup().getPageWidth());

    // Set the watermark image as the left header image
    sheet.getPageSetup().setLeftHeaderImage(imgWtrmrk);

    // Set the left header to display the page number
    sheet.getPageSetup().setLeftHeader("&G");

    // The watermark will only appear in this mode, it will disappear if the mode is normal
    sheet.setViewMode(ViewMode.Layout);
}

private static BufferedImage drawText(String text, Font font, Color textColor, 
                                    Color backColor, double height, double width) {
    // Create a new bitmap image with specified width and height
    BufferedImage img = new BufferedImage((int) width, (int) height, TYPE_INT_ARGB);

    // Create a Graphics object from the image
    Graphics2D loGraphic = img.createGraphics();

    // Measure the size of the text using the specified font
    FontMetrics loFontMetrics = loGraphic.getFontMetrics(font);
    int liStrWidth = loFontMetrics.stringWidth(text);
    int liStrHeight = loFontMetrics.getHeight();

    // Set rotation point
    loGraphic.setColor(backColor);
    loGraphic.fillRect(0, 0, (int) width, (int) height);
    loGraphic.translate(((int) width - liStrWidth) / 2, ((int) height - liStrHeight) / 2);

    // Rotate the drawing surface by -45 degrees
    loGraphic.rotate(Math.toRadians(-45));

    // Translate the drawing origin back to its original position
    loGraphic.translate(-((int) width - liStrWidth) / 2, -((int) height - liStrHeight) / 2);

    loGraphic.setFont(font);

    // Create a brush for the text
    loGraphic.setColor(textColor);

    // Draw text on the image at center position
    loGraphic.drawString(text, ((int) width - liStrWidth) / 2, ((int) height - liStrHeight) / 2);

    loGraphic.dispose();
    // Return the resulting image with the watermark
    return img;
}
```

---

# Excel Header Footer Font
## Change font and size for header and footer in Excel
```java
// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

String text = sheet.getPageSetup().getLeftHeader();

//"Arial Unicode MS" is font name, "18" is font size
text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";

// Update the left header text with a custom string and font size
sheet.getPageSetup().setLeftHeader(text);

// Update the right footer text with a custom string and font size
sheet.getPageSetup().setRightFooter(text);
```

---

# Excel Different Header and Footer
## Set different headers and footers for odd and even pages
```java
// Enable different odd and even page headers and footers
sheet.getPageSetup().setDifferentOddEven((byte)1);

// Set the odd page header text format
sheet.getPageSetup().setOddHeaderString( "&\"Arial\"&12&B&KFFC000 Odd_Header");
// Set the odd page footer text format
sheet.getPageSetup().setOddFooterString ( "&\"Arial\"&12&B&KFFC000 Odd_Footer");
// Set the even page header text format
sheet.getPageSetup().setEvenHeaderString ( "&\"Arial\"&12&B&KFF0000 Even_Header");
// Set the even page footer text format
sheet.getPageSetup().setEvenFooterString ( "&\"Arial\"&12&B&KFF0000 Even_Footer");

// Change the view mode of the worksheet to Layout view
sheet.setViewMode(ViewMode.Layout);
```

---

# Excel Header Footer Configuration
## Set different header and footer for the first page
```java
// Enable different header and footer for the first page only
sheet.getPageSetup().setDifferentFirst((byte)1);

// Set the header text for the first page
sheet.getPageSetup().setFirstHeaderString("Different First page");

// Set the footer text for the first page
sheet.getPageSetup().setFirstFooterString("Different First footer");

// Set the left header text for all pages
sheet.getPageSetup().setLeftHeader("Demo of Spire.XLS");
// Set the center footer text for all pages
sheet.getPageSetup().setCenterFooter("Footer by Spire.XLS");
```

---

# Excel Header and Footer with Images
## Add images to header and footer sections in Excel worksheets
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Create an Image object from a specified image file
BufferedImage image = ImageIO.read(new File("data/logo.png"));

// Set the left header image and text for the page setup
sheet.getPageSetup().setLeftHeaderImage(image);
sheet.getPageSetup().setLeftHeader("&G");

// Set the center footer image and text for the page setup
sheet.getPageSetup().setCenterFooterImage(image);
sheet.getPageSetup().setCenterFooter("&G");

// Set the view mode of the sheet to Layout
sheet.setViewMode(ViewMode.Layout);
```

---

# Excel Header and Footer Setting
## Set custom headers and footers in Excel worksheets
```java
// Get the first worksheet from the Workbook
Worksheet Worksheet = workbook.getWorksheets().get(0);

// Set the left header of the page to a specific text with a specific font
Worksheet.getPageSetup().setLeftHeader("&\"Arial Unicode MS\"&14 Spire.XLS for .NET ");

// Set the center footer of the page to a specific text
Worksheet.getPageSetup().setCenterFooter("Footer Text");

// Set the view mode of the worksheet to layout
Worksheet.setViewMode(ViewMode.Layout);
```

---

# Excel Hyperlink Management
## Add hyperlinks to text in Excel cells
```java
// Create a new instance of Workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a hyperlink to a cell range (D10) in the worksheet
HyperLink UrlLink = sheet.getHyperLinks().add(sheet.getCellRange("D10"));

// Set the display text for the hyperlink using the text in cell D10
UrlLink.setTextToDisplay(sheet.getCellRange("D10").getText());

// Set the type of the hyperlink to URL
UrlLink.setType(HyperLinkType.Url);

// Set the URL address for the hyperlink
UrlLink.setAddress("http://en.wikipedia.org/wiki/Chicago");

// Add another hyperlink to a different cell range (E10) in the worksheet
HyperLink MailLink = sheet.getHyperLinks().add(sheet.getCellRange("E10"));

// Set the display text for the hyperlink using the text in cell E10
MailLink.setTextToDisplay(sheet.getCellRange("E10").getText());

// Set the type of the hyperlink to URL
MailLink.setType(HyperLinkType.Url);

// Set the email address for the hyperlink
MailLink.setAddress("mailto:Nancy.Aqua@gmail.com");
```

---

# Excel Image Hyperlink
## Add a hyperlink to an image in Excel worksheet
```java
// Create a new instance of Workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add an image picture to the worksheet at row 2, column 1 (B)
ExcelPicture picture = sheet.getPictures().add(2, 1, "imagePath");

// Set the hyperlink for the picture
picture.setHyperLink("https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java.html", true);
```

---

# Excel Hyperlink Type Extraction
## Extract hyperlink addresses and types from an Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Iterate over each hyperlink in the worksheet
for (HyperLink item : (Iterable<HyperLink>) sheet.getHyperLinks()) {
    // Get the address of the hyperlink
    String address = item.getAddress();

    // Get the type of the hyperlink
    HyperLinkType type = item.getType();
}
```

---

# Excel Image Hyperlink Extraction
## Get hyperlink address from an image in Excel worksheet
```java
// Get the first worksheet
Worksheet worksheet = workbook.getWorksheets().get(0);

// Get the first picture
ExcelPicture picture = worksheet.getPictures().get(0);

// Get the hyperlink address of this picture
String address = picture.getHyperLink().getAddress();
```

---

# Spire.XLS Hyperlink to External File
## Create hyperlink in Excel worksheet to link to an external file
```java
// Create a new instance of Workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the cell range at row 1, column 1 (A1)
CellRange range = sheet.getRange().get(1, 1);

// Add a hyperlink to the cell range
HyperLink hyperlink = sheet.getHyperLinks().add(range);

// Set the type of the hyperlink to File
hyperlink.setType(HyperLinkType.File);

// Set the text to display for the hyperlink
hyperlink.setTextToDisplay("Link To External File");

// Set the file address for the hyperlink
hyperlink.setAddress("data/AddDataTable.xlsx");
```

---

# Excel Hyperlink to Another Sheet Cell
## Create a hyperlink in an Excel cell that links to a specific cell in another worksheet
```java
// Create a new instance of Workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the cell range A1 in the worksheet
CellRange range = sheet.getRange().get("A1");

// Add a hyperlink to the cell range A1
HyperLink hyperlink = sheet.getHyperLinks().add(range);

// Set the type of the hyperlink to Workbook
hyperlink.setType(HyperLinkType.Workbook);

// Set the text to display for the hyperlink as "Link to Sheet2 cell C5"
hyperlink.setTextToDisplay("Link to Sheet2 cell C5");

// Set the address of the hyperlink to "Sheet2!C5" to link to cell C5 in Sheet2
hyperlink.setAddress("Sheet2!C5");
```

---

# Excel Hyperlink Modification
## Modify hyperlink text and address in Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the collection of hyperlinks in the worksheet
HyperLinksCollection links = sheet.getHyperLinks();

// Set the text to display for the first hyperlink to "Spire.XLS for JAVA"
links.get(0).setTextToDisplay("Spire.XLS for JAVA");

// Set the address of the first hyperlink to "https://www.e-iceblue.com/Introduce/xls-for-java.html"
links.get(0).setAddress("https://www.e-iceblue.com/Introduce/xls-for-java.html");
```

---

# Excel Hyperlink Reader
## Read hyperlinks from an Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Access hyperlinks in the worksheet
sheet.getHyperLinks().get(0).getAddress();
sheet.getHyperLinks().get(1).getAddress();
sheet.getHyperLinks().get(2).getAddress();
```

---

# Excel Hyperlink Removal
## Core functionality for removing hyperlinks from Excel worksheet cells
```java
// Get the first worksheet from the Workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the collection of hyperlinks in the worksheet
HyperLinksCollection links = sheet.getHyperLinks();

// Clear the contents and formatting of cells B1, B2, and B3
sheet.getCellRange("B1").clearAll();
sheet.getCellRange("B2").clearAll();
sheet.getCellRange("B3").clearAll();

// Remove the hyperlinks at index 0, 0, and 0 from the HyperLinksCollection
sheet.getHyperLinks().removeAt(0);
sheet.getHyperLinks().removeAt(0);
sheet.getHyperLinks().removeAt(0);
```

---

# Retrieve External File Hyperlinks
## Extracts information about external file hyperlinks from an Excel worksheet

```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Create a StringBuilder to store the content
StringBuilder content = new StringBuilder();

// Iterate over each hyperlink in the worksheet
for (HyperLink item : (Iterable<HyperLink>) sheet.getHyperLinks()) {
    // Get the address of the hyperlink
    String address = item.getAddress();

    // Get the name of the worksheet containing the hyperlink
    String sheetName = item.getRange().getWorksheetName();

    // Get the range of cells associated with the hyperlink
    CellRange range = item.getRange();

    // Append the cell information, sheet name, and address to the content StringBuilder
    content.append(String.format("Cell[%o,%o] in sheet \"" + sheetName + "\" contains File URL: %s", range.getRow(), range.getColumn(), address));
    content.append("\r\n");
}
```

---

# write hyperlinks in excel
## add different types of hyperlinks to excel cells
```java
// Set the text of cell B9 to "Home page"
sheet.getCellRange("B9").setText("Home page");

// Add a hyperlink to cell B10 with the address "http://www.e-iceblue.com"
HyperLink hylink1 = sheet.getHyperLinks().add(sheet.getCellRange("B10"));
hylink1.setAddress("http://www.e-iceblue.com");

// Set the text of cell B11 to "Support"
sheet.getCellRange("B11").setText("Support");

// Add a hyperlink to cell B12 with the email address "support@e-iceblue.com"
HyperLink hylink2 = sheet.getHyperLinks().add(sheet.getCellRange("B12"));
hylink2.setAddress("mailto:support@e-iceblue.com");

// Set the text of cell B13 to "Forum"
sheet.getCellRange("B13").setText("Forum");

// Add a hyperlink to cell B14 with the address "https://www.e-iceblue.com/forum/"
HyperLink hylink3 = sheet.getHyperLinks().add(sheet.getCellRange("B14"));
hylink3.setAddress("https://www.e-iceblue.com/forum/");
```

---

# Spire.XLS Custom Object Integration
## Add custom objects to Excel using marker designer
```java
// Define a Student class as a nested static class
public static class Student {
    // Declare private instance variables Name and Age
    private String Name;
    private int Age;

    // Create a constructor that takes name and age as parameters
    public Student(String name, int age) {
        // Initialize the Name and Age instance variables with the provided values
        this.Name = name;
        this.Age = age;
    }
}

// Set the value of cell A1 to "&=Student.Name"
sheet.getCellRange("A1").setValue("&=Student.Name");

// Set the value of cell B1 to "&=Student.Age"
sheet.getCellRange("B1").setValue("&=Student.Age");

// Create an ArrayList to store Student objects
ArrayList<Student> list = new ArrayList<Student>();

// Add Student objects to the list
list.add(new Student("John", 16));

// Add the "Student" parameter to the workbook's marker designer and assign the list as its value
workbook.getMarkerDesigner().addParameter("Student", list);

// Apply the marker design to the workbook
workbook.getMarkerDesigner().apply();
```

---

# Spire.XLS Variable Array Implementation
## Add variable array to Excel worksheet using marker designer
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the value of cell range A1 to "&=Array"
sheet.getCellRange("A1").setValue("&=Array");

// Add a parameter named "Array" with an array of strings as its value
workbook.getMarkerDesigner().addParameter("Array", new String[] { "Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email" });

// Apply the marker design to the workbook
workbook.getMarkerDesigner().apply();

// Calculate all the values in the workbook
workbook.calculateAllValue();

// Auto-fit the rows and columns in the allocated range of the worksheet
sheet.getAllocatedRange().autoFitRows();
sheet.getAllocatedRange().autoFitColumns();
```

---

# Spire.XLS Cell Style Copying
## Using Marker Designer to copy cell styles in Excel
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a DataTable to the marker designer
workbook.getMarkerDesigner().addDataTable("data", dataTable);

// Apply the marker design to the workbook
workbook.getMarkerDesigner().apply();
```

---

# Detect Blank Values in Excel
## Using marker designer to detect blank values in Excel data
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Create a new DataTable object
DataTable dt = new DataTable();
dt.setTableName("data");

// Create a new DataColumn with the name "value"
DataColumn column = new DataColumn("value");
dt.getColumns().add(column);

// Create rows with values and an empty string
DataRow row1 = dt.newRow();
row1.setObject(0, 120);

DataRow row2 = dt.newRow();
row2.setObject(0, 55);

DataRow row3 = dt.newRow();
row3.setObject(0, "");

// Add the rows to the DataTable
dt.getRows().add(row1);
dt.getRows().add(row2);
dt.getRows().add(row3);

// Add the DataTable to the workbook's marker designer
workbook.getMarkerDesigner().addDataTable("data", dt);

// Apply the marker design
workbook.getMarkerDesigner().apply();

// Calculate all the formulas in the workbook
workbook.calculateAllValue();
```

---

# Spire.XLS Marker Designer
## Apply marker designer with parameters and data table
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a parameter named "Variable1" with a value of 1234.5678 to the workbook's marker designer
workbook.getMarkerDesigner().addParameter("Variable1", 1234.5678);

// Add a DataTable named "Country" to the workbook's marker designer
workbook.getMarkerDesigner().addDataTable("Country", GetData(inputFile2));

// Apply the marker design to the workbook
workbook.getMarkerDesigner().apply();

// Automatically adjust the row height of the allocated range in the worksheet
sheet.getAllocatedRange().autoFitRows();

// Automatically adjust the column width of the allocated range in the worksheet
sheet.getAllocatedRange().autoFitColumns();
```

---

# Spire.XLS MarkerDesigner Data Direction
## Set data direction using MarkerDesigner in Excel
```java
// Create a new DataTable object
DataTable dt = new DataTable();

// Set the table name of the DataTable as "data"
dt.setTableName("data");

// Create a new DataColumn with the column name "value" and add it to the DataTable's columns collection
dt.getColumns().add(new DataColumn("value"));

// Create three new DataRow objects for the DataTable
DataRow drName1 = dt.newRow();
DataRow drName2 = dt.newRow();
DataRow drName3 = dt.newRow();

// Set the value of the "value" column in each DataRow
drName1.setString("value", "Text1");
drName2.setString("value", "Text2");
drName3.setString("value", "Text3");

// Add the DataRows to the DataTable's rows collection
dt.getRows().add(drName1);
dt.getRows().add(drName2);
dt.getRows().add(drName3);

// Add the DataTable to the MarkerDesigner in the workbook
workbook.getMarkerDesigner().addDataTable("data", dt);

// Apply the marker design to the workbook
workbook.getMarkerDesigner().apply();
```

---

# Format Named Range Cells in Excel
## Apply formatting to cells in a named range
```java
// Get the first named range from the workbook
INamedRange namedRange = workbook.getNameRanges().get(0);

// Get the range referred to by the named range
IXLSRange range = namedRange.getRefersToRange();

// Set the color of the range to yellow
range.getStyle().setColor(Color.yellow);

// Set the font of the range to bold
range.getStyle().getFont().isBold(true);
```

---

# Get All Named Ranges in Excel Workbook
## This code demonstrates how to retrieve all named ranges from an Excel workbook and process their names

```java
// Get the collection of named ranges from the workbook
INameRanges ranges = workbook.getNameRanges();

// Iterate over each named range in the collection
for (INamedRange nameRange : (Iterable<INamedRange>) ranges)
{
    // Get the name of the current named range
    String rangeName = nameRange.getName();
    // Additional processing can be done here with the named range
}
```

---

# Get Named Range Address
## Retrieves the address of a named range from an Excel workbook
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first named range from the workbook
INamedRange namedRange = workbook.getNameRanges().get(0);

// Get the address of the range referred to by the named range
String address = namedRange.getRefersToRange().getRangeAddress();
```

---

# Get Specific Named Range
## Retrieve named ranges from Excel workbook by index and by name
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the name of the named range at index 1
String name1 = workbook.getNameRanges().get(1).getName();

// Get the name of the named range with the name "NameRange3"
String name2 = workbook.getNameRanges().get("NameRange3").getName();

// Dispose of the workbook object and release any associated resources
workbook.dispose();
```

---

# Merge Named Range Cells
## Merge cells in a named range in an Excel workbook
```java
// Get the first named range from the workbook's collection of named ranges
INamedRange namedRange = workbook.getNameRanges().get(0);

// Get the range referred to by the named range
IXLSRange range = namedRange.getRefersToRange();

// Merge the cells within the range
range.merge();
```

---

# spire.xls named ranges management
## create and configure named ranges in Excel
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a new named range to the workbook and assign it to the variable NamedRange
INamedRange namedRange = workbook.getNameRanges().add("NewNamedRange");

// Set the range referred to by the named range to be cell range A8 to E12 in the worksheet
namedRange.setRefersToRange(sheet.getCellRange("A8:E12"));
```

---

# Remove Named Ranges in Excel
## This code demonstrates how to remove named ranges from an Excel workbook using Spire.XLS for Java
```java
// Remove the named range at index 0 from the workbook's collection of named ranges
workbook.getNameRanges().removeAt(0);

// Remove the named range with the name "NameRange2" from the workbook's collection of named ranges
workbook.getNameRanges().remove("NameRange2");
```

---

# Spire.XLS Named Range Renaming
## Rename a named range in Excel workbook
```java
// Create a new Workbook object to represent an Excel workbook
Workbook workbook = new Workbook();

// Retrieve the first named range in the workbook and set its name to "RenameRange"
workbook.getNameRanges().get(0).setName("RenameRange");

// Clean up and release any resources associated with the workbook
workbook.dispose();
```

---

# Spire.XLS Named Range
## Create and configure a named range in Excel worksheet
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a new named range to the worksheet and assign it to the variable namedRange
INamedRange namedRange = sheet.getNames().add("Range1");

// Set the range referred to by the named range to be cell range A1 to D19 in the worksheet
namedRange.setRefersToRange(sheet.getCellRange("A1:D19"));
```

---

# Excel Named Range Formula
## Create a named range and use it in a formula
```java
// Add a new named range to the workbook and assign it to the variable NamedRange
INamedRange namedRange = workbook.getNameRanges().add("MyNamedRange");

// Set the range referred to by the named range to be cell range B10 to B12 in the worksheet
namedRange.setRefersToRange(sheet.getCellRange("B10:B12"));

// Set the formula of cell B13 to calculate the sum of the named range "MyNamedRange"
sheet.getCellRange("B13").setFormula("=SUM(MyNamedRange)");

// Set the numeric values for cells B10, B11, and B12
sheet.getCellRange("B10").setNumberValue(10);
sheet.getCellRange("B11").setNumberValue(20);
sheet.getCellRange("B12").setNumberValue(30);
```

---

# Extract EMF OLE Objects from Excel
## This code demonstrates how to extract EMF OLE objects from an Excel worksheet
```java
// Create a Workbook object
Workbook workbook = new Workbook();

// Load the workbook from a file
workbook.loadFromFile("data/EmfOle.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Check if the worksheet has OLE objects
if (sheet.hasOleObjects()) {

    // Iterate over all OLE objects
    for (int i = 0; i < sheet.getOleObjects().size(); i++) {

        // Get the current OLE object
        IOleObject oleObject = sheet.getOleObjects().get(i);

        // Get the type of the current OLE object
        OleObjectType oleObjectType = sheet.getOleObjects().get(i).getObjectType();

        // Process the OLE object based on its type
        switch (oleObjectType) {
            case Emf:
                // If the OLE object is of type EMF, get its data
                byte[] emfData = oleObject.getOleData();
                break;
        }
    }
}
```

---

# Extract OLE Objects from Excel
## Extract and save OLE objects from an Excel worksheet based on their type
```java
// Create a new Workbook object and load from file
Workbook workbook = new Workbook();
workbook.loadFromFile("path/to/excel/file.xlsx");

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Check if the worksheet has any OLE objects
if (sheet.hasOleObjects()) {
    // Iterate through each OLE object in the worksheet
    for (int i = 0; i < sheet.getOleObjects().size(); i++) {
        // Get the current OLE object
        IOleObject Object = sheet.getOleObjects().get(i);

        // Get the type of the OLE object
        OleObjectType type = sheet.getOleObjects().get(i).getObjectType();

        // Perform different actions based on the type of the OLE object
        switch (type) {
            case WordDocument:
                // Extract the OLE data and save it as a Word document file
                byteArrayToFile(Object.getOleData(), "output/extractOLE.docx");
                break;
            case PowerPointSlide:
                // Extract the OLE data and save it as a PowerPoint slide file
                byteArrayToFile(Object.getOleData(), "output/extractOLE.pptx");
                break;
            case AdobeAcrobatDocument:
                // Extract the OLE data and save it as a PDF document file
                byteArrayToFile(Object.getOleData(), "output/extractOLE.pdf");
                break;
        }
    }
}

// Method to write a byte array to a file
public static void byteArrayToFile(byte[] data, String destPath) {
    // Create a File object with the specified destination path
    File dest = new File(destPath);

    try (
            // Create an InputStream from the byte array using ByteArrayInputStream
            InputStream is = new ByteArrayInputStream(data);

            // Create an OutputStream for writing to the file, with buffering, using FileOutputStream
            OutputStream os = new BufferedOutputStream(new FileOutputStream(dest, false));
    ) {
        // Create a byte array for flushing data
        byte[] flush = new byte[1024];

        int len = -1;
        // Read data from the input stream and write it to the output stream in chunks
        while ((len = is.read(flush)) != -1) {
            os.write(flush, 0, len);
        }

        // Flush the output stream to ensure all data is written
        os.flush();
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

---

# Insert OLE Objects in Excel
## Core functionality for embedding OLE objects in Excel worksheets

```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

// Generate an image for the OLE object
BufferedImage image = GenerateImage(fileName);

// Add an OLE object to the worksheet, embedding the file and using the generated image
IOleObject oleObject = worksheet.getOleObjects().add(fileName, image, OleLinkType.Embed);

// Set the location of the OLE object
oleObject.setLocation(worksheet.getCellRange("B4"));

// Set the object type of the OLE object
oleObject.setObjectType(OleObjectType.ExcelWorksheet);

// Generate an image from a given file name
private static BufferedImage GenerateImage(String fileName) {
    // Create a new Workbook object
    Workbook book = new Workbook();

    // Load the workbook from the specified file
    book.loadFromFile(fileName);

    // Set the margins of the first worksheet to 0
    book.getWorksheets().get(0).getPageSetup().setLeftMargin(0);
    book.getWorksheets().get(0).getPageSetup().setRightMargin(0);
    book.getWorksheets().get(0).getPageSetup().setTopMargin(0);
    book.getWorksheets().get(0).getPageSetup().setBottomMargin(0);

    // Convert the first worksheet to an image
    return book.getWorksheets().get(0).toImage(1, 1, 19, 5);
}
```

---

# Excel Paper Dimensions Retrieval
## Get dimensions of different paper sizes in Excel
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the paper size of the worksheet's page setup to A2Paper
sheet.getPageSetup().setPaperSize(PaperSizeType.A2Paper);

// Set the paper size of the worksheet's page setup to PaperA3
sheet.getPageSetup().setPaperSize(PaperSizeType.PaperA3);

// Set the paper size of the worksheet's page setup to PaperA4
sheet.getPageSetup().setPaperSize(PaperSizeType.PaperA4);

// Set the paper size of the worksheet's page setup to PaperLetter
sheet.getPageSetup().setPaperSize(PaperSizeType.PaperLetter);

// Release any resources used by the workbook
workbook.dispose();
```

---

# Excel Page Setup
## Set Excel page order type
```java
// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the order type of the pages to "OverThenDown"
pageSetup.setOrder(OrderType.OverThenDown);
```

---

# Set Excel Paper Size
## Set the paper size of an Excel worksheet to A4
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the paper size of the worksheet's page setup to PaperA4
sheet.getPageSetup().setPaperSize(PaperSizeType.PaperA4);
```

---

# Excel Page Setup First Page Number
## Set the first page number for worksheet printing
```java
// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the first page number of the worksheet's page setup to 2
sheet.getPageSetup().setFirstPageNumber(2);
```

---

# Excel Header and Footer Margins Setup
## Set header and footer margins in Excel worksheet
```java
// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the header margin to 2 inches
pageSetup.setHeaderMarginInch(2);

// Set the footer margin to 2 inches
pageSetup.setFooterMarginInch(2);
```

---

# Excel Sheet Margin Setup
## Set page margins for an Excel worksheet
```java
// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the bottom margin to 2 units
pageSetup.setBottomMargin(2);

// Set the left margin to 1 unit
pageSetup.setLeftMargin(1);

// Set the right margin to 1 unit
pageSetup.setRightMargin(1);

// Set the top margin to 3 units
pageSetup.setTopMargin(3);
```

---

# Excel Printing Options Setup
## Configure various printing options for Excel worksheet
```java
// Get the PageSetup object from the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the option to print gridlines on the page setup
pageSetup.isPrintGridlines(true);

// Set the option to print headings on the page setup
pageSetup.isPrintHeadings(true);

// Set the page setup to black and white
pageSetup.setBlackAndWhite(true);

// Set the type of comments to print on the page setup
pageSetup.setPrintComments(PrintCommentType.InPlace);

// Set the type of errors to print on the page setup
pageSetup.setPrintErrors(PrintErrorsType.NA);

// Set the page setup to draft mode
pageSetup.setDraft(true);
```

---

# Excel Page Orientation Setup
## Set page orientation to landscape in Excel worksheet
```java
// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the orientation of the page to Landscape
sheet.getPageSetup().setOrientation(PageOrientationType.Landscape);
```

---

# Excel Print Area Setup
## Set print area for Excel worksheet
```java
// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the PageSetup object from the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the print area to be "A1:E5"
pageSetup.setPrintArea("A1:E5");
```

---

# Excel Print Quality Setting
## Set print quality for Excel worksheet
```java
// Retrieve the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the print quality of the worksheet to 180
sheet.getPageSetup().setPrintQuality(180);
```

---

# Spire.XLS Print Title Setup
## Set print title columns and rows for Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the print title columns to "$A:$B"
pageSetup.setPrintTitleColumns("$A:$B");

// Set the print title rows to "$1:$2"
pageSetup.setPrintTitleRows("$1:$2");
```

---

# Excel Worksheet Page Setup
## Set worksheet fit to page properties
```java
// Get the first worksheet in the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the number of pages tall to fit to 1 page
sheet.getPageSetup().setFitToPagesTall(1);

// Set the number of pages wide to fit to 1 page
sheet.getPageSetup().setFitToPagesWide(1);
```

---

# Excel Sheet Page Centering
## Code to center an Excel worksheet both horizontally and vertically on the page when printing
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the center horizontally option to true
pageSetup.setCenterHorizontally(true);

// Set the center vertically option to true
pageSetup.setCenterVertically(true);
```

---

# Clear Pivot Table Fields
## This code demonstrates how to clear all data fields from a pivot table in Excel
```java
// Get the worksheet named "PivotTable"
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Get the first PivotTable in the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Clear all data fields in the PivotTable
pt.getDataFields().clear();

// Calculate the data in the PivotTable
pt.calculateData();
```

---

# Pivot Table Consolidation Functions
## Set subtotal types for pivot table data fields
```java
// Get the worksheet named "PivotTable" from the workbook
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Get the first PivotTable from the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Set the subtotal type of the first data field to Average
pt.getDataFields().get(0).setSubtotal(SubtotalTypes.Average);

// Set the subtotal type of the second data field to Maximum
pt.getDataFields().get(1).setSubtotal(SubtotalTypes.Max);

// Calculate the data in the PivotTable
pt.calculateData();
```

---

# Excel Pivot Table Creation
## Create a pivot table in Excel using Spire.XLS for Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Define a CellRange object for the data range A1:C7
CellRange dataRange = sheet.getCellRange("A1:C7");

// Add a PivotCache with the data range to the workbook
PivotCache cache = workbook.getPivotCaches().add(dataRange);

// Add a PivotTable with the cache to the worksheet at cell E10
PivotTable pt = sheet.getPivotTables().add("Pivot Table", sheet.getCellRange("E10"), cache);

// Get the PivotField for "Product"
PivotField pf = null;
if (pt.getPivotFields().get("Product") instanceof PivotField) {
    pf = (PivotField) pt.getPivotFields().get("Product");
}
pf.setAxis(AxisTypes.Row);

// Get the PivotField for "Month"
PivotField pf2 = null;
if (pt.getPivotFields().get("Month") instanceof PivotField) {
    pf2 = (PivotField) pt.getPivotFields().get("Month");
}
pf2.setAxis(AxisTypes.Row);

// Add a data field to the PivotTable for "Count" with a custom name, subtotal type "Sum"
pt.getDataFields().add(pt.getPivotFields().get("Count"), "SUM of Count", SubtotalTypes.Sum);

// Set the built-in style of the PivotTable to PivotStyleMedium12
pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium12);

// Calculate the data in the PivotTable
pt.calculateData();

// Autofit columns 5 and 6 in the worksheet
sheet.autoFitColumn(5);
sheet.autoFitColumn(6);
```

---

# Spire.XLS Pivot Table Field Name Customization
## Customize pivot table field names including row, column, and data fields
```java
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Access the first pivot table in the worksheet
XlsPivotTable pivotTable = (XlsPivotTable)sheet.getPivotTables().get(0);

// Set a custom name for the row field
pivotTable.getRowFields().get(0).setCustomName("rowName");

// Set a custom name for the column field
pivotTable.getColumnFields().get(0).setCustomName("colName");

// Set a custom name for the data field
pivotTable.getDataFields().get(0).setCustomName("DataName");

// Calculate the pivot table data
pivotTable.calculateData();
```

---

# Disable Pivot Table Ribbon
## Code to disable the wizard/ribbon for a pivot table in an Excel file
```java
// Get the worksheet named "PivotTable" from the workbook
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Get the first PivotTable in the worksheet
XlsPivotTable pt = (XlsPivotTable) sheet.getPivotTables().get(0);

// Disable the wizard for the PivotTable
pt.setEnableWizard(false);
```

---

# Spire.XLS Pivot Table Row Expansion
## Expand or collapse rows in a pivot table by hiding or showing item details
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);
// Get the first pivot table on the worksheet
XlsPivotTable pivotTable = (XlsPivotTable) sheet.getPivotTables().get(0);
// Calculate the data for the pivot table
pivotTable.calculateData();

// Hide the item detail with the value "3501" in the "Vendor No" pivot field
((XlsPivotField) pivotTable.getPivotFields().get("Vendor No")).hideItemDetail("3501", true);

// Show the item detail with the value "3502" in the "Vendor No" pivot field
((XlsPivotField) pivotTable.getPivotFields().get("Vendor No")).hideItemDetail("3502", false);
```

---

# Spire.XLS Pivot Table Data Field Formatting
## Format pivot table data field to display as percentage of column
```java
// Get the first pivot table from the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Get the first data field from the pivot table
PivotDataField pivotDataField = pt.getDataFields().get(0);
// Set the display format of the data field to "Percentage Of Column"
pivotDataField.setShowDataAs(PivotFieldFormatType.PercentageOfColumn);
```

---

# Pivot Table Formatting
## Format the appearance of a pivot table in Excel
```java
// Get the worksheet named "PivotTable"
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Get the first PivotTable in the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Set the built-in style of the PivotTable to PivotStyleLight10
pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleLight10);

// Enable the grid drop zone for the PivotTable
pt.getOptions().setShowGridDropZone(true);

// Set the row layout type of the PivotTable to Compact
pt.getOptions().setRowLayout(PivotTableLayoutType.Compact);
```

---

# Get Pivot Table Refresh Information
## Extract refresh information including date and user from a pivot table in Excel
```java
// Get the first pivot table from the worksheet
XlsPivotTable pivotTable = (XlsPivotTable) worksheet.getPivotTables().get(0);

// Get the refresh date and refreshed by information from the pivot table's cache
DateTime dateTime = pivotTable.getCache().getRefreshDate();
String refreshedBy = pivotTable.getCache().getRefreshedBy();

// Create a result string with refreshed by and refreshed date information
String result = "Pivot table refreshed by:  " + refreshedBy + "\r\nPivot table refreshed date: " + dateTime.toString();
```

---

# Excel Pivot Table Grouping and Ungrouping
## Demonstrates how to group and ungroup data in Excel pivot tables
```java
// Get the first pivot table from the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);
// Get the PivotField object for the "Count" field
PivotField r1 = (PivotField)pt.getPivotFields().get("Count");
// Manually group the values in the "Count" field based on a range of values
pt.setManualGroupField(r1, 7, 15, EnumSet.of(PivotGroupByType.RangeOfValues), 2);

// Get the first pivot table from the worksheet
XlsPivotTable pt2 = (XlsPivotTable)sheet2.getPivotTables().get(0);
// Get the PivotField object for the "Count" field
PivotField r2 = (PivotField)pt2.getPivotFields().get("Count");
// Ungroup the values in the "Count" field
pt2.setUngroup(r2);
```

---

# Pivot Table Date Grouping
## Group pivot table data by date range with specified interval
```java
// Get the first pivot table in the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Get the first row field in the pivot table
IPivotField field = pt.getRowFields().get(0);

// Set the start and end dates for grouping
Date start = new  Date("2023/1/5");
Date end = new  Date("2023/3/2");

// Set the group by type to days
PivotGroupByTypes[] types = new PivotGroupByTypes[] { PivotGroupByTypes.Days };

// Create a new group with the specified start and end dates, group by type, and interval
field.createGroup(start, end, types, 10);

// Calculate the pivot table data
pt.calculateData();

// Refresh the pivot table cache
pt.getCache().isRefreshOnLoad(true);
```

---

# Hide All Items in Pivot Table
## This code demonstrates how to hide all items in a specific pivot table field using Spire.XLS for Java.
```java
// Get the first pivot table from the worksheet
XlsPivotTable pivotTable = (XlsPivotTable)sheet.getPivotTables().get(0);
// Get the PivotField object for the "Product" field
PivotField pivotField = (PivotField)pivotTable.getPivotFields().get("Product");
// Hide all items in the "Product" field
pivotField.hideAllItem(true);

// Calculate the data in the pivot table
pivotTable.calculateData();
```

---

# Excel Pivot Table Layout Configuration
## Set the layout of an Excel pivot table to Tabular format
```java
// Get the first pivot table from the worksheet
XlsPivotTable xlsPivotTable = (XlsPivotTable)worksheet.getPivotTables().get(0);

// Set the report layout of the pivot table to Tabular
xlsPivotTable.getOptions().setReportLayout(PivotTableLayoutType.Tabular);
```

---

# Refresh Pivot Table
## Refresh pivot table by updating data and setting refresh on load property
```java
// Update data that will be reflected in the pivot table
sheet.getRange().get("D2").setValue("999");

// Get the PivotTable
XlsPivotTable pt = (XlsPivotTable) workbook.getWorksheets().get(0).getPivotTables().get(0);

// Set the refresh on load property of the pivot table's cache to true
pt.getCache().isRefreshOnLoad(true);
```

---

# Spire.XLS for Java Pivot Table Formatting
## Set format options for pivot tables in Excel
```java
// Get the first PivotTable from the worksheet
XlsPivotTable pt = (XlsPivotTable)sheet.getPivotTables().get(0);

// Enable automatic formatting for the PivotTable
pt.getOptions().isAutoFormat(true);

// Show row grand totals in the PivotTable
pt.setShowRowGrand(true);

// Show column grand totals in the PivotTable
pt.setShowColumnGrand(true);

// Display the string "null" for cells with null values in the PivotTable
pt.setDisplayNullString(true);
pt.setNullString("null");

// Set the page field order in the PivotTable to DownThenOver
pt.setPageFieldOrder(PagesOrderType.DownThenOver);
```

---

# Excel Pivot Table Field Formatting
## Set formatting options for a pivot table field including sort type, subtotal settings, and auto show
```java
// Get the first PivotTable from the Worksheet and cast it to XlsPivotTable
XlsPivotTable pt = (XlsPivotTable) sheet.getPivotTables().get(0);

// Get the first PivotField from the PivotTable and cast it to PivotField
PivotField pf = (PivotField) pt.getPivotFields().get(0);

// Set the sort type of the PivotField to ascending
pf.setSortType(PivotFieldSortType.Ascending);

// Enable top subtotal for the PivotField
pf.setSubtotalTop(true);

// Set the subtotal type of the PivotField to Count
pf.setSubtotals(SubtotalTypes.Count);

// Enable auto show for the PivotField
pf.isAutoShow(true);
```

---

# Excel Pivot Table Conditional Formatting
## Set conditional formatting for pivot table fields in Excel
```java
// Get the Worksheet named "PivotTable" from the workbook
Worksheet worksheet = workbook.getWorksheets().get("PivotTable");

// Get the first PivotTable from the Worksheet
PivotTable table = (PivotTable)worksheet.getPivotTables().get(0);

// Get the collection of PivotConditionalFormats from the PivotTable
PivotConditionalFormatCollection pcfs = table.getPivotConditionalFormats();

// Add a new PivotConditionalFormat for the first data field in the PivotTable
PivotConditionalFormat pc = pcfs.addPivotConditionalFormat(table.getDataFields().get(0));

// Add a new condition to the PivotConditionalFormat
IConditionalFormat cf= pc.addCondition();

// Set the format type of the condition to NotContainsBlanks
cf.setFormatType(ConditionalFormatType.NotContainsBlanks);

// Set the fill pattern of the condition to Solid
cf.setFillPattern(ExcelPatternType.Solid);

// Set the background color of the condition to blue
cf.setBackColor(Color.blue);
```

---

# Excel Pivot Table with Repeat Labels
## Create a pivot table and configure repeat labels for pivot fields
```java
// Get the range of cells from the original worksheet
CellRange dataRange = sheet.getRange().get("A1:C9");

// Create a PivotCache using the data range
PivotCache cache = workbook.getPivotCaches().add(dataRange);

// Add a PivotTable to the second worksheet using the PivotCache
PivotTable pt = sheet2.getPivotTables().add("Pivot Table", sheet.getCellRange("A1"), cache);

// Get the pivot field for "VendorNo"
IPivotField r1 = pt.getPivotFields().get("VendorNo");

// Set the axis for the pivot field to Row
r1.setAxis(AxisTypes.Row);

// Set the row header caption for the PivotTable to "VendorNo"
pt.getOptions().setRowHeaderCaption("VendorNo");

// Disable subtotals for the "VendorNo" field
r1.setSubtotals(SubtotalTypes.None);

// Enable repeating item labels for the "VendorNo" field
r1.isRepeatItemLabels(true);

// Get the pivot field for "Desc"
IPivotField r2 = pt.getPivotFields().get("Desc");

// Set the axis for the pivot field to Row
r2.setAxis(AxisTypes.Row);

// Set the row layout for the PivotTable to Tabular
pt.getOptions().setRowLayout(PivotTableLayoutType.Tabular);

// Add a data field to the PivotTable for "OnHand" with the caption "Sum of onHand" and summary type as Sum
pt.getDataFields().add(pt.getPivotFields().get("OnHand"), "Sum of onHand", SubtotalTypes.Sum);

// Set the built-in style of the PivotTable to PivotStyleMedium12
pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium12);
```

---

# Pivot Table Data Source Update
## Update pivot table data source and refresh the pivot table
```java
// Get the worksheet named "Data"
Worksheet data = workbook.getWorksheets().get("Data");

// Set the text value of cell A2 in the "Data" worksheet to "NewValue"
data.getRange().get("A2").setText("NewValue");

// Set the numeric value of cell D2 in the "Data" worksheet to 28000
data.getRange().get("D2").setNumberValue(28000);

// Get the worksheet named "PivotTable"
Worksheet sheet = workbook.getWorksheets().get("PivotTable");

// Get the first pivot table from the "PivotTable" worksheet
XlsPivotTable pt = (XlsPivotTable) sheet.getPivotTables().get(0);

// Enable refresh on load for the pivot table cache
pt.getCache().isRefreshOnLoad(true);

// Calculate the data for the pivot table
pt.calculateData();
```

---

# Excel Page Setup for Printing
## Configure page settings for printing Excel documents
```java
// Get the PageSetup for the worksheet
PageSetup pageSetup = worksheet.getPageSetup();

// Set the print area for the worksheet to cells A1:E19
pageSetup.setPrintArea("A1:E19");

// Set the columns A:E as titles to repeat on each printed page
pageSetup.setPrintTitleColumns("$A:$E");

// Set the rows 1:2 as titles to repeat on each printed page
pageSetup.setPrintTitleRows("$1:$2");

// Enable printing of gridlines
pageSetup.isPrintGridlines(true);

// Enable printing of headings
pageSetup.isPrintHeadings(true);

// Set the worksheet to be printed in black and white
pageSetup.setBlackAndWhite(true);

// Set the option to print comments in place
pageSetup.setPrintComments(PrintCommentType.InPlace);

// Set the print quality to 150 dpi
pageSetup.setPrintQuality(150);

// Set the order of printing to over then down
pageSetup.setOrder(OrderType.OverThenDown);

// Get the default PrinterJob
PrinterJob loPrinterJob = PrinterJob.getPrinterJob();

// Get the default PageFormat
PageFormat loPageFormat = loPrinterJob.defaultPage();

// Get the Paper object from the PageFormat
Paper loPaper = loPageFormat.getPaper();

// Set the imageable area of the Paper to cover the entire page
loPaper.setImageableArea(0, 0, loPageFormat.getWidth(), loPageFormat.getHeight());

// Set the number of copies to 1
loPrinterJob.setCopies(1);

// Set the paper of the PageFormat to the modified Paper object
loPageFormat.setPaper(loPaper);

// Set the printable object to the workbook and PageFormat to the modified PageFormat
loPrinterJob.setPrintable(workbook, loPageFormat);
```

---

# Excel Document Printing
## Configure and print an Excel document using PrinterJob
```java
// Create a new Workbook object
Workbook loDoc = new Workbook();

// Get the default PrinterJob
PrinterJob loPrinterJob = PrinterJob.getPrinterJob();

// Get the default PageFormat
PageFormat loPageFormat = loPrinterJob.defaultPage();

// Get the Paper object from the PageFormat
Paper loPaper = loPageFormat.getPaper();

// Set the imageable area of the paper to cover the entire page
loPaper.setImageableArea(0, 0, loPageFormat.getWidth(), loPageFormat.getHeight());

// Set the number of copies to 1
loPrinterJob.setCopies(1);

// Set the paper of the PageFormat to the modified Paper object
loPageFormat.setPaper(loPaper);

// Set the printable object to the Workbook and PageFormat to the modified PageFormat
loPrinterJob.setPrintable(loDoc, loPageFormat);

try {
    // Print the Workbook using the PrinterJob
    loPrinterJob.print();
} catch (PrinterException e) {
    e.printStackTrace();
}
// Clean up resources and release memory
loDoc.dispose();
```

---

# Excel Digital Signature Management
## Add and remove digital signatures from Excel workbooks
```java
// Add digital signature
// Create a CertificateAndPrivateKey object using the certificate file and password
CertificateAndPrivateKey cap = new CertificateAndPrivateKey(certificatePath, password);

// Add a digital signature to the Workbook using the certificate, comment, and current date
workbook.addDigitalSignature(cap, comment, new Date());

// Remove digital signature
// Remove all digital signatures from the workbook
workbook.removeAllDigitalSignatures();
```

---

# Check Excel Worksheet Password Protection
## Verify if a password is correct for a protected worksheet
```java
// Create a Workbook instance
Workbook workbook = new Workbook();

// Load the Excel document
workbook.loadFromFile("data/checkPassProtected.xlsx");

// Get the first worksheet
Worksheet worksheet = workbook.getWorksheets().get(0);

// Verify the given password is correct or not
boolean isCorrect = worksheet.checkProtectionPassword("e-iceblue");
```

---

# Excel Protection Detection
## Detects if an Excel workbook is password protected
```java
String input = "data/protectedWorkbook.xlsx";
// Detect if the Excel workbook is password protected
boolean value = Workbook.isPasswordProtected(input);
```

---

# Excel Formula Hiding
## Hide formulas in Excel worksheet and protect it with password
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Hide formulas in the allocated range of the worksheet
sheet.getAllocatedRange().isFormulaHidden(true);

// Protect the worksheet with a password "e-iceblue"
sheet.protect("e-iceblue");
```

---

# Lock Specific Cells in Excel
## This code demonstrates how to lock specific cells in a new Excel worksheet using Spire.XLS for Java
```java
public class lockSpecificCellInNewExcel {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new Workbook();
        // Create an empty sheet in the workbook
        workbook.createEmptySheet();
        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Iterate through rows 0 to 254
        for (int i = 0; i < 255; i++) {
            // Disable locking for each row's style
            sheet.getRows()[i].getStyle().setLocked(false);
        }

        // Set the text of column 3 (fourth column) to "Locked"
        sheet.getColumns()[3].setText("Locked");
        // Lock the style of column 3 to prevent modifications
        sheet.getColumns()[3].getStyle().setLocked(true);

        // Protect the worksheet with password "123" and apply all protection options
        sheet.protect("123", EnumSet.of(SheetProtectionType.All));
    }
}
```

---

# Excel Column Locking
## Lock specific column in Excel sheet with password protection
```java
// Create a workbook
Workbook workbook = new Workbook();

// Create an empty worksheet
workbook.createEmptySheet();

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Loop through all the columns in the worksheet and unlock them
for (int i = 0; i < 255; i++) {
    sheet.getRows()[i].getStyle().setLocked(false);
}

// Lock the fourth column in the worksheet
sheet.getColumns()[3].setText("Locked");
sheet.getColumns()[3].getStyle().setLocked(true);

// Set the password
sheet.protect("123", EnumSet.of(SheetProtectionType.All));
```

---

# Lock Specific Row in Excel
## Lock a specific row in an Excel worksheet while keeping other rows unlocked
```java
// Create a new workbook
Workbook workbook = new Workbook();
// Create an empty sheet in the workbook
workbook.createEmptySheet();
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Iterate through rows 0 to 254
for (int i = 0; i < 255; i++) {
    // Disable locking for each row's style
    sheet.getRows()[i].getStyle().setLocked(false);
}

// Set the text of row 2 to "Locked"
sheet.getRows()[2].setText("Locked");
// Lock the style of row 2 to prevent modifications
sheet.getRows()[2].getStyle().setLocked(true);

// Protect the worksheet with password "123" and apply all protection options
sheet.protect("123", EnumSet.of(SheetProtectionType.All));
```

---

# Excel Cell Protection
## Lock and unlock specific cells and protect worksheet with password
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set cell B3 as locked
sheet.getRange().get("B3").getStyle().setLocked(true);

// Set cell C3 as unlocked
sheet.getRange().get("C3").getStyle().setLocked(false);

// Protect the worksheet with password "TestPassword" and allow all types of protection
sheet.protect("TestPassword", EnumSet.of(SheetProtectionType.All));
```

---

# Worksheet Protection with Editable Range
## Protect worksheet while allowing specific ranges to be editable
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add an editable range named "EditableRanges" for cells B4 to E12
sheet.addAllowEditRange("EditableRanges", sheet.getCellRange("B4:E12"));

// Protect the worksheet with password "TestPassword" and allow all types of protection
sheet.protect("TestPassword", EnumSet.of(SheetProtectionType.All));
```

---

# Spire.XLS Workbook Protection
## Protect an Excel workbook with a password
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Protect the entire workbook with password "e-iceblue"
workbook.protect("e-iceblue");
```

---

# Unlock Protected Excel Sheet
## Unprotect a password-protected worksheet in Excel using Spire.XLS for Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Unprotect the worksheet with password "e-iceblue"
sheet.unprotect("e-iceblue");
```

---

# Excel Sheet Protection Removal
## Remove password protection from an Excel worksheet
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Remove protection from the worksheet
sheet.unprotect();

// Dispose of the workbook object to release resources
workbook.dispose();
```

---

# Extract Text from Excel TextBox
## This code demonstrates how to extract text from a textbox in an Excel worksheet using Spire.XLS for Java

```java
// Create a Workbook and load an Excel file
Workbook workbook = new Workbook();
workbook.loadFromFile("data/template_Xls_5.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the first TextBox and extract its text
XlsTextBoxShape shape = (XlsTextBoxShape) sheet.getTextBoxes().get(0);
String extractedText = shape.getText();
```

---

# Get TextBox by Name
## Retrieve a text box from Excel worksheet by its name and get its text content
```java
// Add a new TextBoxShape to the Worksheet at position (2, 2) with width 18 and height 65
ITextBoxShape textBox = sheet.getTextBoxes().addTextBox(2, 2, 18, 65);

// Set the name of the TextBoxShape as "FirstTextBox"
textBox.setName("FirstTextBox");

// Set the text content for the TextBoxShape
textBox.setText("Spire.XLS for Java is a professional Java Excel API that enables developers to create, manage, manipulate, convert and print Excel worksheets without using Microsoft Office or Microsoft Excel.");

// Retrieve the TextBoxShape named "FirstTextBox" from the Worksheet
ITextBoxShape FindTextBox = sheet.getTextBoxes().get("FirstTextBox");

// Get the text content of the TextBoxShape
String text = FindTextBox.getText();
```

---

# Excel Text Box Manipulation
## Modify text box content and alignment in Excel worksheet
```java
// Get the first TextBox shape from the worksheet
ITextBox tb = sheet.getTextBoxes().get(0);

// Set the text content of the TextBox
tb.setText("Spire.XLS for Java");

// Set the horizontal alignment of the TextBox to Center
tb.setHAlignment(CommentHAlignType.Center);

// Set the vertical alignment of the TextBox to Center
tb.setVAlignment(CommentVAlignType.Center);
```

---

# Excel Textbox Border Removal
## Remove borderline of textbox in Excel using Spire.XLS
```java
// Add a TextBox shape to the chart at position (50, 50) with width 100 and height 600
XlsTextBoxShape textbox1 = (XlsTextBoxShape)chart.getTextBoxes().addTextBox(50, 50, 100, 600);
textbox1.setText("The solution with borderline");

// Add another TextBox shape to the chart at position (1000, 50) with width 100 and height 600
XlsTextBoxShape textbox2 = (XlsTextBoxShape)chart.getTextBoxes().addTextBox(1000, 50, 100, 600);
textbox2.setText("The solution without borderline");

// Set the line weight of textbox2 to 0, effectively removing the border line around it
textbox2.getLine().setWeight(0);
```

---

# Replace Text in Excel TextBox
## This code demonstrates how to replace text in TextBox shapes within an Excel worksheet
```java
// Specify the tag values to search for and their corresponding replacements
String tag = "TAG_1,TAG_2";
String replace = "Spire.XLS for JAVA,Spire.XLS for .NET";

// Iterate over each tag and its replacement
for (int i = 0; i < tag.split(",").length; i++) {
    ReplaceTextInTextBox(sheet, "<" + tag.split(",")[i] + ">", replace.split(",")[i]);
}

// Helper method to replace text in TextBox shapes within a Worksheet
private static void ReplaceTextInTextBox(Worksheet sheet, String sFind, String sReplace) {
    // Iterate over each TextBox shape in the Worksheet
    for (int i = 0; i < sheet.getTextBoxes().getCount(); i++) {
        ITextBox tb = sheet.getTextBoxes().get(i);

        // Check if the TextBox has non-empty text
        if (tb.getText() != "" && tb.getText() != null) {
            // Check if the text contains the specified search string
            if (tb.getText().contains(sFind)) {
                // Replace the search string with the replacement string
                tb.setText(tb.getText().replace(sFind, sReplace));
            }
        }
    }
}
```

---

# Excel Text Box Styling
## Set font and background properties for text box in Excel
```java
// Get the TextBox shape from the Worksheet
XlsTextBoxShape shape = (XlsTextBoxShape) sheet.getTextBoxes().get(0);

// Create a new ExcelFont object
ExcelFont font = workbook.createFont();

// Set the font properties
font.setFontName("Century Gothic");
font.setSize(10);
font.isBold(true);
font.setColor(Color.blue);

// Apply the font to the text within the TextBox shape
(new RichText(shape.getRichText())).setFont(0, shape.getText().length() - 1, font);

// Set the fill type of the TextBox shape to SolidColor
shape.getFill().setFillType(ShapeFillType.SolidColor);

// Set the foreground color of the TextBox shape to BlueGray
shape.getFill().setForeKnownColor(ExcelColors.BlueGray);
```

---

# Excel Textbox Internal Margin Configuration
## This code demonstrates how to create a textbox in an Excel worksheet and configure its internal margins.
```java
// Add a TextBox shape to the worksheet at position (4, 2) with width 100 and height 300
XlsTextBoxShape textbox = (XlsTextBoxShape) sheet.getTextBoxes().addTextBox(4, 2, 100, 300);

// Set the text content of the TextBox
textbox.setText("Insert TextBox in Excel and set the margin for the text");

// Set the horizontal alignment of the TextBox to Center
textbox.setHAlignment(CommentHAlignType.Center);

// Set the vertical alignment of the TextBox to Center
textbox.setVAlignment(CommentVAlignType.Center);

// Set the inner left margin of the TextBox to 1 point
textbox.setInnerLeftMargin(1);

// Set the inner right margin of the TextBox to 3 points
textbox.setInnerRightMargin(3);

// Set the inner top margin of the TextBox to 1 point
textbox.setInnerTopMargin(1);

// Set the inner bottom margin of the TextBox to 1 point
textbox.setInnerBottomMargin(1);
```

---

# Excel Text Box with Text Wrapping
## Demonstrates how to enable text wrapping in an Excel text box
```java
// Get the first Worksheet from the Workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the TextBox shape at index 0 from the Worksheet
XlsTextBoxShape shape = (XlsTextBoxShape) sheet.getTextBoxes().get(0);

// Enable text wrapping for the TextBox shape
shape.isWrapText(true);
```

---

# Activate Worksheet in Excel Workbook
## This code demonstrates how to activate a specific worksheet in an Excel workbook
```java
// Get the second worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(1);

// Activate the worksheet
sheet.activate();
```

---

# Excel Page Break Management
## Add horizontal and vertical page breaks in Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add a horizontal page break at cell E4 in the worksheet
sheet.getHPageBreaks().add(sheet.getRange().get("E4"));

// Add a vertical page break at cell C4 in the worksheet
sheet.getVPageBreaks().add(sheet.getRange().get("C4"));
```

---

# Adding Worksheet to Excel Workbook
## This code demonstrates how to add a new worksheet to an Excel workbook and set text in a cell
```java
// Add a new worksheet to the Workbook and assign it to the 'sheet' variable
Worksheet sheet = workbook.getWorksheets().add("AddedSheet");
// Set the text "This is a new sheet" in cell C5 of the newly added worksheet
sheet.getRange().get("C5").setText("This is a new sheet.");
```

---

# Excel Worksheet Style Application
## Apply custom style to Excel worksheet
```java
// Create a new CellStyle object named "newStyle"
CellStyle style = workbook.getStyles().addStyle("newStyle");

// Set the color of the style to CYAN
style.setColor(Color.CYAN);

// Set the font color of the style to white
style.getFont().setColor(Color.white);

// Set the font size of the style to 15
style.getFont().setSize(15);

// Make the font bold in the style
style.getFont().isBold(true);

// Apply the style to the worksheet
sheet.applyStyle(style);
```

---

# Copy Worksheet Between Excel Files
## Copy a worksheet from one Excel workbook to another using Spire.XLS for Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Get the PageSetup object for the worksheet
PageSetup pageSetup = sheet.getPageSetup();

// Set the print title rows to be "$1:$5"
pageSetup.setPrintTitleRows("$1:$5");

// Create a new Workbook object
Workbook workbook1 = new Workbook();

// Get the first worksheet from the second workbook
Worksheet sheet1 = workbook1.getWorksheets().get(0);

// Copy the contents of the original worksheet to the new worksheet
sheet1.copyFrom(sheet);
```

---

# Spire.XLS Worksheet Copying
## Copy a worksheet within a workbook
```java
// Add a new worksheet to the workbook and assign it to the 'sheet1' variable
Worksheet sheet1 = workbook.getWorksheets().add("MySheet");

// Get the range of cells that are allocated with data in the source worksheet
CellRange sourceRange = sheet.getAllocatedRange();

// Copy the source range from the source worksheet to the destination worksheet ('sheet1')
// Starting from the first row and column of the source worksheet, and overwrite existing data
sheet.copy(sourceRange, sheet1, sheet.getFirstRow(), sheet.getFirstColumn(), true);
```

---

# Copy Visible Worksheets
## Copy only visible worksheets from one workbook to another
```java
// Create a new workbook for the copied sheets
Workbook workbookNew = new Workbook();

// Set the version of the new workbook
workbookNew.setVersion(ExcelVersion.Version2013);

// Clear any existing worksheets in the new workbook
workbookNew.getWorksheets().clear();

// Iterate through each worksheet in the original workbook
for (Object worksheet : workbook.getWorksheets()) {
    // Convert the object to Worksheet type
    Worksheet sheet = (Worksheet) worksheet;

    // Check if the visibility of the sheet is set to "Visible"
    if (sheet.getVisibility() == WorksheetVisibility.Visible) {
        // Add a copy of the visible sheet to the new workbook
        workbookNew.getWorksheets().addCopy(sheet);
    }
}

// Release resources associated with the workbooks
workbook.dispose();
workbookNew.dispose();
```

---

# Spire.XLS Worksheet Copy
## Copy a worksheet from one workbook to another
```java
// Get the first worksheet from the source workbook
Worksheet srcWorksheet = sourceWorkbook.getWorksheets().get(0);

// Create a new worksheet in the target workbook with the name "added"
Worksheet targetWorksheet = targetWorkbook.getWorksheets().add("added");

// Copy the contents of the source worksheet to the target worksheet
targetWorksheet.copyFrom(srcWorksheet);
```

---

# Detect Empty Worksheet
## Check if Excel worksheets are empty using Spire.XLS
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet worksheet1 = workbook.getWorksheets().get(0);

// Check if the first worksheet is empty
boolean detect1 = worksheet1.isEmpty();

// Get the second worksheet from the workbook
Worksheet worksheet2 = workbook.getWorksheets().get(1);

// Check if the second worksheet is empty
boolean detect2 = worksheet2.isEmpty();
```

---

# Excel Worksheet Data Filling
## Fill data into worksheet cells and set formatting
```java
// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.getWorksheets().get(0);

// Set header cells to bold
worksheet.getRange().get("A1").getStyle().getFont().isBold(true);
worksheet.getRange().get("B1").getStyle().getFont().isBold(true);
worksheet.getRange().get("C1").getStyle().getFont().isBold(true);

// Fill column A with month names
worksheet.getRange().get("A1").setText("Month");
worksheet.getRange().get("A2").setText("January");
worksheet.getRange().get("A3").setText("February");
worksheet.getRange().get("A4").setText("March");
worksheet.getRange().get("A5").setText("April");

// Fill column B with payment values
worksheet.getRange().get("B1").setText("Payments");
worksheet.getRange().get("B2").setNumberValue(251);
worksheet.getRange().get("B3").setNumberValue(515);
worksheet.getRange().get("B4").setNumberValue(454);
worksheet.getRange().get("B5").setNumberValue(874);

// Fill column C with sample text
worksheet.getRange().get("C1").setText("Sample");
worksheet.getRange().get("C2").setText("Sample1");
worksheet.getRange().get("C3").setText("Sample2");
worksheet.getRange().get("C4").setText("Sample3");
worksheet.getRange().get("C5").setText("Sample4");

// Set column width
worksheet.setColumnWidth(2, 10);
```

---

# Export Filtered Excel Values to CSV
## This code demonstrates how to export filtered values from an Excel worksheet to a CSV file
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Load the Excel file from the specified input path
workbook.loadFromFile(input);

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Save the worksheet to the specified output path, with space as delimiter and without auto-filtering
sheet.saveToFile(output, " ", false);
```

---

# Excel Freeze Panes
## Freeze panes in an Excel worksheet at a specified row and column
```java
// Freeze panes at row 2, column 1 (second row, first column)
sheet.freezePanes(2, 1);
```

---

# Get Group Boxes from Excel Worksheet
## Demonstrates how to retrieve and iterate through group boxes in an Excel worksheet using Spire.XLS for Java
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

// Get the collection of group boxes in the worksheet
IGroupBoxes groupBoxes = worksheet.getGroupBoxes();

// Iterate through each group box in the collection
for (int i = 0; i < groupBoxes.getCount(); i++) {
    // Get the name of the current group box
    String name = groupBoxes.get(i).getName();
}
```

---

# Extract Fonts from Excel Workbook
## This code demonstrates how to extract and list all fonts used in an Excel workbook
```java
// Create an ArrayList to store the fonts used in the workbook
ArrayList<IFont> fonts = new ArrayList<IFont>();

// Iterate through each worksheet in the workbook
for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
    // Get the current worksheet
    Worksheet sheet = workbook.getWorksheets().get(i);

    // Iterate through each row in the worksheet
    for (int r = 0; r < sheet.getRows().length; r++) {
        // Iterate through each cell in the row
        for (int c = 0; c < sheet.getRows()[r].getCellList().size(); c++) {
            // Get the font of the current cell and add it to the fonts ArrayList
            fonts.add(sheet.getRows()[r].getCellList().get(c).getStyle().getFont());
        }
    }
}

// Create a StringBuilder to store the information about the fonts
StringBuilder strB = new StringBuilder();
for (int i = 0; i < fonts.size(); i++) {
    // Get the font at the current index
    IFont font = fonts.get(i);

    // Append the font name and size information to the StringBuilder
    strB.append(String.format("FontName:" + font.getFontName() + "; FontSize:" + font.getSize() + "\n"));
}
```

---

# Get Excel Worksheet Page Count
## Retrieve page count for each worksheet in an Excel file
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Load the Excel file from the specified path
workbook.loadFromFile("data/worksheetSample2.xlsx");

// Get the page information for each worksheet and store it in a list of maps
List<Map<Integer, PageColRow>> pageInfoList = workbook.getSplitPageInfo();

// Iterate through each worksheet
for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
    // Get the name of the current worksheet
    String sheetname = workbook.getWorksheets().get(i).getName();

    // Get the page count for the current worksheet
    int pagecount = pageInfoList.get(i).size();
}
```

---

# Get Paper Size from Excel Worksheets
## This code demonstrates how to retrieve the paper size (width and height) from each worksheet in an Excel workbook.
```java
// Iterate through each worksheet in the workbook
for (Object worksheet : workbook.getWorksheets()) {
    // Get the current worksheet
    Worksheet sheet = (Worksheet) worksheet;

    // Get the page width and height of the worksheet's page setup
    double width = sheet.getPageSetup().getPageWidth();
    double height = sheet.getPageSetup().getPageHeight();
}
```

---

# Get Worksheet Names
## Extracts all worksheet names from an Excel workbook
```java
// Create a StringBuilder to store the worksheet names
StringBuilder stringBuilder = new StringBuilder();

// Iterate through each worksheet in the workbook
for (Object worksheet : workbook.getWorksheets()) {
    // Get the current worksheet
    Worksheet sheet = (Worksheet) worksheet;

    // Append the worksheet name to the StringBuilder
    stringBuilder.append(sheet.getName() + "\r\n");
}
```

---

# Spire.XLS Worksheet Visibility Control
## Hide or show worksheets in an Excel workbook
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Set the visibility of the worksheet named "Sheet1" to Hidden
workbook.getWorksheets().get("Sheet1").setVisibility(WorksheetVisibility.Hidden);

// Set the visibility of the second worksheet to Visible
workbook.getWorksheets().get(1).setVisibility(WorksheetVisibility.Visible);
```

---

# Hide Excel Worksheet Tabs
## This code demonstrates how to hide worksheet tabs in an Excel workbook using Spire.XLS for Java
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Hide the worksheet tabs in the workbook
workbook.setShowTabs(false);
```

---

# Hide Zero Values in Excel Worksheet
## This code demonstrates how to hide zero values in an Excel worksheet using Spire.XLS for Java
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the display of zeros in the worksheet to false
sheet.isDisplayZeros(false);
```

---

# Excel Document Property Link to Content
## Demonstrates how to add a custom document property and set its LinkToContent property in an Excel workbook
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the collection of custom document properties and add a new property with name "Test" and value "MyNamedRange"
workbook.getCustomDocumentProperties().add("Test", "MyNamedRange");

// Retrieve the collection of custom document properties
ICustomDocumentProperties properties = workbook.getCustomDocumentProperties();

// Get the specific document property named "Test"
DocumentProperty property = (DocumentProperty) properties.get("Test");

// Set the linkToContent property of the document property to true
property.setLinkToContent(true);
```

---

# Move Excel Worksheet
## Move a worksheet to a specified position within a workbook
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Move the worksheet to the index position 2
sheet.moveWorksheet(2);
```

---

# Excel Page Break Preview
## Set zoom scale for page break preview in Excel worksheet
```java
// Set the zoom scale for page break preview to 80%
sheet.setZoomScalePageBreakView(80);
```

---

# Remove Page Breaks in Excel
## This code demonstrates how to remove vertical and horizontal page breaks in an Excel worksheet
```java
// Get the worksheet
Worksheet sheet = workbook.getWorksheets().get(0);

// Clear all vertical page breaks in the worksheet
sheet.getVPageBreaks().clear();

// Remove the horizontal page break at index 0 in the worksheet
sheet.getHPageBreaks().removeAt(0);

// Set the worksheet view mode to Preview
sheet.setViewMode(ViewMode.Preview);
```

---

# Spire.XLS Worksheet Management
## Move worksheet to a different position in workbook
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Move the worksheet to the index position 2
sheet.moveWorksheet(2);
```

---

# Setting Page Breaks in Excel Worksheets
## This code demonstrates how to add horizontal and vertical page breaks to an Excel worksheet and set the view mode to Preview
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Add horizontal page breaks at specific cell ranges in the worksheet
sheet.getHPageBreaks().add(sheet.getRange().get("A8"));
sheet.getHPageBreaks().add(sheet.getRange().get("A14"));

// Uncomment the following lines to add vertical page breaks at specific cell ranges in the worksheet
//sheet.getVPageBreaks().add(sheet.getRange().get("B1"));
//sheet.getVPageBreaks().add(sheet.getRange().get("C1"));

// Set the view mode of the first worksheet to Preview
workbook.getWorksheets().get(0).setViewMode(ViewMode.Preview);
```

---

# Setting Worksheet Tab Colors
## This code demonstrates how to set different tab colors for worksheets in an Excel workbook
```java
// Get the first worksheet from the Workbook
Worksheet worksheet = workbook.getWorksheets().get(0);
// Set the tab color of the worksheet to red
worksheet.setTabColor(Color.red);

// Get the second worksheet from the Workbook
worksheet = workbook.getWorksheets().get(1);
// Set the tab color of the worksheet to green
worksheet.setTabColor(Color.green);

// Get the third worksheet from the Workbook
worksheet = workbook.getWorksheets().get(2);
// Set the tab color of the worksheet to cyan
worksheet.setTabColor(Color.CYAN);
```

---

# Excel Worksheet View Mode Setting
## Set worksheet view mode to Preview
```java
// Create a new Workbook object.
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook and set its view mode to Preview.
workbook.getWorksheets().get(0).setViewMode(ViewMode.Preview);
```

---

# Excel Grid Lines Visibility Control
## Show or hide grid lines in Excel worksheets
```java
// Get the first worksheet from the workbook.
Worksheet sheet1 = workbook.getWorksheets().get(0);

// Get the second worksheet from the workbook.
Worksheet sheet2 = workbook.getWorksheets().get(1);

// Set the grid lines visibility to false for sheet1.
sheet1.setGridLinesVisible(false);

// Set the grid lines visibility to true for sheet2.
sheet2.setGridLinesVisible(true);
```

---

# Excel Worksheet Tab Display
## Show or hide worksheet tabs in Excel workbook
```java
// Create a new Workbook object.
Workbook workbook = new Workbook();

// Set the show tabs option to true, which displays the worksheet tabs in the Excel application.
workbook.setShowTabs(true);

// Dispose of system resources associated with the workbook.
workbook.dispose();
```

---

# Split Worksheet into Panes
## This code demonstrates how to split an Excel worksheet into multiple panes
```java
// Get the first worksheet from the Workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the index of the first visible column to 2
sheet.setFirstVisibleColumn(2);
// Set the index of the first visible row to 5
sheet.setFirstVisibleRow(5);
// Set the vertical split position at 4000th row
sheet.setVerticalSplit(4000);
// Set the horizontal split position at 5000th column
sheet.setHorizontalSplit(5000);

// Set the active pane to be the bottom right pane (pane 1)
sheet.setActivePane(1);
```

---

# unfreeze excel panes
## remove frozen panes from an Excel worksheet
```java
// Get the first worksheet from the Workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Remove all panes from the worksheet
sheet.removePanes();
```

---

# Spire.XLS Worksheet Protection Verification
## Check if a worksheet is password protected
```java
// Get the first worksheet from the Workbook
Worksheet worksheet = workbook.getWorksheets().get(0);

// Check if the first worksheet is password protected
boolean detect = worksheet.isPasswordProtected();
```

---

# Excel Worksheet Zoom Factor Control
## Set zoom level for Excel worksheet
```java
// Create a new Workbook object.
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook.
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the zoom factor of the worksheet to 85%.
sheet.setZoom(85);
```

---

# Access Excel Document Properties
## Demonstrates how to access custom document properties in an Excel workbook
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Get the custom document properties of the workbook
ICustomDocumentProperties properties = workbook.getCustomDocumentProperties();

// Retrieve the "Editor" document property
DocumentProperty property1 = (DocumentProperty) properties.get("Editor");

// Retrieve the document property at index 0
DocumentProperty property2 = (DocumentProperty) properties.get(0);
```

---

# Add Custom Properties to Excel Workbook
## This code demonstrates how to add custom document properties to an Excel workbook
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Add a custom document property named "_MarkAsFinal" with a boolean value of true
workbook.getCustomDocumentProperties().add("_MarkAsFinal", true);

// Add a custom document property named "The Editor" with a string value of "E-iceblue"
workbook.getCustomDocumentProperties().add("The Editor", "E-iceblue");

// Add a custom document property named "Phone number" with an integer value of 81705109
workbook.getCustomDocumentProperties().add("Phone number", 81705109);

// Add a custom document property named "Revision number" with a double value of 7.12
workbook.getCustomDocumentProperties().add("Revision number", 7.12);

// Add a custom document property named "Revision date" with the current date and time
workbook.getCustomDocumentProperties().add("Revision date", new Date());
```

---

# Excel Workbook Decryption
## Decrypt password-protected Excel workbooks
```java
// Check if the specified Excel file is password protected
boolean value = Workbook.isPasswordProtected(fileName);

// If the file is password protected
if (value) {
    // Create a new Workbook object
    Workbook workbook = new Workbook();

    // Set the open password for the workbook
    workbook.setOpenPassword("eiceblue");

    // Load and open the password-protected workbook
    workbook.loadFromFile(fileName);

    // Remove the protection from the workbook
    workbook.unProtect();

    // Save the decrypted workbook to the specified output path in Excel 2013 format
    workbook.saveToFile(output, ExcelVersion.Version2013);

    // Clean up and release any resources used by the workbook
    workbook.dispose();
}
```

---

# Detect Excel Version
## Detect the version of Excel files using Spire.XLS for Java
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Load the Excel file from the specified file path
workbook.loadFromFile(filePath);

// Get the version of the loaded workbook
ExcelVersion version = workbook.getVersion();

// Clean up and release any resources used by the workbook
workbook.dispose();
```

---

# Detect VBA Macros in Excel
## Check if an Excel workbook contains VBA macros
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Check if the workbook contains VBA macros
boolean hasMacros = workbook.hasMacros();
```

---

# Excel Workbook Encryption
## Protect workbook with password
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Protect the workbook with a password ("eiceblue")
workbook.protect("eiceblue");
```

---

# Excel Workbook Properties Extraction
## Extract built-in and custom document properties from an Excel workbook
```java
// Get the built-in document properties of the workbook
BuiltInDocumentProperties properties1 = workbook.getDocumentProperties();

// Iterate through each built-in property
for (int i = 0; i < properties1.getCount(); i++) {
    // Get the name and value of the property
    String name = properties1.get(i).getName();
    String value = properties1.get(i).getValue().toString();
}

// Get the custom document properties of the workbook
ICustomDocumentProperties properties2 = workbook.getCustomDocumentProperties();

// Iterate through each custom property
for (int i = 0; i < properties2.getCount(); i++) {
    // Get the name and value of the property
    String name = properties2.get(i).getName();
    String value = properties2.get(i).getValue().toString();
}
```

---

# Excel Workbook with Macros
## Load and save Excel file with macros
```java
// Create a new instance of Workbook
Workbook workbook = new Workbook();

// Load an existing Excel file with macros from the specified path
workbook.loadFromFile("data/macroSample.xls");

// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the text "This is a simple test!" to cell A5 in the worksheet
sheet.getRange().get("A5").setText("This is a simple test!");

// Specify the output file path for saving the modified workbook
String output = "output/loadAndSaveFileWithMacro_result.xls";

// Save the workbook as an Excel 97-2003 file format
workbook.saveToFile(output, ExcelVersion.Version97to2003);

// Clean up and release resources used by the workbook
workbook.dispose();
```

---

# Merge Excel Files
## Combine multiple Excel files into a single workbook
```java
// Create a list of stored file paths
List<String> files = new ArrayList<String>();

// Create a new Excel workbook
Workbook newbook = new Workbook();

// Set the version of the new workbook
newbook.setVersion(ExcelVersion.Version2013);

// Empty the worksheets in the new workbook
newbook.getWorksheets().clear();

// Create a temporary Excel workbook
Workbook tempbook = new Workbook();

// Go through the file list
for (String file : files) {

    // Load files from the temporary workbook
    tempbook.loadFromFile(file);

    // Traverse the worksheets in the temporary workbook and copy them to the new workbook
    for (Object workSheet : tempbook.getWorksheets()) {
        Worksheet sheet = (Worksheet) workSheet;
        newbook.getWorksheets().addCopy(sheet, WorksheetCopyType.CopyAll);
    }
}
```

---

# Open Encrypted Excel File
## Try multiple passwords to open an encrypted Excel file
```java
// Iterate through the password array
for (int i = 0; i < passwords.length; i++) {
    try {
        // Create Workbook objects
        Workbook workbook = new Workbook();

        // Set the open password
        workbook.setOpenPassword(passwords[i]);

        // Load the Workbook object from the file
        workbook.loadFromFile(filePath);

        // Free the resources of the Workbook object
        workbook.dispose();
    } catch (Exception ex) {
        // Handle exception for incorrect password
    }
}
```

---

# Excel Workbook Opening Methods
## Demonstrates various ways to open Excel files of different formats

```java
// Load workbook from file path
Workbook workbook = new Workbook();
workbook.loadFromFile(filePath);
workbook.dispose();

// Load workbook from stream
InputStream stream = new FileInputStream(filePath);
Workbook workbookFromStream = new Workbook();
workbookFromStream.loadFromStream(stream);
stream.close();
workbookFromStream.dispose();

// Load Excel 97-2003 format workbook
Workbook workbook97 = new Workbook();
workbook97.loadFromFile(filePath, ExcelVersion.Version97to2003);
workbook97.dispose();

// Load workbook from XML file
Workbook workbookXML = new Workbook();
workbookXML.loadFromXml(xmlFilePath);
workbookXML.dispose();

// Load workbook from CSV file with custom parameters
Workbook workbookCSV = new Workbook();
workbookCSV.loadFromFile(csvFilePath, delimiter, startRow, startColumn);
workbookCSV.dispose();
```

---

# Excel Stream Reading
## Load workbook from input stream
```java
// Create a new workbook object
Workbook workbook = new Workbook();

// Open the input file stream for reading the Excel file
FileInputStream fileStream = new FileInputStream("data/readStream.xlsx");

// Load the workbook from the input stream
workbook.loadFromStream(fileStream);
```

---

# Java Excel Custom Properties Removal
## Remove custom properties from Excel workbook
```java
// Get the custom document properties of the workbook
ICustomDocumentProperties customDocumentProperties = workbook.getCustomDocumentProperties();

// Remove the custom document property with the name "Editor"
customDocumentProperties.remove("Editor");
```

---

# Excel Workbook Save in Multiple Formats
## Core functionality for saving Excel workbooks in various file formats
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Save the workbook to "result.xls" in Excel 97-2003 format
workbook.saveToFile("output/result.xls", ExcelVersion.Version97to2003);

// Save the workbook to "result.xlsx" in Excel 2010 format
workbook.saveToFile("output/result.xlsx", ExcelVersion.Version2010);

// Save the workbook to "result.xlsb" in XLSB 2010 format
workbook.saveToFile("output/result.xlsb", ExcelVersion.Xlsb2010);

// Save the workbook to "result.ods" in ODS (Open Document Spreadsheet) format
workbook.saveToFile("output/result.ods", ExcelVersion.ODS);

// Save the workbook to "result.pdf" in PDF format
workbook.saveToFile("output/result.pdf", FileFormat.PDF);

// Save the workbook to "result.xml" in XML format
workbook.saveToFile("output/result.xml", FileFormat.XML);

// Save the workbook to "result.xps" in XPS format
workbook.saveToFile("output/result.xps", FileFormat.XPS);
```

---

# Save Workbook to Stream
## Demonstrates how to save an Excel workbook to a file output stream
```java
Workbook workbook = new Workbook();
FileOutputStream fileStream = new FileOutputStream("output/saveStream_result.xlsx");
workbook.saveToStream(fileStream, FileFormat.Version2013);
fileStream.close();
workbook.dispose();
```

---

# Excel Calculation Mode
## Set Excel calculation mode to manual
```java
// Create a new Workbook object
Workbook workbook = new Workbook();

// Set the calculation mode of the workbook to Manual
workbook.setCalculationMode(ExcelCalculationMode.Manual);
```

---

# Excel Page Margins Setting
## Set page margins for Excel worksheet
```java
// Get the first worksheet from the workbook
Worksheet sheet = workbook.getWorksheets().get(0);

// Set the top, bottom, left, and right margins of the page setup for the worksheet
sheet.getPageSetup().setTopMargin(0.3);
sheet.getPageSetup().setBottomMargin(1);
sheet.getPageSetup().setLeftMargin(0.2);
sheet.getPageSetup().setRightMargin(1);

// Set the header and footer margins in inches for the page setup of the worksheet
sheet.getPageSetup().setHeaderMarginInch(0.1);
sheet.getPageSetup().setFooterMarginInch(0.5);
```

---

# Excel Tracked Changes Management
## Accept or reject all tracked changes in an Excel workbook
```java
//Create a new Document object
Workbook workbook = new Workbook();

// Accept the changes or reject the changes.
//workbook.acceptAllTrackedChanges();
workbook.rejectAllTrackedChanges();

//Dispose
workbook.dispose();
```

---

# Enable Track Changes in Excel
## This code demonstrates how to enable track changes feature in an Excel workbook
```java
//Create a new Document object
Workbook workbook = new Workbook();

//Enable tracking changes
workbook.setTrackedChanges(true);
```

---

# Create Slicer From Pivot Table
## This code demonstrates how to create slicers from a pivot table in Excel using Spire.XLS for Java

```java
// Get pivot table collection
PivotTablesCollection pivotTables = worksheet.getPivotTables();

//Add a PivotTable to the worksheet
CellRange dataRange = worksheet.getCellRange("A1:C9");
PivotCache cache = wb.getPivotCaches().add(dataRange);

//Cell to put the pivot table
PivotTable pt = worksheet.getPivotTables().add("TestPivotTable", worksheet.getCellRange("A12"), cache);

//Drag the fields to the row area.
PivotField pf = (PivotField)pt.getPivotFields().get("fruit");
pf.setAxis(AxisTypes.Row);
PivotField pf2 =  (PivotField)pt.getPivotFields().get("year");
pf2.setAxis(AxisTypes.Column);

//Drag the field to the data area.
pt.getDataFields().add(pt.getPivotFields().get("amount"), "SUM of Count", SubtotalTypes.Sum);

//Set PivotTable style
pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium10);

pt.calculateData();

//Get slicer collection
XlsSlicerCollection slicers = worksheet.getSlicers();

int index = slicers.add(pt, "E12", 0);

XlsSlicer xlsSlicer = slicers.get(index);
xlsSlicer.setName("xlsSlicer");
xlsSlicer.setWidth(100);
xlsSlicer.setHeight(120);
xlsSlicer.setStyleType(SlicerStyleType.SlicerStyleLight2);
xlsSlicer.isPositionLocked(true);

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache = xlsSlicer.getSlicerCache();
slicerCache.setCrossFilterType(SlicerCacheCrossFilterType.ShowItemsWithNoData);

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.getSlicerCache().getSlicerCacheItems();
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems.get(0);
xlsSlicerCacheItem.isSelected(false);

XlsSlicerCollection slicers_2 = worksheet.getSlicers();

IPivotField r1 = pt.getPivotFields().get("year");
int index_2 = slicers_2.add(pt, "I12", r1);

XlsSlicer xlsSlicer_2 = slicers.get(index_2);
xlsSlicer_2.setRowHeight(40);
xlsSlicer_2.setStyleType(SlicerStyleType.SlicerStyleLight3);
xlsSlicer_2.isPositionLocked(false);

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache_2 = xlsSlicer_2.getSlicerCache();
slicerCache_2.setCrossFilterType(SlicerCacheCrossFilterType.ShowItemsWithDataAtTop);

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems_2 = xlsSlicer_2.getSlicerCache().getSlicerCacheItems();
XlsSlicerCacheItem xlsSlicerCacheItem_2 = slicerCacheItems_2.get(1);
xlsSlicerCacheItem_2.isSelected(false);
pt.calculateData();
```

---

# Create Excel Slicer From Table
## This code demonstrates how to create slicers from a table in Excel using different style types
```java
// Get slicer collection
XlsSlicerCollection slicers = worksheet.getSlicers();

//Create a table with the data from the specific cell range.
IListObject table = worksheet.getListObjects().create("Super Table", worksheet.getCellRange("A1:C9"));

int count = 3;
int index = 0;
for(SlicerStyleType type : SlicerStyleType.values()) {
    count += 5;
    String range = "E" + count;
    index = slicers.add(table, range.toString(), 0);

    //Style setting
    XlsSlicer xlsSlicer = slicers.get(index);
    xlsSlicer.setName("slicers_" + count);
    xlsSlicer.setStyleType(type);
}
```

---

# Excel Slicer Modification
## Modify slicer properties, style, and cache settings in Excel worksheets
```java
// Get the first worksheet in the workbook
Worksheet worksheet = wb.getWorksheets().get(0);

// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.getSlicers();

// Get the first slicer from the slicer collection
XlsSlicer xlsSlicer = slicers.get(0);

// Set the style of the slicer to a dark theme (style type 4)
xlsSlicer.setStyleType(SlicerStyleType.SlicerStyleDark4);

// Change the caption (title) of the slicer
xlsSlicer.setCaption("Modified Slicer");

// Lock the position of the slicer to prevent it from being moved in the worksheet
xlsSlicer.isPositionLocked(true);

// Get the collection of cache items associated with the slicer
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.getSlicerCache().getSlicerCacheItems();

// Get the first cache item in the collection
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems.get(0);

// Deselect the cache item
xlsSlicerCacheItem.isSelected(false);

// Get the display value of the cache item
String displayValue = xlsSlicerCacheItem.getDisplayValue();

// Get the slicer cache associated with the slicer
XlsSlicerCache slicerCache = xlsSlicer.getSlicerCache();

// Set the cross-filter type to show items even if they have no associated data
slicerCache.setCrossFilterType(SlicerCacheCrossFilterType.ShowItemsWithNoData);
```

---

# Read Excel Slicer Information
## Extract and display slicer properties from Excel worksheet
```java
// Create a new Workbook instance
Workbook wb = new Workbook();

// Load an existing Excel file
wb.loadFromFile("data/SlicerTemplate.xlsx");

// Get the first worksheet in the workbook
Worksheet worksheet = wb.getWorksheets().get(0);

// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.getSlicers();

// Iterate through each slicer to extract information
for (int i = 0; i < slicers.getCount(); i++) {
    XlsSlicer xlsSlicer = slicers.get(i);
    
    // Get slicer properties
    xlsSlicer.getName();
    xlsSlicer.getCaption();
    xlsSlicer.getNumberOfColumns();
    xlsSlicer.getColumnWidth();
    xlsSlicer.getRowHeight();
    xlsSlicer.isShowCaption();
    xlsSlicer.isPositionLocked();
    xlsSlicer.getWidth();
    xlsSlicer.getHeight();

    // Get slicer cache information
    XlsSlicerCache slicerCache = xlsSlicer.getSlicerCache();
    slicerCache.getSourceName();
    slicerCache.isTabular();
    slicerCache.getName();

    // Get slicer cache item information
    XlsSlicerCacheItemCollection slicerCacheItems = slicerCache.getSlicerCacheItems();
    XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems.get(1);
    xlsSlicerCacheItem.isSelected();
}

// Dispose of the workbook object to release resources
wb.dispose();
```

---

# Remove Slicer from Excel Worksheet
## Remove all slicers from an Excel worksheet using Spire.XLS for Java
```java
// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.getSlicers();

// Clear all slicers from the collection
slicers.clear();
```

---

