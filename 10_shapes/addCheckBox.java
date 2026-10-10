package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.XlsRange;

public class addCheckBox {
    public static void main(String[] args) {
        // Create a new instance of the Workbook
        Workbook workbook = new Workbook();
        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Get the specific cell range at A1
        XlsRange range = sheet.getRange().get("A1");
        // Insert an embedded Office365 checkbox into the cell A1
        range.insertEmbedCheckBox();
        // Set the state of the embedded checkbox to checked (true)
        range.setEmbedCheckBoxCheckState(true);

        // Save the workbook to a new file named 'addCheckBox_office365.xlsx'
        workbook.saveToFile("addCheckBox_office365.xlsx");
        // Release resources and dispose of the workbook object
        workbook.dispose();
    }
}
