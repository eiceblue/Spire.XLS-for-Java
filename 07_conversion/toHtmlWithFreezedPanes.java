package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.HTMLOptions;

public class toHtmlWithFreezedPanes {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Load an existing workbook with a pivot table from a file
        workbook.loadFromFile("Data\\Freezedpanes.xlsx");

        //Set to true to maintain freeze panes effect in HTML, false to ignore
        HTMLOptions.Default.isSaveFreezedPanes(true);

        // Specify the output file name for the result
        String result = "ToHtmlWithFreezedPanes.html";

        // Save the modified workbook to the specified file using Excel 2010 format
        workbook.saveToHtml(result);

        // Dispose of the workbook object to release resources
        workbook.dispose();

    }
}
