package demoQ2;

import com.spire.xls.*;

public class XMATCHFunction {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Load an existing workbook with a pivot table from a file
        workbook.loadFromFile("Data/XMATCHFunction.xlsx");

        // Get the first worksheet in the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Set the formula for cell C4 to use XMATCH function
        sheet.getCellRange("C4").setFormula("=XMATCH(\"Lili\", A2:A5)");

        //Calculate all cells
        workbook.calculateAllValue();

        // Specify the output file name for the result
        String result = "XMATCHFunction_result.xlsx";

        // Save the modified workbook to the specified file using Excel 2010 format
        workbook.saveToFile(result, ExcelVersion.Version2016);

        // Dispose of the workbook object to release resources
        workbook.dispose();
    }
}
