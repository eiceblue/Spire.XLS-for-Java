package demoQ2;

import com.spire.xls.*;

public class ExcelToJSON {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load the Excel file from the specified path
        workbook.loadFromFile("Data/DataExport.xlsx");

        // Save the workbook data as a JSON file
        workbook.saveToFile("ExcelToJSON.json", FileFormat.Json);

        // Release all resources used by the workbook
        workbook.dispose();
    }
}
