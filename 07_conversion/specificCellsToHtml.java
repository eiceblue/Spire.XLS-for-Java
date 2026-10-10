package demoQ2;

import com.spire.xls.*;
import java.io.*;

public class specificCellsToHtml {
    public static void main(String[] args) throws IOException {
        //Create a workbook
        Workbook workbook = new Workbook();

        //Load the document from disk
        workbook.loadFromFile("Data/ConversionSample1.xlsx");

        //Get the first worksheet in Excel file
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Get the specific cell range A1:B3 from the worksheet
        CellRange cell = sheet.getCellRange("A1:E7");

        // Extract the HTML representation of the selected cell range
        String html = cell.getHtmlString();

        // Specify the output file name for the HTML result
        String result = "SpecificCellsToHtml_out.html";

        // Write the HTML content to the output file
        writeStringToFile(html, result);

        // Dispose of the workbook object to release resources
        workbook.dispose();
    }

    // Custom method to write a string to a file
    public static void writeStringToFile(String content, String fileName) throws IOException {
        FileWriter fWriter = new FileWriter(fileName, true);
        try {
            // Write the content to the file
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        }finally {
            fWriter.flush();
            fWriter.close();
        }
    }
}
