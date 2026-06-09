import com.spire.xls.*;

public class useFInvFormula {
    public static void main(String[] args) {
        // Create a new workbook object
        Workbook workbook = new Workbook();
        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Use F.INV formula
        sheet.getCellRange("A1").setFormula("=F.INV(0.05, 4, 5)");
        workbook.calculateAllValue();

        // Specify the file path where the result will be saved
        String result = "output/useFInvFormula_result.xlsx";

        // Save the workbook to the specified file path in Excel 2013 format
        workbook.saveToFile(result, ExcelVersion.Version2013);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
