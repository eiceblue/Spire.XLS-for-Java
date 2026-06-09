import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.equation.IXlsEquation;

public class addExportUpdateMathEquations {
    public static void main(String[] args) {
        // Create a new workbook object
        Workbook workbook = new Workbook();

        // Load an existing Excel file from the specified path
        workbook.loadFromFile("data/addExportUpdateMathEquations.xlsx");

        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);
        
        // Insert Equations
        sheet.getEquations().addEquation(20, 0, 100, 100, "x_{1}^{2}");
       
        //Get the Equations
        String mathML = sheet.getEquations().get(0).exportMathML();
        String LaTex = sheet.getEquations().get(1).exportLaTex();
        System.out.println("MathML:"+mathML);
        System.out.println("LaTex:"+LaTex);

        // Update the special math equation
        IXlsEquation equation1 = sheet.getEquations().get(0);
        equation1.updateByLaTexText("\\begin{pmatrix} \r\n  1 & 0 \\\\ \r\n  0 & 1  \r\n \\end{pmatrix} \r\n \\left(x-1\\right)\\left(x+3\\right) \\left(x-1\\right)\\left(x+3\\right)");

        // Specify the file path where the result will be saved
        String result = "output/addExportUpdateMathEquations_result.xlsx";
        
        // Save the workbook to the specified file path in Excel 2013 format
        workbook.saveToFile(result, ExcelVersion.Version2013);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
