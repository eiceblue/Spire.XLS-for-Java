import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;

public class excelToMarkdown {
    public static void main(String[] args) throws Exception{
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load an existing Excel file into the workbook
        workbook.loadFromFile("data/excelToMarkdown.xlsx");

        //Save to Markdown document
        workbook.saveToFile("output/ExcelToMarkdown_out.md", FileFormat.Markdown);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
