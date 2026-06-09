import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;

public class markdownToExcel {
    public static void main(String[] args) throws Exception{
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load an existing Excel file into the workbook
        workbook.loadFromMarkdown("data/markdownToExcel.md");

        //Save to Markdown document
        workbook.saveToFile("output/markdownToExcel.xlsx", FileFormat.Version2013);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
