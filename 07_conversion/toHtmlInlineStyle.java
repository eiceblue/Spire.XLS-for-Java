import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.HTMLOptions;

public class toHtmlInlineStyle {
    public static void main(String[] args) throws Exception{
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load an existing Excel file from the specified path
        workbook.loadFromFile("data/toHtmlInlineStyle.xlsx");

        // Get the first worksheet in the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Set the html options
        HTMLOptions options = new HTMLOptions();
        options.setImageEmbedded(true);

        //Setting "HTMLOptions.StyleDefineType.Inline"
        options.setStyleDefine(HTMLOptions.StyleDefineType.Inline);

        //Save to HTML document
        sheet.saveToHtml("output/toHtmlInlineStyle_result.html",options);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
