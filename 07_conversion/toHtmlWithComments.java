import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.HTMLOptions;

public class toHtmlWithComments {
    public static void main(String[] args) {
        String inputFile = "data/ToHtmlWithComment.xlsx";
        String outputFile = "output/ToHtmlWithComment_out.html";

        //Create a workbook
        Workbook workbook = new Workbook();

        //Load an excel document
        workbook.loadFromFile(inputFile);

        //Get the first sheet
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Keep comments
        HTMLOptions options = new HTMLOptions();
        options.isSaveComment(true);

        //Save to HTML file
        sheet.saveToHtml(outputFile,options);
    }
}
