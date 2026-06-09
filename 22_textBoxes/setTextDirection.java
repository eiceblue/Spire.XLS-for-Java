import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.conditionalformatting.TextDirectionType;
import com.spire.xls.core.spreadsheet.shapes.*;

public class setTextDirection {
    public static void main(String[] args)throws Exception {
        // Create a new Workbook object
        Workbook workbook = new Workbook();

        // Get the first worksheet in the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Add a TextBox shape to the worksheet at position (4, 2) with width 100 and height 300
        XlsTextBoxShape textbox = (XlsTextBoxShape) sheet.getTextBoxes().addTextBox(4, 2, 100, 300);

        // Set the text content of the TextBox
        textbox.setText("مبرمج , اختبار .");

        // Set the horizontal alignment of the TextBox to Left
        textbox.setHAlignment(CommentHAlignType.Left);

        // Set the inner left margin of the TextBox to 1 point
        textbox.setInnerLeftMargin(1);

        // Set the inner right margin of the TextBox to 3 points
        textbox.setInnerRightMargin(3);

        // Set the inner top margin of the TextBox to 1 point
        textbox.setInnerTopMargin(1);

        // Set the inner bottom margin of the TextBox to 1 point
        textbox.setInnerBottomMargin(1);

        // Set the text direction of the textbox to Right-to-Left
        textbox.setTextDirection(TextDirectionType.RightToLeft);

        // Specify the path for the output file
        String result = "output/setTextDirection_result.xlsx";
        workbook.saveToFile(result, ExcelVersion.Version2013);
    }
}
