import com.spire.xls.*;

import java.io.FileInputStream;

public class setCustomFontStream {
    public static void main(String[] args) throws Exception{
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load an existing Excel file into the workbook
        workbook.loadFromFile("data/Sample.xlsx");

        // Create a FileInputStream to read the custom font file
        FileInputStream stream = new FileInputStream("data/PT_Serif-Caption-Web-Regular.ttf");

        // Set the custom font stream(s) to be used in the workbook
        workbook.setCustomFontStreams(new FileInputStream[]{stream});

        // Save the workbook to a PDF file with the specified font applied
        workbook.saveToFile("CustomFontStreams.pdf", FileFormat.PDF);

        // Release resources used by the workbook
        workbook.dispose();
    }
}
