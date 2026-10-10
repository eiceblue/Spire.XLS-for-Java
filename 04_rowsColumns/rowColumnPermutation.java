package demoQ2;

import com.spire.xls.*;
import java.util.*;

public class rowColumnPermutation {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook workbook = new Workbook();
        // Load the source Excel file containing the original data
        workbook.loadFromFile("Data/rowColumnPermutation.xlsx");
        // Get the first worksheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Set copy options: Transpose the data and copy all formatting/content
        EnumSet<CopyRangeOptions> options = EnumSet.of(CopyRangeOptions.Transpose, CopyRangeOptions.All);
        // Add a descriptive label for the first transpose operation
        sheet.getRange().get("E1").setText("Transpose A1:C4 to E2:H3：");
        // Transpose range A1:C4 (4 rows x 3 cols) to E2:H3 (3 rows x 4 cols)
        sheet.getRange().get("A1:C4").copy(sheet.getRange().get("E2:H3"), options);
        // Add a descriptive label for the second transpose operation
        sheet.getRange().get("E6").setText("Transpose A1:B5 to E7：");
        // Transpose range A1:B5 (5 rows x 2 cols) to E7:G8 (2 rows x 5 cols)
        sheet.getRange().get("A1:B5").copy(sheet.getRange().get("E7"), options);

        // Update options: Transpose the data and copy only the cell styles
        options = EnumSet.of(CopyRangeOptions.Transpose, CopyRangeOptions.CopyStyles);
        // Add a descriptive label
        sheet.getRange().get("D11").setText("Transpose A11:B11 to D12：");
        // Transpose the horizontal range A11:B11 (1 row x 2 cols) to vertical range D12:D13 (2 rows x 1 col)
        sheet.getRange().get("A11:B11").copy(sheet.getRange().get("D12"), options);

        // Update options: Transpose the data and copy only the conditional formats
        options = EnumSet.of(CopyRangeOptions.Transpose, CopyRangeOptions.CopyConditionalFormats);
        // Add a descriptive label
        sheet.getRange().get("D16").setText("Transpose B17:B20 to D17：");
        // Transpose the vertical range B17:B20 (4 rows x 1 col) to horizontal range D17:G17 (1 row x 4 cols)
        sheet.getRange().get("B17:B20").copy(sheet.getRange().get("D17"), options);

        // Update options: Transpose the data and copy only the cell styles
        options = EnumSet.of(CopyRangeOptions.Transpose, CopyRangeOptions.CopyStyles);
        // Add a descriptive label
        sheet.getRange().get("D19").setText("Transpose B17:B20 to D20：");
        // Transpose the same vertical range B17:B20 to another horizontal position D20:G20
        sheet.getRange().get("B17:B20").copy(sheet.getRange().get("D20"), options);

        // Update options: Transpose the data only, without copying any formatting
        options = EnumSet.of(CopyRangeOptions.Transpose);
        // Add a descriptive label
        sheet.getRange().get("D22").setText("Transpose B17:B20 to D23：");
        // Transpose the vertical range B17:B20 to horizontal range D23:G23 with values only
        sheet.getRange().get("B17:B20").copy(sheet.getRange().get("D23"), options);

        // Save the modified workbook to a new file
        workbook.saveToFile("rowColumnPermutation_result.xlsx");
        // Release resources and dispose of the workbook object
        workbook.dispose();
    }
}
