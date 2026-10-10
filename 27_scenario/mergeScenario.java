package demoQ2;

import com.spire.xls.*;

public class mergeScenario {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook wb = new Workbook();

        // Load the Excel file from the specified path
        wb.loadFromFile("Data/ScenarioSample2_withScenarios.xlsx");

        // Get the first worksheet (index 0) from the workbook
        Worksheet worksheet1 = wb.getWorksheets().get(0);

        // Get the second worksheet (index 1) from the workbook
        Worksheet worksheet2 = wb.getWorksheets().get(1);

        // Merge the scenarios from worksheet2 into worksheet1
        worksheet1.getScenarios().merge(worksheet2);

        // Save the merged workbook to a new file in Excel 2016 format
        wb.saveToFile("mergeScenario.xlsx", ExcelVersion.Version2016);

        // Release all resources used by the workbook
        wb.dispose();
    }
}
