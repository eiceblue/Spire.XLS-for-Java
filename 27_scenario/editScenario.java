package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.scenario.*;

import java.util.ArrayList;
import java.util.List;

public class editScenario {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook wb = new Workbook();

        // Load the existing Excel file from the specified path
        wb.loadFromFile("Data/ScenarioSample2.xlsx");

        // Get the first worksheet (index 0) from the workbook
        Worksheet worksheet = wb.getWorksheets().get(0);

        // Access the collection of scenarios in the worksheet
        XlsScenarioCollection scenarios = worksheet.getScenarios();

        // Retrieve the existing scenario named "Current Quantity" from the collection
        XlsScenario scenario1 = scenarios.get("Current Quantity");

        // Initialize a list of new values to replace the original scenario data
        List<Object> increasedQuantity_Values = new ArrayList<Object>() {{
            add(1000); add(5000); add(4500); add(3900); add(10000); add(8900);
        }};

        // Modify the scenario by updating its target cell range (D5:D10) with the new values
        scenario1.setVariableCells(worksheet.getRange().get("D5:D10"), increasedQuantity_Values);

        // Apply the updated scenario to the worksheet so the new values take effect
        scenario1.show();

        // Save the modified workbook to a new file in Excel 2016 format
        wb.saveToFile("editScenario.xlsx", ExcelVersion.Version2016);

        // Dispose of the workbook object to release file handles and free up memory resources
        wb.dispose();
    }
}
