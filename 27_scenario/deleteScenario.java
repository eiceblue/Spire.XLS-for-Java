package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.scenario.XlsScenarioCollection;

public class deleteScenario {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook wb = new Workbook();

        // Load the existing Excel file containing scenarios
        wb.loadFromFile("Data/ScenarioSample2.xlsx");

        // Get the first worksheet (index 0) from the workbook
        Worksheet worksheet = wb.getWorksheets().get(0);

        // Access the collection of scenarios in the worksheet
        XlsScenarioCollection scenarios = worksheet.getScenarios();

        // Delete the first scenario in the collection by its index (0)
        scenarios.removeScenarioAt(0);

        // Delete a specific scenario by its name
        scenarios.removeScenarioByName("Current Quantity");

        // Initialize an empty string to store the verification output
        String content = "";

        // Append the total count of remaining scenarios to the output string
        content += "Count:" + scenarios.getCount() + "\n";

        // Check if the deleted scenario "Current Quantity" still exists (expected: false) and append the result
        content += "ContainsScenario:" + scenarios.containsScenario("Current Quantity") + "\n";

        // Check if the scenario "Decreased Quantity" exists (expected: true) and append the result
        content += "ContainsScenario:" + scenarios.containsScenario("Decreased Quantity") + "\n";

        // Print the verification results to the console
        System.out.println(content);

        // Save the modified workbook to a new file in Excel 2016 format
        wb.saveToFile("deleteScenario.xlsx", ExcelVersion.Version2016);

        // Dispose of the workbook object to release file handles and free up memory resources
        wb.dispose();
    }
}
