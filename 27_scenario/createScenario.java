package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.scenario.XlsScenarioCollection;
import java.util.*;

public class createScenario {
    public static void main(String[] args) {
        // Create a new Workbook instance to handle the Excel file
        Workbook wb = new Workbook();

        // Load the existing Excel file from the specified path
        wb.loadFromFile("Data/ScenarioSample1.xlsx");

        // Get the first worksheet (index 0) from the workbook
        Worksheet worksheet = wb.getWorksheets().get(0);

        // Access the collection of scenarios in the worksheet
        XlsScenarioCollection scenarios = worksheet.getScenarios();

        // Initialize list objects with different values for scenarios
        // These lists represent the data that will be applied to specific cell ranges under different scenarios

        // Values for "Current % of Change" scenario
        List<Object> currentChangePercentage_Values = new ArrayList<Object>() {{
            add(0.23); add(0.8); add(1.1); add(0.5); add(0.35); add(0.2);
        }};

        // Values for "Increased % of Change" scenario
        List<Object> increasedChangePercentage_Values = new ArrayList<Object>() {{
            add(0.45); add(0.56); add(0.9); add(0.5); add(0.58); add(0.43);
        }};

        // Values for "Decreased % of Change" scenario
        List<Object> decreasedChangePercentage_Values = new ArrayList<Object>() {{
            add(0.3); add(0.2); add(0.5); add(0.3); add(0.5); add(0.23);
        }};

        // Values for "Current Quantity" scenario
        List<Object> currentQuantity_Values = new ArrayList<Object>() {{
            add(1500); add(3000); add(5000); add(4000); add(500); add(4000);
        }};

        // Values for "Increased Quantity" scenario
        List<Object> increasedQuantity_Values = new ArrayList<Object>() {{
            add(1000); add(5000); add(4500); add(3900); add(10000); add(8900);
        }};

        // Values for "Decreased Quantity" scenario
        List<Object> decreasedQuantity_Values = new ArrayList<Object>() {{
            add(1000); add(2000); add(3000); add(3000); add(300); add(4000);
        }};

        // Add scenarios to the worksheet, each targeting a specific cell range with its corresponding set of values
        // The first parameter is the scenario name, the second is the target cell range, and the third is the list of values

        // Add percentage change scenarios to range F5:F10
        scenarios.add("Current % of Change", worksheet.getRange().get("F5:F10"), currentChangePercentage_Values);
        scenarios.add("Increased % of Change", worksheet.getRange().get("F5:F10"), increasedChangePercentage_Values);
        scenarios.add("Decreased % of Change", worksheet.getRange().get("F5:F10"), decreasedChangePercentage_Values);

        // Add quantity scenarios to range D5:D10
        scenarios.add("Current Quantity", worksheet.getRange().get("D5:D10"), currentQuantity_Values);
        scenarios.add("Increased Quantity", worksheet.getRange().get("D5:D10"), increasedQuantity_Values);
        scenarios.add("Decreased Quantity", worksheet.getRange().get("D5:D10"), decreasedQuantity_Values);

        // Save the modified workbook to a new file in Excel 2013 format
        wb.saveToFile("createScenario.xlsx", ExcelVersion.Version2016);

        // Dispose of the workbook object to release file handles and free up memory resources
        wb.dispose();
    }
}
