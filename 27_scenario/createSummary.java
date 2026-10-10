package demoQ2;

import com.spire.xls.*;
import com.spire.xls.core.spreadsheet.scenario.*;
import java.util.*;

public class createSummary {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook wb = new Workbook();

        // Load the Excel file from the specified path
        wb.loadFromFile("Data/ScenarioSample1.xlsx");

        // Get the first worksheet (index 0) from the workbook
        Worksheet worksheet = wb.getWorksheets().get(0);

        // Access the collection of scenarios in the worksheet
        XlsScenarioCollection scenarios = worksheet.getScenarios();

        // Initialize a list of values representing the current change percentages for each scenario row
        List<Object> currentChangePercentage_Values = new ArrayList<Object>() {{ add(0.23); add(0.8); add(1.1); add(0.5); add(0.35); add(0.2); }};

        // Initialize a list of values representing the increased change percentages for each scenario row
        List<Object> increasedChangePercentage_Values = new ArrayList<Object>() {{ add(0.45); add(0.56); add(0.9); add(0.5); add(0.58); add(0.43); }};

        // Initialize a list of values representing the decreased change percentages for each scenario row
        List<Object> decreasedChangePercentage_Values = new ArrayList<Object>() {{ add(0.3); add(0.2); add(0.5); add(0.3); add(0.5); add(0.23); }};

        // Initialize a list of values representing the current quantity for each scenario row
        List<Object> currentQuantity_Values = new ArrayList<Object>() {{ add(1500); add(3000); add(5000); add(4000); add(500); add(4000); }};

        // Initialize a list of values representing the increased quantity for each scenario row
        List<Object> increasedQuantity_Values = new ArrayList<Object>() {{ add(1000); add(5000); add(4500); add(3900); add(10000); add(8900); }};

        // Initialize a list of values representing the decreased quantity for each scenario row
        List<Object> decreasedQuantity_Values = new ArrayList<Object>() {{ add(1000); add(2000); add(3000); add(3000); add(300); add(4000); }};

        // Add a scenario named "Current % of Change" applying current percentage values to cells F5:F10
        scenarios.add("Current % of Change", worksheet.getRange().get("F5:F10"), currentChangePercentage_Values);

        // Add a scenario named "Increased % of Change" applying increased percentage values to cells F5:F10
        scenarios.add("Increased % of Change", worksheet.getRange().get("F5:F10"), increasedChangePercentage_Values);

        // Add a scenario named "Decreased % of Change" applying decreased percentage values to cells F5:F10
        scenarios.add("Decreased % of Change", worksheet.getRange().get("F5:F10"), decreasedChangePercentage_Values);

        // Add a scenario named "Current Quantity" applying current quantity values to cells D5:D10
        scenarios.add("Current Quantity", worksheet.getRange().get("D5:D10"), currentQuantity_Values);

        // Add a scenario named "Increased Quantity" applying increased quantity values to cells D5:D10
        scenarios.add("Increased Quantity", worksheet.getRange().get("D5:D10"), increasedQuantity_Values);

        // Add a scenario named "Decreased Quantity" applying decreased quantity values to cells D5:D10
        scenarios.add("Decreased Quantity", worksheet.getRange().get("D5:D10"), decreasedQuantity_Values);

        // Generate a scenario summary report and place it starting at cell L7
        worksheet.getScenarios().summary(worksheet.getRange().get("L7"));

        // Save the workbook to a new file in Excel 2016 format
        wb.saveToFile("createSummary.xlsx", ExcelVersion.Version2016);

        // Release all resources held by the workbook object
        wb.dispose();
    }
}
