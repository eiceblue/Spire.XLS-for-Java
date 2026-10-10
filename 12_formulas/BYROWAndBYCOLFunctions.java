package demoQ2;

import com.spire.xls.*;

public class BYROWAndBYCOLFunctions {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Load an existing Excel file from the specified path
        workbook.loadFromFile("Data/BYROWAndBYCOLFunctions.xlsx");

        // Get the first sheet from the workbook
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Calculate average score for each row (Columns B to F) using BYROW function
        sheet.getCellRange("G2").setFormula("=BYROW(B2:F2, LAMBDA(row, AVERAGE(row)))");
        sheet.getCellRange("G3").setFormula("=BYROW(B3:F3, LAMBDA(row, AVERAGE(row)))");
        sheet.getCellRange("G4").setFormula("=BYROW(B4:F4, LAMBDA(row, AVERAGE(row)))");
        sheet.getCellRange("G5").setFormula("=BYROW(B5:F5, LAMBDA(row, AVERAGE(row)))");
        sheet.getCellRange("G6").setFormula("=BYROW(B6:F6, LAMBDA(row, AVERAGE(row)))");
        sheet.getCellRange("G7").setFormula("=BYROW(B7:F7, LAMBDA(row, AVERAGE(row)))");

        // Calculate average for each column (subject) using BYCOL function
        sheet.getCellRange("B8").setFormula("=BYCOL(B2:B7, LAMBDA(col, AVERAGE(col)))");
        sheet.getCellRange("C8").setFormula("=BYCOL(C2:C7, LAMBDA(col, AVERAGE(col)))");
        sheet.getCellRange("D8").setFormula("=BYCOL(D2:D7, LAMBDA(col, AVERAGE(col)))");
        sheet.getCellRange("E8").setFormula("=BYCOL(E2:E7, LAMBDA(col, AVERAGE(col)))");
        sheet.getCellRange("F8").setFormula("=BYCOL(F2:F7, LAMBDA(col, AVERAGE(col)))");
        sheet.getCellRange("G8").setFormula("=BYCOL(G2:G7, LAMBDA(col, AVERAGE(col)))");

        // Combined BYROW and BYCOL usage examples - Comprehensive Statistics section
        sheet.getCellRange("I1").setValue("Comprehensive Statistics");

        // Directly get max value
        sheet.getCellRange("I3").setValue("Highest Average");
        sheet.getCellRange("J3").setFormula("=MAX(BYROW(G2:G7, LAMBDA(row, row)))");

        // Directly get min value
        sheet.getCellRange("I4").setValue("Lowest Average");
        sheet.getCellRange("J4").setFormula("=MIN(BYROW(G2:G7, LAMBDA(row, row)))");

        // Calculate overall average across all subjects and students
        sheet.getCellRange("I5").setValue("Overall Subject Average");
        sheet.getCellRange("J5").setFormula("=BYCOL(B2:F7, LAMBDA(col, AVERAGE(col)))");

        // Nested scenario: BYROW nested with BYCOL
        // This formula doubles each value in the range B2:E7 and sums them by row
        sheet.getCellRange("I7").setFormula("=BYROW(B2:E7, LAMBDA(row, SUM(BYCOL(row, LAMBDA(col, col*2)))))");

        // Recalculate all formulas to ensure values are up to date
        sheet.CalculateAllValue();

        // Save the modified workbook to the specified file using Excel 2010 format
        String result = "BYROWAndBYCOLFunctions_out.xlsx";
        workbook.saveToFile(result, ExcelVersion.Version2010);

        // Dispose of the workbook object to release resources
        workbook.dispose();
    }
}
