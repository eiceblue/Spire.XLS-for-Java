package demoQ2;

import com.spire.xls.Workbook;
import com.spire.xls.core.spreadsheet.MarkdownOptions;

public class toMarkdownExportOptions {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new Workbook();

        //Load the document from disk
        workbook.loadFromFile("Data/ToMarkdownExportOptions.xlsx");

        // Create export options for Markdown format
        MarkdownOptions options = new MarkdownOptions();

        // Set whether to save images with relative paths
        options.isSavePicInRelativePath(true);

        // Set whether to save hyperlinks as Markdown reference format
        options.isSaveHyperlinkAsRef(true);

        // Save the workbook as Markdown with the specified options
        String result = "ToMarkdownExportOptions_out.md";
        workbook.saveToMarkdown(result, options);

        // Dispose of the workbook object to release resources
        workbook.dispose();
    }
}
