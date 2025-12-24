import com.spire.xls.*;


public class getCellOfEmbeddedImage {
    public static void main(String[] args) {
        // Create a new Workbook instance
        Workbook workbook = new Workbook();

        // Load an existing Excel file into the workbook
        workbook.loadFromFile("data/EmbeddedImage.xlsx");

        // Get the first worksheet from the workbook's worksheets collection
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Retrieve an array of embedded images from the worksheet's cells
        ExcelPicture[] cellImages = worksheet.getCellImages();

        // Print the name (address) of the cell that contains the first embedded image
        System.out.println(cellImages[0].getEmbedCellName());
    }
}
