package bs.java.utility.excelToYaml;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {

    public static final String FilePath = "C:\\Storage\\workspace\\java\\mycode\\java-util-projects\\util-projects\\java-utility-parent\\excel-yaml-utility\\src\\main\\resources\\";
    public static final String FileName = "sample.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(FilePath+ FileName));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
      /*  Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }*/

        // 2. Or you can use a for-each loop
      /*  System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }*/



        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
/*        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }*/

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();



        }

        printData(sheet);





        // 3. Or you can use Java 8 forEach loop with lambda
       /* System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });*/

        // Closing the workbook
        workbook.close();
    }

    public static void printData (Sheet sheet){
        DataFormatter dataFormatter = new DataFormatter();

        int rowCount = sheet.getPhysicalNumberOfRows();
        int rowIndex = 0;
        int colCount = 0;
        Row row = null;
        Cell cell = null;
        for(int rowindex = 0; rowIndex<rowCount; rowIndex ++){
            row = sheet.getRow(rowIndex);
            colCount = row.getPhysicalNumberOfCells();
            System.out.println(rowIndex);
            for(int colIndex=0; colIndex<colCount ; colIndex++){
                cell = row.getCell(colIndex);
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(colIndex + "-" + cellValue + "\t");
            }
        }
    }

}