package bs.java.utility.excelToYaml;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.Yaml;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;


/*
https://www.baeldung.com/java-snake-yaml
https://www.logicbig.com/tutorials/misc/yaml/java-to-yaml.html

 */
public class ExcelReaderUtility {

    public static final String FilePath = "C:\\excelToYaml\\";
    public static final String FileName = "excelFile.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        String fullExcelFileName = FilePath + FileName;
        String fullYamlFileName = FilePath + "excelToYaml.yaml";
        if(args!=null && args.length>2) {
            if (args[0] != null && !args[0].equals("")) {
                fullExcelFileName = args[0];
            }
            if (args[1] != null && !args[1].equals("")) {
                fullYamlFileName = args[1];
            }
        }

        processExcelFile(fullExcelFileName, fullYamlFileName);


    }

    public static void processExcelFile( String fullExcelFileName, String fullYamlFileName) throws IOException, InvalidFormatException{

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(fullExcelFileName));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, Map<String, String>> dataMap = new LinkedHashMap<>();
        Map<Integer, String> fieldNameMap = new LinkedHashMap<>();
        Map<String, String> fieldPropertiesMap = new LinkedHashMap<>();

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        int rowIndex = 0, colIndex = 0;
        for (Row row: sheet) {

            constructYamlMap(row, fieldNameMap, fieldPropertiesMap, dataMap, rowIndex);

            rowIndex++;
        }

        System.out.println("fieldNameMap " + fieldNameMap.toString());
        System.out.println("fieldPropertiesMap " + fieldPropertiesMap.toString());
        System.out.println("dataMap " + dataMap.toString());

        DumperOptions options = new DumperOptions();
        options.setDefaultFlowStyle(DumperOptions.FlowStyle.BLOCK);
        options.setPrettyFlow(true);

        Yaml yaml = new Yaml(options);
        //StringWriter writer = new StringWriter();
        //yaml.dump(dataMap, writer);
        String output = yaml.dump(dataMap);
        System.out.println("**************************");
        System.out.println();
        System.out.println(output);

        writeToFile(fullYamlFileName, output);
        // Closing the workbook
        workbook.close();
    }

    public static void writeToFile(String fileName, String fileContent)
            throws IOException {
        FileOutputStream outputStream = new FileOutputStream(fileName);
        byte[] strToBytes = fileContent.getBytes();
        outputStream.write(strToBytes);

        outputStream.close();
    }

    public static void constructYamlMap (Row row,
                                     Map<Integer, String> fieldNameMap,
                                     Map<String, String> fieldPropertiesMap,
                                     Map<String, Map<String, String>> dataMap,
                                     int rowIndex) {

        fieldPropertiesMap = new LinkedHashMap<>();
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        String propertyName = "";
        System.out.println();

        Cell cell = null;
        int  colCount = row.getPhysicalNumberOfCells();
        System.out.println(rowIndex);
        for(int colIndex=1; colIndex<colCount ; colIndex++){
            cell = row.getCell(colIndex);
            String cellValue = dataFormatter.formatCellValue(cell);
            //System.out.print(colIndex + "-" + cellValue + "\t");
            if(null==cellValue || cellValue.equals("")){
                break;
            }
            if(rowIndex==0){
                    fieldNameMap.put(colIndex, cellValue);
            }else{

                if(colIndex==1){
                    propertyName = cellValue;
                } else {
                    fieldPropertiesMap.put(fieldNameMap.get(colIndex), cellValue);
                }
            }
        }
        if(rowIndex!=0) {
            dataMap.put(propertyName, fieldPropertiesMap);
        }
    }
}