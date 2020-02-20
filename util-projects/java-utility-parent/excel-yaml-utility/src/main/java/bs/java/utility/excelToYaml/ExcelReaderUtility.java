package bs.java.utility.excelToYaml;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Scanner;


/*
https://www.baeldung.com/java-snake-yaml
https://www.logicbig.com/tutorials/misc/yaml/java-to-yaml.html

 */
public class ExcelReaderUtility {

    public static final String FilePath = "C:\\excelToYaml\\";
    public static final String ExcelFileName = "excelFile.xlsx";
    public static final String ExcelToYamlFileName = "excelToYaml.yaml";
    public static final String yamlFileName = "yamlFile.yaml";
    public static final String yamlToExcelFileName = "yamlToExcel.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        String fullExcelFileName = FilePath + ExcelFileName;
        String fullExcelToYamlFileName = FilePath + ExcelToYamlFileName;
        String fullYamlFileName = FilePath + yamlFileName;
        String fullYamlToExcelFileName = FilePath + yamlToExcelFileName;

        Scanner scanner = new Scanner(System.in);

        boolean loopin = true;

        while (loopin){
            System.out.println("Read Excel File and Convert to Yaml: 1 ");
            System.out.println("Read Yaml File and Convert to Excel: 2 ");
            System.out.println("Exit: 3 ");
            System.out.print("Enter the Option: ");
            String userOption = scanner.next();

            if(userOption.equalsIgnoreCase("1")){
                loopin = false;
                processExcelToYamlFile(fullExcelFileName, fullExcelToYamlFileName);
            }else if(userOption.equalsIgnoreCase("2")){
                loopin = false;
                processYamlToExcelFile(fullYamlFileName, fullYamlToExcelFileName);
            } else if(userOption.equalsIgnoreCase("3")){
                loopin = false;
                System.exit(0);
            } else{
                System.out.println("Invalid Option!!!! ");

            }
        }




    }

    public static void processExcelToYamlFile(String fullExcelFileName, String fullYamlFileName) throws IOException, InvalidFormatException{

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

            constructYamlMapFromExcelRow(row, fieldNameMap, fieldPropertiesMap, dataMap, rowIndex);

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

    public static void processYamlToExcelFile(String fullYamlFileName, String fullYamlToExcelFileName) throws IOException {

        Map<String, Map<String, String>>  yamlData =  readYamlFile(fullYamlFileName);
        createExcelFileFromYamlData(fullYamlToExcelFileName, yamlData );
    }

    public static Map<Integer, String> createFieldsMap (){
        Map<Integer, String> fieldsMap = new LinkedHashMap<>();
        fieldsMap.put(1, "Field Name");
        fieldsMap.put(2, "type");
        fieldsMap.put(3, "description");
        fieldsMap.put(4, "format");
        fieldsMap.put(5, "minLength");
        fieldsMap.put(6, "maxLength");
        fieldsMap.put(7, "example");
        fieldsMap.put(8, "required");
        fieldsMap.put(9, "$ref");
        return fieldsMap;
    }

    public static Map<String, Integer> createFieldsIntegerMap (){
        Map<String, Integer> fieldsIntegerMap = new LinkedHashMap<>();
        fieldsIntegerMap.put( "Field Name",1);
        fieldsIntegerMap.put( "type",2);
        fieldsIntegerMap.put( "description",3);
        fieldsIntegerMap.put( "format",4);
        fieldsIntegerMap.put( "minLength",5);
        fieldsIntegerMap.put( "maxLength",6);
        fieldsIntegerMap.put( "example",7);
        fieldsIntegerMap.put( "required",8);
        fieldsIntegerMap.put( "$ref",9);
        return fieldsIntegerMap;
    }
    private static int rowIndex = 0, colIndex = 0;
    private static void createExcelFileFromYamlData(String fullYamlToExcelFileName, Map<String, Map<String, String>> yamlData) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Yaml To Excel");
        System.out.println("createExcelFileFromYamlData ");
        Map<Integer, String> fieldsMap = createFieldsMap();
        Map<String, Integer> fieldsIntegerMap = createFieldsIntegerMap ();
        rowIndex = 0;
        colIndex = 0;
        Row rowheader = sheet.createRow(rowIndex++);
        System.out.println("Headers Map ");
        System.out.println("**************");
        Cell cellFirstCell = rowheader.createCell(0);
        cellFirstCell.setCellValue("S.NO");
        fieldsMap.forEach((k, v) -> {
            System.out.println("Key = " + k + ", Value = " + v);
            Cell cellHeader = rowheader.createCell(++colIndex);
            cellHeader.setCellValue((String) v);
        });
        System.out.println("**************");
        colIndex = 0;
        System.out.println("Data");

        yamlData.forEach((k,v) -> {
            Row rowfields = sheet.createRow(rowIndex);
            System.out.println("Key = " + k + ", Value = " + v);
            System.out.println("**************");
            Cell cellFieldName = rowfields.createCell(0);
            cellFieldName.setCellValue(rowIndex );
            cellFieldName = rowfields.createCell(1);
            cellFieldName.setCellValue((String) k);
            v.forEach((key,value) -> {
                System.out.println("Key = " + key + ", Value = " + value);
                int fieldColPosition = fieldsIntegerMap.get(key);
                Cell cell = rowfields.createCell(fieldColPosition);
                cell.setCellValue((String) value);
                colIndex++;
            });
            System.out.println("**************");
            rowIndex++;
        });


        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(fullYamlToExcelFileName);
            workbook.write(outputStream);
        } catch (Exception e) {
            System.out.println("**************");
            System.out.println(e.getMessage() + " " + e.getCause());
            e.printStackTrace();
        }

    }


    public  static Map<String, Map<String, String>> readYamlFile (String fullSourceYamlFIleName) throws IOException {
        Yaml yaml = new Yaml();
        FileInputStream fileInputStream = readFromFile(fullSourceYamlFIleName);
        Map<String, Map<String, String>> YamlData = yaml.load(fileInputStream);
        System.out.println("Reading YAML File");
        System.out.println(YamlData);
        fileInputStream.close();
        return YamlData;
    }

    public static void writeToFile(String fileName, String fileContent)
            throws IOException {
        FileOutputStream outputStream = new FileOutputStream(fileName);
        byte[] strToBytes = fileContent.getBytes();
        outputStream.write(strToBytes);

        outputStream.close();
    }

    public static FileInputStream readFromFile(String fileName)
            throws IOException {
        FileInputStream fileInputStream = new FileInputStream(fileName);


        return fileInputStream;
    }
    private static  int  colStaticCount = 0;
    public static void constructYamlMapFromExcelRow(Row row,
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
        if(rowIndex==0) {
            colStaticCount = row.getPhysicalNumberOfCells();
        }
        System.out.println("colStaticCount=" + colStaticCount);
        for(int colIndex=1; colIndex<colStaticCount ; colIndex++){
            cell = row.getCell(colIndex);
            System.out.print("cell.getStringCellValue=" + cell + "\t");

            String cellValue = dataFormatter.formatCellValue(cell);
            System.out.print(colIndex + "-" + cellValue + "\t");
            if(null==cellValue || cellValue.equals("")){
                continue;
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