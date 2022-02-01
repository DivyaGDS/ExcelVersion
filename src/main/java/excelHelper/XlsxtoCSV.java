package excelHelper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxtoCSV {

    static void xlsx(File inputFile, File outputFile) {
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();

        try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLSX file
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = null;

            String ext = FilenameUtils.getExtension(inputFile.toString());

            if (ext.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (ext.equalsIgnoreCase("xls")) {
                workbook = new HSSFWorkbook(fis);
            }

            // Get first sheet from the workbook

            int numberOfSheets = workbook.getNumberOfSheets();
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {
                    row = rowIterator.next();
                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {

                        cell = cellIterator.next();

                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                data.append(cell.getBooleanCellValue() + ",");

                                break;
                            case NUMERIC:
                                data.append(cell.getNumericCellValue() + ",");

                                break;
                            case STRING:
                                data.append(cell.getStringCellValue() + ",");
                                break;

                            case BLANK:
                                data.append("" + ",");
                                break;
                            default:
                                data.append(cell + ",");

                        }
                    }
                    data.append('\n'); // appending new line after each row
                }

            }
            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    // testing the application

    public static void main(String[] args) {
        // int i=0;
        // reading file from desktop
        File inputFile = new File("C:\\Users\\UB217ZA\\Downloads\\Github Clone NGTP\\ExcelVersion\\src\\test\\resources\\testdata\\pageobjects\\PageObjects.xlsx"); //provide your path
        // writing excel data to csv
        File outputFile = new File("C:\\Users\\UB217ZA\\Downloads\\Github Clone NGTP\\ExcelVersion\\src\\test\\resources\\testdata\\pageobjects\\PageObjectsConverted.csv");  //provide your path
        xlsx(inputFile, outputFile);
        System.out.println("Conversion of " + inputFile + " to flat file: "
                + outputFile + " is completed");
    }
}
