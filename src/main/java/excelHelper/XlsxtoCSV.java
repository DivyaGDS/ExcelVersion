//Testing Xlsx commit from excel test

//package excelHelper;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.Iterator;
//
//import org.apache.commons.io.FilenameUtils;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class XlsxtoCSV {
//
//    public static void main(String[] args) throws IOException {
//        ExcelToCsv(args[0],args[1]);
//        System.out.println("Excel file is Converted to CSV !!");
//    }
//
//
//    /**
//     * This method is used to Decompress the Excel - Extract excel file to specified folder directory
//     *
//     * @param sourceExcel          : String : path to the excel file with extension
//     * @param destLocationFolder : String : path to the folder in which extracted excel file data to be stored
//     */
//
//    public static void ExcelToCsv(String sourceExcel, String destLocationFolder) throws IOException {
//        // For storing data into CSV files
//        StringBuffer data = new StringBuffer();
//
//        try {
//            FileOutputStream fos = new FileOutputStream(destLocationFolder);
//            // Get the workbook object for XLSX file
//            FileInputStream fis = new FileInputStream(sourceExcel);
//            Workbook workbook = null;
//
//            String ext = FilenameUtils.getExtension(sourceExcel.toString());
//
//            if (ext.equalsIgnoreCase("xlsx")) {
//                workbook = new XSSFWorkbook(fis);
//            } else if (ext.equalsIgnoreCase("xls")) {
//                workbook = new HSSFWorkbook(fis);
//            }
//
//            // Get first sheet from the workbook
//            int numberOfSheets = workbook.getNumberOfSheets();
//            Row row;
//            Cell cell;
//            // Iterate through each rows from first sheet
//            for (int i = 0; i < numberOfSheets; i++) {
//                data.append("Sheet Name :"+workbook.getSheetName(i)+"\n");
//                Sheet sheet = workbook.getSheetAt(i);
//                Iterator<Row> rowIterator = sheet.iterator();
//
//                for(Row row1 : sheet) {
//                    for(int cn=0; cn<row1.getLastCellNum(); cn++) {
//                        // If the cell is missing from the file, generate a blank one
//                        // (Works by specifying a MissingCellPolicy)
//                        Cell cell1 = row1.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//                    }
//                }
//
//                while (rowIterator.hasNext()) {
//                    row = rowIterator.next();
//                    // For each row, iterate through each columns
//                    Iterator<Cell> cellIterator = row.cellIterator();
//                    while (cellIterator.hasNext()) {
//                        cell = cellIterator.next();
//                        switch (cell.getCellType()) {
//                            case BOOLEAN:
//                                data.append(cell.getBooleanCellValue() + ",");
//                                break;
//                            case NUMERIC:
//                                data.append(cell.getNumericCellValue() + ",");
//                                break;
//                            case BLANK:
//                                data.append(" " + ",");
//                                break;
//                            case STRING:
//                                String testcell = cell.getStringCellValue();
//                                if (cell.getStringCellValue().contains(",")) {
//                                    testcell = cell.getStringCellValue().replaceAll(",", "COMMA");
//                                }
//                                data.append(testcell).append(",");
//                                break;
//                            default:
//                                data.append(cell + ",");
//                        }
//                    }
//                    data.append('\n'); // appending new line after each row
//                }
//
//            }
//            fos.write(data.toString().getBytes());
//            fos.close();
//
//        } catch (Exception ioe) {
//            ioe.printStackTrace();
//        }
//    }
//
//}
