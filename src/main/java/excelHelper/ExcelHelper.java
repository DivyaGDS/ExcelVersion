package excelHelper;

//import org.apache.commons.io.FilenameUtils;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class ExcelHelper {
    public static final byte[] BUFFEBYTESIZE = new byte[1024];
    public static final int BUFFERSIZE = 2048;

    public static void main(String[] args) throws IOException {

        switch(args[0]) {
            case "Compression":
                createExcelFile(args[1],args[2]);
                System.out.println("Excel file is Compressed !!");
                break;
            case "Decompression":
                extractExcelFile(args[1],args[2]);
                System.out.println("Excel file is Decompressed !!");
                break;
//            case "ConvertToCSV":
//                ExcelToCsv(args[1],args[2]);
//                System.out.println("Excel file is Converted To Csv !!");
//                break;
//            case "ConvertToExcel":
//                CSVToExcel(args[1],args[2]);
//                System.out.println("CSV file is Converted To Excel !!");
//                break;
            
        }
    }


    /**
     * This method is used to Compress the Excel - Create excel file from the decompressed XML files
     *
     * @param sourceFolderLocation          : String : path to the excel file with extension
     * @param targetExcelFileName : String : path to the folder in which extracted excel file data to be stored
     */
    public static void createExcelFile(String sourceFolderLocation, String targetExcelFileName) throws IOException {

        Path source = Paths.get(sourceFolderLocation);

        try (
                ZipOutputStream zos = new ZipOutputStream(
                        new FileOutputStream(targetExcelFileName))
        ) {
            Files.walkFileTree(source, new SimpleFileVisitor<>() {
                @Override
                public FileVisitResult visitFile(Path file,
                                                 BasicFileAttributes attributes) {
                    if (attributes.isSymbolicLink()) {
                        return FileVisitResult.CONTINUE;
                    }

                    try (FileInputStream fis = new FileInputStream(file.toFile())) {
                        Path targetFile = source.relativize(file);
                        zos.putNextEntry(new ZipEntry(targetFile.toString()));

                        int len;
                        while (( len = fis.read(BUFFEBYTESIZE)) > 0) {
                            zos.write(BUFFEBYTESIZE, 0, len);
                        }
                        zos.closeEntry();
//                        LOGGER.info("Zip file processed: {}", file);

                    } catch (IOException e) {
//                        LOGGER.error(e.getMessage());
                    }
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult visitFileFailed(Path file, IOException exc) {
//                    LOGGER.error("Unable to zip : {}, {}}", file, exc);
                    return FileVisitResult.CONTINUE;
                }
            });

        }
    }



    /**
     * This method is used to Decompress the Excel - Extract excel file to specified folder directory
     *
     * @param sourceExcel          : String : path to the excel file with extension
     * @param destLocationFolder : String : path to the folder in which extracted excel file data to be stored
     */
    public static void extractExcelFile(String sourceExcel, String destLocationFolder) throws IOException {
        File file = new File(sourceExcel);
//        Logger LOGGER = LoggerFactory.getLogger("StepImplementation jar Initialization");

        try (ZipFile zip = new ZipFile(file)) {
            new File(destLocationFolder).mkdir();
            Enumeration<? extends ZipEntry> zipFileEntries = zip.entries();

            // Process each entry
            while (zipFileEntries.hasMoreElements()) {
                // grab a zip file entry
                ZipEntry entry = zipFileEntries.nextElement();
                String currentEntry = entry.getName();
                File destFile = new File(destLocationFolder, currentEntry);
                File destinationParent = destFile.getParentFile();

                // create the parent directory structure if needed
                destinationParent.mkdirs();
                if (!entry.isDirectory()) {
//                    LOGGER.info("destFile file created from extracted excel file: {}", destFile);
                    BufferedInputStream is = new BufferedInputStream(zip.getInputStream(entry));
                    int currentByte;
                    // establish buffer for writing file
                    byte[] data = new byte[BUFFERSIZE];

                    // write the current file to disk
                    FileOutputStream fos = new FileOutputStream(destFile);
                    try (BufferedOutputStream dest = new BufferedOutputStream(fos, BUFFERSIZE)) {

                        // read and write until last byte is encountered
                        while ((currentByte = is.read(data, 0, BUFFERSIZE)) != -1) {
                            dest.write(data, 0, currentByte);
                        }
                        dest.flush();
                        is.close();
                    }
                }
            }
        }
    }


    /**
     * This method is used to Decompress the Excel - Extract excel file to specified folder directory
     *
     * @param sourceExcel          : String : path to the excel file with extension
     * @param destLocationFolder : String : path to the folder in which extracted excel file data to be stored
     */

//    public static void ExcelToCsv(String sourceExcel, String destLocationFolder) throws IOException {
//        // For storing data into CSV files
//        StringBuffer data = new StringBuffer();
//
//        try {
//            FileOutputStream fos = new FileOutputStream(destLocationFolder);
//            // Get the workbook object for XLSX file
//            FileInputStream fis = new FileInputStream(sourceExcel);
//            XSSFWorkbook workbook = new XSSFWorkbook(fis);
//
//            String ext = FilenameUtils.getExtension(sourceExcel.toString());
//
////            if (ext.equalsIgnoreCase("xlsx")) {
////                workbook = new XSSFWorkbook(fis);
////            } else if (ext.equalsIgnoreCase("xls")) {
////                workbook = new HSSFWorkbook(fis);
////            }
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
//    /**
//     * This method is used to Decompress the Excel - Extract excel file to specified folder directory
//     *
//     * @param sourceExcel        : String : path to the excel file with extension
//     * @param destLocationFolder : String : path to the folder in which extracted excel file data to be stored
//     */
//
//    public static void CSVToExcel(String sourceExcel, String destLocationFolder) throws IOException {
//        FileOutputStream fos = new FileOutputStream(destLocationFolder);
//        String thisLine;
//        BufferedReader myInput = new BufferedReader(new FileReader(sourceExcel));
//        int i = 0;
//        XSSFWorkbook workbook = new XSSFWorkbook();
//
////        if (excelFileType.equalsIgnoreCase("xlsx")) {
////            workbook = new XSSFWorkbook();
////        } else if (excelFileType.equalsIgnoreCase("xls")) {
////            workbook = new HSSFWorkbook();
////        }
//
////        XSSFWorkbook workbook = new XSSFWorkbook();
//        Sheet sheet = null;
//        ArrayList al = null;
//        ArrayList arList = null;
//
//
//        while ((thisLine = myInput.readLine()) != null) {
//            if (thisLine.contains("Sheet Name :")) {
//                String sheetName = thisLine.split("Sheet Name :")[1];
//                sheet = workbook.createSheet(sheetName);
//                arList = new ArrayList();
//            } else {
//                al = new ArrayList();
//                String strar[] = thisLine.split(",");
//                for (int j = 0; j < strar.length; j++) {
//                    al.add(strar[j]);
//                }
//                arList.add(al);
//
//                for (int k = 0; k < arList.size(); k++) {
//                    ArrayList ardata = (ArrayList) arList.get(k);
//                    Row row = sheet.createRow((short) 0 + k);
//                    for (int p = 0; p < ardata.size(); p++) {
//                        Cell cell = row.createCell((short) p);
//                        String data = ardata.get(p).toString();
//                        if (data.contains("COMMA")) {
//                            cell.setCellValue(data.replaceAll("COMMA", ","));
//                        } else {
//                            cell.setCellValue(data);
//                        }
//                    }
//                }
//            }
//        }
//
//        workbook.write(fos);
//        fos.close();
//        System.out.println("Your excel file has been generated");
//    }
//
//
}