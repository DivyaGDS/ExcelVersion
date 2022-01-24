package ExcelHelper;

import java.io.*;
import java.util.Enumeration;
import java.util.zip.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * The Excel compression decompression
 */
public class ExcelDecompression {

    public static final byte[] BUFFEBYTESIZE = new byte[1024];
    public static final int BUFFERSIZE = 2048;

    public static void main(String[] args) throws IOException {
        new ExcelDecompression().extractExcelFile("C:\\Users\\UB217ZA\\Downloads\\Github Clone NGTP\\etaf-helpers\\src\\test\\resources\\testdata\\pageobjects\\PageObjects.xlsx","C:\\Users\\UB217ZA\\Downloads\\Github Clone NGTP\\etaf-helpers\\src\\test\\resources\\testdata\\pageobjects\\pageobjectsdecompressedfolder");
        System.out.println("Excel file is Decompressed !!");
    }

    /**
     * This method is used to extract excel file data to specified folder directory
     *
     * @param excelFile          : String : path to the excel file with extension
     * @param destLocationFolder : String : path to the folder in which extracted excel file data to be stored
     */
    public static void extractExcelFile(String excelFile, String destLocationFolder) throws IOException {
        File file = new File(excelFile);
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

}