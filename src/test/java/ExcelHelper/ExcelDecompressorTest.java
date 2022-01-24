package ExcelHelper;

import org.junit.Test;


import java.io.IOException;

public class ExcelDecompressorTest {
    
    @Test
    public void test(){
        
    }

    @Test
    public void testUnzipExcelFile() throws IOException {
        String sourceExcelFilePath = "src/test/resources/testdata/pageobjects/PageObjects.xlsx";
        String destDirectoryPath = "src/test/resources/testdata/pageobjects/pageobjectsdecompressedfolder";
        ExcelDecompression.extractExcelFile(sourceExcelFilePath, destDirectoryPath);
    }

    @Test
    public void testCreateExcelFromDir() throws IOException {
        String sourceDirectoryPath = "src/test/resources/testdata/pageobjects/pageobjectsdecompressedfolder";
        String destExcelFilePath = "src/test/resources/testdata/pageobjects/PageObjects.xlsx";
//        ExcelDecompression.zipFolderToCreateExcel(sourceDirectoryPath, destExcelFilePath);
    }


}
