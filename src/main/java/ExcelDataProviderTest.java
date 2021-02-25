import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelDataProviderTest {

    @Test(dataProvider = "excel-data")
    public void test(String firstColumnHeader, String secondColumnHeader, String thirdColumnHeader) {
        System.out.println(firstColumnHeader);
        System.out.println(secondColumnHeader);
        System.out.println(thirdColumnHeader);
    }

    @DataProvider(name = "excel-data")
    public Object[][] excelDP() throws IOException {
        Object[][] arrObj = getExcelData("src/main/resources/Data.xlsx", "FirstSheet");
        return arrObj;
    }

    public String[][] getExcelData(String fileName, String sheetName) {
        String[][] data = null;
        try {
            FileInputStream fis = new FileInputStream(fileName);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet(sheetName);
            XSSFRow row = sh.getRow(0);
            int noOfRows = sh.getPhysicalNumberOfRows();
            int noOfCols = row.getLastCellNum();
            Cell cell;
            data = new String[noOfRows - 1][noOfCols];
            for (int i = 1; i < noOfRows; i++) {
                for (int j = 0; j < noOfCols; j++) {
                    row = sh.getRow(i);
                    cell = row.getCell(j);
                    data[i - 1][j] = cell.getStringCellValue();
                }
            }
        } catch (Exception e) {
            System.out.println("The exception is: " + e.getMessage());
        }
        return data;
    }
}
