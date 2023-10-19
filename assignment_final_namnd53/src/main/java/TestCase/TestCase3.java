package TestCase;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TestCase3 {
    private ChromeDriver driver;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private String excelFilePath = "C:\\Users\\ACER\\Documents\\assignment_final_namnd53\\Tiki_TestCase.xlsx";

    @BeforeClass
    public void setUp() {
        WebDriverManager.chromedriver().setup();
        // Initialize Chrome browser
        driver = new ChromeDriver();

        // Open file Excel
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath)) {
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void verifyProductDetail() throws InterruptedException {
        // Step 1: Open browser and navigate web Tiki
        driver.get("https://tiki.vn/");
        driver.manage().window().maximize();

        // Step 2: Check display logo
        WebElement logo = driver.findElement(By.xpath("//img[@alt='tiki-logo']"));
        boolean isLogo = logo.isDisplayed();

        //update result to file Excel
        updateTestResult(isLogo);

        // Close Browser
//        driver.quit();
    }

    private void updateTestResult(boolean isLogo) {

        int rowIndex = 3;
        int columnIndex = 4;

        Row row = sheet.getRow(rowIndex); // Get the row
        if (row == null) {
            row = sheet.createRow(rowIndex); // Create row if it does not exist
        }
        Cell resultCell = row.createCell(columnIndex);//Create 1 cell
        if (isLogo==true) {
            resultCell.setCellValue("Pass");
            CellStyle passStyle = workbook.createCellStyle();
            passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            resultCell.setCellStyle(passStyle);
        } else {
            resultCell.setCellValue("Fail");
            CellStyle failStyle = workbook.createCellStyle();
            failStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            resultCell.setCellStyle(failStyle);
        }

        // Save file Excel
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
