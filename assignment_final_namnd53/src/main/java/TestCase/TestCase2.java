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

public class TestCase2 {
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

        // Step 2: search and click button category_It
        Thread.sleep(2000);
        WebElement category_IT = driver.findElement(By.xpath("//div[@title=\"Laptop - Máy Vi Tính - Linh kiện\"]"));
        category_IT.click();

        // Step 3: Check results

        // display category
        Thread.sleep(3000);
        WebElement category_laptop = driver.findElement(By.xpath("//a[contains(text(),'Laptop') and contains(@class,'item')]"));
        boolean isLapTop = category_laptop.isDisplayed();

        WebElement category_equip_office = driver.findElement(By.xpath("//a[contains(text(),'Thiết Bị Văn Phòng - Thiết Bị Ngoại Vi') and contains(@class,'item')]"));
        boolean isEquip = category_equip_office.isDisplayed();

        // display rating
        WebElement rating_star = driver.findElement(By.xpath("//a[@href='/laptop-may-vi-tinh-linh-kien/c1846?rating=5']"));
        boolean isRating = rating_star.isDisplayed();

        // display brand
        WebElement brand = driver.findElement(By.xpath("//h4[@class='title' and contains(text(),'Thương hiệu')]"));
        boolean isBrand = brand.isDisplayed();

        // display provide
        WebElement provide = driver.findElement(By.xpath("//h4[@class='title' and contains(text(),'Nhà cung cấp')]"));
        boolean isProvide = provide.isDisplayed();

        //update result to file Excel
        updateTestResult(isLapTop,isEquip,isRating,isBrand,isProvide);

        // Close Browser
//        driver.quit();
    }

    private void updateTestResult(boolean isLapTop,boolean isEquip, boolean isRating,boolean isBrand,boolean isProvide) {

        int rowIndex = 2;
        int columnIndex = 4;

        Row row = sheet.getRow(rowIndex); // Get the row
        if (row == null) {
            row = sheet.createRow(rowIndex); // Create row if it does not exist
        }
        Cell resultCell = row.createCell(columnIndex);//Create 1 cell
        if (isLapTop==true && isEquip==true && isRating==true && isBrand==true && isProvide==true) {
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
