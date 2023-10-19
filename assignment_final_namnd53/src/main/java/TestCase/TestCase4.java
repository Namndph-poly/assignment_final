package TestCase;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class TestCase4 {
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

        Thread.sleep(1000);
        // Step 2: Find and click button "Giao đến"
        WebElement button_address = driver.findElement(By.xpath("//div[@class='delivery-zone__heading']"));
        button_address.click();

        //Step 3 : Find and select radio "Chọn khu vực giao hàng khác"
        Thread.sleep(3000);
        List<WebElement> radio_address_other = driver.findElements(By.xpath("//button[@class='RadioButton__Button-sc-r7isam-0 iSOqmm']"));
        WebElement  radio_other = radio_address_other.get(1);
        radio_other.click();

        //Step 4 : Display combox city
        Thread.sleep(1000);

        WebElement combo_city = driver.findElement(By.xpath("//*[@class='location-type' and contains(text(),'Tỉnh/Thành phố')]"));
        boolean isCity = combo_city.isDisplayed();

        //Step 5 : Display combox distric
        Thread.sleep(1000);

        WebElement combo_distric = driver.findElement(By.xpath("//*[@class='location-type' and contains(text(),'Quận')]"));
        boolean isDistric = combo_distric.isDisplayed();

        //Step 6 : Display combox wards
        Thread.sleep(1000);

        WebElement combo_wards = driver.findElement(By.xpath("//*[@class='location-type' and contains(text(),'Phường')]"));
        boolean isWards = combo_wards.isDisplayed();


        //update result to file Excel
        updateTestResult(isCity,isDistric,isWards);

        // Close Browser
//        driver.quit();
    }

    private void updateTestResult(boolean isCity,boolean isDistric,boolean isWards) {

        int rowIndex = 4;
        int columnIndex = 4;

        Row row = sheet.getRow(rowIndex); // Get the row
        if (row == null) {
            row = sheet.createRow(rowIndex); // Create row if it does not exist
        }
        Cell resultCell = row.createCell(columnIndex);//Create 1 cell
        if (isCity == true && isDistric==true && isWards==true) {
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
