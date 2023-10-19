package TestCase;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TestCase5 {
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

        // Step 2: Find txt_search and sendKeys "Iphone 11"
        Thread.sleep(2000);
        WebElement searchBox = driver.findElement(By.xpath("//input[@class='styles__InputRevamp-sc-6cbqeh-2 IXqBC']"));
        searchBox.sendKeys("iphone11");

        Thread.sleep(2000);

        WebElement button_search = driver.findElement(By.xpath("//button[contains(@class,'styles__ButtonRevamp-sc-6cbqeh-3 LdVUr')]"));
        button_search.click();


        // Step 3: Select first product on result page
        Thread.sleep(2000);
        WebElement firstProduct = driver.findElement(By.xpath("//a[contains(@class,'style__Product') and @href='/apple-iphone-11-p184036446.html?spid=32033721']"));
        firstProduct.click();

        // Step 4: Check results

        // display title product
        Thread.sleep(2000);
        WebElement productName = driver.findElement(By.xpath("//h1[contains(text(),'Apple iPhone 11')]"));
        boolean isIphone = productName.getText().contains("Apple iPhone 11");

        // display image product
        Thread.sleep(2000);
        WebElement productImg = driver.findElement(By.xpath("//img[@alt='Apple iPhone 11']"));
        boolean isImg = productImg.isDisplayed();

        // display price product
        Thread.sleep(2000);
        WebElement productPrice = driver.findElement(By.xpath("//div[contains(@class,'product-price__current-price')]"));
        boolean isPrice = productPrice.isDisplayed();


        ////update result to file Excel
        updateTestResult(isIphone,isImg,isPrice);

        // Close Browser
//        driver.quit();
    }

    private void updateTestResult(boolean isIphone,boolean isImg, boolean isPrice) {

        int rowIndex = 5;
        int columnIndex = 4;

        Row row = sheet.getRow(rowIndex); // Get the row
        if (row == null) {
            row = sheet.createRow(rowIndex); // Create row if it does not exist
        }
        Cell resultCell = row.createCell(columnIndex);//Tạo 1 ô
        if (isIphone==true && isImg==true && isPrice==true) {
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
