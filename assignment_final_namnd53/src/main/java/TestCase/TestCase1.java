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

public class TestCase1 {
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

        // Step 2: search and click button Account
        Thread.sleep(2000);
        WebElement button_account = driver.findElement(By.xpath("//span[contains(text(),'Tài khoản')]"));
        button_account.click();

        // Step 3: Check results

        // display txt
        Thread.sleep(2000);
        WebElement txt_PhoneNumber = driver.findElement(By.xpath("//input[@type='tel']"));
        boolean isPhone = txt_PhoneNumber.isDisplayed();

        // display button login
        Thread.sleep(1000);
        WebElement btn_login = driver.findElement(By.xpath("//button[contains(text(),'Tiếp Tục')]"));
        boolean isButtonLogin = btn_login.isDisplayed();

        // display button email
        Thread.sleep(1000);
        WebElement btn_email = driver.findElement(By.xpath("//p[contains(text(),'Đăng nhập bằng email')]"));
        boolean isEmail = btn_email.isDisplayed();

        // display button facebook
        Thread.sleep(1000);
        WebElement btn_faceBook = driver.findElement(By.xpath("//img[@alt='facebook']"));
        boolean isFaceBook = btn_faceBook.isDisplayed();

        // display button google
        Thread.sleep(1000);
        WebElement btn_google = driver.findElement(By.xpath("//img[@alt='google']"));
        boolean isGoogle = btn_google.isDisplayed();

        //update result to file Excel
        updateTestResult(isPhone,isButtonLogin,isEmail,isFaceBook,isGoogle);

        // Close Browser
//        driver.quit();
    }

    private void updateTestResult(boolean isPhone,boolean isButtonLogin, boolean isEmail,boolean isFaceBook,boolean isGoogle) {

        int rowIndex = 1;
        int columnIndex = 4;

        Row row = sheet.getRow(rowIndex); // Get the row
        if (row == null) {
            row = sheet.createRow(rowIndex); // Create row if it does not exist
        }
        Cell resultCell = row.createCell(columnIndex);//Create 1 cell
        if (isPhone==true && isButtonLogin==true && isEmail==true && isFaceBook==true && isGoogle==true) {
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
