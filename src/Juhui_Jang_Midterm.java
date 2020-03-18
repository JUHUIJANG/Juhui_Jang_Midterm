import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Reporter;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Juhui_Jang_Midterm {
    public String baseUrl = "http://demo.guru99.com/test/login.html";
    public String chromeDriverPath = "D:\\2020winter\\Testing_Automated\\Juhui_Jang_Test\\chromedriver_v80_win32\\chromedriver.exe";
    public WebDriver driver;
    public String filePath = "D:\\2020winter\\Testing_Automated\\Juhui_Jang_Test";
    public String fileName = "Juhui_Jang_Test.xlsx";
    public String sheetName = "Result";
    public String emailCellValue;
    public String passCellValue;

    @Test(priority = 0)
    public void writeExcelFile() throws IOException {
        System.out.println("writeExcelFile started");
        Reporter.log("writeExcelFile started");

        //Set email and password to be added into the excel file
        String email = "jjang18@my.centennialcollege.ca";
        String password = "pass1234";

        //Create a File object
        File file = new File(filePath+"\\"+fileName);
        //Create a workbook
        Workbook midtermWorkbook = null;
        //Get only file extension from the fileName by splitting it in substring using index of dot
        String fileExtension = fileName.substring(fileName.indexOf("."));

        //If the file is xlsx file
        if(fileExtension.equals(".xlsx")){
            //Create XSSFWorkbook object
            midtermWorkbook = new XSSFWorkbook();
        }
        //Else if the file is xls file
        else if(fileExtension.equals(".xls")){
            //Create HSSFWorkbook object
            midtermWorkbook = new HSSFWorkbook();
        }

        //Create an excel sheet in the workbook
        Sheet sheet = midtermWorkbook.createSheet(sheetName);

        //Create a new row
        Row row = sheet.createRow(0);

        Cell emailCell = row.createCell(0);
        emailCell.setCellValue(email);

        Cell passCell = row.createCell(1);
        passCell.setCellValue(password);

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        //Create a FileOutputStream object to write data into the excel file
        FileOutputStream outputStream = new FileOutputStream(file);
        //Write data
        midtermWorkbook.write(outputStream);
        //Close the output stream
        outputStream.close();

        System.out.println(fileName + "excel file has been created in "+filePath);
        Reporter.log(fileName + "excel file has been created in "+filePath);
    }

    @Test(priority = 1)
    public void readExcelFile() throws IOException {
        System.out.println("readExcelFile started");
        Reporter.log("readExcelFile started");
        //Create a File object
        File file = new File(filePath+"\\"+fileName);
        //Create a FileInputStream object to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        //Create a workbook
        Workbook midtermWorkbook = null;

        //Get only file extension from the fileName by splitting it in substring using index of dot
        String fileExt = fileName.substring(fileName.indexOf("."));

        //If the file is xlsx file
        if(fileExt.equals(".xlsx")){
            //Create XSSFWorkbook object
            midtermWorkbook = new XSSFWorkbook(inputStream);
        }
        //Else if the file is xls file
        else if(fileExt.equals(".xls")){
            //Create HSSFWorkbook object
            midtermWorkbook = new HSSFWorkbook(inputStream);
        }

        //Read sheet in the workbook
        Sheet sheet = midtermWorkbook.getSheet(sheetName);

        Row row = sheet.getRow(0);
        emailCellValue = row.getCell(0).getStringCellValue();
        passCellValue = row.getCell(1).getStringCellValue();

        System.out.println("Email cell value read from the excel: " +emailCellValue);
        Reporter.log("Email cell value read from the excel:" +emailCellValue);
        System.out.println("Password cell value read from the excel: " +passCellValue);
        Reporter.log("Password cell value read from the excel:" +passCellValue);
    }

    @Test(priority = 2)
    public void doLogin(){
        System.out.println("doLogin started");
        Reporter.log("doLogin started");

        System.setProperty("webdriver.chrome.driver", chromeDriverPath);
        driver = new ChromeDriver();

        // open chrome browser and maximize the window
        driver.manage().window().maximize();
        driver.get(baseUrl);
        System.out.println(baseUrl+" opened");
        Reporter.log(baseUrl+" opened");

        WebElement elementEmail = driver.findElement(By.id("email"));
        elementEmail.sendKeys(emailCellValue);
        System.out.println(emailCellValue + "entered into the email field");
        Reporter.log(emailCellValue + "entered into the email field");

        WebElement elementPass = driver.findElement(By.id("passwd"));
        elementEmail.sendKeys(passCellValue);
        System.out.println(emailCellValue + "entered into the password field");
        Reporter.log(emailCellValue + "entered into the password field");

        driver.findElement(By.id("SubmitLogin")).click();
        System.out.println("Login button clicked");
        Reporter.log("Login button clicked");
    }
}
