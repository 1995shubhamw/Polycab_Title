package package1;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class Titles 
{
 //
    private WebDriver driver;
    private WebDriverWait wait;
    private Workbook workbook;
    private Sheet sheet;
    private int rowNum = 0;
    private String excelFilePath = "C:\\Users\\ShubhamWatile\\eclipse-workspace\\Polycab_Titles\\Title.xlsx";  // Existing Excel file in the project folder

    @BeforeClass
    public void setUp() 
    {
        System.setProperty("webdriver.gecko.driver", "C:\\Testing\\Drivers\\geckodriver-v0.33.0-win64\\geckodriver.exe");
        driver = new FirefoxDriver();
        wait = new WebDriverWait(driver, Duration.ofSeconds(50));
        driver.get("https://dev.polycab.com");
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='rejectBtn']"))).click();

        try (FileInputStream fis = new FileInputStream(excelFilePath)) 
        {
            workbook = new XSSFWorkbook(fis);
        } 
        catch (IOException e) 
        {
            e.printStackTrace();
        }

        String sheetName = this.getClass().getSimpleName();
        sheet = workbook.getSheet(sheetName);
        if (sheet == null) 
        {
            sheet = workbook.createSheet(sheetName);
        }

        if (sheet.getLastRowNum() == 0) {
            Row headerRow = sheet.createRow(rowNum++);
            headerRow.createCell(0).setCellValue("URL");
            headerRow.createCell(1).setCellValue("Title");
        } 
        else 
        {
            rowNum = sheet.getLastRowNum() + 1;  // Continue from the last row if data exists
        }
    }

    @Test
    public void testExtractUrlsAndTitles() throws InterruptedException, IOException
    {
        JavascriptExecutor js = (JavascriptExecutor) driver;

        List<WebElement> links = driver.findElements(By.tagName("a"));
        for (int i = 0; i < links.size(); i++) {
            WebElement link = driver.findElements(By.tagName("a")).get(i);
            String url = link.getAttribute("href");

            if (url != null && !url.isEmpty() && !url.startsWith("#")) { // Check for valid URLs
                driver.get(url);

                // Handle cookie pop-up again (if present on new pages)
                try {
                    WebElement acceptCookiesButtonOnNewPage = driver.findElement(By.xpath("//button[@id='rejectBtn']"));
                    if (acceptCookiesButtonOnNewPage.isDisplayed())
                    {
                        acceptCookiesButtonOnNewPage.click();
                    }
                } catch (Exception e)
                {
                   // System.out.println("Cookie pop-up not found or already handled on new page.");
                }

                String title = driver.getTitle();

                // Print URL and title
                System.out.println("URL: " + url + " - Title: " + title);

                // Write URL and title to Excel
                Row row = sheet.createRow(rowNum++);
                Cell urlCell = row.createCell(0);
                urlCell.setCellValue(url);
                Cell titleCell = row.createCell(1);
                titleCell.setCellValue(title);

                driver.navigate().back();
                Thread.sleep(2000);  // Adjust sleep based on page load time
            }
        }

        Assert.assertTrue(links.size() > 0, "No URLs were found.");
    }

    @AfterClass
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }

        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}


//Commiting to qc branch
