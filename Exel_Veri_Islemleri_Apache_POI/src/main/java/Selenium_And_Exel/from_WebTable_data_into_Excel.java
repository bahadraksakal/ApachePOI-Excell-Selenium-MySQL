package Selenium_And_Exel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class from_WebTable_data_into_Excel {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        System.setProperty("webdriver.chrome.driver","D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\chromedriver.exe");
        WebDriver driver=new ChromeDriver();

        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        driver.manage().window().maximize();

        driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
        String path="population.xlsx";
        
        FileInputStream fis = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.createSheet("sayfa2"); // get ile al sayfa2 yi oluşturdum 2. defa çalışırsa hata alırsın
        XSSFRow  row = sheet.createRow(0);
        
        row.createCell(0).setCellValue("Country");
        row.createCell(1).setCellValue("Population");
        row.createCell(2).setCellValue("% of world");
        row.createCell(3).setCellValue("Date");
        row.createCell(4).setCellValue("Source");
        
        WebElement table=driver.findElement(By.xpath("//*[@id=\"mw-content-text\"]/div[1]/table/tbody"));
	int rows=table.findElements(By.xpath("tr")).size(); // rows present in web table
        //driver.findElement(By.xpath("//*[@id=\"mw-content-text\"]/div[1]/table/tbody/tr"));
        
        for (int r = 1; r <= rows; r++) {
            
            String country=table.findElement(By.xpath("tr["+r+"]/td[1]")).getText();
            String population=table.findElement(By.xpath("tr["+r+"]/td[3]")).getText();
            String perOfWorld=table.findElement(By.xpath("tr["+r+"]/td[4]")).getText();
            String date=table.findElement(By.xpath("tr["+r+"]/td[5]")).getText();
            String source=table.findElement(By.xpath("tr["+r+"]/td[6]")).getText();
            
            System.out.println(country+population+perOfWorld+date+source);
            
            XSSFRow  row_temp = sheet.createRow(r);
            row_temp.createCell(0).setCellValue(country);
            row_temp.createCell(1).setCellValue(population);
            row_temp.createCell(2).setCellValue(perOfWorld);
            row_temp.createCell(3).setCellValue(date);
            row_temp.createCell(4).setCellValue(source);
        }
        fis.close();
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);        
        workbook.close();
        fos.close();
        driver.close();
        System.out.println("Web scrapping is done succesfully.....");
        
    }
}
