/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Selenium_And_Exel;
//TESTER PAKETE TAŞINMIŞTIR
//import java.util.concurrent.TimeUnit;
//import org.openqa.selenium.By;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.WebElement;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.testng.Assert;
//import org.testng.annotations.AfterClass;
//import org.testng.annotations.BeforeClass;
//import org.testng.annotations.DataProvider;
//import org.testng.annotations.Test;

/**
 *
 * @author bahad
 */

//public class Data_Driven_Testing_Selenium_Exel{
//    WebDriver webDriver;
//    
//    @BeforeClass
//    public void setup(){
//        System.setProperty("webdriver.chrome.driver","D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\chromedriver.exe");
//        webDriver= new ChromeDriver();
//        webDriver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
//        webDriver.manage().window().maximize();
//    }
//    
//    
//    @Test(dataProvider = "LoginData")
//    public void LoginTest(String user,String Passw ,String explanation){
//        webDriver.get("https://admin-demo.nopcommerce.com/login");
//		
//        WebElement txtEmail=webDriver.findElement(By.id("Email")); // input username companent..
//        txtEmail.clear(); // default gelen yazıyı temizle
//        txtEmail.sendKeys(user);
//
//
//        WebElement txtPassword=webDriver.findElement(By.id("Password")); // input passw component
//        txtPassword.clear();
//        txtPassword.sendKeys(Passw);
//
//        webDriver.findElement(By.xpath("/html/body/div[6]/div/div/div/div/div[2]/div[1]/div/form/div[3]/button")).click(); //Login  button
//        System.out.println(user+Passw+explanation);
//        
//        String exp_title="Dashboard / nopCommerce administration"; // web sayfasının giris basarılı title si
//        String act_title=webDriver.getTitle(); // o an ki title sayfadan çekiyoruz.
//
//        if(explanation.equals("Valid"))
//        {
//                if(exp_title.equals(act_title))
//                {
//                        webDriver.findElement(By.linkText("Logout")).click();
//                        Assert.assertTrue(true);
//                }
//                else
//                {
//                        Assert.assertTrue(false);
//                }
//        }
//        else if(explanation.equals("Invalid"))
//        {
//                if(exp_title.equals(act_title))
//                {
//                        webDriver.findElement(By.linkText("Logout")).click();
//                        Assert.assertTrue(false);
//                }
//                else
//                {
//                        Assert.assertTrue(true);
//                }
//        }
//    }
//    
//    @DataProvider(name="LoginData")
//    public String [][] getData() 
//    {
//        String loginData[][]= {
//                {"admin@yourstore.com","admin","Valid"},
//                {"admin@yourstore.com","adm","Invalid"},
//                {"adm@yourstore.com","admin","Invalid"},
//                {"adm@yourstore.com","adm","Invalid"}
//        };
//        return loginData;
//    }
//    
//    @AfterClass
//    public void closeDriven(){
//        webDriver.close();
//    }
//}
