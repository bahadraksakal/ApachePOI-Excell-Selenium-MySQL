/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Data_from_Database_Read_and_Write;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bahad
 */
public class DataBase_to_Exel {
    public static void main(String[] args) throws SQLException, FileNotFoundException, IOException {
        //exelpath
        final String File_Path="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\locations.xlsx";
        //connect dataBase
        final Connection connect =DriverManager.getConnection("jdbc:mysql://localhost:3306/database_tablo", "root", "Levent_1315");
        
        //statement//query
        Statement statement = connect.createStatement();
        ResultSet resultSet = statement.executeQuery("SELECT * FROM database_tablo.uye;");
        //Tablonun Başlıkları
        String id="iduye";
        String name="usarname";
        String passw="passworld";
        String mail = "email";
        //Exel
        // var olan veriyi almak için bir ınput.
        FileInputStream fis = new FileInputStream(File_Path);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.createSheet("Uyeler");
        
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue(id);
        row.createCell(1).setCellValue(name);
        row.createCell(2).setCellValue(passw);
        row.createCell(3).setCellValue(mail);
        
        int r=1;        
        while (resultSet.next()) {
            double uyeID=resultSet.getDouble(id);
            String uyeName=resultSet.getString(name);
            String uyePassw=resultSet.getString(passw);
            String uyeMail=resultSet.getString(mail);
            
            row=sheet.createRow(r++);
            
            row.createCell(0).setCellValue(uyeID);
            row.createCell(1).setCellValue(uyeName);
            row.createCell(2).setCellValue(uyePassw);
            row.createCell(3).setCellValue(uyeMail);
            
        }
        fis.close();
        FileOutputStream fos = new FileOutputStream(File_Path);
        workbook.write(fos);
        
        workbook.close();
        fos.close();
        connect.close();
        
        System.out.println("İşlem Başarılı");        
        
    }
    
    
}
