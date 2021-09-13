/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Data_from_Database_Read_and_Write;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bahad
 */
public class Exel_To_DataBase {
    public static void main(String[] args) throws SQLException, FileNotFoundException, IOException {
        //exelpath
        final String File_Path="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\locations.xlsx";
        //connect dataBase
        final Connection connect =DriverManager.getConnection("jdbc:mysql://localhost:3306/database_tablo", "root", "Levent_1315");
        //statement//query
        Statement statement = connect.createStatement();
        //create a new table in the database 'places'
//        String sql="create table database_tablo.Ulkeler (LOCATION_ID decimal(4,0), STREET_ADDRESS varchar(40),POSTAL_CODE varchar(12),CITY varchar(30),STATE_PROVINCE varchar(25),COUNTRY_ID varchar(2))";
//        statement.execute(sql);
        //bir defa olulturulur.
        //Excel
        FileInputStream fis=new FileInputStream(File_Path);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet sheet=workbook.getSheet("Locations Data");
        
        Iterator rowIterator = sheet.rowIterator();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            XSSFRow row = (XSSFRow) rowIterator.next();            
                
            int locId=(int) row.getCell(0).getNumericCellValue();
            String streatAdd=row.getCell(1).getStringCellValue();
            String postalCode=row.getCell(2).getStringCellValue();
            String city=row.getCell(3).getStringCellValue();
            String stateProv=row.getCell(4).getStringCellValue();
            String countryId=row.getCell(5).getStringCellValue();
            String sql="insert into database_tablo.ulkeler (LOCATION_ID, STREET_ADDRESS, POSTAL_CODE , CITY, STATE_PROVINCE, COUNTRY_ID)"+
                 "value ('"+locId+"','"+streatAdd+"','"+postalCode+"','"+city+"','"+stateProv+"','"+countryId+"')";
            statement.execute(sql);
//          statement.execute("commit");
        }   
        
        workbook.close();
        fis.close();
        connect.close();
    }
}
