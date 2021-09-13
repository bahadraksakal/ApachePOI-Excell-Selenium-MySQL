/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exel_veri_islemleri_apache_poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bahad
 */
public class ExcelToHashMap {
    public static void main(String[] args) throws IOException {
		
        FileInputStream fis=new FileInputStream("student.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet sheet=workbook.getSheet("Student data");

        int rows=sheet.getLastRowNum();

        HashMap<Integer,String> data=new HashMap<>();

        //Reading data from excel to HashMap
        for(int r=0;r<=rows;r++)
        {
                int key= (int) sheet.getRow(r).getCell(0).getNumericCellValue();
                String value=sheet.getRow(r).getCell(1).getStringCellValue();
                data.put(key, value);

        }

        //Read data from HashMap

        for(Map.Entry entry:data.entrySet())
        {
                System.out.println(entry.getKey()+"   "+entry.getValue());
        }
		
    }
}
