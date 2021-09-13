/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exel_veri_islemleri_apache_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author bahad
 */
public class Exel_Islemler_Read {
    private static final String FileName="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\data_deneme.xlsx";
    
    public static void main(String[] args) {
        
        FileInputStream input_stream_xlsx=null;
        XSSFWorkbook workbook=null;
        XSSFSheet sheet_1=null;
        try {
            input_stream_xlsx = new FileInputStream(FileName);
            workbook = new XSSFWorkbook(input_stream_xlsx);
            sheet_1 = workbook.getSheet("Sayfa1"); // --> aynı iade XSSFSheet sheet_1 = workbook.getSheetAt(0);
            int rows=sheet_1.getLastRowNum();
            int cols=sheet_1.getRow(1).getLastCellNum();
            
            System.out.println("xxxxxxxxxxxxxxxxxxxxxxxx");
            for (int i = 0; i <=rows; i++) {
                XSSFRow row = sheet_1.getRow(i);
                for (int j = 0; j < cols; j++) {
                    XSSFCell cell = row.getCell(j);
                    switch(cell.getCellType()){
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                                break;
                        case NUMERIC:
                                System.out.println(cell.getNumericCellValue());
                                break;
                        case BOOLEAN:
                                System.err.println(cell.getBooleanCellValue());
                                break;
                    }
                   
                }
                System.out.println();
            }
            
        } catch (FileNotFoundException e) {
            System.out.println("Dosya Bulunamadı");
        }catch(IOException e){
            System.out.println("Yazma- Okuma Hatası");
        }
        
        // ITERATOR KULLANARAK
        
        System.out.println("\n----ITERATOR İLE VERİ OKUMA----\n");
        
        Iterator iterator = sheet_1.iterator();
        while (iterator.hasNext()) {
            
            XSSFRow row = (XSSFRow)iterator.next();
            Iterator cellIterator = row.cellIterator();
            
            while (cellIterator.hasNext()) {
                
                XSSFCell cell = (XSSFCell) cellIterator.next();
                
                switch(cell.getCellType()){
                case STRING:   System.out.println(cell.getStringCellValue());   break;
                case NUMERIC:  System.out.println(cell.getNumericCellValue());  break;
                case BOOLEAN:  System.err.println(cell.getBooleanCellValue());   break;
                }        
                
            }     
        } 
    }
}
