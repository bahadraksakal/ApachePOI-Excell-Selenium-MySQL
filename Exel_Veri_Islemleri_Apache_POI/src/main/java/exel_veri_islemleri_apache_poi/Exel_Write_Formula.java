
package exel_veri_islemleri_apache_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Exel_Write_Formula {
    public static void main(String[] args) throws FileNotFoundException,IOException{
        String file_path = "D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\Numbers.xlsx";
        
        FileOutputStream outputStream = new FileOutputStream(file_path);
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sayfa0");
        XSSFRow row = sheet.createRow(0);
        
        row.createCell(0).setCellValue(10);
        row.createCell(1).setCellValue(20);
        row.createCell(2).setCellValue(30);
    
        row.createCell(4).setCellFormula("A1*B1*C1");
        
        workbook.write(outputStream);
        outputStream.close();
        
        System.out.println("\n---Hücrelere Formül Ekle---\n");
        
        String path="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\books.xlsx";
		
        FileInputStream fis = new FileInputStream(path);

        XSSFWorkbook workbook_2 =new XSSFWorkbook(fis);

        XSSFSheet sheet_2 =workbook_2.getSheetAt(0);

        sheet_2.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)"); // c2 den c6 ya kadar topla.
        
        workbook.close();
        fis.close();

        FileOutputStream fos=new FileOutputStream(path);
        workbook_2.write(fos);

        workbook_2.close();
        fos.close();

        System.out.println("Done!!!");
    }
}
