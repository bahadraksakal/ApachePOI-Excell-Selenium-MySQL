package exel_veri_islemleri_apache_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//workbook --> sheet --> row --> cell
public class Exel_Write {
    
    private static final String exel_file_path = "D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\employee.xlsx";
    private static final String exel_file_path2 = "D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\employee2.xlsx";
    public static void main(String[] args) {
        
        XSSFWorkbook workbook = new XSSFWorkbook();        
        XSSFSheet sheet = workbook.createSheet("Bilgi");
        Object sheet_bilgi_data [][]= { 
            {"Çalışan ID" , "Adı" , "Mesleği"},
            {001 , "Bahadır Aksakal" , "Bilgisayar Mühendisliği"},
            {002 , "Ayşe Kırgız" , "Seyis"},
            {003 , "Zeynep Taşcı" , "Veteriner Hekim"}
        };
        int rows= sheet_bilgi_data.length;
        int cols = sheet_bilgi_data[0].length;
        
        System.out.println("rows:   "+rows); // 4
        System.out.println("cols:   "+cols);//3
        
        for (int i = 0; i < rows; i++) {
            
            XSSFRow row  = sheet.createRow(i);
            
            for (int j = 0; j < cols; j++) {
                
                XSSFCell cell = row.createCell(j);
                Object value = sheet_bilgi_data[i][j];
                
                if (value instanceof String) {
                    cell.setCellValue((String)value);
                }else if (value instanceof Integer) {
                    cell.setCellValue((int)value);
                }else if (value instanceof Boolean) {
                    cell.setCellValue((boolean)value);
                }else{
                    System.err.println("Veri Türü Tanımlı değil");
                }        
            }
        }
        
        try {
            FileOutputStream f_ouOutputStream = new FileOutputStream(exel_file_path);
            workbook.write(f_ouOutputStream);
            f_ouOutputStream.close();
        } catch (FileNotFoundException ex) {
            System.err.println("Dosya Bulunamadı");
        } catch(IOException ex){
            System.err.println("Yazma İşlemi hatası");
        }
         System.out.println("Employee.xlxs write is succesfull");
        
         System.out.println("\n---ForEach-Ve-ArrayList---\n");
        
        XSSFWorkbook workbook_2=new XSSFWorkbook();
        XSSFSheet sheet_2=workbook_2.createSheet("Emp Info 2");		
        ArrayList<Object[]> empdata=new ArrayList<Object[]>();	     
	empdata.add(new Object[]{"Empid","Name","Job"});
        empdata.add(new Object[]{101,"David","Enginner"});
        empdata.add(new Object[]{102,"Smith","Manager"});
        empdata.add(new Object[]{103,"Scott","Analyst"});		
		
        /// using for...each loop	
        
        int rownum=0;
        for(Object[] emp:empdata)
        {
           XSSFRow row_2 =sheet_2.createRow(rownum++);
           int cellnum=0; 

            for(Object value:emp)
            {
                XSSFCell cell_2=row_2.createCell(cellnum++);
                
                if(value instanceof String)
                    cell_2.setCellValue((String)value);
                if(value instanceof Integer)
                    cell_2.setCellValue((Integer)value);
                if(value instanceof Boolean)
                    cell_2.setCellValue((Boolean)value);	
            }
        }
       
        FileOutputStream outstream_2;
        try {
            outstream_2 = new FileOutputStream(exel_file_path2);
            workbook_2.write(outstream_2);
            outstream_2.close();
        } catch (FileNotFoundException ex) {
            System.out.println("Dosya Bulunamadı"+ex);;
        }catch(IOException ex){
            System.out.println("Okuma Yazma Hatası"+ex);
        }
        
        
        System.out.println("Employee.xls file written successfully...");

}
    
}
