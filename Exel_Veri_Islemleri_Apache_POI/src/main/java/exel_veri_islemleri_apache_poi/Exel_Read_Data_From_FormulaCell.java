package exel_veri_islemleri_apache_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Exel_Read_Data_From_FormulaCell {
    private static final String exel_file_temp_formula_read ="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\readformula.xlsx";
    
    public static void main(String[] args) throws FileNotFoundException,IOException{
        
        FileInputStream exel_input = new FileInputStream(exel_file_temp_formula_read);
        XSSFWorkbook workbook = new XSSFWorkbook(exel_input);
        XSSFSheet sheet= workbook.getSheet("Sheet1");
        
        Iterator rows_iterator = sheet.rowIterator();
        
        while(rows_iterator.hasNext()){
            XSSFRow row = (XSSFRow)rows_iterator.next();
            Iterator celliterator= row.cellIterator();
            while (celliterator.hasNext()) {
                XSSFCell cell = (XSSFCell)celliterator.next();
                switch(cell.getCellType()){
                    case STRING: System.out.print("String:    "+cell.getStringCellValue()); break;
                    case FORMULA: System.out.print("Formula:  "+cell.getCellFormula() + " | Değeri:    "+cell.getNumericCellValue()); break;
                    case NUMERIC: System.out.print("Numaric:  "+cell.getNumericCellValue()); break;
                    case BOOLEAN: System.out.print("Boolean:  "+cell.getBooleanCellValue()); break;
                    // NOT case Numaric ile formüllenmiş numaric değerler okunmaz bunu okumak için
                    // yine aynı  cell.getNumericCellValue() kullanılır fakat, case FORMULA:  yapılır.
                }
                System.out.print("|");
            }
            System.out.println();
        }
        
        
                
    }
}
