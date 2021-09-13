
package exel_veri_islemleri_apache_poi;

import java.awt.Font;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Cell_Background_Foreground_Colors {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        String exel_Path="D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\styles.xlsx";
        //değişikliklerin kayıtlı olabilmesi için
          FileInputStream fis = new FileInputStream(exel_Path);
        //var olan verilerin üstüne yazılabilir fakat var olan bir sheet in üstüne yazılamaz.
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
//        XSSFSheet sheet = workbook.createSheet("Sayfa2");
        XSSFSheet sheet = workbook.getSheet("Sayfa1");
        XSSFRow row= sheet.createRow(0);
        //background ayarlama
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillBackgroundColor(IndexedColors.DARK_TEAL.getIndex());
        cellStyle.getFont().setFontHeightInPoints((short) 25);        
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.getFont().setColor(IndexedColors.BRIGHT_GREEN1.getIndex());
        XSSFCell cell = row.createCell(0);
        row.createCell(0).setCellValue("Hoşgeldiniz");
        cell.setCellStyle(cellStyle);
        
        //ForeGround ayaralama // ayrı ayrı fıntlar oluşturmayıp get font ile cell üzerinden alırsak aynı fontun üzerinde değişiklik yapacağımız için cell cell özelleştirme yapamayız.
        
        XSSFCellStyle foregCellStyle = workbook.createCellStyle();
        foregCellStyle.setFillForegroundColor(IndexedColors.GOLD.getIndex());
        foregCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        foregCellStyle.setFont(font);
       
        foregCellStyle.getFont().setFontHeightInPoints((short) 14);
        
        XSSFCell cell2=row.createCell(1);
        cell2.setCellValue("foreground");
        cell2.setCellStyle(foregCellStyle);
        
        fis.close();
        
        FileOutputStream fos = new FileOutputStream(exel_Path);
        workbook.write(fos);
        
        workbook.close();
        fos.close();
    }
}
