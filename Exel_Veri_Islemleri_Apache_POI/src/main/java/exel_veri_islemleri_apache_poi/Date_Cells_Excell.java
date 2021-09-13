/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exel_veri_islemleri_apache_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bahad
 */
public class Date_Cells_Excell {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        final String exel_file_path = "D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\Date_Cells.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook();        
        XSSFSheet sheet = workbook.createSheet("Tarihler ve Saat");
        XSSFCreationHelper creationHelper = new XSSFCreationHelper(workbook);
        XSSFRow row = sheet.createRow(0);
        //format0 numeric type // bu yapı ile alınan date exelde işleme tabi tutulmalıdır .
        sheet.createRow(1).createCell(0).setCellValue(new Date());
        
        //format1: dd-mm-yyyy
        XSSFCellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy"));
        
        XSSFCell cell1 = row.createCell(0);
        cell1.setCellValue(new Date());
        cell1.setCellStyle(cellStyle1);
        
        //format2 mm-dd-yyyy
        XSSFCellStyle cellStyle2 = workbook.createCellStyle();
        cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy"));
        
        XSSFCell cell2 = row.createCell(1);
        cell2.setCellValue(new Date());
        cell2.setCellStyle(cellStyle2);
        
        //format3 mm-dd-yyyy hh:mm:ss
        XSSFCellStyle cellStyle3 = workbook.createCellStyle();
        cellStyle3.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy hh:mm:ss"));
        
        XSSFCell cell3 = row.createCell(2);
        cell3.setCellValue(new Date());
        cell3.setCellStyle(cellStyle3);
        
        //format4 hh:mm:ss
        XSSFCellStyle cellStyle4 = workbook.createCellStyle();
        cellStyle4.setDataFormat(creationHelper.createDataFormat().getFormat("hh:mm:ss"));
        
        XSSFCell cell4 = row.createCell(3);
        cell4.setCellValue(new Date());
        cell4.setCellStyle(cellStyle4);
        
        
        FileOutputStream fos = new FileOutputStream(exel_file_path);
        workbook.write(fos);
        
        workbook.close();
        fos.close();
        
        
        
        
        
        
        
    }
}
