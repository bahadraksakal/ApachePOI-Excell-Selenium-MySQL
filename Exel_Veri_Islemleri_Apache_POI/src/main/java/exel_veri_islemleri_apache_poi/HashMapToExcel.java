/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exel_veri_islemleri_apache_poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bahad
 */
public class HashMapToExcel {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet= workbook.createSheet("Student data");
		
		
		Map<Integer,String> data=new HashMap<>();
		data.put(101,"John1");
		data.put(102,"Smith2");
		data.put(103,"Scott3");
		data.put(104,"Kim4");
		data.put(105,"Mary5");
		
		int rowno=0;
		
               
		for(Map.Entry entry:data.entrySet())
		{
			XSSFRow row=sheet.createRow(rowno++);
			
			row.createCell(0).setCellValue((int)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());
		}
		
		
		FileOutputStream fos=new FileOutputStream("D:\\Belgeler\\NetBeansProjects\\Exel_Veri_Islemleri_Apache_POI\\student.xlsx");
		
		workbook.write(fos);
		fos.close();
		System.out.println("Excel written succesfully");
		
	}

}
