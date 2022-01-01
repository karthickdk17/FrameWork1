package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Details {
	
	public static void main(String[] args) throws IOException {
		File file =new File("C:\\Users\\Dell\\eclipse-workspace\\FrameWork1\\Excel\\Excel1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Data");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);	
			
			int type = cell.getCellType();
			if(type==1) {
				String data1 = cell.getStringCellValue();
				System.out.println(data1);
				
				}
			if(type==0) {
				
				double d = cell.getNumericCellValue();
				long l=(long)d;
				String data2 = String.valueOf(l);
				System.out.println(data2);
			}
			
			
			}
			
		}
		
	
	}

}
