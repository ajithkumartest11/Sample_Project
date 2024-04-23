package com.selenium.exceldummy;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_excel {
	
	public static void main(String[] args) {
		
		FileInputStream in = null;
		try {
			in = new FileInputStream("D:\\Chandresh\\Selenium\\Selenium Framework\\empinfo.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(in);
			XSSFSheet sheet = workbook.getSheet("Employees");
			
			//specific row & cell value
			//System.out.println("particluar value>>>>>>>>"+" "+sheet.getRow(1).getCell(0));
			//System.out.println("particluar value>>>>>>>>"+" "+sheet.getRow(1).getCell(2));
			
			Iterator<Row> itr = sheet.iterator();
			while(itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> itr2 = row.iterator();
				while(itr2.hasNext()) {
					Cell cell = itr2.next();
					System.out.print(cell.getStringCellValue() + "\t\t");
					
				}
				System.out.println("");
			}
			
		}catch (Exception e) {
			e.printStackTrace();
		}finally {
			if(in != null) {
				try {
					in.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
	}

}
