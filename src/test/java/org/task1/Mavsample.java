package org.task1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Mavsample {
	public static void main(String[]args) throws IOException {
		File f = new File("C:\\Users\\Indiaseah Quality\\eclipse-workspace\\TrialMaven\\src\\test\\resources\\ExcelRead.xlsx");
		
				
		Workbook w = new XSSFWorkbook();
		Sheet s=w.createSheet("ExcelWrite");
		Row r= s.createRow(0);
		Cell c=r.createCell(0);
		c.setCellValue("dhoni");
	FileOutputStream o=new FileOutputStream(f);
	w.write(o);
	System.out.println("done");
	}
		//Sheet s= w.getSheet("Sheet1");
		//for (int i=0; i <s.getPhysicalNumberOfRows(); i++) {
			//Row r= s.getRow(i);
			//for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
				//Cell c=r.getCell(j);
				//int cellType=c.getCellType();
				//if(cellType==1) {
				//	String name= c.getStringCellValue();
					//System.out.println(name);
				//}
				//else if(cellType==0) {
				//	if(DateUtil.isCellDateFormatted(c)) {
					//	Date d=c.getDateCellValue();
						//SimpleDateFormat sd= new SimpleDateFormat("MM/dd/yyyy");
						//String name=sd.format(d);
						//System.out.println(name);
					//}
				//	else {
					//	double d= c.getNumericCellValue();
					//	long l=(long)d;
						//String name= String.valueOf(l);
						// System.out.println(name);
						
						//}
				//}
		//	}
		//}
		
	//}

}
