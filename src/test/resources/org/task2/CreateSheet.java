package org.task2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateSheet {
	public static void main(String[]args) throws IOException {
		File f = new File("C:\\Users\\Indiaseah Quality\\eclipse-workspace\\TrialMaven\\src\\test\\resources\\ExcelRead.xlsx");
		FileInputStream f1=new FileInputStream(f);
		Workbook w = new XSSFWorkbook(f1);
		Sheet s=w.createSheet("ExcelWrite");
		Row r= s.createRow(0);
		Cell c=r.createCell(0);
		c.setCellValue("dhoni");
	FileOutputStream o=new FileOutputStream(f);
	w.write(o);
	System.out.println("done");
	}
}
