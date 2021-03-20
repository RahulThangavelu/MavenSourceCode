package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("unused")
public class DataDrivenReadUpdate {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\\\Users\\\\RAHUL\\\\eclipse-workspace\\\\FrameWorks\\\\Excel\\\\Book1.xlsx");
		
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s = w.getSheet("Sheet0");
        Row r = s.getRow(4);
        Cell c = r.getCell(3);
        int ct = c.getCellType();
        System.out.println(ct);
        String scv = c.getStringCellValue();
        if(scv.equals("Rahul Thangavel")) {
        	c.setCellValue("Santhi Sabitha");
        	FileOutputStream o = new FileOutputStream(f);
        	w.write(o);
        	
        	
        	System.out.println("Read and Update");
        }
        
	}

}
