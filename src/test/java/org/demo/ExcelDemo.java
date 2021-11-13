package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDemo {


public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\Udhay\\eclipse-workspace\\MavenDemo\\Excel\\Untitled 1.xlsx");
	FileInputStream fis=new FileInputStream(f);
	Workbook w=new XSSFWorkbook(fis);
	Sheet sh = w.getSheet("Sheet1");
	for(int i=0;i<sh.getPhysicalNumberOfRows();i++) {
		Row r = sh.getRow(i);
		for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
			Cell c = r.getCell(j);
			int type = c.getCellType();
			if(type==1) {
				String val = c.getStringCellValue();
				System.out.println(val);
			}
		    if(type==0) {
                boolean b = DateUtil.isCellDateFormatted(c);
            if(b) {
	            Date d = c.getDateCellValue();
	            SimpleDateFormat sim=new SimpleDateFormat("dd/MM/YYYY");
	            String date = sim.format(d);
	            System.out.println(date);
	            
	            
			}
            else {
            	double num = c.getNumericCellValue();
            	long l = (long) num;
            	String v = String.valueOf(l);
            System.out.println(v);
            }
		}
		
	}

}	
}
}

	