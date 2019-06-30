package test.org.exc;
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

public class change {
public static void main(String[] args) throws IOException {
	File f1=new File("C:\\Users\\RAGHU\\eclipse-workspace\\exc\\ex\\sample.xlsx");
	FileInputStream ft=new FileInputStream(f1);
	Workbook w=new XSSFWorkbook(ft);
	Sheet s=w.getSheet("raghu");
	Row r=s.getRow(0);Cell c=r.getCell(0);
	double t=c.getNumericCellValue();
	long l=(long) t;
	System.out.println(l);
	if(l==10)
	
		c.setCellValue(20);
	System.out.println("done");
	
FileOutputStream o=new FileOutputStream(f1);
w.write(o);
}
}
