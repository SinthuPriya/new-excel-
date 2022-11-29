package ExcelCreation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceloperation {

	
	public static void main(String[] args) throws IOException {
    File f=new File("F:\\Users\\Sinthuja\\eclipse-workspace\\EXCELOPERATION\\excel\\New Microsoft Excel Worksheet.xlsx");
	FileInputStream  stream= new FileInputStream(f);
	XSSFWorkbook workbook=new XSSFWorkbook(stream);
	XSSFSheet sheet=workbook.getSheetAt(0);
	String entry1=sheet.getRow(1).getCell(1).getStringCellValue();
	System.out.println("The data in the box is" +entry1);
	System.out.println("sindhu");
	workbook.close();
	
		
		

	}

}
