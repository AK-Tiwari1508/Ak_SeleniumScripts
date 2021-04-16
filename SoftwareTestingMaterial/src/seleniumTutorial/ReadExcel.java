package seleniumTutorial;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//How to read excel files using Apache POI
public class ReadExcel {
	public static void main (String [] args) throws IOException{
			FileInputStream fis = new FileInputStream("F:\\Selenium\\Apachi POI\\Test.xlsx");
			try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
				XSSFSheet sheet = wb.getSheetAt(0);
				System.out.println(sheet.getRow(0).getCell(0) +" "+ sheet.getRow(0).getCell(1));
				System.out.println(sheet.getRow(1).getCell(0) +" "+ sheet.getRow(1).getCell(1));
			}	
	}		
}