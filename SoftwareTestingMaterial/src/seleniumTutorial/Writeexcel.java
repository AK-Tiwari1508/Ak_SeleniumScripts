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

public class Writeexcel {
	public static void main (String [] args) throws IOException{
		//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
		FileInputStream fis = new FileInputStream("F:\\Selenium\\Apachi POI\\Test1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		//call the getSheet() method of Workbook and pass the Sheet Name here. 
		//In this case I have given the sheet name as ?TestData? 
                //or if you use the method getSheetAt(), you can pass sheet number starting from 0. Index starts with 0.
		XSSFSheet sheet = workbook.getSheetAt(0);
		//XSSFSheet sheet = workbook.getSheetAt(0);
		//Now create a row number and a cell where we want to enter a value. 
		//Here im about to write my test data in the cell B2. It reads Column B as 1 and Row 2 as 1. Column and Row values start from 0.
		//The below line of code will search for row number 2 and column number 2 (i.e., B) and will create a space. 
                //The createCell() method is present inside Row class.
         Row row = sheet.createRow(1);
		Cell cell = row.createCell(0);
		//Now we need to find out the type of the value we want to enter. 
                //If it is a string, we need to set the cell type as string 
                //if it is numeric, we need to set the cell type as number
		cell.setCellValue("SoftwareTestingMaterial.com");
		FileOutputStream fos = new FileOutputStream("F:\\Selenium\\Apachi POI\\Test1.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("END OF WRITING DATA IN EXCEL");

}
}