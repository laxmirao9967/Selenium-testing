package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		//location of file
		String excelFilePath=".\\Datafiles\\Apachefiles.xlsx";
		//open the file in reading mode
		FileInputStream inputstream=new FileInputStream(excelFilePath);
		//to get the workbook from the file
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheet("Sheet1");
		//for index  workbook.getSheetAt(0);
		
		
		//USING FOR LOOP
		/*int rows=sheet.getLastRowNum();
		//System.out.println(rows);
		int cols=sheet.getRow(1).getLastCellNum();
		//System.out.println(cols);
		
		for(int r=0;r<=rows;r++)
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell=row.getCell(c);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print(" | ");
				
			}
			System.out.println();
		}
		*/
		///Iterator
		//iterator method will return all the rows and we can iterator all the rows i.e repeat all the rows and cells
		
		Iterator iterator=sheet.iterator();
		
		while(iterator.hasNext())
		{
			XSSFRow row=(XSSFRow) iterator.next();
			
			Iterator cellIterator=row.cellIterator();
			while(cellIterator.hasNext())
			{
				XSSFCell cell=(XSSFCell) cellIterator.next();
				
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print(" | ");
				
			}
			System.out.println();
		}
		
		
		

	}

}
