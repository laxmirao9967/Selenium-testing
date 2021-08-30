package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell2 {

	public static void main(String[] args) throws IOException {
		String path=".\\Datafiles\\Books.xlsx";
		FileInputStream fin=new FileInputStream(path);
		XSSFWorkbook workbook=new XSSFWorkbook(fin);
		
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");
		
		fin.close();
		
		FileOutputStream fos=new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		
		System.out.println("Updated successfully");
		


	}

}
