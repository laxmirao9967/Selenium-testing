package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemo1 {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet Info");
		
		Object empdata[][]= { {"EmpID","Name","Job"},
				{101,"David","Engineer"},
				{102,"Smith","Manager"},
				{103,"Scott","Analyst"}	
		};
		
	
		//Using for loop
		/*int rows=empdata.length;
		int cols=empdata[0].length;
		System.out.println(rows); //4
		System.out.println(cols); //3
		
		for(int r=0;r<rows;r++)
		{
			XSSFRow row=sheet.createRow(r);//to create a row
			
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell=row.createCell(c);
				Object value=empdata[r][c];
				
				if(value instanceof String)
					cell.setCellValue((String)value);//we are type casting bczs value is in object.
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
		*/
		
		//Using for ...each loop
		int rowcount=0;
		for(Object emp[]:empdata)
		{
			XSSFRow row=sheet.createRow(rowcount++);
			int columncount=0;
			for(Object value:emp)
			{
				XSSFCell cell=row.createCell(columncount++);
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
		
//		String filePath=".\\Datafiles\\employee.xlsx";
		String filePath=".\\Datafiles\\employee_using_foreach_loop.xlsx";
		FileOutputStream outstream=new FileOutputStream(filePath);
		workbook.write(outstream);

		outstream.close();
		System.out.println("Employee.xlsx file written successfully");
		

	}

}

