package exceldata;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class exceldata {
	
	
	//This is new code after github
	@Test
	public ArrayList<String> getdata(String testcase) throws IOException
	{
	
		FileInputStream fs = new FileInputStream("C:\\Users\\shreyas.kolambakar\\Documents\\Selenium resource\\datadriven\\data.xlsx"); 
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		
		int sheets = wb.getNumberOfSheets();
		
		ArrayList<String> a = new ArrayList();
		
		for (int i=0;i<sheets;i++)
		{
			
			if(wb.getSheetName(i).equalsIgnoreCase("data"))
			{
			XSSFSheet sheet = wb.getSheetAt(i); 
			Iterator<Row> rows = sheet.iterator(); //sheet is collection of rows
			Row firstrow = rows.next();
			Row secondrow = rows.next();
			Iterator<Cell> cell = secondrow.cellIterator(); //row is collection of cells
			int k = 0;
			int column = 0;
			while(cell.hasNext()) //checks if cell has value
			{
				Cell value = cell.next();
				if(value.getStringCellValue().equalsIgnoreCase("TC"))
				{
					column = k;
					
				}
				k++;
			}
			
			System.out.println(column);
			
			while(rows.hasNext())
			{
			 
				Row r = rows.next();
				if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcase))
                  {
	               Iterator<Cell> ce= r.cellIterator();
	               while(ce.hasNext())
	               {
	            	   
	            	   Cell c = ce.next();
	            	   if(c.getCellTypeEnum()==CellType.STRING)
	            	   {
	            		   a.add(c.getStringCellValue()); 
	            		   
	            	   }
	            	   else
	            	   {
	            		a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
	            		
	            	   }
	            	   
	               }
                  }
				
				
			}
			
			
			
			}
		}
		return a;
	}

	

}
