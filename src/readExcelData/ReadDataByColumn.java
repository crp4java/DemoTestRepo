package readExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataByColumn {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file = new File("C:\\Users\\Thousif\\Desktop\\ReadData.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		XSSFRow row = sheet.getRow(0);
		XSSFCell cell = null;
		
		int colNum = 0;
		
		for(int i=0; i<row.getLastCellNum(); i++)
		{
			if(row.getCell(i).getStringCellValue().equals("UserName"))
				colNum=i;
		}
		row = sheet.getRow(1);
		cell = row.getCell(colNum);
		String pass = String.valueOf(cell.getStringCellValue());
		System.out.println("Value from Excel is :"+pass);

	}

}
