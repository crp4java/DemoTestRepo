package readExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataRead {
	
	public void readExcel(String filePath,String fileName,String sheetName) throws Exception{
		File file = new File(filePath+"\\"+fileName);
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = null;
		String fileExtensionName = fileName.substring(fileName.indexOf('.'));
		if(fileExtensionName.equals(".xlsx")){
			wb=new XSSFWorkbook(fis);
		}
		else if(fileExtensionName.equals(".xls")){
			wb = new HSSFWorkbook(fis);
		}
		
		Sheet Sheet =wb.getSheet(sheetName);
		int rowCount = Sheet.getLastRowNum()-Sheet.getFirstRowNum();
		for(int i=0; i<rowCount+1; i++){
			Row row = Sheet.getRow(i);
			for(int j=0; j<row.getLastCellNum();j++){
				System.out.println(row.getCell(j).getStringCellValue()+"     ");
			}
			System.out.println();
		}
		
		
		
		
	}
	
	

	public static void main(String[] args) throws Exception 
	{
		DataRead dr = new DataRead();
		
		dr.readExcel("C:\\Users\\Thousif\\Desktop", "TestData.xlsx","Sheet1");
	
		
	}

}
