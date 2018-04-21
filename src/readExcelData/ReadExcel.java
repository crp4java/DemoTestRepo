package readExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		
		File src = new File("C:\\Users\\Thousif\\Desktop\\TestData.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheetAt(0);
		String data0 =  sheet1.getRow(0).getCell(0).getStringCellValue();
		System.out.println(data0);
		String data1 = sheet1.getRow(0).getCell(1).getStringCellValue();
		System.out.println(data1);
		
		wb.close();

	}

}
