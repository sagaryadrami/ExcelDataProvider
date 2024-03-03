package DataDriven.ExcelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Excel {

	
//	@Test
//	public void getExcel() throws IOException {
//		FileInputStream fis = new FileInputStream("C:\\Users\\Sagar yadrami\\OneDrive\\Desktop\\exceldriven.xlsx");
//		
//		XSSFWorkbook wb = new XSSFWorkbook(fis);
//		XSSFSheet sheet = wb.getSheetAt(0);
//		int rowCount = sheet.getPhysicalNumberOfRows();
//		XSSFRow row = sheet.getRow(0);
//		int columnCount = row.getLastCellNum();
//		Object data[][]=new Object[rowCount-1][columnCount];
//		for(int i=0;i<rowCount;i++) {
//			 row = sheet.getRow(i);
//			 System.out.println("outer loop started");
//			 for(int j=0;j<columnCount;j++) {
//			System.out.println(	 row.getCell(j));
//			 }
//			
//			 System.out.println("outer loop ended");
//		}
//	}
	
	
	
	
	
	
	
	
	DataFormatter f = new DataFormatter();
	

	@Test(dataProvider="drivetest")
	public void testcasedata(String communication,String greetings,String id) {
		System.out.println(communication+greetings+id);
	}

	@DataProvider(name="drivetest")
	public Object[][] getdata() throws IOException {
		FileInputStream fis = new FileInputStream("C:\\\\Users\\\\Sagar yadrami\\\\OneDrive\\\\Desktop\\\\exceldriven.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowcount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		short columncount = row.getLastCellNum();
		Object Data[][]=new Object[rowcount-1][columncount];
		for(int i=0;i<rowcount-1;i++) {
			row=sheet.getRow(i+1);
			if(row!=null) {
				for(int j=0;j<columncount;j++) {
					XSSFCell cell = row.getCell(j);
					Data[i][j]=f.formatCellValue(cell);
				}
			}
			
		}return Data;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
