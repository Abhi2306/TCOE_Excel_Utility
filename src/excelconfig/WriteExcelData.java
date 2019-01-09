package excelconfig;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelData {

	public static void WriteExcel(String FilePath, String FileName, String SheetName) throws IOException {
		
		File file = new File(FilePath+ "\\" +FileName);
		
		FileInputStream InputStream = new FileInputStream(file);
		Workbook wb = null;
		
		if (FileName.substring(FileName.indexOf(".")).equalsIgnoreCase(".xlsx")) {
			
			wb = new XSSFWorkbook(InputStream);
		} else if(FileName.substring(FileName.indexOf(".")).equalsIgnoreCase(".xls")){
			
			wb = new HSSFWorkbook(InputStream);
		}
		
		Sheet sheet = wb.getSheet(SheetName);
		
		Row row = sheet.getRow(0);
		
		int countOfRow = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		Row newRow = sheet.createRow(countOfRow+1);
		
		int count=0;
		
		for(int i=0;i<row.getLastCellNum();i++) {
			
			Cell cell = newRow.createCell(i);
			
			cell.setCellValue("Abhi_"+i);
			
			count++;
		}
		
		InputStream.close();
		
		FileOutputStream outputStream = new FileOutputStream(file);
		
		wb.write(outputStream);
		
		outputStream.close();
		
		if(count>0) {
			System.out.println("Data is successfully written on the sheet..!!");
		}else {
			System.out.println("Please check the code..!!");
		}
		
	}
}
