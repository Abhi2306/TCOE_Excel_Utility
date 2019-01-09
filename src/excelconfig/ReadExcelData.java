package excelconfig;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {

	/*public static void main(String[] args) throws IOException {
		
		String FilePath = "C:\\Users\\abhishek.khatod\\Documents\\eclipse_workspace\\yash.TCOEFrameWork\\src\\main\\java\\Resources";
		String FileName = "ReadData.xlsx";
		String SheetName = "FirstSheet";
		System.out.println(ReadExcelData.ReadExcel(FilePath, FileName, SheetName));
	}*/
	
	public static Map<String, String> ReadExcel(String FilePath, String FileName, String SheetName) throws IOException {
		
		//Map for storing the data from excel
		HashMap<String, String> ExcelData = new HashMap<String, String>();
		
		File file = new File(FilePath + "\\" + FileName);
		FileInputStream inputStream = new FileInputStream(file);
		Workbook wb = null;

		//To check the extension of file and create object accordingly
		if (FileName.substring(FileName.indexOf(".")).equalsIgnoreCase(".xlsx")) {
			wb = new XSSFWorkbook(inputStream);
		} else if(FileName.substring(FileName.indexOf(".")).equalsIgnoreCase(".xls")){
			wb = new HSSFWorkbook(inputStream);
		}

		// To get the sheet where data has to be read from
		Sheet sheet = wb.getSheet(SheetName);
		int count=0;
		for (int i = 0;i<sheet.getLastRowNum(); i++) {

			// To get the row
			Row row = sheet.getRow(i);
			
			//System.out.println("row : "+row.getCell(0).getStringCellValue());
			for(int j=1;j<row.getLastCellNum();j++) {
				
				ExcelData.put(row.getCell(0).getStringCellValue(), row.getCell(j).getStringCellValue());
				
				//loop for skipping blank cells
				if(row.getCell(j).getStringCellValue() != null) {
					
					System.out.println(row.getCell(0).getStringCellValue()+" --> "+row.getCell(j).toString());
					count++;
				}else {
					continue;
				}
			}
			System.out.println();
		}
		
		inputStream.close();
		
		if(count>0) {
			System.out.println("Data is successfully read from the sheet..!!");
		}else {
			System.out.println("Please make sure sheet is not blank");
		}
		
		
		for(Map.Entry<String, String> Map_Data : ExcelData.entrySet()) {
			
			System.out.println(Map_Data.getKey()+" : "+Map_Data.getValue());
		}
		
		return ExcelData;
	}
}
