package src.utilitymanger;

import java.io.*;
import java.util.*;

public class ExcelManager{
    
    XSSFWorkbook workbook;
    XSSFSheet sheet;

    public static List<Map<String,String>> getTestDetails(String sheetname){
		List<Map<String,String>> list = null;

		try(FileInputStream fs = new FileInputStream(System.getProperty("user.dir")+"/src/main/java/")) {
			XSSFWorkbook workbook = new XSSFWorkbook(fs);
			XSSFSheet sheet = workbook.getSheet(sheetname);

			int lastrownum = sheet.getLastRowNum();
			int lastcolnum = sheet.getRow(0).getLastCellNum();

			Map<String,String> map =null;
			list = new ArrayList<>();

			for(int i=1; i<=lastrownum;i++) { 
				map = new HashMap<>(); 
				for(int j=0;j<lastcolnum;j++) {
					String key= sheet.getRow(0).getCell(j).getStringCellValue();
					String value = sheet.getRow(i).getCell(j).getStringCellValue();
					map.put(key, value);
				}
				list.add(map);
			}

		} catch (FileNotFoundException e) {
			throw new InvalidPathForExcelException("Excel File you trying to read is not found");
		} catch (IOException e) {
			throw new FrameworkException("Some io exception happened  while reading excel file");
		}
		System.out.println(list);
		return list;
	}
    
}