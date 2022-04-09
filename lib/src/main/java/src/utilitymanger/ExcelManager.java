package src.utilitymanger;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelManager {

    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;

    public static void setExcelFile() throws Exception {
        try {
            FileInputStream ExcelFile = new FileInputStream(System.getProperty("user.dir") + "/src/main/resources/TestData.xlsx");
            ExcelWBook = new XSSFWorkbook(ExcelFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getCellData(int RowNum, int ColNum, String SheetName) throws Exception {
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
        Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
		String data = null;
		if(Cell.getCellType()== CellType.STRING)
			data = Cell.getStringCellValue();
		else if(Cell.getCellType()==CellType.NUMERIC)
			data = String.valueOf(Cell.getNumericCellValue());
        return data;
    }

    public static int getRowCount(String SheetName) {
        int iNumber = 0;
        try {
            ExcelWSheet = ExcelWBook.getSheet(SheetName);
            iNumber = ExcelWSheet.getLastRowNum() + 1;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return iNumber;
    }

    public static int getRowContains(String sTestCaseName, int colNum, String SheetName) throws Exception {
        int iRowNum = 0;
        try {
            //ExcelWSheet = ExcelWBook.getSheet(SheetName);
            int rowCount = ExcelManager.getRowCount(SheetName);
            for (; iRowNum < rowCount; iRowNum++) {
                if (ExcelManager.getCellData(iRowNum, colNum, SheetName).equalsIgnoreCase(sTestCaseName)) {
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return iRowNum;
    }


    @SuppressWarnings("static-access")
    public static void setCellData(String Result, int RowNum, int ColNum, String SheetName) throws Exception {
        try {

            ExcelWSheet = ExcelWBook.getSheet(SheetName);
            Row = ExcelWSheet.getRow(RowNum);
            Cell = Row.getCell(ColNum);
            if (Cell == null) {
                Cell = Row.createCell(ColNum);
                Cell.setCellValue(Result);
            } else {
                Cell.setCellValue(Result);
            }
            FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir") + "/src/main/resources/TestData.xlsx");
            ExcelWBook.write(fileOut);
            //fileOut.flush();
            fileOut.close();
            ExcelWBook = new XSSFWorkbook(new FileInputStream(System.getProperty("user.dir") + "/src/main/resources/TestData.xlsx"));
        } catch (Exception e) {
            e.printStackTrace();

        }
    }


}