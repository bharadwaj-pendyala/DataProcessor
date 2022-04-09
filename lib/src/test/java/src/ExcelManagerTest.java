package src;

import org.testng.annotations.Test;
import src.utilitymanger.ExcelManager;

public class ExcelManagerTest {

    @Test
    public void testExcel() throws Exception {
        ExcelManager.setExcelFile();
        System.out.println("Value is:" +ExcelManager.getCellData(2,3, "Sheet1"));
    }

}
