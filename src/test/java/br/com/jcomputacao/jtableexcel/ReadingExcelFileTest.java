package br.com.jcomputacao.jtableexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * 15/06/2015 23:43:32
 * @author murilo
 */
public class ReadingExcelFileTest {

    public ReadingExcelFileTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    @Test
    public void hello() throws FileNotFoundException, IOException {
        //File excel = new File("cellStyleExample.xlsx");
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        InputStream is = classloader.getResourceAsStream("cellStyleExample.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        int sheets = wb.getNumberOfSheets();
        for (int sheet = 0; sheet < sheets; sheet++) {
            String sheetName = wb.getSheetName(sheet);
            XSSFSheet ws = wb.getSheet(sheetName);

            int rowNum = ws.getLastRowNum() + 1;
            int colNum = ws.getRow(0).getLastCellNum();
            String[][] data = new String[rowNum][colNum];

            for (int i = 0; i < rowNum; i++) {
                XSSFRow row = ws.getRow(i);
                if (row != null) {
                    for (int j = 0; j < colNum; j++) {
                        XSSFCell cell = row.getCell(j);
                        XSSFCellStyle style = cell.getCellStyle();
                        String value = cell.toString();
                        data[i][j] = value;
                        System.out.println("The value in sheet " + sheetName + ", line " + i + " and column " + j + " is " + value);
                        System.out.println("Fill Pattern       : " + style.getFillPatternEnum());
                        System.out.println("Background Color   : " + style.getFillBackgroundColor());
                        System.out.println("Foreground Color   : " + style.getFillForegroundColor());
                        System.out.println("Font theme Color   : " + style.getFont().getThemeColor());
                        System.out.println("Font family        : " + style.getFont().getFamily());
                        System.out.println("Font Color         : " + style.getFont().getColor());
                    }
                }
            }
        }

        System.out.println("Ok");
    }

}