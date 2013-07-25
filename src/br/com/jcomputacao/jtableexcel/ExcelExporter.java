package br.com.jcomputacao.jtableexcel;

import java.io.IOException;
import java.io.OutputStream;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 16/09/2011 11:42:53
 * @author Murilo
 */
public class ExcelExporter {

    private final TableModel tableModel;
    private final OutputStream destination;
    private String sheetName;
    private boolean xlsx = true;

    public ExcelExporter(TableModel model, OutputStream destination) {
        this.tableModel = model;
        this.destination = destination;
    }
    
    public void setExportToXls() {
        this.xlsx = false;
    }
    
    public void setExportToXlsx() {
        this.xlsx = true;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void execute() throws IOException {
        if(xlsx) {
            executeXlsx();
        } else {
            executeXls();
        }
    }
    
    private void executeXls() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        if (this.sheetName == null || this.sheetName.trim().equals("")) {
            this.sheetName = "JTable";
        }
        HSSFSheet sheet = workbook.createSheet(this.sheetName);

        createHeader(sheet);
        createBody(sheet);

        workbook.write(destination);
    }
    
    private void executeXlsx() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        if (this.sheetName == null || this.sheetName.trim().equals("")) {
            this.sheetName = "JTable";
        }
        XSSFSheet sheet = workbook.createSheet(this.sheetName);

        createHeader(sheet);
        createBody(sheet);

        workbook.write(destination);
    }

    private void createHeader(HSSFSheet sheet) {
        int cols = tableModel.getColumnCount();
        HSSFRow row = sheet.createRow(0);

        for (int i = 0; i < cols; i++) {
            String columnName = tableModel.getColumnName(i);
            HSSFCell cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new HSSFRichTextString(columnName));
        }
    }
    
    private void createHeader(XSSFSheet sheet) {
        int cols = tableModel.getColumnCount();
        XSSFRow row = sheet.createRow(0);

        for (int i = 0; i < cols; i++) {
            String columnName = tableModel.getColumnName(i);
            XSSFCell cell = row.createCell(i);
            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new XSSFRichTextString(columnName));
        }
    }
    
    private void createBody(HSSFSheet sheet) {
        int cols = tableModel.getColumnCount();
        int rows = tableModel.getRowCount();

        for (int i = 0; i < rows; i++) {
            HSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < cols; j++) {
                HSSFCell cell = row.createCell(j);

                Object value = tableModel.getValueAt(i, j);
                if (value != null) {
                    defineCell(cell, value);
                }
            }
        }
    }

    private void createBody(XSSFSheet sheet) {
        int cols = tableModel.getColumnCount();
        int rows = tableModel.getRowCount();

        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < cols; j++) {
                XSSFCell cell = row.createCell(j);

                Object value = tableModel.getValueAt(i, j);
                if (value != null) {
                    defineCell(cell, value);
                }
            }
        }
    }
    
    private void defineCell(HSSFCell cell, Object value) {
        if (value instanceof Double || value instanceof Float
                || value instanceof Long || value instanceof Integer) {
            cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
            cell.setCellValue(new Double(value.toString()));
        } else if (value instanceof Boolean) {
            cell.setCellType(HSSFCell.CELL_TYPE_BOOLEAN);
            cell.setCellValue(Boolean.valueOf(value.toString()));
//        } else if (value instanceof java.util.Date || value instanceof java.util.Calendar) {
//            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
//            cell.setCellValue(new Double(value.toString()));
        } else {
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new HSSFRichTextString(value.toString()));
        }
    }

    private void defineCell(XSSFCell cell, Object value) {
        if (value instanceof Double || value instanceof Float
                || value instanceof Long || value instanceof Integer) {
            cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
            cell.setCellValue(new Double(value.toString()));
        } else if (value instanceof Boolean) {
            cell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);
            cell.setCellValue(Boolean.valueOf(value.toString()));
//        } else if (value instanceof java.util.Date || value instanceof java.util.Calendar) {
//            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
//            cell.setCellValue(new Double(value.toString()));
        } else {
            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new XSSFRichTextString(value.toString()));
        }
    }

}
