package br.com.jcomputacao.jtableexcel;

import java.util.List;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.util.ArrayList;
import java.util.Date;
import java.util.Locale;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
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

    private static final Locale LOCALE = new Locale("pt", "BR");
    private static final DateTimeFormatter FORMAT_DATE = DateTimeFormatter
            .ofPattern("dd/MM/yyyy", LOCALE);
    private static final DateTimeFormatter FORMAT_TIME = DateTimeFormatter
            .ofPattern("HH:mm:ss", LOCALE);
    
    private final TableModel tableModel;
    private final OutputStream destination;
    private String sheetName;
    private boolean xlsx = true;
    private List<TableModel> tableModels;
    private List<String> sheetNames;
    private CellStyle localDateCellStyle;
    private CellStyle localTimeCellStyle;

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
    
    public void addSheet(TableModel model, String sheetName) {
        if (tableModels == null) {
            this.tableModels = new ArrayList<TableModel>();
        }
        tableModels.add(model);
        if (sheetNames == null) {
            this.sheetNames = new ArrayList<String>();
        }
        sheetNames.add(sheetName);
    }
    
    public void addSheet(TableModel model) {
        String thisSheetName;
        if (this.tableModels == null) {
            thisSheetName = "Folha 2";
        } else {
            thisSheetName = "Folha " + tableModels.size() + 1;
        }
        addSheet(model, thisSheetName);
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
        createCellStyle(workbook);
        HSSFSheet sheet = workbook.createSheet(this.sheetName);
        createHeader(tableModel, sheet);
        createBody(tableModel, sheet);
        
        if (tableModels != null && !tableModels.isEmpty()) {
            for (int i = 0; i < tableModels.size(); i++) {
                TableModel thisTableModel = tableModels.get(i);
                String thisSheetName = sheetNames.get(i);
                sheet = workbook.createSheet(thisSheetName);
                createHeader(thisTableModel, sheet);
                createBody(thisTableModel, sheet);
            }
        }

        workbook.write(destination);
    }
    
    private void executeXlsx() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        if (this.sheetName == null || this.sheetName.trim().equals("")) {
            this.sheetName = "JTable";
        }
        createCellStyle(workbook);
        XSSFSheet sheet = workbook.createSheet(this.sheetName);
        createHeader(tableModel, sheet);
        createBody(tableModel, sheet);
        
                if (tableModels != null && !tableModels.isEmpty()) {
            for (int i = 0; i < tableModels.size(); i++) {
                TableModel thisTableModel = tableModels.get(i);
                String thisSheetName = sheetNames.get(i);
                sheet = workbook.createSheet(thisSheetName);
                
                createHeader(thisTableModel, sheet);
                createBody(thisTableModel, sheet);
            }
        }

        workbook.write(destination);
    }

    private void createCellStyle(Workbook workbook) {
        CreationHelper createHelper = workbook.getCreationHelper();
        localDateCellStyle = workbook.createCellStyle();
        localDateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
        localTimeCellStyle = workbook.createCellStyle();
        localTimeCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("[HH]:mm:ss"));
    }

    private void createHeader(TableModel thisTableModel, HSSFSheet sheet) {
        int cols = thisTableModel.getColumnCount();
        HSSFRow row = sheet.createRow(0);

        for (int i = 0; i < cols; i++) {
            String columnName = thisTableModel.getColumnName(i);
            HSSFCell cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new HSSFRichTextString(columnName));
        }
    }
    
    private void createHeader(TableModel thisTableModel, XSSFSheet sheet) {
        int cols = thisTableModel.getColumnCount();
        XSSFRow row = sheet.createRow(0);

        for (int i = 0; i < cols; i++) {
            String columnName = thisTableModel.getColumnName(i);
            XSSFCell cell = row.createCell(i);
            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new XSSFRichTextString(columnName));
        }
    }
    
    private void createBody(TableModel thisTableModel, HSSFSheet sheet) {
        int cols = thisTableModel.getColumnCount();
        int rows = thisTableModel.getRowCount();

        for (int i = 0; i < rows; i++) {
            HSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < cols; j++) {
                HSSFCell cell = row.createCell(j);

                Object value = thisTableModel.getValueAt(i, j);
                if (value != null) {
                    defineCell(cell, value);
                }
            }
        }
    }

    private void createBody(TableModel thisTableModel, XSSFSheet sheet) {
        int cols = thisTableModel.getColumnCount();
        int rows = thisTableModel.getRowCount();

        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < cols; j++) {
                XSSFCell cell = row.createCell(j);

                Object value = thisTableModel.getValueAt(i, j);
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
        } else if (value instanceof LocalDate) {
            cell.setCellStyle(localDateCellStyle);
            cell.setCellValue(DateUtil.getExcelDate(asDate((LocalDate) value)));
        } else if (value instanceof LocalTime) {
            cell.setCellStyle(localTimeCellStyle);
            cell.setCellValue(DateUtil.convertTime(((LocalTime) value).format(FORMAT_TIME)));
        } else {
            if (isDateValid((String) value)) {
                cell.setCellStyle(localDateCellStyle);
                cell.setCellValue(DateUtil.getExcelDate(asDate(LocalDate.parse((String) value, FORMAT_DATE))));
            } else if (isTimeValid((String) value)) {
                cell.setCellStyle(localTimeCellStyle);
                cell.setCellValue(DateUtil.convertTime((String) value));
            } else {
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                cell.setCellValue(new HSSFRichTextString(value != null ? value.toString() : ""));
            }
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
        } else if (value instanceof LocalDate) {
            cell.setCellStyle(localDateCellStyle);
            cell.setCellValue(DateUtil.getExcelDate(asDate((LocalDate) value)));
        } else if (value instanceof LocalTime) {
            cell.setCellStyle(localTimeCellStyle);
            cell.setCellValue(DateUtil.convertTime(((LocalTime) value).format(FORMAT_TIME)));
        } else {
            if (isDateValid((String) value)) {
                cell.setCellStyle(localDateCellStyle);
                cell.setCellValue(DateUtil.getExcelDate(asDate(LocalDate.parse((String) value, FORMAT_DATE))));
            } else if (isTimeValid((String) value)) {
                cell.setCellStyle(localTimeCellStyle);
                cell.setCellValue(DateUtil.convertTime((String) value));
            } else {
                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                cell.setCellValue(new XSSFRichTextString(value != null ? value.toString() : ""));
            }
        }
    }
    
    private Date asDate(LocalDate localDate) {
        return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public static boolean isDateValid(String date) {
        if (date == null) {
            return false;
        }
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/uuuu").withResolverStyle(ResolverStyle.STRICT);
        try {
            LocalDate d = LocalDate.parse(date, formatter);
            return true;
        } catch (DateTimeParseException ex) {
            return false;
        }
    }

    public static boolean isTimeValid(String time) {
        if (time == null) {
            return false;
        }
        DateTimeFormatter formatter = verifyTime(time);
        try {
            if (formatter == null) {
                return false;
            }
            LocalTime t = LocalTime.parse(time, formatter);
            return true;
        } catch (DateTimeParseException ex) {
            return false;
        }
    }

    private static DateTimeFormatter verifyTime(String time) {
        DateTimeFormatter formatter = null;
        int aux = time.length();
        switch (aux) {
            case 5:
                formatter = DateTimeFormatter.ofPattern("HH:mm").withResolverStyle(ResolverStyle.STRICT);
                break;
            case 7:
            case 8:
                formatter = DateTimeFormatter.ofPattern("HH:mm:ss").withResolverStyle(ResolverStyle.STRICT);
                break;
        }
        return formatter;
    }
}