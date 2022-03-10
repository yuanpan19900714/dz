package com.yp.hbl.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

    public static List<List<String>> readExcel(String filePath) {
        List<List<String>> dataLst = null;
        InputStream is = null;
        try {
            // 验证文件是否合法
            if (validateExcel(filePath)) {
                return null;
            }
            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003);

        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            FileUtil.closeInputStream(is);
        }

        return dataLst;
    }

    private static List<List<String>> read(InputStream inputStream, boolean isExcel2003) {
        List<List<String>> dataLst = null;
        try {
            // 根据版本选择创建Workbook的方式
            Workbook wb;
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }
            dataLst = read(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataLst;
    }

    /**
     * @param wb Workbook
     */
    private static List<List<String>> read(Workbook wb) {
        int totalRows;
        int totalCells = 0;
        List<List<String>> dataLst = new ArrayList<>();
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到Excel的行数
        totalRows = sheet.getPhysicalNumberOfRows();
        // 得到Excel的列数
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        // 遍历行
        readSheet(totalRows, totalCells, dataLst, sheet);
        return dataLst;
    }

    private static void readSheet(int totalRows, int totalCells, List<List<String>> dataLst, Sheet sheet) {
        for (int r = 0; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            List<String> rowLst = new ArrayList<>();
            // 遍历列
            for (int c = 0; c < totalCells; c++) {
                Cell cell = row.getCell(c);
                String cellValue = "";
                if (null != cell) {
                    // 以下是判断数据的类型
                    switch (cell.getCellType()) {
                        case NUMERIC: // 数字
                            cellValue = cell.getNumericCellValue() + "";
                            break;
                        case STRING: // 字符串
                            cellValue = cell.getStringCellValue();
                            break;
                        case BOOLEAN: // Boolean
                            cellValue = cell.getBooleanCellValue() + "";
                            break;
                        case FORMULA: // 公式
                            cellValue = cell.getCellFormula() + "";
                            break;
                        case BLANK: // 空值
                            cellValue = "";
                            break;
                        case ERROR: // 故障
                            cellValue = "非法字符";
                            break;
                        default:
                            cellValue = "未知类型";
                            break;
                    }
                }
                rowLst.add(cellValue);
            }
            // 保存第r行的第c列
            dataLst.add(rowLst);
        }
    }

    public static boolean validateExcel(String filePath) {
        // 检查文件名是否为空或者是否是Excel格式的文件
        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {
            return true;
        }
        // 检查文件是否存在
        File file = new File(filePath);
        return !file.exists();
    }

    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    public static void writeExcel(List<List<String>> dataList, String finalXlsxPath) {
        OutputStream out = null;
        try {
            // 读取Excel文档
            File finalXlsxFile = new File(finalXlsxPath);
            Workbook workBook = getWorkbok(finalXlsxFile);
            // sheet 对应一个工作页
            Sheet sheet = workBook.getSheetAt(0);

            // 删除原有数据，除了属性列
            int rowNumber = sheet.getLastRowNum();
            for (int i = 1; i <= rowNumber; i++) {
                Row row = sheet.getRow(i);
                sheet.removeRow(row);
            }

            // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(finalXlsxPath);
            workBook.write(out);

            // 写一行
            for (int j = 0; j < dataList.size(); j++) {
                Row row = sheet.createRow(j);
                // 写一列
                List<String> datas = dataList.get(j);
                for (int k = 0; k < datas.size(); k++) {
                    row.createCell(k).setCellValue(datas.get(k));
                }
            }

            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(finalXlsxPath);
            workBook.write(out);
            System.out.println("数据导出成功");
        } catch (Exception e) {
            System.out.println("请创建一个空文件");
            e.printStackTrace();
        } finally {
            FileUtil.closeOutputStream(out);
        }

    }

    public static final String EXCEL_XLS = "xls";
    public static final String EXCEL_XLSX = "xlsx";

    public static Workbook getWorkbok(File file) throws IOException {
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(EXCEL_XLS)) { // Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) { // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    public static Workbook createWorkbook(File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(EXCEL_XLS)) { // Excel&nbsp;2003
            wb = new HSSFWorkbook();
        } else if (file.getName().endsWith(EXCEL_XLSX)) { // Excel 2007/2010
            wb = new XSSFWorkbook();
        } else {
            System.out.println("文件类型错误：" + file.getName());
        }
        OutputStream os = new FileOutputStream(file);
        if (wb != null) {
            wb.write(os);
        }
        return wb;
    }

    public static boolean isMergedCell(Sheet sheet, int rowIndex, int columnIndex) {
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (rowIndex >= firstRow && rowIndex <= lastRow) {
                if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    public static void setRowBorder(Row row, int startIndex, int endIndex, CellStyle borderStyle) {
        for (int i = startIndex; i < endIndex + 1; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            CellStyle cellStyle = cell.getCellStyle();
            cellStyle.cloneStyleFrom(borderStyle);
            cell.setCellStyle(cellStyle);
        }
    }

    public Sheet createSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        return sheet;
    }

    public Row createRow(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        return row;
    }

    public Cell createStringCell(Row row, int columnIndex, String cellValue, CellStyle cellStyle) {
        Cell cell = row.createCell(columnIndex);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(cellValue);
        return cell;
    }

}