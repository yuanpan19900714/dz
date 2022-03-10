package com.yp.hbl.service;

import com.yp.hbl.entity.Constant;
import com.yp.hbl.entity.Consumer;
import com.yp.hbl.entity.Dz;
import com.yp.hbl.util.ExcelUtil;
import com.yp.hbl.util.FileUtil;
import com.yp.hbl.util.NumberUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;


public class DzExcel {
    public static List<Dz> readExcel(String filePath, String sheetName) {
        List<Dz> dataLst = null;
        InputStream is = null;
        try {
            //验证文件是否合法
            if (ExcelUtil.validateExcel(filePath)) {
                return null;
            }
            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (ExcelUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003, sheetName);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            FileUtil.closeInputStream(is);
        }
        return dataLst;
    }

    private static List<Dz> read(InputStream inputStream, boolean isExcel2003, String sheetName) {
        List<Dz> dataLst = null;
        Workbook wb;
        try {
            //根据版本选择创建Workbook的方式
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }
            if (Constant.dataCompareWorkbook.SRC_SHEET_NAME.equals(sheetName)) {
                dataLst = readSrc(wb);
            }
            if (Constant.dataCompareWorkbook.HIS_SHEET_NAME.equals(sheetName)) {
                dataLst = readHis(wb);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataLst;
    }

    private static List<Dz> readHis(Workbook wb) {
        int totalRows;
        List<Dz> dataLst = new ArrayList<>();
        // 得到历史数据sheet
        Sheet sheet = wb.getSheet(Constant.dataCompareWorkbook.HIS_SHEET_NAME);
        // 得到Excel的行数
        totalRows = sheet.getPhysicalNumberOfRows();
        //从第二行开始读取数据
        for (int r = 1; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            if (row.getCell(0) == null) {
                continue;
            }
            Dz dz = new Dz();
            //第1列，店名编号
            dz.setShopNameNum(row.getCell(0).getStringCellValue());
            //第2列，店名
            dz.setShopName(row.getCell(1).getStringCellValue());
            //第3列，单号
            dz.setOrderNum(row.getCell(2).getStringCellValue());
            //第4列，日期
            dz.setDate(new SimpleDateFormat("yyyy-MM-dd").format(row.getCell(3).getDateCellValue()));
            //第5列，产品编号
            dz.setProductNum(row.getCell(4).getStringCellValue());
            //第6列，产品名称
            dz.setProductName(row.getCell(5).getStringCellValue());
            //第7列，单价
            dz.setUntiPrice(row.getCell(6).getNumericCellValue());
            //第8列，数量
            dz.setQuantity((int) (row.getCell(7).getNumericCellValue()));
            //第9列，金额
            dz.setAmount(row.getCell(8).getNumericCellValue());
            //第10列，对方价
            dz.setOtherPrice(row.getCell(9).getNumericCellValue());
            //第11列，对方单价
            dz.setOtherUnitPrice(NumberUtil.getRound(row.getCell(10).getNumericCellValue(), 2));
            //第12列，差价
            dz.setDisparity(NumberUtil.getRound(row.getCell(11).getNumericCellValue(), 2));
            //第13列，差额
            dz.setBalance(NumberUtil.getRound(row.getCell(12).getNumericCellValue(), 2));
            // 保存第r行
            dataLst.add(dz);
        }
        return dataLst;
    }

    private static List<Dz> readSrc(Workbook wb) {
        int totalRows;
        List<Dz> dataLst = new ArrayList<>();
        // 得到源数据sheet
        Sheet sheet = wb.getSheet(Constant.dataCompareWorkbook.SRC_SHEET_NAME);
        // 得到Excel的行数
        totalRows = sheet.getPhysicalNumberOfRows();
        //从第二行开始读取数据
        for (int r = 1; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            if (row.getCell(0) == null) {
                continue;
            }
            Dz dz = new Dz("", "", "");
            //第1列，店名编号
            dz.setShopNameNum(row.getCell(0).getStringCellValue());
            //第2列，店名
            String shopName = row.getCell(1).getStringCellValue();
            dz.setShopName(Dz.dealShopName(shopName));
            //第3列，单号
            dz.setOrderNum(row.getCell(2).getStringCellValue());
            //第4列，日期
            dz.setDate(new SimpleDateFormat("yyyy-MM-dd").format(row.getCell(3).getDateCellValue()));
            //第5列，产品编号
            dz.setProductNum(row.getCell(4).getStringCellValue());
            //第6列，产品名称
            dz.setProductName(row.getCell(5).getStringCellValue());
            //第7列，单价
            dz.setUntiPrice(row.getCell(6).getNumericCellValue());
            //第8列，数量
            dz.setQuantity((int) (row.getCell(7).getNumericCellValue()));
            //第9列，金额
            dz.setAmount(row.getCell(8).getNumericCellValue());
            // 保存第r行
            dataLst.add(dz);
        }
        return dataLst;
    }

    public static void writeExcel(String destFilePath, String destSheetName, List<Dz> hisDataList, List<Dz> srcDataList) {
        OutputStream out = null;
        try {
            // 读取Excel文档
            File destFile = new File(destFilePath);
            Workbook workBook = ExcelUtil.getWorkbok(destFile);
            // sheet 对应一个工作页
            Sheet sheet = workBook.getSheet(destSheetName);
            //删除原有数据，从第2行开始，J至M列
            int rowNumber = sheet.getLastRowNum();
            for (int i = 1; i <= rowNumber; i++) {
                Row row = sheet.getRow(i);
                for (int k = 9; k < 13; k++) {
                    Cell cell = row.getCell(k);
                    if (cell != null) {
                        row.removeCell(cell);
                    }
                }
            }
            // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(destFilePath);
            workBook.write(out);

            //初始化单元格格式对象
            CellStyle cellStyle = workBook.createCellStyle();
            //创建格式化对象
            DataFormat dataFormat = workBook.createDataFormat();
            //保留2位小数
            cellStyle.setDataFormat(dataFormat.getFormat("#,##0.00"));
            //写一行
            for (int i = 0; i < srcDataList.size(); i++) {
                Dz srcDz = srcDataList.get(i);
                String srcShopNameNum = srcDz.getShopNameNum();
                String srcProductNum = srcDz.getProductNum();
                List<Dz> comparedDataList = new ArrayList<>();
                for (Dz hisDz : hisDataList) {
                    String hisShopNameNum = hisDz.getShopNameNum();
                    String hisProductNum = hisDz.getProductNum();
                    if (srcShopNameNum == null || srcProductNum == null || hisShopNameNum == null || hisProductNum == null) {
                        continue;
                    }
                    if (srcShopNameNum.equals(hisShopNameNum) && srcProductNum.equals(hisProductNum)) {
                        comparedDataList.add(hisDz);
                    }
                }
                //按日期倒序排列
                comparedDataList.sort((a, b) -> b.getDate().compareTo(a.getDate()));
                Dz comDz = new Dz();
                if (comparedDataList.size() > 0) {
                    comDz = comparedDataList.get(0);
                }
                //取最后日期的对方价、对方单价、差价、差额等值写入excel(从第2行开始)
                Row row = sheet.getRow(i + 1);
                int rowNum = i + 2;
                row.createCell(9).setCellValue(comDz.getOtherPrice());
                srcDz.setOtherPrice(comDz.getOtherPrice());

                Cell cell10 = row.createCell(10);
                cell10.setCellFormula("J" + rowNum + "*" + Dz.BS);
                cell10.setCellStyle(cellStyle);
                srcDz.setOtherUnitPrice(comDz.getOtherUnitPrice());

                Cell cell11 = row.createCell(11);
                cell11.setCellFormula("G" + rowNum + "-" + "K" + rowNum);
                cell11.setCellStyle(cellStyle);
                srcDz.setDisparity(comDz.getDisparity());

                Cell cell12 = row.createCell(12);
                cell12.setCellFormula("H" + rowNum + "*(G" + rowNum + "-" + "K" + rowNum + ")");
                cell12.setCellStyle(cellStyle);
                srcDz.setBalance(comDz.getBalance());
            }
            //使excel内部公式生效
            workBook.setForceFormulaRecalculation(true);
            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(destFile);
            workBook.write(out);
            System.out.println("生成对账成功");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            FileUtil.closeOutputStream(out);
        }
    }

    public static List<Consumer> readExcel1(String filePath, String sheetName) {
        List<Consumer> dataLst = null;
        InputStream is = null;
        try {
            //验证文件是否合法
            if (ExcelUtil.validateExcel(filePath)) {
                return null;
            }
            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (ExcelUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read1(is, isExcel2003, sheetName);

        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            FileUtil.closeInputStream(is);
        }
        return dataLst;
    }

    private static List<Consumer> read1(InputStream is, boolean isExcel2003, String sheetName) {
        List<Consumer> dataLst = null;
        Workbook wb;
        try {
            //根据版本选择创建Workbook的方式
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {
                wb = new XSSFWorkbook(is);
            }
            dataLst = readConsumer(wb, sheetName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataLst;
    }

    private static List<Consumer> readConsumer(Workbook wb, String sheetName) {
        int totalRows;
        List<Consumer> dataLst = new ArrayList<>();
        // 得到客户数据sheet
        Sheet sheet = wb.getSheet(sheetName);
        // 得到Excel的行数
        totalRows = sheet.getPhysicalNumberOfRows();
        //从第二行开始读取数据
        for (int r = 1; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            if (row.getCell(0) == null) {
                continue;
            }
            Consumer consumer = new Consumer();
            //第1列，区域
            consumer.setArea(row.getCell(0).getStringCellValue());
            //第2列，客户编号
            consumer.setConsumerNum(row.getCell(1).getStringCellValue());
            //第3列，客户名称
            consumer.setConsumerName(row.getCell(2).getStringCellValue());
            //第4列，订单号码
            consumer.setOrderNum(row.getCell(3).getStringCellValue());
            //第5列，收货日期
            consumer.setDeliveryDate(new SimpleDateFormat("yyyy-MM-dd").format(row.getCell(4).getDateCellValue()));
            //第6列，客户未税金额
            consumer.setConsumerNoTaxAmount(row.getCell(5).getNumericCellValue());
            // 保存第r行
            dataLst.add(consumer);
        }
        return dataLst;
    }

    public static void setCellStyle(String filePath) {
        InputStream is = null;
        try {
            //验证文件是否合法
            if (ExcelUtil.validateExcel(filePath)) {
                return;
            }
            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (ExcelUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            setCellStyle(is, isExcel2003, filePath);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            FileUtil.closeInputStream(is);
        }

    }

    private static void setCellStyle(InputStream is, boolean isExcel2003, String filePath) {
        Workbook wb;
        try {
            //根据版本选择创建Workbook的方式
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {
                wb = new XSSFWorkbook(is);
            }
            setCellStyle(wb, filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setCellStyle(Workbook wb, String filePath) {
        //创建格式化对象
        DataFormat dataFormat = wb.createDataFormat();
        CellStyle textStyle = wb.createCellStyle();
        textStyle.setDataFormat(dataFormat.getFormat("@"));
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(dataFormat.getFormat("yyyy-m-d"));
        Sheet sheet1 = wb.getSheet(Constant.dataCompareWorkbook.SRC_SHEET_NAME);//公司数据分项
        setSheetCellStyle(textStyle, dateStyle, sheet1);
        Sheet sheet2 = wb.getSheet(Constant.dataCompareWorkbook.CONSUMER_SHEET_NAME);//客户数据
        int rowNum2 = sheet2.getLastRowNum();
        for (int i = 1; i < rowNum2; i++) {
            Row row = sheet2.getRow(i);
            Cell cell = row.getCell(1);//客户编号
            cell.setCellStyle(textStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.STRING);

            cell = row.getCell(3);//订单号码
            cell.setCellStyle(textStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.STRING);

            cell = row.getCell(4);//收货日期
            cell.setCellStyle(dateStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.NUMERIC);
        }
        sheet2.autoSizeColumn(4);
        Sheet sheet3 = wb.getSheet(Constant.dataCompareWorkbook.HIS_SHEET_NAME);//上期数据
        setSheetCellStyle(textStyle, dateStyle, sheet3);


        // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
        FileOutputStream out;
        try {
            out = new FileOutputStream(filePath);
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setSheetCellStyle(CellStyle textStyle, CellStyle dateStyle, Sheet sheet) {
        int rowNum = sheet.getLastRowNum();
        for (int i = 1; i < rowNum; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);//店名编号
            cell.setCellStyle(textStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.STRING);
            cell = row.getCell(2);//单号
            cell.setCellStyle(textStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.STRING);
            cell = row.getCell(3);//日期
            cell.setCellStyle(dateStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.NUMERIC);
            cell = row.getCell(4);//编号
            cell.setCellStyle(textStyle);//设置单元格格式为"文本"
            //cell.setCellType(CellType.STRING);
        }
        sheet.autoSizeColumn(3);
    }
}
