package com.yp.hbl.service;

import com.yp.hbl.entity.*;
import com.yp.hbl.util.ExcelUtil;
import com.yp.hbl.util.FileUtil;
import com.yp.hbl.util.NumberUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.stream.Collectors;

public class DzWorkbook {
    private static final String[] columns = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
    private static Map<String, List<String>> workbookMap;
    public Map<String, CellStyle> cellStyleMap;

    static Map<String, List<String>> getWorkbookMap() {
        return workbookMap;
    }

    public static TreeMap<Integer, List<DzData>> getDzWorkbookData(List<Dz> srcDataList, List<Consumer> consumerDataList) {
        Map<String, List<Dz>> shopOrderMap = new HashMap<>();
        Map<String, List<Consumer>> consumerMap = new HashMap<>();
        if (srcDataList != null) {
            shopOrderMap = srcDataList.stream().collect(Collectors.groupingBy(s -> s.getShopName() + "_" + s.getShopNameNum() + "_" + s.getDate() + "_" + s.getOrderNum()));
        }
        if (consumerDataList != null) {
            consumerMap = consumerDataList.stream().collect(Collectors.groupingBy(s -> s.getConsumerName() + "_" + s.getConsumerNum() + "_" + s.getOrderNum()));
        }
        TreeMap<String, List<Dz>> sortedShopOrderMap = new TreeMap<>(Comparator.naturalOrder());
        sortedShopOrderMap.putAll(shopOrderMap);
        List<DzData> dzDataList = new ArrayList<>();
        for (String shopDateOrder : sortedShopOrderMap.keySet()) {
            String shopName = shopDateOrder.split("_")[0];
            String shopNameNum = shopDateOrder.split("_")[1];
            String date = shopDateOrder.split("_")[2];
            String orderNum = shopDateOrder.split("_")[3];
            DzData dzData = new DzData();
            double companyAmount = 0;
            double consumerAmount = 0;
            double consumerTaxAmount = 0;
            double discountAmount = 0;
            boolean isSame = true;
            for (Dz dz : sortedShopOrderMap.get(shopDateOrder)) {
                if (dz.getShopName().equals(shopName) && dz.getShopNameNum().equals(shopNameNum) && dz.getDate().equals(date) && dz.getOrderNum().equals(orderNum)) {
                    companyAmount += dz.getAmount();
                }
            }
            companyAmount = NumberUtil.getRound(companyAmount, 2);
            if (consumerMap.containsKey(shopName + "_" + shopNameNum + "_" + orderNum)) {
                if (consumerDataList != null) {
                    for (Consumer consumer : consumerDataList) {
                        if (consumer.getConsumerName().equals(shopName) && consumer.getConsumerNum().equals(shopNameNum) && consumer.getOrderNum().equals(orderNum)) {
                            consumerAmount = consumer.getConsumerNoTaxAmount();
                        }
                    }
                }
                consumerTaxAmount = NumberUtil.getRound(consumerAmount * Dz.BS, 2);
                discountAmount = NumberUtil.getRound(consumerTaxAmount * Consumer.DISCOUNT, 2);
            } else {
                isSame = false;
            }
            double billAmount = NumberUtil.getRound(consumerTaxAmount - discountAmount, 2);
            double balanceAmount = NumberUtil.getRound(companyAmount - consumerTaxAmount, 2);

            dzData.setShopNo(Constant.ShopNo.getShopNo(shopName));
            dzData.setShopNum(shopNameNum);
            dzData.setShopName(shopName);
            dzData.setDate(date);
            dzData.setOrderNum(orderNum);
            dzData.setCompanyAmount(companyAmount);
            dzData.setConsumerAmount(consumerAmount);
            dzData.setConsumerTaxAmount(consumerTaxAmount);
            dzData.setDiscountAmount(discountAmount);
            dzData.setBillAmount(billAmount);
            dzData.setBalanceAmount(balanceAmount);
            dzData.setSame(isSame);

            dzDataList.add(dzData);
        }
        dzDataList.sort(Comparator.comparing(a -> (a.getShopNo() + "_" + a.getOrderNum() + a.getDate())));
        Map<Integer, List<DzData>> dzDataMap = dzDataList.stream().collect(Collectors.groupingBy(DzData::getShopNo));
        TreeMap<Integer, List<DzData>> sortedDzDataMap = new TreeMap<>(Comparator.naturalOrder());
        sortedDzDataMap.putAll(dzDataMap);
        return sortedDzDataMap;
    }

    public static void writeDzWorkbook(String workbookName, TreeMap<Integer, List<DzData>> dzDataMap) {
        workbookMap = new HashMap<>();
        deleteAllDzWorkbook(workbookName);
        String defaultWorkbook = Constant.dzWorkbook.WORKBOOK_NAME;
        if (dzDataMap.size() > Constant.dzWorkbook.MAX_SHEET_NUM) {
            TreeMap<Integer, List<DzData>> dzDataCopyMap = new TreeMap<>(dzDataMap);
            int size = dzDataMap.size();
            int maxSheetNum = Constant.dzWorkbook.MAX_SHEET_NUM;
            int bookCount = (int) Math.ceil((double) size / (double) maxSheetNum);
            for (int i = 1; i < bookCount + 1; i++) {
                TreeMap<Integer, List<DzData>> dzMap = new TreeMap<>();
                for (int j = 0; j < maxSheetNum; j++) {
                    if (dzDataCopyMap.size() > 0) {
                        int firstKey = dzDataCopyMap.firstKey();
                        dzMap.put(firstKey, dzDataCopyMap.get(firstKey));
                        dzDataCopyMap.remove(firstKey);
                    }
                }
                String workbookName_i = workbookName.replace(defaultWorkbook, defaultWorkbook + i);
                List<String> shopList = getShopListFromMap(dzMap, new ArrayList<>());
                workbookMap.put(workbookName_i, shopList);
                writeDz2Workbook(workbookName_i, dzMap);
            }
        } else {
            List<String> shopList = getShopListFromMap(dzDataMap, new ArrayList<>());
            workbookMap.put(workbookName, shopList);
            writeDz2Workbook(workbookName, dzDataMap);
        }
    }

    private static List<String> getShopListFromMap(TreeMap<Integer, List<DzData>> dzDataMap, List<String> shopList) {
        for (int key : dzDataMap.keySet()) {
            List<DzData> dzList = dzDataMap.get(key);
            if (dzList.size() > 0) {
                DzData dz = dzList.get(0);
                String shopName = dz.getShopName();
                if (shopName.contains("-") && shopName.contains("店")) {
                    shopName = shopName.substring(shopName.lastIndexOf("-") + 1, shopName.lastIndexOf("店") + 1);
                }
                shopList.add(shopName);
            }
        }
        return shopList;
    }

    private static void writeDz2Workbook(String destFilePath, TreeMap<Integer, List<DzData>> dzDataMap) {
        OutputStream out = null;
        try {
            // 创建新的对账工作簿
            File destFile = new File(destFilePath);
            Workbook workBook = ExcelUtil.createWorkbook(destFile);
            CreationHelper creationHelper = workBook.getCreationHelper();
            //初始化单元格格式对象
            CellStyle cellStyle = workBook.createCellStyle();
            // 创建sheet页
            workBook.createSheet(Constant.dzWorkbook.SUM_SHEET_NAME);//创建汇总页
            int size = dzDataMap.size();
            TreeMap<Integer, List<DzData>> dzDataCopyMap = new TreeMap<>(dzDataMap);
            for (int j = 0; j < size; j++) {
                if (dzDataCopyMap.size() > 0) {
                    int firstKey = dzDataCopyMap.firstKey();
                    List<DzData> dzDataList = dzDataCopyMap.get(firstKey);
                    if (dzDataList.size() > 0) {
                        DzData dzData = dzDataList.get(0);
                        String shopName = dzData.getShopName();
                        if (shopName.contains("-") && shopName.contains("店")) {
                            shopName = shopName.substring(shopName.lastIndexOf("-") + 1, shopName.lastIndexOf("店") + 1);
                        }
                        workBook.createSheet(shopName);
                    }
                    dzDataCopyMap.remove(firstKey);
                }
            }
            //写入数据
            Sheet sumSheet = workBook.getSheetAt(0);
            Row firstRow = sumSheet.createRow(0);
            firstRow.createCell(0).setCellValue(Constant.dzWorkbook.SUM_SHEET_NAME);
            sumSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));//合并单元格

            Font font = workBook.createFont();
            font.setFontName("微软雅黑");
            font.setFontHeight((short) 14);
            font.setFontHeightInPoints((short) 14);
            cellStyle.setFont(font);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            firstRow.getCell(0).setCellStyle(cellStyle);

            Row secondRow = sumSheet.createRow(1);
            secondRow.createCell(0).setCellValue("序号");
            secondRow.createCell(1).setCellValue("客户编号");
            secondRow.createCell(2).setCellValue("店名");
            secondRow.createCell(3).setCellValue("公司金额");
            secondRow.createCell(4).setCellValue("客户金额");
            secondRow.createCell(5).setCellValue("客户含税金额");
            secondRow.createCell(6).setCellValue("折扣");
            secondRow.createCell(7).setCellValue("开票金额");
            secondRow.createCell(8).setCellValue("差异");
            //设置表头格式
            setHeadRowCellStyle(secondRow);

            dzDataCopyMap.putAll(dzDataMap);

            CellStyle num2PointStyle = workBook.createCellStyle();
            DataFormat num2PointFormat = workBook.createDataFormat();
            num2PointStyle.setDataFormat(num2PointFormat.getFormat("#,##0.00"));

            CellStyle linkStyle = workBook.createCellStyle();
            Font linkFont = workBook.createFont();
            linkFont.setUnderline((byte) 1);
            linkFont.setColor(IndexedColors.BLUE.index);
            linkStyle.setFont(linkFont);

            CellStyle sumRowStyle = workBook.createCellStyle();
            Font font1 = workBook.createFont();
            font1.setBold(true);
            sumRowStyle.setAlignment(HorizontalAlignment.RIGHT);
            sumRowStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
            sumRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            sumRowStyle.setDataFormat(num2PointFormat.getFormat("#,##0.00"));
            sumRowStyle.setFont(font1);
            for (int j = 0; j < size; j++) {
                if (dzDataCopyMap.size() > 0) {
                    int firstKey = dzDataCopyMap.firstKey();
                    List<DzData> dzDataList = dzDataCopyMap.get(firstKey);
                    int listSize = dzDataList.size();
                    if (listSize > 0) {
                        DzData dz = dzDataList.get(0);
                        String shopName = dz.getShopName();
                        if (shopName.contains("-") && shopName.contains("店")) {
                            shopName = shopName.substring(shopName.lastIndexOf("-") + 1, shopName.lastIndexOf("店") + 1);
                        }
                        Sheet sheet = workBook.getSheet(shopName);
                        Row row0 = sheet.createRow(0);
                        row0.createCell(0).setCellValue(dz.getShopNo());
                        row0.createCell(1).setCellValue(dz.getShopNum());
                        row0.createCell(2).setCellValue(dz.getShopName());
                        Row row1 = sheet.createRow(1);
                        row1.createCell(0).setCellValue("店名");
                        row1.createCell(1).setCellValue("日期");
                        row1.createCell(2).setCellValue("订单号码");
                        row1.createCell(3).setCellValue("公司金额");
                        row1.createCell(4).setCellValue("客户金额");
                        row1.createCell(5).setCellValue("客户含税金额");
                        row1.createCell(6).setCellValue("折  扣");
                        row1.createCell(7).setCellValue("开票金额");
                        row1.createCell(8).setCellValue("差  异");
                        row0.createCell(10).setCellValue("返回");
                        XSSFHyperlink backHyperlink = (XSSFHyperlink) creationHelper.createHyperlink(HyperlinkType.DOCUMENT);
                        backHyperlink.setAddress("'" + Constant.dzWorkbook.SUM_SHEET_NAME + "'!A1");
                        row0.getCell(10).setHyperlink(backHyperlink);
                        row0.getCell(10).setCellStyle(linkStyle);
                        setHeadRowCellStyle(row1);

                        XSSFDrawing xssfDrawing = (XSSFDrawing) sheet.createDrawingPatriarch();
                        XSSFComment xssfComment = xssfDrawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) 1, 2, (short) 4, 4));
                        xssfComment.setString(new XSSFRichTextString("不重复的订单整行用红底显示"));

                        for (int i = 0; i < listSize; i++) {
                            DzData dzData = dzDataList.get(i);
                            int rowIndex = i + 2;
                            int rowNum = rowIndex + 1;
                            Row row = sheet.createRow(rowIndex);
                            row.createCell(0).setCellValue(dzData.getShopName());
                            row.createCell(1).setCellValue(dzData.getDate());
                            row.createCell(2).setCellValue(dzData.getOrderNum());
                            row.createCell(3).setCellValue(dzData.getCompanyAmount());
                            row.createCell(4).setCellValue(dzData.getConsumerAmount());
                            Cell cell5 = row.createCell(5);
                            cell5.setCellFormula("E" + rowNum + "*" + Dz.BS);
                            Cell cell6 = row.createCell(6);
                            cell6.setCellFormula("F" + rowNum + "*" + Consumer.DISCOUNT);
                            Cell cell7 = row.createCell(7);
                            cell7.setCellFormula("F" + rowNum + "-" + "G" + rowNum);
                            Cell cell8 = row.createCell(8);
                            cell8.setCellFormula("D" + rowNum + "-" + "F" + rowNum);
                            for (int k = 5; k < 9; k++) {
                                Cell cell = row.getCell(k);
                                cell.setCellStyle(num2PointStyle);
                            }
                            if (!dzData.isSame()) {
                                CellStyle redRowStyle = workBook.createCellStyle();
                                redRowStyle.cloneStyleFrom(num2PointStyle);
                                redRowStyle.setFillForegroundColor(IndexedColors.RED.index);
                                redRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                for (int k = 0; k < 9; k++) {
                                    Cell cell = row.getCell(k);
                                    cell.setCellStyle(redRowStyle);
                                    cell.setCellComment(xssfComment);
                                }
                            }
                        }
                        Row row = sheet.createRow(listSize + 2);
                        row.createCell(0).setCellValue("小计");
                        sheet.addMergedRegion(new CellRangeAddress(listSize + 2, listSize + 2, 0, 2));
                        row.getCell(0).setCellStyle(sumRowStyle);
                        for (int i = 3; i < 9; i++) {
                            Cell cell = row.createCell(i);
                            cell.setCellStyle(sumRowStyle);
                            cell.setCellFormula("SUM(" + columns[i] + "3:" + columns[i] + (listSize + 2) + ")");
                        }
                        setDzWorkbookAutoSize(sheet);
                        int rowIndex = sheet.getLastRowNum() + 1;
                        int rowNum = j + 3;

                        Row sumRow = sumSheet.createRow(j + 2);
                        sumRow.createCell(0).setCellValue(j + 1);
                        sumRow.createCell(1).setCellValue(dz.getShopNum());
                        sumRow.createCell(2).setCellValue(shopName);
                        XSSFHyperlink hyperlink = (XSSFHyperlink) creationHelper.createHyperlink(HyperlinkType.DOCUMENT);
                        hyperlink.setAddress("'" + shopName + "'!A1");
                        sumRow.getCell(2).setHyperlink(hyperlink);
                        sumRow.getCell(2).setCellStyle(linkStyle);
                        sumRow.createCell(3).setCellFormula("'" + shopName + "'!D" + rowIndex);//必须加单引号，不然公式中有括号会报错
                        sumRow.createCell(4).setCellFormula("'" + shopName + "'!E" + rowIndex);
                        sumRow.createCell(5).setCellFormula("E" + rowNum + "*" + Dz.BS);
                        sumRow.createCell(6).setCellFormula("F" + rowNum + "*" + Consumer.DISCOUNT);
                        sumRow.createCell(7).setCellFormula("F" + rowNum + "-" + "G" + rowNum);
                        sumRow.createCell(8).setCellFormula("D" + rowNum + "-" + "F" + rowNum);
                        for (int i = 3; i < 9; i++) {
                            Cell cell = sumRow.getCell(i);
                            cell.setCellStyle(num2PointStyle);
                        }
                    }
                    dzDataCopyMap.remove(firstKey);
                }
            }
            setDzWorkbookAutoSize(sumSheet);
            //使excel内部公式生效
            workBook.setForceFormulaRecalculation(true);
            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(destFile);
            workBook.write(out);
            System.out.println("生成对账工作簿[" + destFile + "]成功");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            FileUtil.closeOutputStream(out);
        }
    }

    private static void deleteAllDzWorkbook(String destFilePath) {
        String parentFolder = destFilePath.substring(0, destFilePath.lastIndexOf(File.separator));
        File file = new File(parentFolder);
        List<String> fileList = FileUtil.getAllFileName(file, new ArrayList<>());
        if (fileList != null) {
            for (String fileName : fileList) {
                if (fileName.contains(Constant.dzWorkbook.WORKBOOK_NAME)) {
                    if (new File(fileName).delete()) {
                        System.out.println("删除文件[" + fileName + "]");
                    }
                }
            }
        }
    }

    static void setDzWorkbookAutoSize(Sheet sheet) {
        for (int i = 0; i < 11; i++) {
            sheet.autoSizeColumn(i, true);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 17 / 10);
        }
    }

    static void setHeadRowCellStyle(Row row) {
        Workbook wb = row.getSheet().getWorkbook();
        Font font = wb.createFont();
        font.setFontName("微软雅黑");
        font.setBold(true);
        font.setFontHeight((short) 10);
        font.setFontHeightInPoints((short) 10);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
            Cell cell = row.getCell(i);
            cell.setCellStyle(cellStyle);
        }
    }

}
