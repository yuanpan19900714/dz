package com.yp.hbl.service;

import com.yp.hbl.entity.Constant;
import com.yp.hbl.entity.Dz;
import com.yp.hbl.entity.FkData;
import com.yp.hbl.util.ExcelUtil;
import com.yp.hbl.util.FileUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.stream.Collectors;


public class FkExcel {
    private static Map<String, CellStyle> cellStyleMap;
    private static Map<Integer, String> dateRangeMap;

    private static Map<Integer, String> getDateRangeMap() {
        return dateRangeMap;
    }

    public static void writeFkWorkbook(String destFilePath, TreeMap<Integer, List<FkData>> fkDataMap) {
        deleteAllFkWorkbook(destFilePath);
        if (fkDataMap.size() > Constant.fkWorkbook.MAX_SHEET_NUM) {
            TreeMap<Integer, List<FkData>> fkDataCopyMap = new TreeMap<>(fkDataMap);
            int size = fkDataMap.size();
            int maxSheetNum = Constant.dzWorkbook.MAX_SHEET_NUM;
            int bookCount = (int) Math.ceil((double) size / (double) maxSheetNum);
            for (int i = 1; i < bookCount + 1; i++) {
                TreeMap<Integer, List<FkData>> fkMap = new TreeMap<>();
                for (int j = maxSheetNum - 1; j >= 0; j--) {
                    if (fkDataCopyMap.size() > 0) {
                        int firstKey = fkDataCopyMap.firstKey();
                        fkMap.put(firstKey, fkDataCopyMap.get(firstKey));
                        fkDataCopyMap.remove(firstKey);
                    }
                }
                writeFk2Workbook(destFilePath.replace(Constant.fkWorkbook.WORKBOOK_NAME, Constant.fkWorkbook.WORKBOOK_NAME + i), fkMap);
            }
        } else {
            writeFk2Workbook(destFilePath, fkDataMap);
        }
    }

    private static Workbook createWorkbook(File destFile) {
        Workbook workBook = null;
        try {
            workBook = ExcelUtil.createWorkbook(destFile);
            CreationHelper creationHelper = workBook.getCreationHelper();
            cellStyleMap = new HashMap<>();

            CellStyle num2PointStyle = workBook.createCellStyle();
            DataFormat num2PointFormat = workBook.createDataFormat();
            num2PointStyle.setDataFormat(num2PointFormat.getFormat("#,##0.00"));
            cellStyleMap.put("num2PointFormat", num2PointStyle);

            CellStyle linkStyle = workBook.createCellStyle();
            Font linkFont = workBook.createFont();
            linkFont.setUnderline((byte) 1);
            linkFont.setColor(IndexedColors.BLUE.index);
            linkStyle.setFont(linkFont);
            cellStyleMap.put("linkStyle", linkStyle);

            CellStyle headStyle = workBook.createCellStyle();
            Font font = workBook.createFont();
            font.setFontName("????????????");
            font.setFontHeight((short) 16);
            font.setFontHeightInPoints((short) 16);
            font.setBold(true);
            headStyle.setFont(font);
            headStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put("headStyle", headStyle);

            CellStyle headStyle1 = workBook.createCellStyle();
            Font font1 = workBook.createFont();
            font1.setFontName("????????????");
            font1.setFontHeight((short) 12);
            font1.setFontHeightInPoints((short) 12);
            font1.setBold(true);
            headStyle1.setFont(font1);
            headStyle1.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put("headStyle1", headStyle1);

            CellStyle headStyle2 = workBook.createCellStyle();
            Font font2 = workBook.createFont();
            font2.setFontName("????????????");
            font2.setFontHeight((short) 11);
            font2.setFontHeightInPoints((short) 11);
            font2.setBold(true);
            headStyle2.setFont(font2);
            headStyle2.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put("headStyle2", headStyle2);

            CellStyle headStyle3 = workBook.createCellStyle();
            Font font3 = workBook.createFont();
            font3.setFontName("????????????");
            font3.setFontHeight((short) 12);
            font3.setFontHeightInPoints((short) 12);
            font3.setBold(true);
            headStyle3.setFont(font3);
            headStyle3.setAlignment(HorizontalAlignment.LEFT);
            cellStyleMap.put("headStyle3", headStyle3);

            CellStyle dateStyle = workBook.createCellStyle();
            dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy???mm???dd???"));
            cellStyleMap.put("dateStyle", dateStyle);

            CellStyle borderBottomStyle = workBook.createCellStyle();
            borderBottomStyle.setBorderBottom(BorderStyle.DOUBLE);
            borderBottomStyle.setFont(font1);
            borderBottomStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put("borderBottomStyle", borderBottomStyle);

            CellStyle borderBottomStyle1 = workBook.createCellStyle();
            borderBottomStyle1.setBorderBottom(BorderStyle.DOUBLE);
            borderBottomStyle1.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy???mm???dd???"));
            cellStyleMap.put("borderBottomStyle1", borderBottomStyle1);

            CellStyle borderRightStyle = workBook.createCellStyle();
            borderRightStyle.setBorderRight(BorderStyle.DOUBLE);
            cellStyleMap.put("borderRightStyle", borderRightStyle);

            CellStyle alignCenterStyle = workBook.createCellStyle();
            alignCenterStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put("alignCenterStyle", alignCenterStyle);


        } catch (IOException e) {
            e.printStackTrace();
        }
        return workBook;
    }

    private static void writeFk2Workbook(String destFilePath, TreeMap<Integer, List<FkData>> fkDataMap) {
        OutputStream out = null;
        try {
            // ???????????????????????????
            File destFile = new File(destFilePath);
            Workbook workBook = createWorkbook(destFile);
            CreationHelper creationHelper = workBook.getCreationHelper();
            // ??????sheet???
            workBook.createSheet(Constant.fkWorkbook.SUM_SHEET_NAME);//???????????????
            int size = fkDataMap.size();
            TreeMap<Integer, List<FkData>> fkDataCopyMap = new TreeMap<>(fkDataMap);
            makeSheets(workBook, size, fkDataCopyMap);
            //????????????
            Sheet sumSheet = workBook.getSheetAt(0);
            Row firstRow = sumSheet.createRow(0);
            firstRow.createCell(0).setCellValue("??????");
            firstRow.createCell(1).setCellValue("????????????");
            firstRow.createCell(2).setCellValue("??????");
            //??????????????????
            DzWorkbook.setHeadRowCellStyle(firstRow);

            fkDataCopyMap.putAll(fkDataMap);
            for (int j = 0; j < size; j++) {
                if (fkDataCopyMap.size() > 0) {
                    int firstKey = fkDataCopyMap.firstKey();
                    List<FkData> fkDataList = fkDataCopyMap.get(firstKey);
                    int listSize = fkDataList.size();
                    if (listSize > 0) {
                        FkData fk = fkDataList.get(0);
                        String shopName = fk.getShopName();
                        if (shopName.contains("-") && shopName.contains("???")) {
                            shopName = shopName.substring(shopName.lastIndexOf("-") + 1, shopName.lastIndexOf("???") + 1);
                        }
                        Sheet sheet = workBook.getSheet(shopName);
                        //???1???
                        Row row0 = sheet.createRow(0);
                        row0.createCell(0).setCellValue("MMM????????????");
                        row0.getCell(0).setCellStyle(cellStyleMap.get("headStyle"));
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));//???????????????
                        //???2???
                        Row row1 = sheet.createRow(1);
                        XSSFRichTextString value = new XSSFRichTextString("?????????????????????????????????(???????????????)");
                        Font underlineFont = workBook.createFont();
                        underlineFont.setUnderline(Font.U_SINGLE);
                        underlineFont.setFontName("????????????");
                        underlineFont.setFontHeight((short) 16);
                        underlineFont.setFontHeightInPoints((short) 16);
                        underlineFont.setBold(true);
                        value.applyFont(11, 18, underlineFont);
                        row1.createCell(0).setCellValue(value);
                        row1.getCell(0).setCellStyle(cellStyleMap.get("headStyle"));
                        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 9));//???????????????
                        //???3???????????????
                        sheet.createRow(2);
                        //???4???
                        Row row3 = sheet.createRow(3);
                        row3.createCell(0).setCellValue("??????????????????:");
                        row3.getCell(0).setCellStyle(cellStyleMap.get("headStyle1"));
                        row3.createCell(1).setCellValue("?????????");
                        row3.getCell(1).setCellStyle(cellStyleMap.get("headStyle1"));
                        for (int i = 2; i < 8; i++) {
                            row3.createCell(i);
                        }
                        row3.createCell(8).setCellValue("????????????:");
                        row3.getCell(8).setCellStyle(cellStyleMap.get("headStyle1"));
                        row3.createCell(9).setCellFormula("TODAY()");
                        row3.getCell(9).setCellStyle(cellStyleMap.get("dateStyle"));
                        CellRangeAddress region = CellRangeAddress.valueOf("A3:J3");
                        RegionUtil.setBorderBottom(BorderStyle.DOUBLE, region, sheet);
                        //???5???
                        Row row4 = sheet.createRow(4);
                        row4.createCell(0).setCellValue("????????????");
                        row4.getCell(0).setCellStyle(cellStyleMap.get("headStyle1"));
                        row4.createCell(1).setCellValue(fk.getShopNo());
                        row4.createCell(8).setCellValue("????????????");
                        row4.getCell(8).setCellStyle(cellStyleMap.get("headStyle1"));
                        row4.createCell(9).setCellValue(fk.getShopName());
                        //???6???
                        Row row5 = sheet.createRow(5);
                        row5.createCell(0).setCellValue("????????????");
                        row5.getCell(0).setCellStyle(cellStyleMap.get("headStyle1"));
                        row5.createCell(1);
                        String dateRange = getDateRangeMap().get(firstKey);
                        String startDate = dateRange.split("???")[0];
                        String endDate = dateRange.split("???")[1];
                        startDate = startDate.replaceFirst("-", "???").replace("-", "???") + "???";
                        endDate = endDate.replaceFirst("-", "???").replace("-", "???") + "???";
                        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 9));//???????????????
                        row5.getCell(1).setCellValue(startDate + " ??? " + endDate);
                        row5.getCell(1).setCellStyle(cellStyleMap.get("alignCenterStyle"));
                        //???7???
                        Row row6 = sheet.createRow(6);
                        row6.createCell(0).setCellValue("???????????????");
                        row6.getCell(0).setCellStyle(cellStyleMap.get("headStyle1"));
                        String workbookName = getWorkbookNameByShop(shopName);
                        String folder = workbookName.substring(0, workbookName.lastIndexOf(File.separator) + 1);
                        String workbook = workbookName.substring(workbookName.lastIndexOf(File.separator) + 1);
                        //???????????????????????????????????????????????????
                        Sheet dzSheet = ExcelUtil.getWorkbok(new File(workbookName)).getSheet(shopName);
                        int dzIndex = dzSheet.getLastRowNum() + 1;
                        row6.createCell(1).setCellFormula("'" + folder + "[" + workbook + "]" + shopName + "'!$D$" + dzIndex);
                        row6.getCell(1).setCellStyle(cellStyleMap.get("num2PointStyle"));
                        row6.createCell(2).setCellValue("??????????????????????????????");
                        row6.getCell(2).setCellStyle(cellStyleMap.get("headStyle1"));
                        sheet.addMergedRegion(new CellRangeAddress(6, 6, 2, 4));//???????????????
                        region = CellRangeAddress.valueOf("A5:J5");
                        RegionUtil.setBorderTop(BorderStyle.MEDIUM, region, sheet);
                        //????????????????????????????????????????????????????????????????????????
                        row6.createCell(5).setCellFormula("'" + folder + "[" + workbook + "]" + shopName + "'!$F$" + dzIndex);
                        row6.getCell(5).setCellStyle(cellStyleMap.get("num2PointStyle"));
                        row6.createCell(8).setCellValue("????????????");
                        row6.getCell(8).setCellStyle(cellStyleMap.get("headStyle1"));
                        row6.createCell(9).setCellFormula("B7-F7");
                        row6.getCell(9).setCellStyle(cellStyleMap.get("num2PointStyle"));
                        region = CellRangeAddress.valueOf("A6:J6");
                        RegionUtil.setBorderTop(BorderStyle.MEDIUM, region, sheet);
                        region = CellRangeAddress.valueOf("A4:J6");
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        //???8???
                        Row row7 = sheet.createRow(7);
                        row7.createCell(0).setCellValue("????????????");
                        row7.getCell(0).setCellStyle(cellStyleMap.get("headStyle1"));
                        row7.createCell(9).setCellValue("????????????");
                        row7.getCell(9).setCellStyle(cellStyleMap.get("headStyle1"));
                        region = CellRangeAddress.valueOf("A7:J7");
                        RegionUtil.setBorderBottom(BorderStyle.DASHED, region, sheet);
                        RegionUtil.setBorderTop(BorderStyle.MEDIUM, region, sheet);
                        //???9???
                        Row row8 = sheet.createRow(8);
                        row8.createCell(0).setCellValue("????????????");
                        row8.createCell(1).setCellValue("????????????");
                        row8.createCell(2).setCellValue("????????????");
                        row8.createCell(3).setCellValue("????????????");
                        row8.createCell(4).setCellValue("??????");
                        row8.createCell(5).setCellValue("????????????");
                        row8.createCell(6).setCellValue("?????????????????????");
                        row8.createCell(7).setCellValue("?????????????????????");
                        row8.createCell(8).setCellValue("??????");
                        for (int i = 0; i < 9; i++) {
                            row8.getCell(i).setCellStyle(cellStyleMap.get("headStyle2"));
                        }
                        region = CellRangeAddress.valueOf("A8:J8");
                        RegionUtil.setBorderBottom(BorderStyle.HAIR, region, sheet);
                        //????????????
                        for (int i = 0; i < listSize; i++) {
                            FkData fkData = fkDataList.get(i);
                            int rowIndex = i + 9;
                            Row row_i = sheet.createRow(rowIndex);
                            row_i.createCell(0).setCellValue(fkData.getProductNo());
                            row_i.createCell(1).setCellValue(fkData.getProductName());
                            row_i.createCell(2).setCellValue(fkData.getCompanyPrice());
                            row_i.createCell(3).setCellValue(fkData.getConsumerPrice());
                            row_i.createCell(4).setCellValue(fkData.getQuantity());
                            row_i.createCell(5).setCellValue(fkData.getBalanceAmount());
                            row_i.createCell(6).setCellValue(fkData.getInnerBalanceAmount());
                            row_i.createCell(7).setCellValue(fkData.getOuterBalanceAmount());
                            row_i.createCell(8).setCellValue(fkData.getDateRange());
                            for (int k = 2; k < 8; k++) {
                                if (k != 4) {
                                    row_i.getCell(k).setCellStyle(cellStyleMap.get("num2PointStyle"));
                                }
                            }
                        }
                        //?????????
                        int index = listSize + 9;
                        Row row = sheet.createRow(index);
                        row.createCell(0).setCellValue("?????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.createCell(2).setCellFormula("SUM(C9:C" + index + ")");
                        row.createCell(3).setCellFormula("SUM(D9:D" + index + ")");
                        row.createCell(4).setCellFormula("SUM(E9:E" + index + ")");
                        row.createCell(5).setCellFormula("SUM(F9:F" + index + ")");
                        row.createCell(9).setCellValue("?????????");
                        row.getCell(9).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("????????????????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 1));
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.createCell(9).setCellValue("???????????????");
                        row.getCell(9).setCellStyle(cellStyleMap.get("headStyle3"));
                        region = CellRangeAddress.valueOf("I7:I" + index);
                        RegionUtil.setBorderRight(BorderStyle.MEDIUM, region, sheet);
                        region = CellRangeAddress.valueOf("J7:J" + index);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        region = CellRangeAddress.valueOf("A" + index + ":J" + (index + 1));
                        RegionUtil.setBorderTop(BorderStyle.DOUBLE, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("?????????????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle2"));
                        region = CellRangeAddress.valueOf("A" + index + ":J" + index);
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??? ???1.????????????????????????????????????????????????????????????????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??? ???2.??????????????????????????????????????????????????????????????????TDS???TDM???????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??? ???3.???????????????ADF????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??? ???4.???????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(7).setCellValue("????????????????????????");
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 7, 8));//???????????????
                        row.getCell(7).setCellStyle(cellStyleMap.get("headStyle3"));
                        region = CellRangeAddress.valueOf("A" + index + ":J" + index);
                        RegionUtil.setBorderBottom(BorderStyle.DOUBLE, region, sheet);
                        region = CellRangeAddress.valueOf("A" + (index - 4) + ":J" + index);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("????????????????????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle2"));
                        region = CellRangeAddress.valueOf("A" + index + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 9));//???????????????
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??????????????????");
                        row.createCell(2).setCellValue("???????????????");
                        row.createCell(5).setCellValue("???????????????");
                        row.createCell(8).setCellValue("?????????????????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(2).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(5).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(8).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0);
                        row.createCell(2).setCellValue("?????????????????????");
                        row.createCell(5).setCellValue("?????????????????????");
                        row.getCell(2).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(5).setCellStyle(cellStyleMap.get("headStyle3"));
                        region = CellRangeAddress.valueOf("A" + (index - 1) + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        sheet.addMergedRegion(new CellRangeAddress(index - 1, index, 0, 0));//???????????????
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("??????????????????");
                        row.createCell(2).setCellValue("????????????????????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(2).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0);
                        region = CellRangeAddress.valueOf("A" + (index - 1) + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        sheet.addMergedRegion(new CellRangeAddress(index - 1, index, 0, 0));//???????????????
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("?????????????????????");
                        row.createCell(2).setCellValue("????????????????????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        row.getCell(2).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0);
                        region = CellRangeAddress.valueOf("A" + (index - 1) + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        sheet.addMergedRegion(new CellRangeAddress(index - 1, index, 0, 0));//???????????????
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("????????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0);
                        region = CellRangeAddress.valueOf("A" + (index - 1) + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        sheet.addMergedRegion(new CellRangeAddress(index - 1, index, 0, 0));//???????????????
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0).setCellValue("?????????");
                        row.getCell(0).setCellStyle(cellStyleMap.get("headStyle3"));
                        //?????????
                        index += 1;
                        row = sheet.createRow(index);
                        row.createCell(0);
                        sheet.addMergedRegion(new CellRangeAddress(index - 1, index, 0, 0));//???????????????
                        region = CellRangeAddress.valueOf("A" + (index - 1) + ":J" + (index + 1));
                        RegionUtil.setBorderBottom(BorderStyle.DOUBLE, region, sheet);
                        RegionUtil.setBorderRight(BorderStyle.DOUBLE, region, sheet);
                        //???????????????
                        row0.createCell(11).setCellValue("??????");
                        XSSFHyperlink backHyperlink = (XSSFHyperlink) creationHelper.createHyperlink(HyperlinkType.DOCUMENT);
                        backHyperlink.setAddress("'" + Constant.fkWorkbook.SUM_SHEET_NAME + "'!A1");
                        row0.getCell(11).setHyperlink(backHyperlink);
                        row0.getCell(11).setCellStyle(cellStyleMap.get("linkStyle"));

                        //?????????????????????
                        Row sumRow = sumSheet.createRow(j + 1);
                        sumRow.createCell(0).setCellValue(j + 1);
                        sumRow.createCell(1).setCellValue(fk.getShopNum());
                        sumRow.createCell(2).setCellValue(shopName);
                        XSSFHyperlink hyperlink = (XSSFHyperlink) creationHelper.createHyperlink(HyperlinkType.DOCUMENT);
                        hyperlink.setAddress("'" + shopName + "'!A1");
                        sumRow.getCell(2).setHyperlink(hyperlink);
                        sumRow.getCell(2).setCellStyle(cellStyleMap.get("linkStyle"));

                        DzWorkbook.setDzWorkbookAutoSize(sheet);

                    }
                    fkDataCopyMap.remove(firstKey);
                }
            }
            DzWorkbook.setDzWorkbookAutoSize(sumSheet);

            //???excel??????????????????
            workBook.setForceFormulaRecalculation(true);
            // ?????????????????????????????????????????????????????????????????????????????????sheet????????????????????????????????????
            out = new FileOutputStream(destFile);
            workBook.write(out);
            System.out.println("?????????????????????[" + destFile + "]??????");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            FileUtil.closeOutputStream(out);
        }
    }

    private static String getWorkbookNameByShop(String shopName) {
        String workbook = "";
        Map<String, List<String>> workbookMap = DzWorkbook.getWorkbookMap();
        for (String key : workbookMap.keySet()) {
            List<String> shopList = workbookMap.get(key);
            if (shopList.contains(shopName)) {
                workbook = key;
            }
        }
        return workbook;
    }

    private static void makeSheets(Workbook workBook, int size, TreeMap<Integer, List<FkData>> dzDataCopyMap) {
        for (int j = 0; j < size; j++) {
            if (dzDataCopyMap.size() > 0) {
                int firstKey = dzDataCopyMap.firstKey();
                List<FkData> fkDataList = dzDataCopyMap.get(firstKey);
                if (fkDataList.size() > 0) {
                    FkData fkData = fkDataList.get(0);
                    String shopName = fkData.getShopName();
                    if (shopName.contains("-") && shopName.contains("???")) {
                        shopName = shopName.substring(shopName.lastIndexOf("-") + 1, shopName.lastIndexOf("???") + 1);
                    }
                    workBook.createSheet(shopName);
                }
                dzDataCopyMap.remove(firstKey);
            }
        }
    }

    private static void deleteAllFkWorkbook(String fkWorkbookName) {
        String parentFolder = fkWorkbookName.substring(0, fkWorkbookName.lastIndexOf(File.separator));
        File file = new File(parentFolder);
        List<String> fileList = FileUtil.getAllFileName(file, new ArrayList<>());
        if (fileList != null) {
            for (String fileName : fileList) {
                if (fileName.contains(Constant.fkWorkbook.WORKBOOK_NAME)) {
                    if (new File(fileName).delete()) {
                        System.out.println("????????????[" + fileName + "]");
                    }
                }
            }
        }
    }

    public static TreeMap<Integer, List<FkData>> getFkWorkbookData(List<Dz> srcDataList) {
        Map<String, List<Dz>> dzMap = srcDataList.stream().collect(Collectors.groupingBy(s -> s.getShopName() + "_" + s.getShopNameNum()));
        List<FkData> fkDataList = new ArrayList<>();
        dateRangeMap = new HashMap<>();
        for (String key : dzMap.keySet()) {
            String shopName = key.split("_")[0];
            String shopNameNum = key.split("_")[1];
            int shopNo = Constant.ShopNo.getShopNo(shopName);
            List<Dz> dzList = dzMap.get(key);
            dzList.sort(Comparator.comparing(Dz::getDate));
            int size = dzList.size();
            if (size > 0) {
                String startDate = dzList.get(0).getDate();
                String endDate = dzList.get(size - 1).getDate();
                String shopDateRange = startDate + "???" + endDate;
                Map<String, List<Dz>> productMap = dzList.stream().collect(Collectors.groupingBy(s -> s.getProductNum() + "_" + s.getProductName()));
                for (String key1 : productMap.keySet()) {
                    FkData fkData = new FkData();
                    String productNo = key1.split("_")[0];
                    String productName = key1.split("_")[1];
                    List<Dz> productList = productMap.get(key1);
                    int size1 = productList.size();
                    if (size1 > 0) {
                        short quantity = 0;
                        double balanceAmount = 0;
                        double innerBalanceAmount = 0;
                        double outerBalanceAmount = 0;
                        productList.sort(Comparator.comparing(Dz::getDate));
                        double companyPrice = productList.get(0).getUntiPrice();
                        //????????????=[???????????????]??????????????????K????????????
                        double consumerPrice = productList.get(size1 - 1).getOtherUnitPrice();
                        String startDate1 = productList.get(0).getDate();
                        String endDate1 = productList.get(size1 - 1).getDate();
                        String dateRange = startDate1 + "???" + endDate1;
                        for (Dz dz1 : productList) {
                            quantity += dz1.getQuantity();
                            balanceAmount += dz1.getBalance();
                        }

                        fkData.setShopNo(shopNo);
                        fkData.setShopNum(shopNameNum);
                        fkData.setShopName(shopName);
                        fkData.setProductNo(productNo);
                        fkData.setProductName(productName);
                        fkData.setCompanyPrice(companyPrice);
                        fkData.setConsumerPrice(consumerPrice);
                        fkData.setQuantity(quantity);
                        fkData.setBalanceAmount(balanceAmount);
                        fkData.setInnerBalanceAmount(innerBalanceAmount);
                        fkData.setOuterBalanceAmount(outerBalanceAmount);
                        fkData.setDateRange(dateRange);
                        fkDataList.add(fkData);
                    }
                }
                dateRangeMap.put(shopNo, shopDateRange);
            }
        }
        fkDataList.sort(Comparator.comparing(FkData::getShopNo));
        Map<Integer, List<FkData>> dzDataMap = fkDataList.stream().collect(Collectors.groupingBy(FkData::getShopNo));
        TreeMap<Integer, List<FkData>> sortedFkDataMap = new TreeMap<>(Comparator.naturalOrder());
        sortedFkDataMap.putAll(dzDataMap);
        return sortedFkDataMap;
    }
}
