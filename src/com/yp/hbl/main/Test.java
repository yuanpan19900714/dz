package com.yp.hbl.main;

import com.yp.hbl.service.DzExcel;
import com.yp.hbl.util.ExcelUtil;
import com.yp.hbl.util.FileUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class Test {
    public static void main(String[] args) {
//        Calendar calendar = Calendar.getInstance();
//        calendar.setFirstDayOfWeek(Calendar.MONDAY);
//        calendar.setTime(new Date());
//        System.out.println(calendar.get(Calendar.WEEK_OF_YEAR));
//        writeBorderBottom();
        DzExcel.setCellStyle("C:\\Users\\Administrator\\Desktop\\新建文件夹\\数据对照表.xlsx");
    }

    private static void writeBorderBottom() {
        String filePath = "F:\\郝\\VBA\\测试.xlsx";
        OutputStream out = null;
        try {
            // 读取Excel文档
            File destFile = new File(filePath);
            Workbook workBook = ExcelUtil.getWorkbok(destFile);
            // sheet 对应一个工作页
            Sheet sheet = workBook.getSheetAt(0);

            CellRangeAddress region = CellRangeAddress.valueOf("A1:A1");
            RegionUtil.setBorderBottom(BorderStyle.DASH_DOT, region, sheet);
            region = CellRangeAddress.valueOf("B1:B1");
            RegionUtil.setBorderBottom(BorderStyle.DASH_DOT_DOT, region, sheet);
            region = CellRangeAddress.valueOf("C1:C1");
            RegionUtil.setBorderBottom(BorderStyle.DASHED, region, sheet);
            region = CellRangeAddress.valueOf("D1:D1");
            RegionUtil.setBorderBottom(BorderStyle.DOTTED, region, sheet);
            region = CellRangeAddress.valueOf("E1:E1");
            RegionUtil.setBorderBottom(BorderStyle.HAIR, region, sheet);
            region = CellRangeAddress.valueOf("F1:F1");
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);

            CreationHelper creationHelper = workBook.getCreationHelper();
//            creationHelper.
            workBook.createSheet().createDrawingPatriarch();
            //使excel内部公式生效
            workBook.setForceFormulaRecalculation(true);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
//            cell.
            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(destFile);
            workBook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            FileUtil.closeOutputStream(out);
        }
    }


}
