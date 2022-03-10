package com.yp.hbl.main;

import com.yp.hbl.entity.*;
import com.yp.hbl.service.DzExcel;
import com.yp.hbl.service.DzWorkbook;
import com.yp.hbl.service.FkExcel;
import com.yp.hbl.util.ExcelUtil;
import com.yp.hbl.util.FileUtil;

import java.io.File;
import java.util.List;
import java.util.Scanner;
import java.util.TreeMap;

/**
 * 生成对账工作簿
 */
public class GenerateDz {
    public static void main(String[] args) {
        System.out.println("欢迎使用本程序------");
        System.out.println("请输入[" + Constant.dataCompareWorkbook.WORKBOOK_NAME + "]文件所在目录：");
        boolean isDz = false;
        while (!isDz) {
            String srcFileName = "", fatherDir;
            Scanner scanner = new Scanner(System.in);
            String fileName = scanner.nextLine();
            if (fileName.endsWith(File.separator)) {
                fatherDir = fileName.substring(0, fileName.lastIndexOf(File.separator));
            } else {
                fatherDir = fileName;
            }
            String srcFileName1 = fatherDir + File.separator + Constant.dataCompareWorkbook.WORKBOOK_NAME + "." + ExcelUtil.EXCEL_XLSX;
            String srcFileName2 = fatherDir + File.separator + Constant.dataCompareWorkbook.WORKBOOK_NAME + "." + ExcelUtil.EXCEL_XLS;
            if (FileUtil.isExists(srcFileName1)) {
                isDz = true;
                srcFileName = srcFileName1;
            } else if (FileUtil.isExists(srcFileName2)) {
                isDz = true;
                srcFileName = srcFileName2;
            } else {
                System.out.println("路径：[" + fatherDir + "]下，不存在 " + Constant.dataCompareWorkbook.WORKBOOK_NAME);
            }
            List<Dz> hisDataList;
            List<Dz> srcDataList;
            if (isDz) {
                System.out.println("源文件路径：[" + srcFileName + "]");
                //准备阶段，统一源数据表格的单元格格式
                DzExcel.setCellStyle(srcFileName);
                long startTime = System.currentTimeMillis();
                System.out.println("开始生成对账------");
                System.out.println("正在生成对账，请稍等------");
                //第1步，读取历史数据和原始数据，即读取【数据对照表】中[上期数据]和[公司数据分项]sheet页数据
                hisDataList = DzExcel.readExcel(srcFileName, Constant.dataCompareWorkbook.HIS_SHEET_NAME);
                srcDataList = DzExcel.readExcel(srcFileName, Constant.dataCompareWorkbook.SRC_SHEET_NAME);
                //第2步，生成对账，即根据原始数据匹配历史数据店名编号和产品编号，得到对方价、对方单价、差价、差额，并写入excel
                DzExcel.writeExcel(srcFileName, Constant.dataCompareWorkbook.SRC_SHEET_NAME, hisDataList, srcDataList);
                long endTime = System.currentTimeMillis();
                System.out.println("生成对账完毕------");
                System.out.println("生成对账耗时 [" + (endTime - startTime) + "]");
                System.out.println();
                //第3步，读取客户数据
                startTime = System.currentTimeMillis();
                System.out.println("开始生成对账工作簿------");
                System.out.println("正在生成对账工作簿，请稍等------");
                List<Consumer> consumerDataList = DzExcel.readExcel1(srcFileName, Constant.dataCompareWorkbook.CONSUMER_SHEET_NAME);
                //第4步，生成对账工作簿所需数据
                TreeMap<Integer, List<DzData>> dzDataMap = DzWorkbook.getDzWorkbookData(srcDataList, consumerDataList);
//                System.out.println("srcDataList = " + srcDataList);
//                System.out.println("consumerDataList = " + consumerDataList);
                //第5步，生成对账工作簿
                String dzWorkbookName = fatherDir + File.separator + Constant.dzWorkbook.WORKBOOK_NAME + "." + ExcelUtil.EXCEL_XLSX;
                DzWorkbook.writeDzWorkbook(dzWorkbookName, dzDataMap);
                endTime = System.currentTimeMillis();
                System.out.println("生成对账工作簿完毕------");
                System.out.println("生成对账工作簿耗时 [" + (endTime - startTime) + "]");
                System.out.println();
                //第6步，生成反馈工作簿数据
                TreeMap<Integer, List<FkData>> fkDataMap = null;
                if (srcDataList != null) {
                    fkDataMap = FkExcel.getFkWorkbookData(srcDataList);
                }
                startTime = System.currentTimeMillis();
                //第7步，生成反馈工作簿
                System.out.println("开始生成反馈工作簿------");
                System.out.println("正在生成反馈工作簿，请稍等------");
                String fkWorkbookName = fatherDir + File.separator + Constant.fkWorkbook.WORKBOOK_NAME + "." + ExcelUtil.EXCEL_XLSX;
                if (fkDataMap != null) {
                    FkExcel.writeFkWorkbook(fkWorkbookName, fkDataMap);
                }
                endTime = System.currentTimeMillis();
                System.out.println("生成反馈工作簿完毕------");
                System.out.println("生成反馈工作簿耗时 [" + (endTime - startTime) + "]");
                scanner.close();
            }
        }
    }
}
