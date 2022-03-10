package com.yp.hbl.util;

import com.yp.hbl.entity.Constant;

import java.io.*;
import java.util.List;

public class FileUtil {

    public static boolean isExists(String filePath) {
        return new File(filePath).exists();
    }

    public static void closeInputStream(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void closeOutputStream(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static List<String> getAllFileName(File file, List<String> fileList) {
        File[] files = file.listFiles(new MyFileFilter());
        if (files == null) {
            return null;
        }
        for (File f : files) {
            if (f.isDirectory()) {
                fileList.add(f.getPath());
                getAllFileName(f, fileList);
            } else {
                fileList.add(f.getPath());
            }
        }
        return fileList;
    }

    private static class MyFileFilter implements FileFilter {

        @Override
        public boolean accept(File pathname) {
            if (pathname.isDirectory()) {
                return true;
            }
            String name = pathname.getName();
            if (name.endsWith(ExcelUtil.EXCEL_XLSX) && name.contains(Constant.dzWorkbook.WORKBOOK_NAME)) {
                return true;
            }
            return name.endsWith(ExcelUtil.EXCEL_XLSX) && name.contains(Constant.fkWorkbook.WORKBOOK_NAME);
        }
    }
}
