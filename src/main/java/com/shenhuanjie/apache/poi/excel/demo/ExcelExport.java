package com.shenhuanjie.apache.poi.excel.demo;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class ExcelExport {


    public static void main(String[] args) throws IOException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddhhMMss");
        String sheetName = "测试Excel格式";
        String sheetTitle = "测试Excel格式";
        String fileName = simpleDateFormat.format(new Date());
        List columnNames = new LinkedList<>();
        columnNames.add("日期-String");
        columnNames.add("日期-Date");
        columnNames.add("时间戳-Long");
        columnNames.add("客户编码");
        columnNames.add("整数");
        columnNames.add("带小数的正数");
        columnNames.add("带小数的正数");
        ExportExcel2007 exportExcel2007 = new ExportExcel2007();
        for (int k = 0; k < 10; k++) {
            exportExcel2007.writeExcelTitle("C://temp", fileName, sheetName+k, columnNames, sheetTitle);
            for (int j = 0; j < 10; j++) {
                List<List> objects = new LinkedList<>();
                for (int i = 0; i < 10; i++) {
                    List dataA = new LinkedList<>();
                    dataA.add("2016-09-05 17:27:25");
                    dataA.add(new Date());
                    dataA.add(1451036631012L);
                    dataA.add("000628");
                    dataA.add(i);
                    dataA.add(1.323 + i);
                    dataA.add(1.323 + i);
                    objects.add(dataA);
                }
                try {
                    exportExcel2007.writeExcelData("C://temp", fileName, sheetName+k, objects);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                objects.clear();
            }
        }
        exportExcel2007.dispose();
    }
}