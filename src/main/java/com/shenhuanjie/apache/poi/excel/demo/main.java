package com.shenhuanjie.apache.poi.excel.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class main {
    public static void main(String[] args) throws IOException {
        //文件路径
        String filePath = "C:\\Users\\shenh\\Documents\\Output\\Excel\\excel-sample.xls";
        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表(Sheet)
        HSSFSheet sheet = workbook.createSheet();
        //创建工作表(Sheet)
        sheet = workbook.createSheet("Test");
        //删除默认工作表(Sheet)
        workbook.removeSheetAt(0);
        FileOutputStream out = new FileOutputStream(filePath);
        //保存Excel文件
        workbook.write(out);
        //关闭文件流
        out.close();
        System.out.println("OK!");
    }
}
