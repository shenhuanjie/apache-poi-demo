package com.shenhuanjie.apache.poi.excel.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class main {
    private static final String ROOT_PATH = "C:\\Users\\shenh\\Documents\\Output\\Excel\\";

    public static void main(String[] args) throws IOException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddhhMMss");
        //文件名
        String fileName = simpleDateFormat.format(new Date()) + ".xls";
        //文件路径
        String filePath = ROOT_PATH + fileName;
        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();

        //操作Workbook和Sheet
        //workbook = getHssfWorkbook(workbook);

        //创建工作表（Sheet)
        HSSFSheet sheet = workbook.createSheet("Test");
        //创建行，从0开始
        HSSFRow row = sheet.createRow(0);

        FileOutputStream out = new FileOutputStream(filePath);
        //保存Excel文件
        workbook.write(out);
        //关闭文件流
        out.close();
        System.out.println("OK!");
    }

    /**
     * 操作Workbook和Sheet
     *
     * @return
     */
    private static HSSFWorkbook getHssfWorkbook(HSSFWorkbook workbook) {
        //创建工作表(Sheet)
        workbook.createSheet("Test");
        return workbook;
    }
}
