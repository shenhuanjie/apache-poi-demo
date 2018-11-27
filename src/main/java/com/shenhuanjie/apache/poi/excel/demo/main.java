package com.shenhuanjie.apache.poi.excel.demo;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class main {
    private static final String FILE_PATH_ = "C:\\Users\\shenh\\Documents\\Output\\Excel\\";

    public static void main(String[] args) throws IOException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddhhMMss");
        // 文件名
        String fileName = simpleDateFormat.format(new Date()) + ".xls";
        // 文件路径
        String filePath = FILE_PATH_ + fileName;
        // 创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();

        /**
         * EXCEL 基本操作
         */

        // 操作Workbook和Sheet
        // createHssfWorkbook(workbook);

        // 创建单元格
        // createHssfCell(workbook);

        // 创建文档摘要信息
        // createHssfCellCommentHssfInformationProperties(workbook);

        // 创建批注
        // createHssfCellComment(workbook);

        // 创建页眉和页脚
        // workbook = createHssfHeaderAndFooter(workbook);

        /**
         * Excel 单元格操作
         */
        // 设置格式
        // workbook = setSheetStyle(workbook);

        // 合并单元格
        // workbook = setSheetRegion(workbook);

        // 单元格对齐
        workbook = setHssfSheetStyle(workbook);

        fileOutputStream(filePath, workbook);


        System.out.println("OK!");
    }

    /**
     * 设置单元格对齐
     *
     * @param workbook
     * @return
     */
    private static HSSFWorkbook setHssfSheetStyle(HSSFWorkbook workbook) {
        HSSFSheet sheet = workbook.createSheet("Test");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("单元格对齐");
        HSSFCellStyle style = workbook.createCellStyle();
        // 水平居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 垂直居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 自动换行
        style.setWrapText(true);
        // 缩进
        style.setIndention((short) 5);
        // 文本旋转，这里的取值是从 -90 到90，而不是0 - 180
        style.setRotation((short) 60);
        cell.setCellStyle(style);
        return workbook;
    }

    /**
     * 合并单元格
     *
     * @param workbook
     */
    private static HSSFWorkbook setSheetRegion(HSSFWorkbook workbook) {
        HSSFSheet sheet = workbook.createSheet("Test");
        HSSFRow row = sheet.createRow(0);
        // 合并列
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("合并列");
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
        sheet.addMergedRegion(region);
        // 合并行
        cell = row.createCell(6);
        cell.setCellValue("合并行");
        region = new CellRangeAddress(0, 5, 6, 6);
        sheet.addMergedRegion(region);
        return workbook;
    }

    /**
     * 设置单元格格式
     *
     * @param workbook
     * @return
     */
    private static HSSFWorkbook setSheetStyle(HSSFWorkbook workbook) {
        // 创建工作表（Sheet）
        HSSFSheet sheet = workbook.createSheet("Test");
        HSSFRow row = sheet.createRow(0);
        // 设置日期格式——使用Excel内嵌的格式
        HSSFCell cell = row.createCell(0);
        cell.setCellValue(new Date());
        HSSFCellStyle style = workbook.createCellStyle();
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
        cell.setCellStyle(style);
        // 设置保留2位小数——使用Excel内嵌的格式
        cell = row.createCell(1);
        cell.setCellValue(12.3456789);
        style = workbook.createCellStyle();
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        cell.setCellStyle(style);
        // 设置货币格式——使用自定义的格式
        cell = row.createCell(2);
        cell.setCellValue(12345.6789);
        style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("Y#,##0"));
        cell.setCellStyle(style);
        // 设置百分比格式——使用自定义格式
        cell = row.createCell(3);
        cell.setCellValue(0.123456789);
        style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        cell.setCellStyle(style);
        // 设置中文大写格式--使用自定义的格式
        cell = row.createCell(4);
        cell.setCellValue(12345);
        style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("[DbNum2][$-804]0"));
        cell.setCellStyle(style);
        // 设置科学计数法格式--使用自定义的格式
        cell = row.createCell(5);
        cell.setCellValue(12345);
        style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00E+00"));
        cell.setCellStyle(style);
        return workbook;
    }

    /**
     * 创建页眉与页脚
     *
     * @param workbook
     */
    private static HSSFWorkbook createHssfHeaderAndFooter(HSSFWorkbook workbook) {
        // 创建工作表（Sheet)
        HSSFSheet sheet = workbook.createSheet("Test");

        // 得到页眉
        HSSFHeader header = sheet.getHeader();
        header.setLeft("页眉左边");
        header.setRight("页眉右边");
        header.setCenter("页眉中间");
        // 得到页脚
        HSSFFooter footer = sheet.getFooter();
        footer.setLeft("页脚左边");
        footer.setRight("页脚右边");
        footer.setCenter("页脚中间");
        return workbook;
    }

    /**
     * 创建标注
     *
     * @param workbook
     */
    private static void createHssfCellComment(HSSFWorkbook workbook) {
        // 创建批注
        HSSFSheet sheet = workbook.createSheet("Test");
        HSSFPatriarch patr = sheet.createDrawingPatriarch();
        // 创建批注位置
        HSSFClientAnchor anchor = patr.createAnchor(0, 0, 0, 0, 5, 1, 8, 3);
        // 创建批注
        HSSFComment comment = patr.createCellComment(anchor);
        // 设置批注内容
        comment.setString(new HSSFRichTextString("这是一个批注段落"));
        // 设置批注作者
        comment.setAuthor("李志伟");
        // 设置批注默认显示
        comment.setVisible(true);
        HSSFCell cell = sheet.createRow(2).createCell(1);
        cell.setCellValue("测试");
        // 把批注赋值给单元格
        cell.setCellComment(comment);
    }

    /**
     * 创建文档摘要信息
     *
     * @param workbook
     */
    private static void createHssfCellCommentHssfInformationProperties(HSSFWorkbook workbook) {
        // 创建文档信息
        workbook.createInformationProperties();
        // 摘要信息
        DocumentSummaryInformation dsi = workbook.getDocumentSummaryInformation();
        // 类别
        dsi.setCategory("Excel文件");
        // 管理者
        dsi.setManager("李志伟");
        // 公司
        dsi.setCompany("——");
        // 摘要信息
        SummaryInformation si = workbook.getSummaryInformation();
        // 主题
        si.setSubject("——");
        // 标题
        si.setTitle("测试文档");
        // 作者
        si.setAuthor("李志伟");
        // 备注
        si.setComments("POI测试文档");
    }

    /**
     * 创建单元格Cell
     *
     * @param workbook
     */
    private static void createHssfCell(HSSFWorkbook workbook) {
        //创建工作表（Sheet)
        HSSFSheet sheet = workbook.createSheet("Test");
        //创建行，从0开始
        HSSFRow row = sheet.createRow(0);
        //创建行的单元格，也是从0开始
        HSSFCell cell = row.createCell(0);
        //设置单元格内容
        cell.setCellValue("李志伟");
        //设置单元格内容，重载
        row.createCell(1).setCellValue(false);
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue(12.345);
    }


    /**
     * 操作Workbook和Sheet
     *
     * @return
     */
    private static void createHssfWorkbook(HSSFWorkbook workbook) {
        //创建工作表(Sheet)
        workbook.createSheet("Test");
    }

    /**
     * 保存Excel文件
     *
     * @param filePath
     * @param workbook
     * @throws IOException
     */
    private static void fileOutputStream(String filePath, HSSFWorkbook workbook) throws IOException {
        FileOutputStream out = new FileOutputStream(filePath);
        //保存Excel文件
        workbook.write(out);
        //关闭文件流
        out.close();

        System.out.println("FILE_PATH_:" + filePath);
    }
}
