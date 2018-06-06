package com.howie.easyexcelmethodencapsulation.excel;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.howie.easyexcelmethodencapsulation.test.ExportInfo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description 工具类
 * @Date 2018-06-06
 * @Time 14:07
 */
public class ExcelUtil {
    /**
     * 读取 Excel(多个 sheet)
     *
     * @param excel  文件
     * @param object 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel object) {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null) {
            return null;
        }
        for (Sheet sheet : reader.getSheets()) {
            if (object != null) {
                sheet.setClazz(object.getClass());
            }
            reader.read(sheet);
        }
        return excelListener.getDatas();
    }

    /**
     * 读取某个 sheet 的 Excel
     *
     * @param excel   文件
     * @param object  实体类映射，继承 BaseRowModel 类
     * @param sheetNo sheet 的序号
     *                当前版本中：
     *                XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
     *                XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel object, int sheetNo) {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null) {
            return null;
        }
        Sheet sheet = new Sheet(sheetNo);
        sheet.setClazz(object.getClass());
        reader.read(sheet);
        return excelListener.getDatas();
    }

    /**
     * 导出 Excel ：一个 sheet，带表头
     *
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param object    映射实体类，Excel 模型
     */
    public static void writeExcel(HttpServletResponse response, List<? extends BaseRowModel> list,
                                  String fileName, String sheetName, BaseRowModel object) throws IOException {
        //创建本地文件
        String filePath = fileName + ".xlsx";
        File dbfFile = new File(filePath);
        if (!dbfFile.exists() || dbfFile.isDirectory()) {
            dbfFile.createNewFile();
        }
        fileName = new String(filePath.getBytes(), "ISO-8859-1");
        response.addHeader("Content-Disposition", "filename=" + fileName);
        OutputStream out = response.getOutputStream();
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            Sheet sheet = new Sheet(1, 0, object.getClass());
            sheet.setSheetName(sheetName);
            writer.write(list, sheet);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

//    /**
//     * 导出 Excel ：一个 sheet，不带表头
//     *
//     * @param response  HttpServletResponse
//     * @param list      数据 list，每个元素为一个 BaseRowModel
//     * @param fileName  导出的文件名
//     * @param sheetname 导入文件的 sheet 名
//     */
//    public static void writeExcel(HttpServletResponse response, List<BaseRowModel> list,
//                                  String fileName, String sheetname) throws IOException {
//        //创建本地文件
//        String filePath = fileName + ".xlsx";
//        File dbfFile = new File(filePath);
//        if (!dbfFile.exists() || dbfFile.isDirectory()) {
//            dbfFile.createNewFile();
//        }
//        fileName = new String(filePath.getBytes(), "ISO-8859-1");
//        response.addHeader("Content-Disposition", "filename=" + fileName);
//        OutputStream out = response.getOutputStream();
//        try {
//            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
//            Sheet sheet = new Sheet(1);
//            sheet.setSheetName(sheetname);
//            writer.write(list, sheet);
//            writer.finish();
//        } catch (Exception e) {
//            e.printStackTrace();
//        } finally {
//            try {
//                out.close();
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }

    /**
     * 返回 ExcelReader
     *
     * @param excel         需要解析的 Excel 文件
     * @param excelListener new ExcelListener()
     */
    private static ExcelReader getReader(MultipartFile excel,
                                         ExcelListener excelListener) {
        String filename = excel.getOriginalFilename();
        if (filename == null || (!filename.toLowerCase().endsWith(".xls") && !filename.toLowerCase().endsWith(".xlsx"))) {
            throw new ExcelException("文件格式错误！");
        }
        ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.XLSX;
        if (filename.toLowerCase().endsWith(".xls")) {
            System.out.println(filename);
            System.out.println(filename.toLowerCase().endsWith(".xls"));
            excelTypeEnum = ExcelTypeEnum.XLS;
        }
        System.out.println(excelTypeEnum);
        InputStream inputStream;
        try {
            inputStream = excel.getInputStream();
            return new ExcelReader(inputStream, excelTypeEnum,
                    null, excelListener);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
