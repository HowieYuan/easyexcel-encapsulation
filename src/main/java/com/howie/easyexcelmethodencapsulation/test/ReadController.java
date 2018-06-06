package com.howie.easyexcelmethodencapsulation.test;

import com.howie.easyexcelmethodencapsulation.excel.ExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description
 * @Date 2018-06-05
 * @Time 16:56
 */
@RestController
public class ReadController {
    @RequestMapping(value = "readExcel", method = RequestMethod.POST)
    public Object readExcel(MultipartFile excel) {
        return ExcelUtil.readExcel(excel, new ImportInfo());
    }

    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response, String fileName, String sheetName) throws IOException {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("hhh");
        model1.setAge("19");
        model1.setAddress("123456789");
        model1.setEmail("asfasaf@sakfhak.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("hhh");
        model2.setAge("19");
        model2.setAddress("123456789");
        model2.setEmail("asfasaf@sakfhak.com");
        list.add(model2);
        ExcelUtil.writeExcel(response, list, fileName, sheetName, new ExportInfo());
    }
}
