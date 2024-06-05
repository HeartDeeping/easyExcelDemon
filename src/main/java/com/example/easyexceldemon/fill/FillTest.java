package com.example.easyexceldemon.fill;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import org.junit.Test;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//测试填充excel模板
@RestController
public class FillTest {
    @Test
    @RequestMapping("/test")
    public void fillExcel() throws Exception {
        String templateFilePath = "C:\\Users\\26537\\Desktop\\测试填充模板.xlsx";
        String fileName = "姓名年龄表.xlsx";

        try  {
            ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFilePath).build();
            Map<String, Object> dataMap = new HashMap<>();
            Map<String, Object> dataMap2 = new HashMap<>();
            dataMap2.put("name","张三");
            dataMap2.put("age",20);

            Map<String, Object> dataMap3 = new HashMap<>();
            dataMap3.put("name","lisi");
            dataMap3.put("age",50);
            dataMap3.put("hobby","游泳");


            dataMap.put("start", "2019年10月9日13:28:28");
            Person p1 = new Person("张三",18);
            Person p2 = new Person("lisi",20);
            Person p3 = new Person("wangwu",25);

            Person p4 = new Person("hh",66);
            Person p5 = new Person("zz",44);
            Person p6 = new Person("aa",77);

            List<Map<String, Object>> personList = new ArrayList<>();
            personList.add(dataMap2);
            personList.add(dataMap3);

            List<Object> gmzyList = new ArrayList<>();
            gmzyList.add(p4);
            gmzyList.add(p5);
            gmzyList.add(p6);

            dataMap.put("data",personList);
            dataMap.put("gmzyList",gmzyList);

            // 创建两个不同的 WriteSheet 实例，分别对应两个不同的 sheet
            WriteSheet writeSheet1 = EasyExcel.writerSheet(0).build();
            WriteSheet writeSheet2 = EasyExcel.writerSheet(1).build();
            WriteSheet writeSheet3 = EasyExcel.writerSheet( 2).build();
            FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
            excelWriter.fill(new FillWrapper("data",personList), fillConfig, writeSheet1)
                    .fill(new FillWrapper("data",personList), fillConfig, writeSheet1)
                    .fill(new FillWrapper("gmzyList",gmzyList), fillConfig, writeSheet2)
                    .fill(dataMap, writeSheet1).fill(dataMap, writeSheet2).fill(dataMap,writeSheet3).finish();
            System.out.println("导出完毕");
        } catch (Exception e) {
            throw new Exception(e);
        }
    }

    public void setCharset(HttpServletResponse response, String agent, String fileName)
            throws UnsupportedEncodingException {
        response.setContentType("application/x-download;charset=UTF-8");
        String finalFileName;
        if (agent.contains("Mozilla") && agent.contains("Firefox")) {// google,火狐浏览器
            finalFileName = new String(fileName.getBytes(), "ISO8859-1");
        } else {
            finalFileName = URLEncoder.encode(fileName, "UTF8");// 其他浏览器
        }
        response.setHeader("Content-Disposition", "attachment; filename=" + finalFileName);
        response.setHeader("Access-Control-Expose-Headers","file_name");
        response.setHeader("file_name",URLEncoder.encode(fileName, "UTF8"));
    }
}

