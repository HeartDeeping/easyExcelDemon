package com.example.easyexceldemon.fill;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

@RestController
public class ExportDataApi {
    @RequestMapping("/export")
    public void exportDataApi(HttpServletResponse response, @RequestBody Map<String, Object> exportData) throws Exception {
        Map<String, Object> dataMap = (Map<String, Object>) exportData.get("dataMap");
        List<Map<String, Object>> data = (List<Map<String, Object>>) dataMap.get("data");
        List<Map<String, Object>> gmzyList = (List<Map<String, Object>>) dataMap.get("gmzyList");
        InputStream inputStream1 = getClass().getClassLoader().getResourceAsStream("static/excel_template/xt/detail_info_yn.xlsx");
        InputStream inputStream2 = getClass().getClassLoader().getResourceAsStream("static/excel_template/xt/detail_info_yn.xlsx");
        InputStream inputStream3 = getClass().getClassLoader().getResourceAsStream("static/excel_template/xt/detail_info_yn.xlsx");
        // 创建监听器
        TemplateKeyListener listener = new TemplateKeyListener(0,7);
        TemplateKeyListener listener1 = new TemplateKeyListener(1,3);

        ExcelReader excelReader = EasyExcel.read(inputStream1, listener).build();
        ExcelReader excelReader1 = EasyExcel.read(inputStream2, listener1).build();

        ReadSheet readSheet0 = EasyExcel.readSheet(0).build();
        ReadSheet readSheet1 = EasyExcel.readSheet(1).build();

        excelReader.read(readSheet0);
        excelReader1.read(readSheet1);

        // 获取键
        List<String> keys = listener.getKeys();
        dealData(keys,data);
        List<String> gmzyKeys = listener1.getKeys();
        dealData(gmzyKeys,gmzyList);

        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(inputStream3).build();
        WriteSheet writeSheet1 = EasyExcel.writerSheet(0).build();
        WriteSheet writeSheet2 = EasyExcel.writerSheet(1).build();
        FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.VERTICAL).build();
        excelWriter.fill(dataMap, fillConfig, writeSheet1)
                .fill(dataMap, fillConfig, writeSheet2)
                .fill(new FillWrapper("data", data), fillConfig, writeSheet1)
                .fill(new FillWrapper("gmzyList", gmzyList), fillConfig, writeSheet2)
                .finish();
        System.out.println("导出完成");
    }

    public void dealData(List<String> keys, List<Map<String, Object>> data) {
        // 添加自增序号
        int i = 1;
        // 添加模版里的key
        for (Map<String, Object> map : data) {
            map.put("index",i++);
            for (String key : keys) {
                if (!map.containsKey(key)) {
                    map.put(key,"");
                }
            }
        }
    }
}
