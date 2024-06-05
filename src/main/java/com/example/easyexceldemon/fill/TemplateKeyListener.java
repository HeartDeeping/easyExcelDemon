package com.example.easyexceldemon.fill;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class TemplateKeyListener extends AnalysisEventListener<Map<Integer, String>> {
    private List<String> keys = new ArrayList<>();
    private int sheetIndex;
    private int rowIndex;

    public TemplateKeyListener(int sheetIndex, int rowIndex) {
        this.sheetIndex = sheetIndex;
        this.rowIndex = rowIndex;
    }

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        if (context.readSheetHolder().getSheetNo() == sheetIndex && context.readRowHolder().getRowIndex() == rowIndex) {
            data.remove(0); // 移除第一列
            data.values().stream()
                    .filter(value -> value != null && value.contains("{") && value.contains("}"))
                    .forEach(value -> {
                        // 提取键名
                        String key = value.substring(value.indexOf('.') + 1, value.length() - 1);
                        keys.add(key);
                    });
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // no-op
    }

    public List<String> getKeys() {
        return keys;
    }
}

