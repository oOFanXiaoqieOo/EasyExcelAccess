package com.xo.util;
/**
 * SheetData类
 * ExcelSheet数据单元，用于easyexcel多sheet生成传递参数使用
 * */
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class SheetData {
    private List<? extends BaseRowModel> data;
    private Sheet sheet;
    public Sheet getSheet() {
        return this.sheet;
    }
    public Collection<?> getData() {
        return this.data;
    }
    public void setData(ArrayList<? extends BaseRowModel> list) {
        this.data = list;
    }
    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }
}
