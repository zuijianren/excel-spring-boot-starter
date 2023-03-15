package com.zuijianren.excel.pojo;

import com.zuijianren.excel.config.SheetConfig;
import lombok.Data;

import java.util.Collection;

/**
 * excel 渲染数据的封装对象
 *
 * @author zuijianren
 * @date 2023/3/14 15:38
 */
@Data
public class ExcelData {

    private SheetConfig sheetConfig;
    private Collection dataList;

    public ExcelData(SheetConfig sheetConfig, Collection<?> dataList) {
        this.sheetConfig = sheetConfig;
        this.dataList = dataList;
    }
}
