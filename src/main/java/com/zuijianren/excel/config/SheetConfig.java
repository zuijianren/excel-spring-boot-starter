package com.zuijianren.excel.config;

import lombok.Builder;
import lombok.Data;

import java.util.Comparator;
import java.util.List;

/**
 * sheet 配置
 *
 * @author zuijianren
 * @date 2023/3/13 11:02
 */
@Builder
@Data
public class SheetConfig {

    /**
     * 表名
     */
    private String sheetName;

    /**
     * 首行是否展示表名
     */
    @Builder.Default
    private boolean showSheetName = true;

    /**
     * 是否展示序号列
     */
    @Builder.Default
    private boolean showSerialNumber = false;

    // todo 样式属性管理

    /**
     * 属性
     */
    private List<PropertyConfig> propertyConfigList;

    /**
     * 对象中是否有字段含有 multi 属性
     */
    private boolean hasMulti;

    /**
     * 根据当前配置 计算列数
     *
     * @return 当前 sheet 表总列数 (需要注意序号列)
     */
    public int getColNum() {
        int result = 0;
        if (this.showSerialNumber) {
            result++;
        }
        for (PropertyConfig propertyConfig : this.propertyConfigList) {
            result = result + propertyConfig.getColNum();
        }
        return result;
    }

    /**
     * 计算行数
     */
    public int getRowNum() {
        return propertyConfigList.stream().map(PropertyConfig::getRowNum).max(Comparator.naturalOrder()).orElse(0);
    }
}
