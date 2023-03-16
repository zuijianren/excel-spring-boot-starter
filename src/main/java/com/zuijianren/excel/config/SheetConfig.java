package com.zuijianren.excel.config;

import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
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

    /**
     * 属性
     */
    private List<PropertyConfig> propertyConfigList;

    /**
     * 对象中是否有字段含有 multi 属性
     */
    private boolean hasMulti;

    /**
     * sheet 名字 配置
     */
    private AbstractCellStyleConfig sheetNameStyleConfig;

    /**
     * 标题单元格样式
     * <p>
     * 如果当前为空 则写入数据时   采用上一级的配置进行写入
     */
    private AbstractCellStyleConfig headCellStyleConfig;

    /**
     * 内容单元格样式
     * <p>
     * 如果当前为空 则写入数据时   采用上一级的配置进行写入
     */
    private AbstractCellStyleConfig contentCellStyleConfig;

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
