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
    private boolean showSheetName;

    /**
     * 是否展示序号列
     */
    private boolean showSerialNumber;

    /**
     * 是否冻结标题行
     */
    private boolean freezeHead;

    /**
     * 属性
     */
    private List<PropertyConfig> propertyConfigList;


    /**
     * sheet 名字样式配置
     */
    private AbstractCellStyleConfig sheetNameStyleConfig;

    /**
     * 序列号样式
     */
    private AbstractCellStyleConfig serialNumberStyleConfig;

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
