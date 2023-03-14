package com.zuijianren.excel.config;

import com.zuijianren.excel.converter.ExcelConverter;
import lombok.Builder;
import lombok.Data;

import java.util.Comparator;
import java.util.List;

/**
 * 属性配置
 *
 * @author zuijianren
 * @date 2023/3/13 11:04
 */
@Builder
@Data
public class PropertyConfig implements Comparable<PropertyConfig> {

    /**
     * 排序
     */
    private int order;

    /**
     * 名字
     */
    private String[] value;

    /**
     * 当前属性是否为集合
     */
    private boolean multi;

    /**
     * 是否是内嵌对象
     * <p>
     * 如果是内嵌对象, 则忽略配置的转换器, 根据子属性配置集合直接加载
     * <p>
     * 否则直接调用转换器方法 转换为 字符串 进行填入
     */
    private boolean nested;

    /**
     * 是否展示当前的名字(即 是否以当前名字包裹下一级的名字)
     * <p>
     * 仅限 nested 为 true 时, 解析当前值
     */
    private boolean showCurrentName;

    /**
     * 子属性配置集合
     * <p>
     * 如果某个 property 为复杂对象, 则需要配置当前项
     */
    private List<PropertyConfig> childPropertyConfigList;

    /**
     * 转换器
     */
    private Class<? extends ExcelConverter<?>> converter;


    @Override
    public int compareTo(PropertyConfig o) {
        return this.order - o.order;
    }

    /**
     * 获取当前属性所占列数
     * <p>
     * 通常情况下为 1, 仅在 nested 为 true 时, 有变化
     *
     * @return 列数
     */
    public int getColNum() {
        // 如果不是 内嵌 类型  则 只占一列
        if (!this.nested) {
            return 1;
        }
        int result = 0;
        for (PropertyConfig propertyConfig : this.childPropertyConfigList) {
            result = result + propertyConfig.getColNum();
        }
        return result;
    }

    /**
     * 获取总行数
     *
     * @return 行数
     */
    public int getRowNum() {
        if (!this.nested) {
            return this.value.length;
        }
        int result = 0;
        if (showCurrentName) {
            result = this.value.length;
        }
        // 获取 下一级的行数
        Integer nextRowNum = childPropertyConfigList.stream().map(PropertyConfig::getRowNum).max(Comparator.naturalOrder()).orElse(0);
        result = result + nextRowNum;
        return result;
    }
}
