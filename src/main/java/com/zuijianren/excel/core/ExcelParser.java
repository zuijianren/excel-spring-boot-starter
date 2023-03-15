package com.zuijianren.excel.core;

import com.zuijianren.excel.annotations.ExcelMultiProperty;
import com.zuijianren.excel.annotations.ExcelProperty;
import com.zuijianren.excel.annotations.ExcelSheet;
import com.zuijianren.excel.config.PropertyConfig;
import com.zuijianren.excel.config.SheetConfig;
import com.zuijianren.excel.exceptions.ParserException;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.util.*;

/**
 * excel 解析器
 * <p>单例模式: 懒汉式 无线程安全问题</p>
 *
 * @author zuijianren
 * @date 2023/3/13 12:57
 */
public class ExcelParser {

    private final Map<Class<?>, SheetConfig> sheetConfigCache = new HashMap<>();

    private static final ExcelParser instance = new ExcelParser();

    public static ExcelParser getInstance() {
        return instance;
    }

    private ExcelParser() {
    }

    /**
     * 获取 sheet 表配置
     *
     * @param clazz 对应对象
     * @return sheet 表配置
     */
    public SheetConfig getSheetConfig(Class<?> clazz) {
        // 获取缓存的配置
        SheetConfig sheetConfig = sheetConfigCache.get(clazz);
        if (sheetConfig != null) {
            return sheetConfig;
        }
        // 如果不存在  则进行解析
        // 校验当前类是否添加 ExcelSheet 注解
        if (!clazz.isAnnotationPresent(ExcelSheet.class)) {
            // 如果未添加  抛出异常
            throw new ParserException("所选择的对象: [" + clazz + "] 不支持导出. 如需要导出, 请添加 ExcelSheet 等相关注解.");
        }
        return parseSheetConfig(clazz);
    }

    /**
     * 解析 sheet 配置
     *
     * @param clazz 对应对象
     * @return sheet 表配置
     */
    private SheetConfig parseSheetConfig(Class<?> clazz) {
        ArrayList<PropertyConfig> propertyConfigList = new ArrayList<>(); // 解析 property

        // 获取 注解 进行解析
        ExcelSheet sheetAnnotation = clazz.getAnnotation(ExcelSheet.class);
        String sheetName = sheetAnnotation.value(); // 获取 sheet 名
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            if (field.isAnnotationPresent(ExcelProperty.class)) {
                PropertyConfig propertyConfig = parsePropertyConfig(field);
                propertyConfigList.add(propertyConfig);
            } else if (field.isAnnotationPresent(ExcelMultiProperty.class)) {
                // todo  multi 属性最大为1   超过应进行报错（不知道如何展示这种情况）
                PropertyConfig propertyConfig = parseMultiPropertyConfig(field);
                propertyConfigList.add(propertyConfig);
            }
            // 忽略未添加 ExcelProperty 或者 ExcelMultiProperty 注解的属性
        }

        // 根据 order 进行排序
        propertyConfigList.sort(PropertyConfig::compareTo);

        // 创建 sheetConfig 对象
        SheetConfig sheetConfig = SheetConfig.builder()
                .sheetName(sheetName)
                .propertyConfigList(propertyConfigList)
                .build();

        // 存入缓存
        sheetConfigCache.put(clazz, sheetConfig);

        return sheetConfig;
    }

    /**
     * 解析 multi property 属性
     *
     * @param field property 对应的字段
     */
    private PropertyConfig parseMultiPropertyConfig(Field field) {

        if (!Collection.class.isAssignableFrom(field.getType())) {
            throw new ParserException(field.getName() + "属性不是集合类, 无法使用ExcelMultiProperty注解.");
        }

        // 获取 泛型
        ParameterizedType type = (ParameterizedType) field.getGenericType();
        Class<?> genericClass = (Class<?>) type.getActualTypeArguments()[0];

        List<PropertyConfig> propertyConfigList = new ArrayList<>();

        if (genericClass.isAnnotationPresent(ExcelSheet.class)) {
            SheetConfig sheetConfig = this.getSheetConfig(genericClass);
            propertyConfigList = sheetConfig.getPropertyConfigList();
        }


        // 获取 注解 进行解析
        ExcelMultiProperty multiPropertyAnnotation = field.getAnnotation(ExcelMultiProperty.class);

        // 创建 propertyConfig 对象
        PropertyConfig propertyConfig = PropertyConfig.builder()
                .order(multiPropertyAnnotation.order())
                .value(multiPropertyAnnotation.value())
                .converter(multiPropertyAnnotation.converter())
                .showCurrentName(multiPropertyAnnotation.showCurrentName())
                .nested(multiPropertyAnnotation.nested())
                .multi(true)
                .childPropertyConfigList(propertyConfigList)
                .field(field)
                .type(genericClass)
                .build();


        return propertyConfig;
    }

    /**
     * 解析 property 属性
     *
     * @param field property 对应的字段
     */
    private PropertyConfig parsePropertyConfig(Field field) {
        // 获取 注解 进行解析
        ExcelProperty propertyAnnotation = field.getAnnotation(ExcelProperty.class);

        // 创建 propertyConfig 对象
        PropertyConfig propertyConfig = PropertyConfig.builder()
                .order(propertyAnnotation.order())
                .value(propertyAnnotation.value())
                .converter(propertyAnnotation.converter())
                .showCurrentName(propertyAnnotation.showCurrentName())
                .nested(propertyAnnotation.nested())
                .multi(false)
                .field(field)
                .type(field.getType())
                .build();

        return propertyConfig;
    }
}
