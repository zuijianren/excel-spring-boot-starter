package com.zuijianren.excel.core;

import com.zuijianren.excel.annotations.ExcelMultiProperty;
import com.zuijianren.excel.annotations.ExcelProperty;
import com.zuijianren.excel.annotations.ExcelSheet;
import com.zuijianren.excel.annotations.style.ExcelContentCellStyle;
import com.zuijianren.excel.annotations.style.ExcelHeadCellStyle;
import com.zuijianren.excel.annotations.style.ExcelSheetNameStyle;
import com.zuijianren.excel.config.PropertyConfig;
import com.zuijianren.excel.config.SheetConfig;
import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
import com.zuijianren.excel.config.style.HeadCellStyleConfig;
import com.zuijianren.excel.exceptions.ParserException;
import lombok.extern.slf4j.Slf4j;

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
@Slf4j
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
     * <p>
     * todo 循环依赖问题解决
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
        return parseSheetConfig(clazz);
    }

    /**
     * 解析 sheet 配置
     *
     * @param clazz 对应对象
     * @return sheet 表配置
     */
    private SheetConfig parseSheetConfig(Class<?> clazz) {

        // 校验当前类是否添加 ExcelSheet 注解
        if (!clazz.isAnnotationPresent(ExcelSheet.class)) {
            // 如果未添加  抛出异常
            throw new ParserException("所选择的对象: [" + clazz + "] 不支持导出. 如需要导出, 请添加 ExcelSheet 等相关注解.");
        }

        // 打印 解析 信息   用于测试当前导出工具实际运行情况
        log.debug("解析类: " + clazz);

        // 样式解析
        AbstractCellStyleConfig sheetNameStyle = Optional.ofNullable(clazz.getAnnotation(ExcelSheetNameStyle.class)).map(this::parseExcelSheetNameStyle).orElse(null);
        AbstractCellStyleConfig headStyle = Optional.ofNullable(clazz.getAnnotation(ExcelHeadCellStyle.class)).map(this::parseExcelHeadCellStyle).orElse(null);
        AbstractCellStyleConfig contentStyle = Optional.ofNullable(clazz.getAnnotation(ExcelContentCellStyle.class)).map(this::parseExcelContentCellStyle).orElse(null);

        // 获取 ExcelSheet 注解 进行解析
        ExcelSheet sheetAnnotation = clazz.getAnnotation(ExcelSheet.class);
        String sheetName = sheetAnnotation.value(); // 获取 sheet 名
        Field[] fields = clazz.getDeclaredFields();

        int multiNum = 0; // multi 属性统计
        Field multiFieldCache = null; // multi 属性缓存  用于提示错误

        ArrayList<PropertyConfig> propertyConfigList = new ArrayList<>(); // 解析 property
        for (Field field : fields) {
            field.setAccessible(true);
            PropertyConfig propertyConfig = null;
            if (field.isAnnotationPresent(ExcelProperty.class)) {
                propertyConfig = parsePropertyConfig(field);
            } else if (field.isAnnotationPresent(ExcelMultiProperty.class)) {
                propertyConfig = parseMultiPropertyConfig(field);
            } else {
                // 忽略未添加 ExcelProperty 或者 ExcelMultiProperty 注解的属性
                continue;
            }
            /*
             * 一个 Sheet 对象仅允许有一个 multi 属性(multi属性中可以继续持有multi属性, 不进行限制 即 一定程度上可以多次一对多)
             * 否则 会展示异常 因此 此处直接报错
             */
            if (propertyConfig.isMulti()) {
                multiNum++;
                if (multiNum >= 2) {
                    throw new ParserException("一个导出类仅允许拥有一个multi属性. 当前导出类存在两个multi属性: [" + field.getName() + ", " + multiFieldCache.getName() + "]");
                }
                multiFieldCache = field;
            }
            // 处理样式
            if (propertyConfig.getHeadCellStyleConfig() == null) {
                propertyConfig.setHeadCellStyleConfig(headStyle);
            }
            if (propertyConfig.getContentCellStyleConfig() == null) {
                propertyConfig.setContentCellStyleConfig(contentStyle);
            }


            propertyConfigList.add(propertyConfig);
        }

        // 根据 order 进行排序
        propertyConfigList.sort(PropertyConfig::compareTo);

        // 创建 sheetConfig 对象
        SheetConfig sheetConfig = SheetConfig.builder()
                .sheetName(sheetName)
                .propertyConfigList(propertyConfigList)
                // 样式配置
                .sheetNameStyleConfig(sheetNameStyle)
                .headCellStyleConfig(headStyle)
                .contentCellStyleConfig(contentStyle)
                // 集合属性
                .hasMulti(multiNum != 0)
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

        // 样式解析
        AbstractCellStyleConfig headStyle = Optional.ofNullable(field.getAnnotation(ExcelHeadCellStyle.class)).map(this::parseExcelHeadCellStyle).orElse(null);
        AbstractCellStyleConfig contentStyle = Optional.ofNullable(field.getAnnotation(ExcelContentCellStyle.class)).map(this::parseExcelContentCellStyle).orElse(null);

        // 获取 注解 进行解析
        ExcelMultiProperty multiPropertyAnnotation = field.getAnnotation(ExcelMultiProperty.class);

        // 获取 泛型
        ParameterizedType type = (ParameterizedType) field.getGenericType();
        Class<?> genericClass = (Class<?>) type.getActualTypeArguments()[0];

        List<PropertyConfig> childPropertyConfigList = new ArrayList<>();
        if (multiPropertyAnnotation.nested()) {
            if (genericClass.isAnnotationPresent(ExcelSheet.class)) {
                SheetConfig sheetConfig = this.getSheetConfig(genericClass);
                childPropertyConfigList = sheetConfig.getPropertyConfigList();
            } else {
                throw new ParserException(field.getName() + "不支持 nested 属性. 需要对应对象使用 ExcelSheet 注解才能使用");
            }
        }

        // 创建 propertyConfig 对象
        return PropertyConfig.builder()
                // 基础属性
                .order(multiPropertyAnnotation.order())
                .value(multiPropertyAnnotation.value())
                .field(field)
                .type(genericClass)
                // 样式属性
                .headCellStyleConfig(headStyle)
                .contentCellStyleConfig(contentStyle)
                // 转换器
                .converter(multiPropertyAnnotation.converter())
                // 内嵌属性相关
                .nested(multiPropertyAnnotation.nested())
                .showCurrentName(multiPropertyAnnotation.showCurrentName())
                .childPropertyConfigList(childPropertyConfigList)
                // 集合属性
                .multi(true)
                .build();
    }

    /**
     * 解析 property 属性
     *
     * @param field property 对应的字段
     */
    private PropertyConfig parsePropertyConfig(Field field) {
        // 获取 注解 进行解析
        ExcelProperty propertyAnnotation = field.getAnnotation(ExcelProperty.class);

        Class<?> type = field.getType(); // 获取属性类型
        boolean hasMulti = false; // 当前 field 是否含有 multi 属性

        // 样式解析
        AbstractCellStyleConfig headStyle = Optional.ofNullable(field.getAnnotation(ExcelHeadCellStyle.class)).map(this::parseExcelHeadCellStyle).orElse(null);
        AbstractCellStyleConfig contentStyle = Optional.ofNullable(field.getAnnotation(ExcelContentCellStyle.class)).map(this::parseExcelContentCellStyle).orElse(null);

        // nested 属性解析
        List<PropertyConfig> childPropertyConfigList = null; // 子属性集合 (如果为内嵌属性, 则赋值)
        if (propertyAnnotation.nested()) {
            if (type.isAnnotationPresent(ExcelSheet.class)) {
                SheetConfig sheetConfig = this.getSheetConfig(type);
                childPropertyConfigList = sheetConfig.getPropertyConfigList();
                hasMulti = sheetConfig.isHasMulti();
            } else {
                throw new ParserException(field.getName() + "不支持 nested 属性. 需要对应对象使用 ExcelSheet 注解才能使用");
            }
        }

        // 创建 propertyConfig 对象
        return PropertyConfig.builder()
                // 基础属性
                .order(propertyAnnotation.order())
                .value(propertyAnnotation.value())
                .field(field)
                .type(type)
                // 样式属性
                .headCellStyleConfig(headStyle)
                .contentCellStyleConfig(contentStyle)
                // 转换器
                .converter(propertyAnnotation.converter())
                // 内嵌属性相关
                .showCurrentName(propertyAnnotation.showCurrentName())
                .nested(propertyAnnotation.nested())
                .childPropertyConfigList(childPropertyConfigList)
                // 集合属性
                .multi(hasMulti)
                .build();
    }

    /**
     * 解析 ExcelContentCellStyle 注解 获取配置
     *
     * @param cellStyle 样式注解对象
     * @return 样式配置
     */
    private AbstractCellStyleConfig parseExcelContentCellStyle(ExcelContentCellStyle cellStyle) {
        return HeadCellStyleConfig.builder()
                // todo 解析
                .build();
    }

    /**
     * 解析 ExcelHeadCellStyle 注解 获取配置
     *
     * @param cellStyle 样式注解对象
     * @return 样式配置
     */
    private AbstractCellStyleConfig parseExcelHeadCellStyle(ExcelHeadCellStyle cellStyle) {
        return HeadCellStyleConfig.builder()
                // todo 解析
                .build();
    }

    /**
     * 解析 ExcelSheetNameStyle 注解 获取配置
     *
     * @param cellStyle 样式注解对象
     * @return 样式配置
     */
    private AbstractCellStyleConfig parseExcelSheetNameStyle(ExcelSheetNameStyle cellStyle) {
        return null;
    }

}
