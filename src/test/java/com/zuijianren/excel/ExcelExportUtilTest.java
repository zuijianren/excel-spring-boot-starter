package com.zuijianren.excel;

import com.zuijianren.excel.config.ExcelConfig;
import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
import com.zuijianren.excel.config.style.ContentCellStyleConfig;
import com.zuijianren.excel.config.style.HeadCellStyleConfig;
import com.zuijianren.excel.config.style.SheetNameStyleConfig;
import com.zuijianren.excel.core.ExcelWriter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;

/**
 * @author zuijianren
 * @date 2023/3/13 13:40
 */
public class ExcelExportUtilTest {

    Student student;
    Teacher teacher;

    @BeforeEach
    public void beforeEach() {
        student = new Student(1, 18, "姜辞旧");
        Student student2 = new Student(2, 19, "姜辞旧2");

        teacher = new Teacher(1, 28, "姜老师");
        Teacher2 teacher2 = new Teacher2(1, 28, "姜老师");
        teacher.setNickNameList(Arrays.asList("小姜", "老姜"));
        teacher.setStudentList(Arrays.asList(student, student2));
        teacher.setTeacher2(teacher2);
    }

    @Test
    @DisplayName("写入类测试")
    public void writerTest() throws IOException {
        AbstractCellStyleConfig sheetNameStyleConfig = SheetNameStyleConfig.builder().build();
        sheetNameStyleConfig.setBgColor(IndexedColors.GREY_25_PERCENT.index);
        sheetNameStyleConfig.setBold(true);
        sheetNameStyleConfig.setBorderStyle(BorderStyle.MEDIUM);
        AbstractCellStyleConfig headCellStyleConfig = HeadCellStyleConfig.builder().build();
        headCellStyleConfig.setBgColor(IndexedColors.GREY_25_PERCENT.index);
        headCellStyleConfig.setBold(true);
        headCellStyleConfig.setBorderStyle(BorderStyle.MEDIUM);
        AbstractCellStyleConfig contentCellStyleConfig = ContentCellStyleConfig.builder().build();
        contentCellStyleConfig.setBgColor(IndexedColors.WHITE.index);
        AbstractCellStyleConfig numberStyleConfig = ContentCellStyleConfig.builder().build();

        ExcelConfig excelConfig = ExcelConfig.builder()
                .sheetNameStyleConfig(sheetNameStyleConfig)
                .serialNumberStyleConfig(numberStyleConfig)
                .headCellStyleConfig(headCellStyleConfig)
                .contentCellStyleConfig(contentCellStyleConfig)
                .build();
        ExcelWriter.createExcelWriter("a.xlsx", excelConfig)
                .write(Student.class, Collections.singletonList(student))
                .write(Teacher.class, Collections.singletonList(teacher))
                .doWrite();
    }

    @Test
    public void poiTest() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssfWorkbook.createSheet();
        XSSFRow row = sheet.createRow(1);
        XSSFCell cell = row.createCell(2);
        cell.setCellValue("a");

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();

        Font titleFont = xssfWorkbook.createFont();
        titleFont.setBold(true); // 设置为粗体
        titleFont.setColor(IndexedColors.PINK.index);
        cellStyle.setFont(titleFont);

        cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.LEAST_DOTS);

        XSSFCell cell2 = row.createCell(3);
        cell2.setCellValue("a");
        cell2.setCellStyle(cellStyle);

//        mergeColCell(sheet, 1, 2, 2);
//        mergeColCell(sheet, 1, 3, 2);

        xssfWorkbook.write(new FileOutputStream("test.xlsx"));
    }

    @Test
    public void getClassTest() {
        int i = 0;
        Object o = i;
        System.out.println(o.getClass());
    }

    @Test
    public void reflectTest() throws NoSuchFieldException, IllegalAccessException {
        Class<? extends Teacher> teacherClass = teacher.getClass();
        Field nickNameList = teacherClass.getDeclaredField("nickNameList");
        nickNameList.setAccessible(true);
        Collection o = (Collection) nickNameList.get(teacher);
        o.forEach(System.out::println);
    }


}
