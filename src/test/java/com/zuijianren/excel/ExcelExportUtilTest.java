package com.zuijianren.excel;

import com.zuijianren.excel.core.ExcelWriter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;

import static com.zuijianren.excel.core.ExcelOperator.mergeColCell;

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
        // 写入
        ExcelWriter excelWriter = ExcelWriter.createExcelWriter("a.xlsx", null);

        excelWriter
//                .write(Student.class, Collections.singletonList(student))
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

        XSSFCell cell2 = row.createCell(3);
        cell2.setCellValue("a");

        mergeColCell(sheet, 1, 2, 2);
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
