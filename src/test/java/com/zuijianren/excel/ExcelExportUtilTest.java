package com.zuijianren.excel;

import com.zuijianren.excel.core.ExcelWriter;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;

/**
 * @author zuijianren
 * @date 2023/3/13 13:40
 */
public class ExcelExportUtilTest {


    @Test
    @DisplayName("写入类测试")
    public void writerTest() throws IOException {
        // 数据准备
        Student student = new Student(1, 18, "姜辞旧");

        Teacher teacher = new Teacher(1, 28, "姜老师");
        teacher.setNickNameList(Arrays.asList("小姜", "老姜"));

        // 写入
        ExcelWriter excelWriter = ExcelWriter.createExcelWriter("a.xlsx", null);

        excelWriter
                .write(Student.class, Collections.singletonList(student))
                .write(Teacher.class, Collections.singletonList(teacher))
                .doWrite();
    }


}
