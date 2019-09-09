package com.tansun.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.Arrays;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelApplicationTests {

    @Test
    public void contextLoads() {
    }

    @Test
    public void test(){

        ExcelWriter ewb = EasyExcel.write().excelType(ExcelTypeEnum.XLSX).file("./test.xlsx").build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();

        ewb.write(Arrays.asList("123", "345"), writeSheet);
    }
}
