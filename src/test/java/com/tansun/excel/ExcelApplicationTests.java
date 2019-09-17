package com.tansun.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.IExcelExportService;
import com.tansun.excel.model.ExcelTestModel;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.util.Arrays;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelApplicationTests {

    @Autowired
    @Qualifier("excelExportCommonImpl")
    IExcelExportService excelExportService;

    @Test
    public void contextLoads() {
    }

    @Test
    public void test() {

        try(OutputStream outputStream = new FileOutputStream(new File("d:\\tmp\\"));){

            // count

            // dataLists

            ExcelExportParams<ExcelTestModel> params = ExcelExportParams
                    .<ExcelTestModel>builder()
                    .excelTypeEnum(ExcelTypeEnum.XLSX)
                    .name("导出测试")
                    .perDealRows(10)
                    .perSheetRows(20)
                    .sheetName("infomation")
                    .outputStream(outputStream)
                    .build();
            excelExportService.export(params);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {

        }


    }
}
