package com.tansun.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.IExcelExportService;
import com.tansun.excel.model.ExcelTestModel;
import com.tansun.excel.service.PersonInfoRepository;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.data.domain.PageRequest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.util.Arrays;
import java.util.List;
import java.util.function.BiFunction;
import java.util.function.LongSupplier;
import java.util.function.Supplier;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelApplicationTests {

    @Autowired
    @Qualifier("excelExportCommonImpl")
    IExcelExportService excelExportService;

    @Autowired
    PersonInfoRepository personInfoRepository;

    @Test
    public void contextLoads() {
    }

    @Test
    public void test() {

        try(OutputStream outputStream = new FileOutputStream(new File("d:\\tmp\\"));){

            // count
            LongSupplier sp = () -> personInfoRepository.count();
            // dataLists
            BiFunction<Integer, Integer, List<ExcelTestModel>> bf = (pageSize, pageCount) ->
                personInfoRepository.findAll(PageRequest.of(pageSize, pageCount)).getContent();

            ExcelExportParams<ExcelTestModel> params = ExcelExportParams
                    .<ExcelTestModel>builder()
                    .excelTypeEnum(ExcelTypeEnum.XLSX)
                    .name("导出测试")
                    //.perDealRows(10)
                    //.perSheetRows(20)
                    .sheetName("infomation")
                    .supplierCount(sp)
                    .functionData(bf)
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
