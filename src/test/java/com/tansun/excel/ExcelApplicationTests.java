package com.tansun.excel;

import com.alibaba.excel.support.ExcelTypeEnum;
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
import java.util.List;
import java.util.function.BiFunction;
import java.util.function.LongSupplier;

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

        File file = new File("d:\\tmp.xlsx");
        if(file.exists()){
            file.delete();
        }
        try(OutputStream outputStream = new FileOutputStream(file);){

            // count
            LongSupplier sp = () -> personInfoRepository.count();
            // dataLists
            BiFunction<Integer, Integer, List<ExcelTestModel>> bf = (pageSize, pageCount) ->
                personInfoRepository.findAll(PageRequest.of(pageSize, pageCount)).getContent();

            ExcelExportParams<ExcelTestModel> params = ExcelExportParams
                    .<ExcelTestModel>builder()
                    .excelTypeEnum(ExcelTypeEnum.XLSX)
                    .name("导出测试")
//                    .perDealRows(10)
//                    .perSheetRows(20)
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
