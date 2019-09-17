package com.tansun.excel.export.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.IExcelExportService;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.List;

@Service
public class ExcelExportCommonImpl<T>  implements IExcelExportService {

    @Override
    public void export(ExcelExportParams params) {
        check(params);
        execute(params);
    }

    private void check(ExcelExportParams params){
        //Preconditions.checkArgument();
    }

    private void execute(ExcelExportParams<T> params) {
        ExcelTypeEnum excelTypeEnum = params.getExcelTypeEnum();

        int sheetCount = 0; // sheet数
        int pDealCount = 0; // 中间sheet查询次数
        int lDealCount = 0; // 最后一个sheet查询次数
        int perSheetRows = params.getPerSheetRows();
        int perDealRows = params.getPerDealRows();
        int totalRows = params.getSupplierCount().getAsInt();

        OutputStream outputStream = params.getOutputStream();

        ExcelWriter excelWriter = EasyExcel.write(outputStream).excelType(excelTypeEnum).build();
        WriteSheet ws = EasyExcel.writerSheet(params.getSheetName()).build();

        // 一个sheet内
        if(totalRows <= perDealRows){
            // 一次处理完成
            sheetCount = 1;
            pDealCount = lDealCount = 1;
            List<T> dataList = (List<T>) params.getSupplierData().get();
            excelWriter.write(dataList, ws);
        }
        // 一个sheet内，处理多次
        if(totalRows <= perSheetRows){
            sheetCount = 1;
            pDealCount = lDealCount = totalRows%perDealRows == 0 ? totalRows/perDealRows : totalRows/perDealRows + 1;
        }

        // 多个sheet
        // 单sheet处理次数
        // 中间sheet处理次数
        //pDealCount = perSheetRows%perDealRows == 0 ? perSheetRows/perDealRows : perSheetRows/perDealRows + 1;
        // 中间sheet,单sheet内，取余数，保证前面sheet数据都等于pSheetRows
        //int pLeft = perSheetRows%perDealRows;
        pDealCount = perSheetRows/perDealRows; // 必须保证可以整除

        // 用于计算sheet数，要求每个sheet必须处理perSheetRows
        int tmpCount = totalRows%perSheetRows;
        // 最后一个sheet
        if(tmpCount == 0){
            lDealCount = pDealCount;
            sheetCount = totalRows/perSheetRows;
        } else {
            lDealCount = tmpCount%perDealRows == 0 ?
                tmpCount/perDealRows :
                tmpCount/perDealRows + 1;
            sheetCount = totalRows/perSheetRows + 1;
        }

        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            ws = EasyExcel.writerSheet(params.getSheetName() + sheetIndex).build();
            if(sheetIndex < sheetCount - 1){
                export(params, sheetIndex, sheetCount, pDealCount, excelWriter, ws);
            } else{
                // 最后一个sheet
                export(params, sheetIndex, sheetCount, lDealCount, excelWriter, ws);
            }
        }


    }

    private void export(ExcelExportParams<T> params, Integer sheetIndex, Integer sheetCount, Integer dealCount, ExcelWriter ew, WriteSheet ws){
        for (int dealIndex = 0; dealIndex < dealCount; dealIndex++){

            List<T> dataLists = params.getFunctionData().apply(sheetIndex * dealCount + dealIndex, dealCount);

            ew.write(dataLists, ws);

        }
    }

}
