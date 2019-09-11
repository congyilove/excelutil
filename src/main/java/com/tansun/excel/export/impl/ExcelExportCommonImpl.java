package com.tansun.excel.export.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.google.common.base.Preconditions;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.IExcelExportService;

import java.io.OutputStream;
import java.util.List;

public class ExcelExportCommonImpl implements IExcelExportService {

    @Override
    public void export(ExcelExportParams params) {
        check(params);
        execute(params);
    }

    private void check(ExcelExportParams params){
        //Preconditions.checkArgument();
    }

    private <T> void execute(ExcelExportParams<T> params) {
        ExcelTypeEnum excelTypeEnum = params.getExcelTypeEnum();

        int perSheetRows = params.getPerSheetRows();
        int perDealRows = params.getPerDealRows();
        int totalRows = params.getSupplierCount().getAsInt();

        OutputStream outputStream = params.getOutputStream();

        ExcelWriter excelWriter = EasyExcel.write(outputStream).excelType(excelTypeEnum).build();
        WriteSheet ws = EasyExcel.writerSheet(params.getSheetName()).build();

        // 一个sheet内
        if(totalRows <= perDealRows){
            // 一次处理完成

            List<T> dataList = (List<T>) params.getSupplierData().get();
            excelWriter.write(dataList, ws);
        }
        // 一个sheet内，处理多次
        if(totalRows <= perSheetRows){
            int count = totalRows%perDealRows == 0 ? totalRows/perDealRows : totalRows/perDealRows + 1;

        }

        // 多个sheet
        // sheet数，要求每个sheet必须处理perSheetRows
        int tmpCount = totalRows%perSheetRows;
        int sheetCount = tmpCount == 0 ? totalRows/perSheetRows : totalRows/perSheetRows + 1;
        // 单sheet处理次数
        int pDealCount = perSheetRows/perDealRows;
        // 单sheet内，最后处理的偏移量
        int pOffset = perSheetRows%perDealRows;

        // 最后一个sheet
        int lDealCount = 0;
        int lOffset = 0;
        if(tmpCount == 0){
            lDealCount = pDealCount;
            lOffset = pOffset;
        } else {
            if(tmpCount%perDealRows == 0) {
                lDealCount = tmpCount/perDealRows;
                lOffset = 0;
            }else{
                lDealCount = tmpCount/perDealRows + 1;
                lOffset = tmpCount%perDealRows;
            }
        }


    }
}
