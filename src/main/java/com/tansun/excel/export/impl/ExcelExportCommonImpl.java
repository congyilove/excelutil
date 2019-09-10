package com.tansun.excel.export.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.google.common.base.Preconditions;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.IExcelExportService;

import java.io.OutputStream;

public class ExcelExportCommonImpl implements IExcelExportService {

    @Override
    public void export(ExcelExportParams params) {
        check(params);
        execute(params);
    }

    private void check(ExcelExportParams params){
        //Preconditions.checkArgument();
    }

    private void execute(ExcelExportParams params) {
        ExcelTypeEnum excelTypeEnum = params.getExcelTypeEnum();

        int perSheetRows = params.getPerSheetRows();
        int perDealRows = params.getPerDealRows();
        int totalRows = params.getSupplierCount().getAsInt();

        if(totalRows <= perDealRows){
            // 一次处理完成
        }
        if(totalRows <= perSheetRows){

        }
        // sheet数
        int sheetCount = totalRows%perSheetRows == 0 ? totalRows/perSheetRows : totalRows/perSheetRows + 1;

        int m = perSheetRows%perDealRows == 0 ? perSheetRows/perDealRows : perSheetRows/perDealRows + 1;

        if(perDealRows <= perSheetRows){
        }

        OutputStream outputStream = params.getOutputStream();

        EasyExcel.write(outputStream).excelType(excelTypeEnum).build();

    }
}
