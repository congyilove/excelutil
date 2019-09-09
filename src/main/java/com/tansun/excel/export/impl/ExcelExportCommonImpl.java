package com.tansun.excel.export.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
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

    }

    private void execute(ExcelExportParams params) {
        ExcelTypeEnum excelTypeEnum = params.getExcelTypeEnum();

        OutputStream outputStream = params.getOutputStream();

        EasyExcel.write(outputStream).excelType(excelTypeEnum).build();

    }
}
