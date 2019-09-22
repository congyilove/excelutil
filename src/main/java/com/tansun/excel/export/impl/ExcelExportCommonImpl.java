package com.tansun.excel.export.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.google.common.base.Preconditions;
import com.tansun.excel.common.ExcelConstants;
import com.tansun.excel.export.ExcelExportParams;
import com.tansun.excel.export.ExportComParams;
import com.tansun.excel.export.IExcelExportService;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.List;
import java.util.Optional;

import static com.google.common.base.Preconditions.checkNotNull;

@Service
public class ExcelExportCommonImpl<T>  implements IExcelExportService<T> {

    @Override
    public void export(ExcelExportParams<T> params) {
        checkNotNull(params);

        check(params);
        ExportComParams<T> exportComParams = calculate(params);
        doExport(exportComParams);

    }

    private void check(ExcelExportParams params){
        checkNotNull(params.getOutputStream(), "没有指定输出流");
        int perDealRows  = params.getPerDealRows();
        int perSheetRows = params.getPerSheetRows();
        // 默认为XLSX格式
        ExcelTypeEnum excelTypeEnum = Optional.of(params.getExcelTypeEnum()).orElse(ExcelTypeEnum.XLSX);
        switch (excelTypeEnum){
            case XLSX:
                if(perDealRows == 0){
                    perDealRows = ExcelConstants.PER_DEAL_ROWS_2007;
                }
                if(perSheetRows == 0){
                    perSheetRows= ExcelConstants.SHEET_MAX_ROWS_2007;
                }
                break;
            case XLS:
                if(perDealRows == 0){
                    perDealRows = ExcelConstants.PER_DEAL_ROWS_2003;
                }
                if(perSheetRows == 0){
                    perSheetRows= ExcelConstants.SHEET_MAX_ROWS_2003;
                }
                break;
            default:
                break;
        }

        Preconditions.checkState(perSheetRows%perDealRows == 0, "必须整除");
        params.setPerDealRows(perDealRows);
        params.setPerSheetRows(perSheetRows);
    }

    private ExportComParams<T> calculate(ExcelExportParams<T> params) {
        int sheetCount; // sheet数
        int pDealCount; // 中间sheet查询次数
        int lDealCount; // 最后一个sheet查询次数
        int perSheetRows = params.getPerSheetRows() ;
        int perDealRows  = params.getPerDealRows();
        // 可以肯定count不会超过2^31
        int totalRows    = (int)params.getSupplierCount().getAsLong();

        if(totalRows <= perSheetRows){
            // 一个sheet内，处理多次
            // 一个sheet内
            if(totalRows <= perDealRows){
                // 一次处理完成
                sheetCount = 1;
                pDealCount = lDealCount = 1;
             } else {
                sheetCount = 1;
                pDealCount = lDealCount = totalRows%perDealRows == 0 ? totalRows/perDealRows : totalRows/perDealRows + 1;
            }
        } else {
            // 多个sheet
            // 单sheet处理次数
            // 中间sheet处理次数
            // pDealCount = perSheetRows%perDealRows == 0 ? perSheetRows/perDealRows : perSheetRows/perDealRows + 1;
            // 中间sheet,单sheet内，取余数，保证前面sheet数据都等于pSheetRows
            //int pLeft = perSheetRows%perDealRows;
            pDealCount = perSheetRows/perDealRows; // TODO 必须保证可以整除

            // 用于计算sheet数，要求每个sheet必须处理perSheetRows
            int tmpCount = totalRows%perSheetRows;
            // 最后一个sheet
            if(tmpCount == 0){
                sheetCount = totalRows/perSheetRows;
                lDealCount = pDealCount;
            } else {
                sheetCount = totalRows/perSheetRows + 1;
                lDealCount = tmpCount%perDealRows == 0 ?
                    tmpCount/perDealRows :
                    tmpCount/perDealRows + 1;
            }
        }
        return ExportComParams.<T>builder().excelExportParams(params).sheetCount(sheetCount).pDealCount(pDealCount).lDealCount(lDealCount).build();
    }

    private void doExport(ExportComParams<T> exportComParams){
        ExcelExportParams<T> params = exportComParams.getExcelExportParams();
        int sheetCount = exportComParams.getSheetCount();
        int pDealCount = exportComParams.getPDealCount();
        int lDealCount = exportComParams.getLDealCount();

        ExcelTypeEnum excelTypeEnum = params.getExcelTypeEnum();
        OutputStream outputStream   = params.getOutputStream();

        ExcelWriter excelWriter = EasyExcel.write(outputStream).excelType(excelTypeEnum). build();

        int dealCount;
        int perDealRows = params.getPerDealRows();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(params.getSheetName() + sheetIndex).sheetNo(sheetIndex).build();
            if(sheetIndex < sheetCount - 1){
                dealCount = pDealCount;
            } else{
                // 最后一个sheet
                dealCount = lDealCount;
            }
            for (int dealIndex = 0; dealIndex < dealCount; dealIndex++){
                List<T> dataLists = params.getFunctionData().apply(sheetIndex * pDealCount + dealIndex, perDealRows);
                excelWriter.write(dataLists, writeSheet);

            }
        }
        excelWriter.finish();

    }

}
