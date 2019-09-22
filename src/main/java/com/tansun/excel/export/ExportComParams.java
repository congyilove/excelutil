package com.tansun.excel.export;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Builder
public class ExportComParams<T> {

    private ExcelExportParams<T> excelExportParams;

    private int sheetCount;

    private int pDealCount;

    private int lDealCount;

}
