package com.tansun.excel.export;

public interface IExcelExportService<T> {

    void export(ExcelExportParams<T> params);

}
