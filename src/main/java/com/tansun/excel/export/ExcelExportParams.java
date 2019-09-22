package com.tansun.excel.export;

import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.io.OutputStream;
import java.util.List;
import java.util.function.BiFunction;
import java.util.function.LongSupplier;

@Setter
@Getter
@Builder
public class ExcelExportParams<T> {

    // excel 类型
    private ExcelTypeEnum excelTypeEnum;
    // excel name
    private String name;
    // 输出流
    private OutputStream outputStream;
    // 模板路径
    private String templatePath;
    // sheet 名
    private String sheetName;
    // model
    private Class<T> clazz;
    // 表头
    private  String[] headers;
    // 每次处理条数
    private int perDealRows;
    // 每个sheet存放数据条数
    private int perSheetRows;

    // 分页处理
    private BiFunction<Integer, Integer, List<T>> functionData;
    // 总数据量
    private LongSupplier supplierCount;


}
