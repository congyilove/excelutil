package com.tansun.excel.common;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.commons.codec.Charsets;

import javax.servlet.http.HttpServletResponse;

// 考虑带模板导出和无模板导出
// 导出到web还是local
// excel的版本
public class ExcelUtils {

    public static void export(){
    //    EasyExcel.write().withTemplate();
    }

    public static void exportToWeb(HttpServletResponse response){

    }

    public static void exportToLocal(){

    }


    /**
     * 将文件输出到浏览器的公共配置
     * @param response 响应
     * @param fileName 文件名（不含拓展名）
     * @param typeEnum excel类型
     */
    private static void prepareResponse(HttpServletResponse response, String fileName, ExcelTypeEnum typeEnum) {
        String fileName2Export = new String((fileName).getBytes(Charsets.UTF_8), Charsets.ISO_8859_1);
        response.setContentType("multipart/form-data");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName2Export + typeEnum.getValue());
    }

}
