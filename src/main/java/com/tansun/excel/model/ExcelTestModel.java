package com.tansun.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class ExcelTestModel {

    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("性别")
    private String sex;
    @ExcelProperty("家庭住址")
    private String address;
    @ExcelProperty("年龄")
    private int age;
    @ExcelProperty("电话")
    private String telephone;
    @ExcelProperty("公司")
    private String company;
    @ExcelProperty("兴趣爱好")
    private String hobby;
    @ExcelProperty("毕业院校")
    private String graduateSchool;
    @ExcelProperty("邮箱")
    private String email;
    @ExcelProperty("是否已婚")
    private boolean isMarried;
}
