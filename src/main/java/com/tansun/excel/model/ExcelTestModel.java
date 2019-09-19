package com.tansun.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.*;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Data
@NoArgsConstructor
@Setter
@Entity
@Table(name="person_info")
public class ExcelTestModel {

    @Id
    private int pk;

    @Column
    @ExcelProperty("姓名")
    private String name;
    @Column
    @ExcelProperty("性别")
    private String sex;
    @Column
    @ExcelProperty("年龄")
    private int age;
    @Column
    @ExcelProperty("家庭住址")
    private String address;
    @Column
    @ExcelProperty("电话")
    private String telephone;
    @Column
    @ExcelProperty("公司")
    private String company;
    @Column
    @ExcelProperty("兴趣爱好")
    private String hobby;
    @Column
    @ExcelProperty("毕业院校")
    private String graduateSchool;
    @Column
    @ExcelProperty("邮箱")
    private String email;
    @Column
    @ExcelProperty("是否已婚")
    private Boolean isMarried;

}
