package com.tansun.excel.service;

import com.tansun.excel.model.ExcelTestModel;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface PersonInfoRepository extends JpaRepository<ExcelTestModel, Integer> { }
