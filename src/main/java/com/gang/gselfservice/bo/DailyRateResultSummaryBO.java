package com.gang.gselfservice.bo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DailyRateResultSummaryBO {

    private int verySatisfiedNum; // 非常满意
    private int satisfiedNum; // 基本满意
    private int notSatisfiedNum; // 不满意
    private int invalidNum; // 无效评价
    private int validNum; // 有效评价
    private int notRateNum; // 未评价
    private double verySatisfiedPercentage; // 非常满意率
}
