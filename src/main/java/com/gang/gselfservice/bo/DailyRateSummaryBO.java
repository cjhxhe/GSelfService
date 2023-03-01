package com.gang.gselfservice.bo;

import com.gang.gselfservice.enums.DailyRateSourceEnum;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DailyRateSummaryBO {

    private DailyRateSourceEnum dailyRateSourceEnum; // 类型

    private int totalNum; // 总评价（纳入监控）户数

    private int ratedNum; // 主动评价 户数

    private DailyRateResultSummaryBO dailyRateResultSummaryBO; // 满意度汇总情况
}
