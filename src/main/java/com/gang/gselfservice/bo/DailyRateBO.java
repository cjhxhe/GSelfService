package com.gang.gselfservice.bo;

import com.gang.gselfservice.enums.DailyRateResultEnum;
import com.gang.gselfservice.enums.DailyRateSourceEnum;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DailyRateBO {
    // 原文件 excel 解析后都为String
    private String no; // 序号
    private String date; // 日期
    private String rateName; // 评价人姓名
    private String rateRole; // 评价人身份
    private String phone; // 手机号码
    private String bizName; // 办理业务
    private String socialCreditCode; // 社会信用代码
    private String taxPayer; // 纳税人名称
    private String sendTime; // 发送时间
    private String comment; // 回复内容
    private DailyRateResultEnum rateResult; // 评价结果
    private String rateTime; // 评价时间
    private DailyRateSourceEnum source; // 来源
    private String taxAuthority; // 税务机关

    public static Map<String, String> colMap;

    static {
        colMap = new HashMap<>();
        colMap.put("序号", "no");
        colMap.put("日期", "date");
        colMap.put("评价人姓名", "rateName");
        colMap.put("评价人身份", "rateRole");
        colMap.put("手机号码", "phone");
        colMap.put("办理业务", "bizName");
        colMap.put("社会信用代码", "socialCreditCode");
        colMap.put("纳税人名称", "taxPayer");
        colMap.put("发送时间", "sendTime");
        colMap.put("回复内容", "comment");
        colMap.put("评价结果", "rateResult");
        colMap.put("评价时间", "rateTime");
        colMap.put("来源", "source");
        colMap.put("税务机关", "taxAuthority");
    }

    public static void setProperty(DailyRateBO payload, String name, String value) {
        try {
            Field field = DailyRateBO.class.getDeclaredField(name);
            field.setAccessible(true);
            if (field.getType() == DailyRateResultEnum.class) {
                field.set(payload, DailyRateResultEnum.getByResult(value));
            } else if (field.getType() == DailyRateSourceEnum.class) {
                field.set(payload, DailyRateSourceEnum.getByTitle(value));
            } else {
                field.set(payload, value);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
