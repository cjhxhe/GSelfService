package com.gang.gselfservice.enums;

public enum DailyRateResultEnum {

    VERY_SATISFIED("非常满意", 1),
    SATISFIED("基本满意", 2),
    NOT_SATISFIED("不满意", 3),
    INVALID("无效", 4),
    BLANK("空白", 99),
    ;

    private String result;
    private int order;

    DailyRateResultEnum(String result, int order) {
        this.result = result;
        this.order = order;
    }

    public String getResult() {
        return result;
    }

    public static DailyRateResultEnum getByResult(String title) {
        for (DailyRateResultEnum value : DailyRateResultEnum.values()) {
            if (value.getResult().equals(title)) {
                return value;
            }
        }
        return BLANK; // 空白(默认)
    }
}
