package com.gang.gselfservice.enums;

public enum DailyRateSourceEnum {
    ETAX_RATE("电子税务局(好差评)", 1, "电子税务局(好差评)"),
    ETAX_SMS("电子税务局(短信)", 2, "电子税务局（满意度）"),
    VISIT_SMS("进厅人员(短信)", 3, "进厅人员（满意度）"),
    MANUAL("人工咨询", 4, "人工咨询（满意度）"),
    BLANK("空白", 99, "");

    private String title;
    private int order;
    private String desc;

    DailyRateSourceEnum(String title, int order, String desc) {
        this.title = title;
        this.order = order;
        this.desc = desc;
    }

    public String getTitle() {
        return title;
    }

    public int getOrder() {
        return order;
    }

    public String getDesc() {
        return desc;
    }

    public static DailyRateSourceEnum getByTitle(String title) {
        for (DailyRateSourceEnum value : DailyRateSourceEnum.values()) {
            if (value.getTitle().equals(title)) {
                return value;
            }
        }
        return BLANK; // 空白(默认)
    }
}
