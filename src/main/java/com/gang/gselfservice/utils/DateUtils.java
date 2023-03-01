package com.gang.gselfservice.utils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class DateUtils {

    public static final String DATETIME_FORMATTER = "yyyy-MM-dd HH:mm:ss";
    public static final String DATETIME_FORMATTER2 = "yyyyMMddHHmmss";
    public static final String DATE_FORMATTER_CN = "yyyy年M月dd日";
    public static final String DATE_FORMATTER = "yyyy-MM-dd";

    public static String getCurrentDate() {
        return getCurrentDate(DATE_FORMATTER_CN);
    }

    public static String getCurrentDate(String pattern) {
        return LocalDate.now().format(DateTimeFormatter.ofPattern(pattern));
    }

    public static Date parseDate(String date, String pattern) {
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
            return simpleDateFormat.parse(date);
        } catch (ParseException e) {
            return null;
        }
    }

    public static String getCurrentDateTime() {
        return getCurrentDateTime(DATETIME_FORMATTER);
    }

    public static String getCurrentDateTime(String pattern) {
        return LocalDateTime.now().format(DateTimeFormatter.ofPattern(pattern));
    }

}
