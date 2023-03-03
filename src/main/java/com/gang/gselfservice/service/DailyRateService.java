package com.gang.gselfservice.service;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import com.gang.gselfservice.bo.DailyRateBO;
import com.gang.gselfservice.bo.DailyRateResultSummaryBO;
import com.gang.gselfservice.bo.DailyRateSummaryBO;
import com.gang.gselfservice.config.ThreadPool;
import com.gang.gselfservice.enums.DailyRateResultEnum;
import com.gang.gselfservice.enums.DailyRateSourceEnum;
import com.gang.gselfservice.exception.BizException;
import com.gang.gselfservice.task.DBUpsertTask;
import com.gang.gselfservice.utils.DateUtils;
import com.gang.gselfservice.utils.ExcelStreamingUtils;
import com.gang.gselfservice.utils.ExcelUtils;
import com.gang.gselfservice.utils.ExcelUtilsInterface;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author yugang
 */
public class DailyRateService {

    private static final String DOC_TEMPLATE = "/templates/template.docx";

    private static final DecimalFormat PERCENTAGE_FORMAT = new DecimalFormat("0.00%");

    // 每日评价汇总列表
    private List<DailyRateSummaryBO> dailyRateSummaryBOList;

    // 需要回访的列表
    private Map<DailyRateSourceEnum, List<DailyRateBO>> needFeedBackList;

    // 表头对应序号
    private Map<String, Integer> colIndexMap;

    // 处理的数据的日期
    private String processingDate;

    /**
     * 导出word 文档
     *
     * @param outputPath
     */
    public void exportToWord(String outputPath) throws Exception {
        Map<String, Object> paramsMap = new HashMap<>();

        // 基本信息
        paramsMap.put("today", getProcessingDate());

        // 详细数据
        dailyRateSummaryBOList.sort((o1, o2) -> {
            DailyRateSourceEnum dailyRateSourceEnum1 = o1.getDailyRateSourceEnum();
            DailyRateSourceEnum dailyRateSourceEnum2 = o2.getDailyRateSourceEnum();
            if (dailyRateSourceEnum1 == null && dailyRateSourceEnum2 == null) {
                return 0;
            } else if (dailyRateSourceEnum1 == null) {
                return -1;
            } else if (dailyRateSourceEnum2 == null) {
                return 1;
            } else if (dailyRateSourceEnum1.getOrder() == dailyRateSourceEnum2.getOrder()) {
                return 0;
            } else {
                return dailyRateSourceEnum1.getOrder() > dailyRateSourceEnum2.getOrder() ? 1 : -1;
            }
        });

        for (DailyRateSummaryBO dailyRateSummaryBO : dailyRateSummaryBOList) {
            String prefix;
            if (dailyRateSummaryBO.getDailyRateSourceEnum() != null) {
                prefix = dailyRateSummaryBO.getDailyRateSourceEnum().name().toLowerCase();
            } else {
                prefix = "total";
            }
            paramsMap.put(String.format("%s_totalNum", prefix), dailyRateSummaryBO.getTotalNum());
            paramsMap.put(String.format("%s_ratedNum", prefix), dailyRateSummaryBO.getRatedNum());
            paramsMap.put(String.format("%s_verySatisfiedNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getVerySatisfiedNum());
            paramsMap.put(String.format("%s_satisfiedNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getSatisfiedNum());
            paramsMap.put(String.format("%s_notSatisfiedNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getNotSatisfiedNum());
            paramsMap.put(String.format("%s_validNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getValidNum());
            paramsMap.put(String.format("%s_verySatisfiedPercentage", prefix), PERCENTAGE_FORMAT.format(dailyRateSummaryBO.getDailyRateResultSummaryBO().getVerySatisfiedPercentage()));
            paramsMap.put(String.format("%s_invalidNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getInvalidNum());
            paramsMap.put(String.format("%s_notRateNum", prefix), dailyRateSummaryBO.getDailyRateResultSummaryBO().getNotRateNum());
        }

        // 回访情况表
        TableRenderData feedbackTable;
        // 表头
        RowRenderData tableHead = Rows.of("来源", "序号", "税号", "名称", "办税身份", "姓名", "联系电话", "回复内容", "评价分类", "回访情况")
                .textFontSize(10).center().create();
        Tables.TableBuilder tableBuilder = Tables.of(tableHead).autoWidth();
        MergeCellRule.MergeCellRuleBuilder mergeCellRuleBuilder = MergeCellRule.builder();

        // 表数据
        int totalRow = 1;
        for (DailyRateSourceEnum sourceEnum : DailyRateSourceEnum.values()) {
            if (sourceEnum == DailyRateSourceEnum.BLANK) {
                // 空白的不管
                continue;
            }
            List<DailyRateBO> rateBOList = needFeedBackList.get(sourceEnum);
            if (CollectionUtils.isEmpty(rateBOList)) {
                RowRenderData emptyRow = Rows.of(sourceEnum.getDesc(), "无", StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY,
                        StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY).textFontSize(10).center().create();
                tableBuilder.addRow(emptyRow);
                totalRow++;
            } else {
                RowRenderData[] rows = new RowRenderData[rateBOList.size()];
                for (int i = 0; i < rateBOList.size(); i++) {
                    DailyRateBO dailyRateBO = rateBOList.get(i);
                    rows[i] = Rows.of(sourceEnum.getDesc(), String.valueOf(i + 1), dailyRateBO.getSocialCreditCode(), dailyRateBO.getTaxPayer(),
                            dailyRateBO.getRateRole(), dailyRateBO.getRateName(), dailyRateBO.getPhone(), dailyRateBO.getComment(),
                            dailyRateBO.getRateResult().getResult(), StringUtils.EMPTY).textFontSize(10).center().create();
                    tableBuilder.addRow(rows[i]);
                    totalRow++;
                }
                // 首列合并(只有一行数据就不合并了)
                if (1 < rateBOList.size()) {
                    mergeCellRuleBuilder.map(MergeCellRule.Grid.of(totalRow - rateBOList.size(), 0), MergeCellRule.Grid.of(totalRow - 1, 0));
                }
            }
        }

        feedbackTable = tableBuilder.mergeRule(mergeCellRuleBuilder.build()).center().create();
        paramsMap.put("feedback_table", feedbackTable);

        // 写word出去
        URL url = DailyRateService.class.getResource(DOC_TEMPLATE);
        XWPFTemplate xwpfTemplate = XWPFTemplate.compile(url.openStream()).render(paramsMap);
        xwpfTemplate.writeAndClose(new FileOutputStream(outputPath));
    }

    private String zeroToBlank(int input) {
        if (0 == input) {
            return StringUtils.EMPTY;
        }
        return String.valueOf(input);
    }

    /**
     * 解析每一行数据
     *
     * @param row
     * @return
     */
    private DailyRateBO parseOneRow(List<String> row) {
        DailyRateBO dailyRateBO = new DailyRateBO();
        for (Map.Entry<String, Integer> entry : colIndexMap.entrySet()) {
            DailyRateBO.setProperty(dailyRateBO, entry.getKey(), row.get(entry.getValue()));
        }
        return dailyRateBO;
    }

    private void parseSheetHeader(List<String> headerLine) {
        Map<String, Integer> oriColIndexMap = new HashMap<>();
        for (int i = 0; i < headerLine.size(); i++) {
            oriColIndexMap.put(StringUtils.trim(headerLine.get(i)), i);
        }
        colIndexMap = new HashMap<>();
        for (Map.Entry<String, Integer> entry : oriColIndexMap.entrySet()) {
            if (DailyRateBO.colMap.containsKey(entry.getKey())) {
                colIndexMap.put(DailyRateBO.colMap.get(entry.getKey()), entry.getValue());
            }
        }
    }

    /**
     * 设置数据的日期
     *
     * @param data
     */
    private void setProcessingDate(DailyRateBO data) {
        processingDate = data.getDate();
    }

    /**
     * 获得处理数据的日期，没有则默认当天
     *
     * @return
     */
    private String getProcessingDate() {
        return StringUtils.defaultString(processingDate, DateUtils.getCurrentDate());
    }

    public static void main(String[] args) throws Exception {
        String filePath = "/Users/yugang/Downloads/满意度数据全库数据2022.01-12.xlsx";
        File file = new File(filePath);
        FileInputStream fis = new FileInputStream(file);
        ExcelStreamingUtils excelUtils = new ExcelStreamingUtils(fis);
        List<List<String>> lists = excelUtils.read(0);
        lists.stream().limit(100).forEach(System.out::println);
    }

    /**
     * 解析整个文档
     *
     * @param filePath
     */
    public void dealDailyRateExcel(String filePath) throws Exception {
        File file = new File(filePath);
        FileInputStream fis = new FileInputStream(file);

        // 兼容两种Excel
        String excelVersion = ExcelUtils.getExcelVersion(filePath);
        ExcelUtilsInterface excelUtils = ExcelUtils.LEGACY_VERSION.equals(excelVersion) ? new ExcelUtils(fis, excelVersion) : new ExcelStreamingUtils(fis);

        // 解析数据
        List<List<String>> lists = excelUtils.read(0);
        if (CollectionUtils.isEmpty(lists)) {
            throw new BizException("请检查Excel,没找到需要处理的数据!");
        }

        // 检查一下Excel sheet数
        List<String> headerLine = lists.get(0);
        int sheetCount = excelUtils.getSheetCount();
        if (1 < sheetCount && 10 > headerLine.size()) {
            throw new BizException("请把需要处理的数据放在第一页!");
        }

        // 解析表头
        parseSheetHeader(headerLine);
        lists.remove(0);

        // 解析数据
        List<DailyRateBO> dailyRateBOList = new ArrayList<>(lists.size());
        for (List<String> row : lists) {
            if (null == row.get(0)) {
                break; // 如果遇到空行则停止
            }
            dailyRateBOList.add(parseOneRow(row));
        }

        if (CollectionUtils.isEmpty(dailyRateBOList)) {
            throw new BizException("没有找到评价数据!");
        } else {
            // 获取数据时间
            setProcessingDate(dailyRateBOList.get(0));
        }

        // 分来源
        Map<DailyRateSourceEnum, List<DailyRateBO>> sourcedList = dailyRateBOList.stream().collect(Collectors.groupingBy(DailyRateBO::getSource));
        dailyRateSummaryBOList = new ArrayList<>(sourcedList.size());
        for (Map.Entry<DailyRateSourceEnum, List<DailyRateBO>> entry : sourcedList.entrySet()) {
            DailyRateSummaryBO dailyRateSummaryBO = new DailyRateSummaryBO();
            dailyRateSummaryBO.setDailyRateSourceEnum(entry.getKey()); // 类型

            List<DailyRateBO> rateBOList = entry.getValue();
            dailyRateSummaryBO.setTotalNum(rateBOList.size()); // 总纳入数量

            Map<DailyRateResultEnum, List<DailyRateBO>> rateResultList = rateBOList.stream().collect(Collectors.groupingBy(DailyRateBO::getRateResult));
            DailyRateResultSummaryBO dailyRateResultSummaryBO = new DailyRateResultSummaryBO(); // 各类数量汇总
            for (Map.Entry<DailyRateResultEnum, List<DailyRateBO>> resultEntry : rateResultList.entrySet()) {
                switch (resultEntry.getKey()) {
                    case VERY_SATISFIED:
                        dailyRateResultSummaryBO.setVerySatisfiedNum(resultEntry.getValue().size());
                        break;
                    case SATISFIED:
                        dailyRateResultSummaryBO.setSatisfiedNum(resultEntry.getValue().size());
                        break;
                    case NOT_SATISFIED:
                        dailyRateResultSummaryBO.setNotSatisfiedNum(resultEntry.getValue().size());
                        break;
                    case INVALID:
                        dailyRateResultSummaryBO.setInvalidNum(resultEntry.getValue().size());
                        break;
                    case BLANK:
                        dailyRateResultSummaryBO.setNotRateNum(resultEntry.getValue().size());
                        break;
                    default:
                        break;
                }
            }

            // 有效评价小计
            dailyRateResultSummaryBO.setValidNum(dailyRateResultSummaryBO.getVerySatisfiedNum()
                    + dailyRateResultSummaryBO.getSatisfiedNum()
                    + dailyRateResultSummaryBO.getNotSatisfiedNum());

            // 非常满意率
            if (dailyRateResultSummaryBO.getValidNum() == 0) {
                dailyRateResultSummaryBO.setVerySatisfiedPercentage(0d);
            } else {
                dailyRateResultSummaryBO.setVerySatisfiedPercentage((dailyRateResultSummaryBO.getVerySatisfiedNum() + 0.0) / dailyRateResultSummaryBO.getValidNum());
            }

            // 主动评价（户数）
            dailyRateSummaryBO.setRatedNum(dailyRateResultSummaryBO.getValidNum() + dailyRateResultSummaryBO.getInvalidNum());

            dailyRateSummaryBO.setDailyRateResultSummaryBO(dailyRateResultSummaryBO);
            dailyRateSummaryBOList.add(dailyRateSummaryBO);
        }

        // 小计（不含好差评）
        DailyRateSummaryBO totalDailyRateSummaryBO = getTotalDailyRateResultSummaryBO(dailyRateSummaryBOList);
        dailyRateSummaryBOList.add(totalDailyRateSummaryBO);

        // 回访列表
        needFeedBackList = getNeedFeedBackList(dailyRateBOList);

        excelUtils.close();

        // 落库
        ThreadPool.getExecutorService().submit(new DBUpsertTask(dailyRateBOList));
    }

    /**
     * 补全所有来源
     */
    public void patchDailyRateSummary() {
        for (DailyRateSourceEnum sourceEnum : DailyRateSourceEnum.values()) {
            Optional<DailyRateSummaryBO> any = dailyRateSummaryBOList.stream().filter(s -> s.getDailyRateSourceEnum() == sourceEnum).findAny();
            if (!any.isPresent()) {
                dailyRateSummaryBOList.add(DailyRateSummaryBO.builder()
                        .dailyRateSourceEnum(sourceEnum)
                        .dailyRateResultSummaryBO(new DailyRateResultSummaryBO())
                        .build());
            }
        }
    }

    /**
     * 小计（不含好差评）
     *
     * @param dailyRateSummaryBOList
     * @return
     */
    private DailyRateSummaryBO getTotalDailyRateResultSummaryBO(List<DailyRateSummaryBO> dailyRateSummaryBOList) {
        // 不含好差评,要去掉
        List<DailyRateSummaryBO> dataList = dailyRateSummaryBOList.stream().filter(r -> r.getDailyRateSourceEnum() != DailyRateSourceEnum.ETAX_RATE).collect(Collectors.toList());

        DailyRateSummaryBO totalDailyRateSummaryBO = new DailyRateSummaryBO();
        int totalSum = dataList.stream().mapToInt(DailyRateSummaryBO::getTotalNum).sum();
        int totalRatedSum = dataList.stream().mapToInt(DailyRateSummaryBO::getRatedNum).sum();
        totalDailyRateSummaryBO.setTotalNum(totalSum);
        totalDailyRateSummaryBO.setRatedNum(totalRatedSum);
        DailyRateResultSummaryBO totalDailyRateResultSummaryBO = new DailyRateResultSummaryBO();
        totalDailyRateResultSummaryBO.setNotRateNum(totalSum - totalRatedSum);
        totalDailyRateResultSummaryBO.setVerySatisfiedNum(dataList.stream().mapToInt(r -> r.getDailyRateResultSummaryBO().getVerySatisfiedNum()).sum());
        totalDailyRateResultSummaryBO.setSatisfiedNum(dataList.stream().mapToInt(r -> r.getDailyRateResultSummaryBO().getSatisfiedNum()).sum());
        totalDailyRateResultSummaryBO.setNotSatisfiedNum(dataList.stream().mapToInt(r -> r.getDailyRateResultSummaryBO().getNotSatisfiedNum()).sum());
        totalDailyRateResultSummaryBO.setInvalidNum(dataList.stream().mapToInt(r -> r.getDailyRateResultSummaryBO().getInvalidNum()).sum());

        // 有效评价小计
        totalDailyRateResultSummaryBO.setValidNum(totalDailyRateResultSummaryBO.getVerySatisfiedNum()
                + totalDailyRateResultSummaryBO.getSatisfiedNum()
                + totalDailyRateResultSummaryBO.getNotSatisfiedNum());

        // 非常满意率
        if (totalDailyRateResultSummaryBO.getValidNum() == 0) {
            totalDailyRateResultSummaryBO.setVerySatisfiedPercentage(0d);
        } else {
            totalDailyRateResultSummaryBO.setVerySatisfiedPercentage((totalDailyRateResultSummaryBO.getVerySatisfiedNum() + 0.0) / totalDailyRateResultSummaryBO.getValidNum());
        }

        // 主动评价（户数）
        totalDailyRateSummaryBO.setRatedNum(totalDailyRateResultSummaryBO.getValidNum() + totalDailyRateResultSummaryBO.getInvalidNum());

        totalDailyRateSummaryBO.setDailyRateResultSummaryBO(totalDailyRateResultSummaryBO);

        return totalDailyRateSummaryBO;
    }

    /**
     * 需要回访的列表 不存在的来源需要加一个占位
     *
     * @param dailyRateBOList
     * @return
     */
    private Map<DailyRateSourceEnum, List<DailyRateBO>> getNeedFeedBackList(List<DailyRateBO> dailyRateBOList) {
        Map<DailyRateSourceEnum, List<DailyRateBO>> map = dailyRateBOList.stream()
                .filter(r -> r.getRateResult() != DailyRateResultEnum.BLANK && r.getRateResult() != DailyRateResultEnum.VERY_SATISFIED)
                .collect(Collectors.groupingBy(DailyRateBO::getSource));
        for (DailyRateSourceEnum sourceEnum : DailyRateSourceEnum.values()) {
            if (null != sourceEnum
                    && !map.containsKey(sourceEnum)
                    && sourceEnum != DailyRateSourceEnum.BLANK) {
                map.put(sourceEnum, Collections.emptyList());
            }
        }

        return map;
    }
}
