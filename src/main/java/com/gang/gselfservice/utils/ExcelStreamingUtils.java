package com.gang.gselfservice.utils;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Excel 工具类 只支持xlsx
 */
public class ExcelStreamingUtils implements ExcelUtilsInterface {

    private Workbook workbook;
    private OutputStream os;
    private String pattern;// 日期格式

    private static final String EXT_XLSX = "xlsx";

    private static final String DEF_VERSION = "2007";

    public void setPattern(String pattern) {
        this.pattern = pattern;
    }

    public ExcelStreamingUtils(Workbook workboook) {
        this.workbook = workboook;
    }

    public ExcelStreamingUtils(InputStream is) {
        workbook = StreamingReader.builder().bufferSize(8192).rowCacheSize(100).open(is);
    }

    public String toString() {
        return "共有 " + getSheetCount() + "个sheet 页！";
    }

    public String toString(int sheetIx) throws IOException {
        return "第 " + (sheetIx + 1) + "个sheet 页，名称： " + getSheetName(sheetIx) + "，共 " + getRowCount(sheetIx) + "行！";
    }

    /**
     * 读取 Excel 第一页所有数据
     *
     * @return
     * @throws Exception
     */
    public List<List<String>> read() throws Exception {
        return read(0, 0, getRowCount(0) - 1);
    }

    /**
     * 读取指定sheet 页所有数据
     *
     * @param sheetIx 指定 sheet 页，从 0 开始
     * @return
     * @throws Exception
     */
    public List<List<String>> read(int sheetIx) throws Exception {
        return read(sheetIx, 0, getRowCount(sheetIx) - 1);
    }

    @Override
    public boolean write(List<List<String>> rowData, String sheetName, boolean isNewSheet) {
        return false;
    }

    @Override
    public boolean write(int sheetIx, List<List<String>> rowData, boolean isAppend) {
        return false;
    }

    /**
     * 读取指定sheet 页指定行数据
     *
     * @param sheetIx 指定 sheet 页，从 0 开始
     * @param start   指定开始行，从 0 开始
     * @param end     指定结束行，从 0 开始
     * @return
     * @throws Exception
     */
    public List<List<String>> read(int sheetIx, int start, int end) throws Exception {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        List<List<String>> list = new ArrayList<List<String>>();

        if (end > getRowCount(sheetIx)) {
            end = getRowCount(sheetIx);
        }

        int rowCount = 0;
        int cols = 0;
        for (Row row : sheet) {
            if (0 == rowCount) {
                cols = row.getLastCellNum(); // 第一行总列数
                if (0 == cols) {
                    // 没数据
                    return list;
                }
            }
            // 没到起始位置
            if (rowCount < start) {
                continue;
            }
            // 读到结束位置了，返回
            if (rowCount > end) {
                return list;
            }

            List<String> rowList = new ArrayList<>();
            for (int j = 0; j < cols; j++) {
                if (row == null) {
                    rowList.add(null);
                    continue;
                }
                rowList.add(getCellValueToString(row.getCell(j)));
            }
            list.add(rowList);

            rowCount++;
        }

        return list;
    }

    /**
     * 指定行是否为空
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定开始行，从 0 开始
     * @return true 不为空，false 不行为空
     * @throws IOException
     */
    public boolean isRowNull(int sheetIx, int rowIndex) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        return sheet.getRow(rowIndex) == null;
    }

    /**
     * 指定单元格是否为空
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定开始行，从 0 开始
     * @param colIndex 指定开始列，从 0 开始
     * @return true 行不为空，false 行为空
     * @throws IOException
     */
    public boolean isCellNull(int sheetIx, int rowIndex, int colIndex) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        if (!isRowNull(sheetIx, rowIndex)) {
            return false;
        }
        Row row = sheet.getRow(rowIndex);
        return row.getCell(colIndex) == null;
    }

    /**
     * 返回sheet 中的行数
     *
     * @param sheetIx 指定 Sheet 页，从 0 开始
     * @return
     */
    public int getRowCount(int sheetIx) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        return sheet.getLastRowNum() + 1;
    }

    /**
     * 返回所在行的列数
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定行，从0开始
     * @return 返回-1 表示所在行为空
     */
    public int getColumnCount(int sheetIx, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);
        return row == null ? -1 : row.getLastCellNum();
    }

    /**
     * 返回 row 和 column 位置的单元格值
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定行，从0开始
     * @param colIndex 指定列，从0开始
     * @return
     */
    public String getValueAt(int sheetIx, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        return getCellValueToString(sheet.getRow(rowIndex).getCell(colIndex));
    }

    /**
     * 返回指定行的值的集合
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定行，从0开始
     * @return
     */
    public List<String> getRowValue(int sheetIx, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);
        List<String> list = new ArrayList<String>();
        if (row == null) {
            list.add(null);
        } else {
            for (int i = 0; i < row.getLastCellNum(); i++) {
                list.add(getCellValueToString(row.getCell(i)));
            }
        }
        return list;
    }

    /**
     * 返回列的值的集合
     *
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param rowIndex 指定行，从0开始
     * @param colIndex 指定列，从0开始
     * @return
     */
    public List<String> getColumnValue(int sheetIx, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        List<String> list = new ArrayList<String>();
        for (int i = rowIndex; i < getRowCount(sheetIx); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                list.add(null);
                continue;
            }
            list.add(getCellValueToString(sheet.getRow(i).getCell(colIndex)));
        }
        return list;
    }

    /**
     * 获取excel 中sheet 总页数
     *
     * @return
     */
    public int getSheetCount() {
        return workbook.getNumberOfSheets();
    }

    /**
     * 获取 sheet名称
     *
     * @param sheetIx 指定 Sheet 页，从 0 开始
     * @return
     * @throws IOException
     */
    public String getSheetName(int sheetIx) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        return sheet.getSheetName();
    }

    /**
     * 获取sheet的索引，从0开始
     *
     * @param name sheet 名称
     * @return -1表示该未找到名称对应的sheet
     */
    public int getSheetIndex(String name) {
        return workbook.getSheetIndex(name);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 关闭流
     *
     * @throws IOException
     */
    public void close() throws IOException {
        if (os != null) {
            os.close();
        }
        workbook.close();
    }

    /**
     * 转换单元格的类型为String 默认的 <br>
     * 默认的数据类型：CELL_TYPE_BLANK(3), CELL_TYPE_BOOLEAN(4),
     * CELL_TYPE_ERROR(5),CELL_TYPE_FORMULA(2), CELL_TYPE_NUMERIC(0),
     * CELL_TYPE_STRING(1)
     *
     * @param cell
     * @return
     */
    private String getCellValueToString(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    if (pattern != null) {
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        strCell = sdf.format(date);
                    } else {
                        strCell = date.toString();
                    }
                    break;
                }
                // 不是日期格式，则防止当数字过长时以科学计数法显示
//                cell.setCellType(CellType.STRING);
                strCell = cell.toString();
                break;
            case STRING:
                strCell = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return strCell;
    }

}