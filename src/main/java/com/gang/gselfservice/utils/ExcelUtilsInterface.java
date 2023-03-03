package com.gang.gselfservice.utils;

import java.io.IOException;
import java.util.List;

public interface ExcelUtilsInterface {

    List<List<String>> read() throws Exception;

    List<List<String>> read(int sheetIx) throws Exception;

    boolean write(List<List<String>> rowData, String sheetName, boolean isNewSheet) throws IOException;

    boolean write(int sheetIx, List<List<String>> rowData, boolean isAppend) throws IOException;

    void close() throws IOException;

    int getRowCount(int sheetIx);

    String getSheetName(int sheetIx) throws IOException;

    int getSheetCount();

    boolean isRowNull(int sheetIx, int rowIndex) throws IOException;
}
