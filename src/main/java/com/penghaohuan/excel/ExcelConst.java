package com.penghaohuan.excel;

/**
 * Excel常量类.
 * @author penghaohuan
 */
public final class ExcelConst {
    public static final int FONT_HEIGHT_TITLE = 14;
    public static final int FONT_HEIGHT_DATA = 11;
    public static final int COLUMN_WIDTH = 256;
    public static final String FONT_NAME = "宋体";
    public static final String MERGE_ADD_ROW = "addrow";
    public static final String MERGE_COLUMN_PAD = "$$$";
    public static final String MERGE_NEED = "NULL";
    public static final String FILE_NAME_MODEL_KEY = "fileName";
    public static final String MUTIPLES_MODEL_KEY = "multiples";
    public static final String SHEET_NAME_MODEL_KEY = "sheetName";
    public static final String TITLE_MODEL_KEY = "title";
    public static final String TITLE_CONTENT_MODEL_KEY = "titleContent";
    public static final String HEAD_MODEL_KEY = "head";
    public static final String SUB_HEAD_MODEL_KEY = "subhead";
    public static final String DATA_MODEL_KEY = "data";
    public static final String FIRST_ROW_MODEL_KEY = "firstRow";
    public static final String FIRST_COL_MODEL_KEY = "firstCol";
    public static final String MERGE_CELLS_MODEL_KEY = "mergeCells";
    public static final String MERGE_FOR_DATA_MODEL_KEY = "mergeForData";
    public static final String SETVALUE_OF_POINTS_MODEL_KEY = "setValueOfPoints";
    public static final String TITLE_OCCUPY_LENGTH_MODEL_KEY = "titleOccupyLength";
    public static final String EXCEL_SUFFIX_OLD = ".xls";
    public static final String EXCEL_SUFFIX_NEW = ".xlsx";
    public static final String STRING_CLASSPATH = "java.lang.String";
    public static final String INTEGER_CLASSPATH = "java.lang.Integer";
    public static final String LONG_CLASSPATH = "java.lang.Long";
    public static final String DATE_CLASSPATH = "java.util.Date";
    public static final String CORRECT_SYMBOL = "%c";
    public static final String PATTERN_DATETIME = "(^((?:19|20)\\d\\d)-([1-9]|0[1-9]|1[012])-([1-9]|0[1-9]|[12][0-9]|3[01])$)|(^((?:19|20)\\d\\d)/([1-9]|0[1-9]|1[012])/([1-9]|0[1-9]|[12][0-9]|3[01])$)";
    public static final String PATTERN_PHONE = "^[0-9]{11}";
    public static final String PATTERN_IDCARD = "(^\\d{15}$)|(^\\d{18}$)|(^\\d{17}(\\d|X|x)$)|(^[a-zA-Z]{5,17}$)|(^[a-zA-Z0-9]{5,17}$)";

    private ExcelConst() {
    }
}
