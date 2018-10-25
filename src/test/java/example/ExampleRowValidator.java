package example;

import com.penghaohuan.excel.handler.ExcelImporter;
import org.apache.commons.lang3.StringUtils;

/**
 * A simple row example validator.
 * used to validate the row what is from excel.
 *
 * @author penghaohuan
 */
public class ExampleRowValidator {

    /**
     *
     * @param value 单元格值
     * @param exceptionMsg 异常信息，包含校验值所在的行、列信息
     * @return 返回原始值+ %c后缀为校验通过，否则为返回的校验异常信息
     */
    public String validate(ExampleVO value, String exceptionMsg) {
        if (value == null || StringUtils.isBlank(value.getNo()) || value.getBirth() == null) {
            return exceptionMsg + "数据缺失";
        }
        final String substring = value.getNo().substring(1);
        return substring.startsWith(String.valueOf(value.getBirth().getYear() + 1900)) ?
                ExcelImporter.CORRECT_SYMBOL : exceptionMsg + "格式错误，编号与出生年份不匹配";
    }
}
