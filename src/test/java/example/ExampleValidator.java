package example;

import com.penghaohuan.excel.handler.ExcelImporter;
import org.apache.commons.lang3.StringUtils;

/**
 * A simple validator.
 * used to validate the content what is from excel.
 *
 * @author penghaohuan
 */
public class ExampleValidator {

    /**
     *
     * @param value 单元格值
     * @param exceptionMsg 异常信息，包含校验值所在的行、列信息
     * @return 返回原始值+ %c后缀为校验通过，否则为返回的校验异常信息
     */
    public String validate(String value, String exceptionMsg) {
        if (StringUtils.isBlank(value)) {
            return ExcelImporter.CORRECT_SYMBOL;
        }
        return value.startsWith("D") ? value + ExcelImporter.CORRECT_SYMBOL : exceptionMsg + "格式错误，未以字母D开头";
    }
}
