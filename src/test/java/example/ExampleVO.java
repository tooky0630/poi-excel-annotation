package example;

import com.penghaohuan.excel.annotation.ExcelDesc;

import java.util.Date;

/**
 * A Simple Example For Importer Usage.
 */
@ExcelDesc(name = "人员信息表", function = "validate", clazz = ExampleRowValidator.class)
public class ExampleVO {

    @ExcelDesc(name = "编号", function = "validate", clazz = ExampleValidator.class, keyAttr = true)
    private String no;

    @ExcelDesc(name = "年龄", isCheckNull = true)
    private Integer age;

    @ExcelDesc(name = "出生年月日", dateFormat = "yyyy/M/d")
    private Date birth;

    @ExcelDesc(name = "手机号码", regularExpression = "^[0-9]{11}", regularExpressionTip = "手机号码格式不正确")
    private String phone;

    public String getNo() {
        return no;
    }

    public void setNo(String no) {
        this.no = no;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getBirth() {
        return birth;
    }

    public void setBirth(Date birth) {
        this.birth = birth;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }
}
