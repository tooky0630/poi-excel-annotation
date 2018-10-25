package com.penghaohuan.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导入导出属性配置.
 * @author penghaohuan
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE, ElementType.FIELD})
public @interface ExcelDesc {

    /**
     * Excel表头列名.
     */
    String name();

    /**
     * 是否是关键字段，如果是，则该字段必须有值
     */
    boolean keyAttr() default false;

    /**
     * 如果字段实际是日期时间字段类型，定义日期时间的格式，如：yyyy/MM/dd
     */
    String dateFormat() default "";

    /**
     * 是否进行非空验证.
     * @return 是否进行非空验证
     */
    boolean isCheckNull() default false;

    /**
     * 正则表达式验证.
     * @return 是否进行正则匹配
     */
    String regularExpression() default "";

    /**
     * 正则表达式验证错误提示信息.
     * @return 验证错误提示信息.
     */
    String regularExpressionTip() default "";

    /**
     * 校验类.
     * 在使用function注解时，可以指定校验类，可以在类和属性上注解
     * 优先级：属性注解校验类>类注解校验类>默认导出实体
     * @return 校验类
     */
    Class clazz() default NoValidateClass.class;

    /**
     * 方法校验,默认到相应导出实体找相应的方法.
     * 如需使用其他类中的方法请配合clazz注解使用，可以在类和属性上注解
     * 优先级：属性注解校验类>类注解校验类>默认导出实体
     *  校验方法示例.
     *  value 待校验值
     *  exceptionMsg 固定异常信息内容
     *  校验通过时value追加%c，失败是返回校验信息
     *  public String demo(final String value, final String exceptionMsg) {
     * //编写校验逻辑，检验成功的需要在value追加%c
     *  if("hello".equals(value)){
     *     return value +" world2"+ ImportExcelUtil.CORRECT_SYMBOL;
     *  }else{
     *   return  exceptionMsg+"内容必须是hello";
     *  }
     * }
     * @return 是否进行方法校验
     */
    String function() default "";

    /**
     * 校验类默认值--未指定校验类标志.
     */
    class NoValidateClass {
    }
}