package com.penghaohuan.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出属性配置.
 * @author penghaohuan
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE, ElementType.FIELD})
public @interface ExportExcelDesc {
    String name();

    short color() default 0;

    boolean bold() default true;

    String fontName() default "";

    short fontHeightInPoints() default 14;

    int columnWidth() default 0;
}
