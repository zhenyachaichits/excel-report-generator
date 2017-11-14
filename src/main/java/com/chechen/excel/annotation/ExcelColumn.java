package com.chechen.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE_USE})
public @interface ExcelColumn {
    String name() default "";

    String[] subColumn() default "";

    String messageSuccess() default "";

    String messageError() default "";

    String emptyCellMessage() default "Empty";

    String optionalCellMessage() default "";
}
