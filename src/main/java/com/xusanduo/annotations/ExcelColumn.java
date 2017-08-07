package com.xusanduo.annotations;

import java.lang.annotation.*;

/**
 * Created by zengyh on 2017/7/26.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelColumn {

    //Excel表单列号
    int columnIndex();

    //Excel表单标题
    String titleName();

    //列宽
    int columnWidth() default 15;
}
