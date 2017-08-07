package com.xusanduo.annotations;

import java.lang.annotation.*;

/**
 * Created by zengyh on 2017/7/26.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {
    //表单名称
    String value();
}
