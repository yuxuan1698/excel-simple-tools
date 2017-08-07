package com.xusanduo.dtos;

import com.xusanduo.annotations.ExcelColumn;
import com.xusanduo.annotations.ExcelSheet;

import java.io.Serializable;

/**
 * Created by xusanduo on 2017/7/26.
 */
@ExcelSheet("goods")
public class GoodsDTO implements Serializable {

    //Excel数据读取列开始序号
    public static final int columnStartNum = 0;
    //Excel数据读取标题行开始序号
    public static final int rowTitleStartNum = 0;
    //Excel数据读取数据行开始序号
    public static final int rowDataStartNum = 1;


    @ExcelColumn(columnIndex = 0, titleName = "名称")
    private String goodsName;
    @ExcelColumn(columnIndex = 1, titleName = "规格")
    private String standard;
    @ExcelColumn(columnIndex = 2, titleName = "单位")
    private String unitName;
    @ExcelColumn(columnIndex = 3, titleName = "分类")
    private String categoryName;
    @ExcelColumn(columnIndex = 4, titleName = "状态")
    private Boolean isValid;


    public String getGoodsName() {
        return goodsName;
    }

    public void setGoodsName(String goodsName) {
        this.goodsName = goodsName;
    }

    public String getStandard() {
        return standard;
    }

    public void setStandard(String standard) {
        this.standard = standard;
    }

    public String getUnitName() {
        return unitName;
    }

    public void setUnitName(String unitName) {
        this.unitName = unitName;
    }

    public String getCategoryName() {
        return categoryName;
    }

    public void setCategoryName(String categoryName) {
        this.categoryName = categoryName;
    }

    public Boolean getIsValid() {
        return isValid;
    }

    public void setIsValid(Boolean valid) {
        isValid = valid;
    }


}
