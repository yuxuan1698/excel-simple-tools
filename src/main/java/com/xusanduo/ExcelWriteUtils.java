package com.xusanduo;

import com.xusanduo.annotations.ExcelColumn;
import com.xusanduo.annotations.ExcelSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * Created by zengyh on 2017/7/31.
 */
public class ExcelWriteUtils {

    //Excel数据读取列开始序号
    public Field columnStartNumField;
    //Excel数据读取标题行开始序号
    public Field rowTitleStartNumField;
    //Excel数据读取数据行开始序号
    public Field rowDataStartNumField;

    //Excel类型
    public enum excelTypeEnum{
        excel2003("2003","application/vnd.ms-excel"),
        excel2007("2007","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        String code;
        String value;
        excelTypeEnum(String code,String value){
            this.code = code;
            this.value = value;
        }
    };

    //创建标题样式
    public CellStyle createTitleStyle(Workbook workbook){
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        Font font = workbook.createFont();
        font.setFontName("微软雅黑 Light");
        font.setFontHeightInPoints((short)11);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(font);
        /*style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);*/

        return style;
    }

    //创建数据样式
    public CellStyle createDataStyle(Workbook workbook){
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)10);
        style.setFont(font);

        return style;
    }

    //设置表单整体样式
    public <T> void initSheetStyle(Sheet sheet, Class<T> classz){
        // 设置表格默认列宽度为15个字节
        //sheet.setDefaultColumnWidth((short) 15);
        //设置每列宽度
        Field[] fields = classz.getDeclaredFields();
        for( Field field : fields ){
            ExcelColumn columnAnnotation = field.getAnnotation(ExcelColumn.class);
            if(columnAnnotation!=null){
               sheet.setColumnWidth(columnAnnotation.columnIndex() + 1,columnAnnotation.columnWidth() * 256);
            }
        }
    }

    public <T> void  wirteBeanToExcelTitle(Class<T> classz, Sheet sheet, CellStyle cellStyle) throws NoSuchFieldException, IllegalAccessException, InstantiationException {
        //获取常量字段
        columnStartNumField = classz.getField("columnStartNum");
        rowTitleStartNumField = classz.getField("rowTitleStartNum");

        //获取常量
        T object = classz.newInstance();
        int columnStartNum = (int) columnStartNumField.get(object);
        int rowTitleStartNum = (int) rowTitleStartNumField.get(object);

        //获取属性
        Field[] fields = classz.getDeclaredFields();

        //创建标题行
        Row title = sheet.createRow(rowTitleStartNum);
        for( Field field : fields ){
            ExcelColumn columnAnnotation = field.getAnnotation(ExcelColumn.class);
            if(columnAnnotation!=null){
                Cell cell = title.createCell(columnAnnotation.columnIndex() + columnStartNum);
                cell.setCellValue(columnAnnotation.titleName());
                cell.setCellStyle(cellStyle);
            }
        }
    }

    public Workbook initWorkBook(String excelType){
        Workbook workbook = null;
        if(excelTypeEnum.excel2003.code.equals(excelType)){
            workbook = new HSSFWorkbook();
        }
        if(excelTypeEnum.excel2007.code.equals(excelType)){
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    public <T> void  writeBeanToExcelData(Class<T> classz, List<T> datas, Sheet sheet, CellStyle cellStyle) throws NoSuchFieldException, IllegalAccessException, InstantiationException {

        //获取常量字段
        columnStartNumField = classz.getField("columnStartNum");
        rowDataStartNumField = classz.getField("rowDataStartNum");

        //获取常量
        T object = classz.newInstance();
        int columnStartNum = (int) columnStartNumField.get(object);
        int rowDataStartNum = (int) rowDataStartNumField.get(object);

        //获取属性
        Field[] fields = classz.getDeclaredFields();
        int rowIndex = rowDataStartNum;
        for(T data : datas) {
            Row dataRow = sheet.createRow(rowIndex);
            for( Field field : fields ){
                ExcelColumn columnAnnotation = field.getAnnotation(ExcelColumn.class);
                if(columnAnnotation!=null){
                    //获取字段值
                    field.setAccessible(Boolean.TRUE);
                    Object fieldValue = field.get(data);

                    //转换数据类型
                    Cell cell = dataRow.createCell(columnStartNum + columnAnnotation.columnIndex());
                    this.castFromFieldValue(cell, field.getType(), fieldValue);
                    cell.setCellStyle(cellStyle);
                }
            }
            rowIndex++;
        }

    }

    public void castFromFieldValue(Cell cell, Class<?> fieldType, Object fieldValue){

        if(fieldValue == null){
            return;
        }
        if(fieldType.getName().equals("java.lang.String")){
            cell.setCellValue(fieldValue.toString());
        }else if(fieldType.getName().equals("java.lang.Float")){
           cell.setCellValue(Float.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Short")){
           cell.setCellValue(Short.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Integer")){
            cell.setCellValue(Integer.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Long")){
            cell.setCellValue(Long.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Double")){
            cell.setCellValue(Double.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Boolean")){
            cell.setCellValue(Boolean.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.util.Date")) {
            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) fieldValue));
        }else if(fieldType.getName().equals("java.lang.Byte")){
            cell.setCellValue(Byte.valueOf(fieldValue.toString()));
        }else if(fieldType.getName().equals("java.lang.Byte") || fieldType.getName().equals("[B") || fieldType.getName().equals("[Ljava.lang.Byte")){
            cell.setCellValue(Byte.valueOf(fieldValue.toString()));
        }
    }


    public <T> ByteArrayOutputStream writeBeanToExcelData(Class<T> classz, List<T> datas, String excelType) throws IllegalAccessException, NoSuchFieldException, InstantiationException, IOException {

        //初始化excel表格
        Workbook workbook = this.initWorkBook(excelType);

        //读取表单名称
        ExcelSheet sheetAnnotation = classz.getAnnotation(ExcelSheet.class);
        String sheetName = sheetAnnotation.value();
        Sheet sheet = workbook.createSheet(sheetName);
        this.initSheetStyle(sheet, classz);

        //创建样式
        CellStyle cellTitleStyle = this.createTitleStyle(workbook);
        CellStyle cellDataStyle = this.createDataStyle(workbook);

        //写入标题
        this.wirteBeanToExcelTitle(classz, sheet, cellTitleStyle);

        //写入数据
        this.writeBeanToExcelData(classz, datas, sheet, cellDataStyle);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);

        return os;
    }
}
