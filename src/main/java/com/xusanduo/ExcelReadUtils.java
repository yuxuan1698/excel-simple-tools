package com.xusanduo;

import com.xusanduo.annotations.ExcelColumn;
import com.xusanduo.annotations.ExcelSheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Stream;

/**
 * POI API 读取Excel文件内容
 * Created by zengyh on 2017/7/26.
 */
public class ExcelReadUtils {

    //Excel数据读取列开始序号
    public Field columnStartNumField;
    //Excel数据读取标题行开始序号
    public Field rowTitleStartNumField;
    //Excel数据读取数据行开始序号
    public Field rowDataStartNumField;

    public static final String excel2003 = "application/vnd.ms-excel";
    public static final String excel2007 = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /**
     * 初始化Excel文件
     * @param fileName
     * @param inputStream
     * @return
     * @throws Exception
     */
    public Workbook initExcelWorkBook(String fileName, String contentType,InputStream inputStream) throws Exception {

        Workbook workbook = null;
        if(fileName.endsWith(".xlsx") || contentType.equals(excel2007)) {
            workbook = new XSSFWorkbook(inputStream);
        }else if(fileName.endsWith(".xls") || contentType.equals(excel2003)){
            workbook = new HSSFWorkbook(inputStream);
        }else{
            throw new Exception("文件格式不正确");
        }

        return workbook;
    }

    /**
     * 初始化Excel文件
     * @param file
     * @return
     * @throws Exception
     */
    public Workbook initExcelWorkBook(File file) throws Exception {

        if( !file.exists() ){
            throw new Exception("文件不存在");
        }
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = null;
        if(file.getName().endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(fileInputStream);
        }else if(file.getName().endsWith(".xls")){
            workbook = new HSSFWorkbook(fileInputStream);
        }else{
            throw new Exception("文件格式不正确");
        }

        return workbook;
    }

    /**
     * 读取Excel表头
     * @param classz
     * @param workbook
     * @param <T>
     * @return
     * @throws Exception
     */
    public <T> String[] readExcelTitle(Class<T> classz,Workbook workbook) throws Exception {

        //读取表单
        ExcelSheet sheetAnnotation = classz.getAnnotation(ExcelSheet.class);
        String sheetName = sheetAnnotation.value();
        Sheet sheet = workbook.getSheet(sheetName);
        if( sheet == null ) {
            sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                throw new Exception("文件格式不正确，未找到表单:" + sheetName);
            }
        }

        //获取常量字段
        columnStartNumField = classz.getField("columnStartNum");
        rowTitleStartNumField = classz.getField("rowTitleStartNum");

        //获取常量
        T object = classz.newInstance();
        int columnStartNum = (int) columnStartNumField.get(object);
        int rowTitleStartNum = (int) rowTitleStartNumField.get(object);

        //获取属性
        Field[] fields = classz.getDeclaredFields();
        //获取表单标题
        long titleCount = Stream.of(fields).filter(f->f.getAnnotation(ExcelColumn.class)!=null).count();
        String[] titles = new String[(int) titleCount];

        //获取标题行
        Row titleRow = sheet.getRow(rowTitleStartNum);
        for( Field field : fields ){
            ExcelColumn columnAnnotation = field.getAnnotation(ExcelColumn.class);
            if(columnAnnotation!=null){
                Cell cell = titleRow.getCell(columnStartNum + columnAnnotation.columnIndex());
                Object value = this.getCellFormatValue(cell);
                if( value == null) throw new Exception("列标题："+columnAnnotation.titleName()+" 为空");
                titles[columnAnnotation.columnIndex()] = value.toString();
            }
        }
        return titles;
    }


    /**
     * 读取表单数据
     * @param classz
     * @param workbook
     * @param <T>
     * @return
     * @throws Exception
     */
    public <T> List<T> readExcelDataToBean(Class<T> classz, Workbook workbook) throws Exception {

        //读取表单
        ExcelSheet sheetAnnotation = classz.getAnnotation(ExcelSheet.class);
        String sheetName = sheetAnnotation.value();
        Sheet sheet = workbook.getSheet(sheetName);

        if( sheet == null ){
            sheet = workbook.getSheetAt(0);
            if ( sheet == null ){
                throw new Exception("文件格式不正确，未找到表单:"+sheetName);
            }
        }

        //获取常量字段
        columnStartNumField = classz.getField("columnStartNum");
        rowDataStartNumField = classz.getField("rowDataStartNum");


        //获取常量
        T object = classz.newInstance();
        int columnStartNum = (int) columnStartNumField.get(object);
        int rowDataStartNum = (int) rowDataStartNumField.get(object);

        //获取属性
        Field[] fields = classz.getDeclaredFields();
        //判断数据行格式
        int lastRowNum = sheet.getLastRowNum();
        if(rowDataStartNum > lastRowNum){
            throw new Exception("数据格式不正确");
        }

        //读取数据
        List<T> results = new ArrayList<T>();
        for( int i = rowDataStartNum; i <= lastRowNum; i++ ){
            T instance = classz.newInstance();
            Row dataRow = sheet.getRow(i);
            if (dataRow == null){
                continue;
            }
            //是否为空行
            Boolean isNullRow = Boolean.TRUE;
            for( Field field : fields ){
                ExcelColumn columnAnnotation = field.getAnnotation(ExcelColumn.class);
                if(columnAnnotation!=null){
                    Cell cell = dataRow.getCell(columnStartNum + columnAnnotation.columnIndex());
                    //单元格为空
                    if( cell == null || this.getCellFormatValue(cell) == null || StringUtils.isBlank(this.getCellFormatValue(cell).toString().trim())) {
                        continue;
                    } else {
                        isNullRow = Boolean.FALSE;
                    }

                    //获取单元格数据类型
                    Class<?> valueType = this.getCellFormatType(cell);
                    if( valueType == null ) {
                        continue;
                    }
                    //读取单元格数据
                    Object value = this.getCellFormatValue(cell);
                    //显式转换数据类型
                    Object fieldValue = null;
                    fieldValue = this.castToFieldValue(field.getType(),valueType,value);
                    //属性字段赋值
                    field.setAccessible(Boolean.TRUE);
                    field.set(instance,fieldValue);
                }
            }
            if ( !isNullRow ) {
                results.add(instance);
            }
        }
        return results;
    }

    /**
     * 转换Excel值类型
     * @param fieldType
     * @param value
     * @return
     */
    public Object castToFieldValue(Class<?> fieldType, Class<?> valueType, Object value){
        if(value==null){
            return null;
        }

        if(fieldType.getName().equals("java.lang.String")){
            if(valueType.getName().equals("double")){
                Double tempValue = Double.valueOf(value.toString());
                if(((int)(tempValue*100))%100>0){
                    return value.toString();
                }else{
                    return String.valueOf(tempValue.intValue());
                }
            }else {
                return value.toString();
            }
        }
        if(fieldType.getName().equals("java.lang.Float")){
            return Double.valueOf(value.toString()).floatValue();
        }
        if(fieldType.getName().equals("java.lang.Short")){
            return Double.valueOf(value.toString()).shortValue();
        }
        if(fieldType.getName().equals("java.lang.Integer")){
            return Double.valueOf(value.toString()).intValue();
        }
        if(fieldType.getName().equals("java.lang.Long")){
            return Double.valueOf(value.toString()).longValue();
        }
        if(fieldType.getName().equals("java.lang.Double")){
            return Double.valueOf(value.toString());
        }
        if(fieldType.getName().equals("java.lang.Boolean")){
            return Boolean.valueOf(value.toString());
        }
        if(fieldType.getName().equals("java.util.Date")) {
            return DateUtil.getJavaDate((double)value);
        }
        if(fieldType.getName().equals("java.lang.Byte")){
            return value.toString().getBytes()[0];
        }
        if(fieldType.getName().equals("java.lang.Byte") || fieldType.getName().equals("[B") || fieldType.getName().equals("[Ljava.lang.Byte")){
            return value.toString().getBytes();
        }

        return null;
    }


    /**
     * 获取Excel单元格
     * @param cell
     * @param <T>
     * @return
     */
    public <T> T getCelllValue(Cell cell){
        Object object = getCellFormatValue(cell);
        if(object==null){
            return null;
        }
        return (T)object;
    }

    /**
     * 获取Excel单元格数据
     * @param cell
     * @return
     */
    public Object getCellFormatValue(Cell cell){

        if(cell==null){
            return null;
        }
        int type = cell.getCellType();
        switch (type){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue();
            case Cell.CELL_TYPE_FORMULA:
                if(DateUtil.isCellDateFormatted(cell)){
                    return cell.getDateCellValue();
                }else {
                    return cell.getNumericCellValue();
                }
            default:
                return null;
        }
    }

    /**
     * 获取Excel单元格数据类型
     * @param cell
     * @return
     */
    public Class<?> getCellFormatType(Cell cell){
        if(cell==null){
            return null;
        }
        int type = cell.getCellType();
        switch (type){
            case Cell.CELL_TYPE_BLANK:
                return String.class;
            case Cell.CELL_TYPE_ERROR:
                return byte[].class;
            case Cell.CELL_TYPE_BOOLEAN:
                return boolean.class;
            case Cell.CELL_TYPE_NUMERIC:
                return double.class;
            case Cell.CELL_TYPE_STRING:
                return String.class;
            case Cell.CELL_TYPE_FORMULA:
                if(DateUtil.isCellDateFormatted(cell)){
                    return Date.class;
                }else {
                    return double.class;
                }
            default:
                return null;
        }
    }

}
