package com.xusanduo;

import com.alibaba.fastjson.JSON;
import com.google.common.collect.Lists;
import com.xusanduo.ExcelReadUtils;
import com.xusanduo.dtos.GoodsDTO;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Created by zengyh on 2017/8/7.
 */
public class ExcelToolDemo {

    //写入数据
    public static void main1(String[] args) throws IOException, NoSuchFieldException, InstantiationException, IllegalAccessException {

        System.out.println(new Boolean(true).toString());
        ExcelWriteUtils excelWriteUtils = new ExcelWriteUtils();
        List<GoodsDTO> goodsDTOs = Lists.newArrayList();

        GoodsDTO goodsDTO = new GoodsDTO();
        goodsDTO.setGoodsName("西瓜");
        goodsDTO.setStandard("10斤/个");
        goodsDTO.setUnitName("个");
        goodsDTO.setCategoryName("水果");
        goodsDTO.setIsValid(Boolean.FALSE);
        goodsDTOs.add(goodsDTO);

        goodsDTO = new GoodsDTO();
        goodsDTO.setGoodsName("冬瓜");
        goodsDTO.setStandard("20斤/个");
        goodsDTO.setUnitName("斤");
        goodsDTO.setCategoryName("蔬菜");
        goodsDTO.setIsValid(Boolean.FALSE);
        goodsDTOs.add(goodsDTO);

        //写入Excel并转化为字节
        ByteArrayOutputStream os = excelWriteUtils.writeBeanToExcelData(GoodsDTO.class,goodsDTOs,"2007");
        //输出文件
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\goods_export.xlsx");
        fileOutputStream.write(os.toByteArray());
    }


    //读取数据
    public static void main(String[] args) {

        File file = new File("D:\\goods-import.xlsx");
        System.out.println(file.getPath());
        ExcelReadUtils excelUtils = new ExcelReadUtils();

        //初始化表单
        Workbook workbook = null;
        try {
            workbook = excelUtils.initExcelWorkBook(file);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("文件格式不正确");
        }


        //读取标题
        String[] titles = new String[0];
        try {
            titles = excelUtils.readExcelTitle(GoodsDTO.class,workbook);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("数据格式不正确");
        }
        System.out.println("title:"+ JSON.toJSONString(titles));


        //读取数据
        List<GoodsDTO> goodsDTOs = null;
        try {
            goodsDTOs = excelUtils.readExcelDataToBean(GoodsDTO.class,workbook);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("数据格式不正确");
        }
        System.out.println("Goods Data:");
        goodsDTOs.stream().forEach(goodsExcelDO -> {
            System.out.println(JSON.toJSONString(goodsExcelDO));
        });

    }

}
