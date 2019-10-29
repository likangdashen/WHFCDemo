package com.whfc.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author: ；kang
 * @Description:
 * @Date:Create：in 2019/10/25 9:14
 * @Version：1.0
 */
public class ExcelPOI {

    /*注意：
    * （1） HSSF低版本 xlx
    * （2） XSSF高版本 xlxs
    * */

    //创建文件并写入数据
    public static void excelWirte(){
        String[] title={"id","name","age"};

        //创建EXCLE工作薄
        HSSFWorkbook workbook = new  HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workbook.createSheet();
        //创建第一行标题行
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = null;
        //插入第一行数据：
        for(int i = 0 ; i <title.length;i++ ){
            //创建列
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }
        //追加数据，创建10行数据
        for(int i = 1 ; i <= 10 ;i++ ){
            HSSFRow rowValue = sheet.createRow(i);
            HSSFCell cell1 = null;
            for(int j = 0 ; j <= 2 ;j++ ){
                cell1 = rowValue.createCell(j);
                cell1.setCellValue( i + "*" + j + "=" +i*j);
            }
        }
        //创建文件写入数据
        File excFile = new File("E:\\test.xlx");
        try {
            excFile.createNewFile();
            FileOutputStream outStream = new FileOutputStream(excFile);
            //将工作薄内容写出到文件,注意是统一输出的
            workbook.write(outStream);
            //写入完成需要关闭流
            outStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    //读取解析EXCEL文件
    public static void excelRead(){
        File readFile = new File("E:\\test.xlx");
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(readFile);
            //
           HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
            //获取sheet
            //workbook.getSheetAt(index);读取索引index上的工作表
            HSSFSheet sheet0 = workbook.getSheet("Sheet0");
            //起始行，与最后一行
            int firstRow = 0;
            int lastRow = sheet0.getLastRowNum();
            for(int i = 0;i <= lastRow; i++ ){
                //获取数据行
                HSSFRow row = sheet0.getRow(i);
                //获取最后单元格列号
                short lastCellNum = row.getLastCellNum();
                for(int j = 0 ; j < lastCellNum; j++){
                    HSSFCell cell = row.getCell(j);
                    String stringCellValue = cell.getStringCellValue();
                    System.out.print(stringCellValue + "   ");
                };
                System.out.println();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
       //excelWirte();
        excelRead();
    }

}
