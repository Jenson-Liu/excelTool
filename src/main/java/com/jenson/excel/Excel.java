package com.jenson.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import com.jenson.excel.exception.SheetException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * row是第几整行
 * rank是第几整列
 * cell是当前列数
 */
/**
 * @author : Jenson.Liu
 * @date : 2019/11/28  11:02 上午
 */
public final class Excel {

    /**
     *  上传的文件
     */
    File excelFile;
    /**
     *  当前需要查询的多个cells
     */
    ArrayList<Integer> cells = new ArrayList<>();
    /**
     *  指定当前sheet
     */
    Object currentSheet;
    /**
     *  当前查询指定列
     */
    int IntegerCurrentCell = -1;

    /**
     * XSSFWorkbook对象
     */
    XSSFWorkbook xssfWorkbook;
    /**
     * sheet对象
     */
    XSSFSheet xssfSheet;
    /**
     * row对象，excel的行（整行包装成row）
     */

    public Excel(File excelFile) {
        this.excelFile = excelFile;
        try {
            xssfWorkbook = new XSSFWorkbook(excelFile);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public Excel(File excelFile, Object currentSheet) {
        this.excelFile = excelFile;
        this.currentSheet = currentSheet;
    }

    /**
     * 设置当前的sheet
     */
    public Excel setSheet(Object sheetName) throws SheetException {
        /**
         * 判断参数类型
         */
        if (sheetName instanceof String){
            xssfSheet = xssfWorkbook.getSheet(sheetName.toString());
            currentSheet = sheetName;
        }else if(sheetName instanceof Integer){
            Iterator<Sheet> iter =  xssfWorkbook.sheetIterator();
            int num = Integer.parseInt(sheetName.toString());
            int i = 0;
            while (iter.hasNext()){
                if (i == num){
                    xssfSheet = (XSSFSheet) iter.next();
                }else {
                    i++;
                }
            }
        }else {
            throw new SheetException("the param of currentSheet is wrong");
        }
        return this;
    }

    /**
     * 需要获取多行
     * 需要调用此方法，设置需要那些cell
     * 数据类型为String
     * @param cellName
     * @return
     */
    public Excel addCell(String cellName){
        Row row = xssfSheet.getRow(0);
        int num = 0;
        for (Cell cell:row){
            if(cellName.equals(cell.getStringCellValue())){
                this.cells.add(num);
            }else {
                num++;
            }
        }
        return this;
    }

    /**
     * 需要获取多行
     * 需要调用此方法，设置需要那些cell
     * 数据类型为int
     * @param cell
     * @return
     */
    public Excel addCell(int cell){
        this.cells.add(cell);
        return this;
    }

    /**
     * 设置获取第几列
     */
    public Excel setCell(String param) throws SheetException {
        if (xssfSheet == null){
            throw new SheetException("the param of Sheet is not set");
        }
        IntegerCurrentCell = 0;
        Row row = xssfSheet.getRow(0);
        for (Cell cell:row){
            if (cell.getStringCellValue().equals(param)){
                return this;
            }else {
                IntegerCurrentCell++;
            }
        }
        return this;
    }

    /**
     * 传入参数为param
     * @param param
     * @return
     * @throws SheetException
     */
    public Excel setCell(int param) throws SheetException {
        IntegerCurrentCell = param;
        return this;
    }

        /**
         * 获取所有的sheet的名称
         * @return
         */
    public ArrayList<String> getAllSheetNames(){
        ArrayList<String> list = new ArrayList<String>();
        XSSFWorkbook wb = null;
        try {
            wb = new XSSFWorkbook(new FileInputStream(excelFile));
            Iterator<Sheet> iter =  wb.sheetIterator();
            while (iter.hasNext()){
                String sheetName = iter.next().getSheetName();
                list.add(sheetName);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 获取指定sheet下指定的指定列的内容
     * 单列
     * @return
     */
    public ArrayList<String> getList() throws SheetException {
        ArrayList<String> list = new ArrayList<>();
        if(IntegerCurrentCell != -1){
            for (Row row:xssfSheet){
                list.add(row.getCell(IntegerCurrentCell).getStringCellValue());
            }
        }else {
            throw new SheetException("the current Cell is not set");
        }
        return list;
    }

    /**
     *
     * @return
     */
    public LinkedHashMap<Integer,ArrayList<String>> getMultipleList() throws SheetException {
        LinkedHashMap<Integer,ArrayList<String>> listLinkedHashMap = new LinkedHashMap<>();
        if(cells.size() >= 0){
            for (int i:cells) {
                listLinkedHashMap.put(i,new ArrayList<>());
            }
            Collections.sort(cells);
                for (Row row:xssfSheet){
                    for (int i:cells) {
                        listLinkedHashMap.get(i).add(row.getCell(i).getStringCellValue());
                    }
                }
        }else {
            throw new SheetException("the current Cell is not set");
        }
        return listLinkedHashMap;
    }



    /**
     * 获取标题头
     * 如果第一行有信息
     */
    public ArrayList<String> getAllRowsOfFirstLine() throws SheetException {
        if (xssfSheet == null){
            throw new SheetException("the param of Sheet is not set");
        }
        ArrayList<String> list = new ArrayList<>();
        Row row = xssfSheet.getRow(0);
        for (Cell cell:row){
            list.add(cell.getStringCellValue());
        }
        return list;
    }
}
