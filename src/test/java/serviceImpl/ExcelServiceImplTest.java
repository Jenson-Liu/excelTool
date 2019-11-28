package serviceImpl;

import com.jenson.excel.Excel;
import com.jenson.excel.exception.SheetException;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Set;

import static org.junit.Assert.*;

/**
 * @author : Jenson.Liu
 * @date : 2019/11/28  10:50 上午
 */
public class ExcelServiceImplTest {

    @Test
    public void getAllSheet() throws SheetException {
        Excel excel = new Excel(new File("/Users/i501695/IdeaProjects/excelTool/src/main/resources/excel/EXPORT.XLSX"));
        ArrayList<String> list = excel.setSheet("Sheet1").setCell(0).getList();
        for (String string:list){
            System.out.println(string);
        }
    }

    @Test
    public void getListBySheet() throws SheetException {
        Excel excel = new Excel(new File("/Users/i501695/IdeaProjects/excelTool/src/main/resources/excel/EXPORT.XLSX"));
        LinkedHashMap<Integer,ArrayList<String>> listLinkedHashMap = excel.setSheet("Sheet1").addCell(0)
                .addCell(1).getMultipleList();
        Set<Integer> set = listLinkedHashMap.keySet();
        for (int key:set){
            for (String string:listLinkedHashMap.get(key)){
                System.out.print(string);
            }
            System.out.println();
        }

    }
}