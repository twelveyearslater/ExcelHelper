package top.masochism.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import top.masochism.bo.Excel;

import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    public static void VLOOKUP(Excel excel, char ch1, int start1, int end1,char ch2, char ch3, int start2, int end2, char ch4) {
        Sheet sheet = getSheet(excel, "test2");
        Map<Object, Object> findMap = new HashMap<>();
        Row row;
        Cell cell;
        for(int i = start2 - 1; i < end2; i++) {
            findMap.put(getValue(sheet, i, ch3 - 'A'), getValue(sheet, i, ch4 - 'A'));
        }
//        cell = sheet.getRow(start2).getCell(ch2 - 'A');
        for(int j = start1 - 1; j < end1; j++) {
            row = sheet.getRow(j);
            cell = row.getCell(ch1 - 'A');
            Object obj = getValue(sheet, j, ch1 - 'A');
            cell.setCellValue((Double)findMap.get(obj));
        }


    }

    public static Sheet getSheet(Excel excel, String name){
        return excel.workbook.getSheet(name);
    }

    public static Object getValue(Sheet sheet, int roNum, int columnNum) {
        Row row = sheet.getRow(roNum);
        Cell cell = row.getCell(columnNum);
        String type = "" + cell.getCellType();
        if("NUMERIC".equals(type)){
            return cell.getNumericCellValue();
        }else if("String".equals(type)) {
            return cell.getStringCellValue();
        }else{
            return null;
        }
    }
}
