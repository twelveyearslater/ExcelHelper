package top.masochism.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import top.masochism.bo.Excel;

import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    public static void VLOOKUP(Excel excel, String sheetName, char ch1, int start1, int end1,char ch2, char ch3, int start2, int end2, char ch4) {
        Sheet sheet = getSheet(excel, sheetName);
        Map<Object, Object> findMap = new HashMap<>();
        Row row;
        Cell cell;
        for(int i = start2 - 1; i < end2; i++) {
            findMap.put(getValue(sheet, i, ch3 - 'A'), getValue(sheet, i, ch4 - 'A'));
        }
        for(int j = start1 - 1; j < end1; j++) {
            row = sheet.getRow(j);
            cell = row.createCell(ch2 - 'A');
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
        }else if("FORMULA".equals(type)){
            return cell.getArrayFormulaRange();
        }else if("BOOLEAN".equals(type)){
            return cell.getBooleanCellValue();
        }else if("ERROR".equals(type)){
            return cell.getErrorCellValue();
        }else{
            return "";
        }
    }
}
