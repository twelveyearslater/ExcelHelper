package top.masochism.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import top.masochism.bo.Excel;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    public static void VLOOKUP(Excel excel, String sheetName, String start1, String end1, String start2, String end2,String target, String need) throws Exception {
        int[] start1Arr = convert(start1);
        int[] end1Arr = convert(end1);
        int[] start2Arr = convert(start2);
        int[] end2Arr = convert(end2);
        int targetColumn = convertColumn(target);
        int needColumn = convertColumn(need);

        VLOOKUP(excel, sheetName, start1Arr, end1Arr, start2Arr, end2Arr, targetColumn, needColumn);
    }

    public static void VLOOKUP(Excel excel, String sheetName, int[] start1, int[] end1, int[] start2, int[] end2, int target, int need) {
        Sheet sheet = getSheet(excel, sheetName);
        Map<Object, Object> findMap = new HashMap<>();
        Row row;
        Cell cell;
        for(int i = start2[0]; i <= end2[0]; i++) {
            findMap.put(getValue(sheet, i, start2[1]), getValue(sheet, i, target));
        }
        for(int j = start1[0]; j <= end1[0]; j++) {
            row = sheet.getRow(j);
            cell = row.createCell(need);
            Object obj = getValue(sheet, j, start1[1]);
            if(findMap.containsKey(obj)) {
                Object obj2 = findMap.get(obj);
                if (obj2.getClass() == Double.class) {
                    cell.setCellValue((Double) findMap.get(obj));
                } else if (obj2.getClass() == String.class) {
                    cell.setCellValue((String) findMap.get(obj));
                } else if (obj2.getClass() == Date.class) {
                    cell.setCellValue((Date) findMap.get(obj));
                } else if (obj2.getClass() == Boolean.class) {
                    cell.setCellValue((Boolean) findMap.get(obj));
                } else {
                    cell.setBlank();
                }
            }else{
                cell.setBlank();
            }
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
        }else if("STRING".equals(type)) {
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

    public static int[] convert(String str) throws Exception {
        int[] arr = new int[2];
        str = str.toUpperCase();
        char[] cArr = str.toCharArray();
        if( cArr[cArr.length - 1] < '0' || cArr[cArr.length - 1] > '9') throw new Exception("数据格式错误");
        boolean flag = true;
        int num = 0;
        int count = 0;
        StringBuilder sbr = new StringBuilder();
        for(int i = cArr.length - 1; i >= 0; i--) {
            char ch = cArr[i];
            if(flag && ch >= '0' && ch <= '9'){
                sbr.insert(0, ch);
            }else{
                if( ch < 'A' || ch > 'Z') throw new Exception("数据格式错误");
                num += Math.pow(26, count) * (ch - 'A' + 1);
                count++;
                flag = false;
            }
        }
        arr[0] = Integer.parseInt(sbr.toString()) - 1;
        arr[1] = num == 0 ? 0 : num - 1;
        return arr;
    }

    public static int convertColumn(String str) throws Exception {
        char[] cArr = str.toCharArray();
        int num = 0;
        int count = 0;
        for(int i = cArr.length - 1; i >= 0; i--) {
            char ch = cArr[i];
            if( ch < 'A' || ch > 'Z') throw new Exception("数据格式错误");
            num += Math.pow(26, count) * (ch - 'A' + 1);
            count++;
        }
        return num == 0 ? 0 : num - 1;
    }

    public static void main(String[] args) throws Exception {
        int[] a = convert("BA209");
        System.out.println(a[0]);
        System.out.println(a[1]);
        System.out.println(convertColumn("BA"));
    }
}
