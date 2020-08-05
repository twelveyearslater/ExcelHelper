package top.masochism.entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import top.masochism.bo.Excel;
import top.masochism.utils.ExcelUtils;

import java.io.IOException;

public class Test {

   /* public static void main(String[] args) throws IOException {
        Excel excel = new Excel("h:\\" , "test.xlsx");
        Sheet test1Sheet = excel.workbook.getSheet("test1");
        Row row = test1Sheet.getRow(0);
        for(int i = 0; i < 6; i++) {
            Cell cell = row.getCell(i);
            System.out.println(cell.getCellType());
            System.out.println(cell.getRowIndex());
            System.out.println(cell.getColumnIndex());
            System.out.println(cell.getCellStyle().getDataFormat());
            System.out.println(("" + cell.getCellType()).equals("NUMERIC") ? cell.getNumericCellValue() : cell.getStringCellValue());
        }
    }*/

    public static void main(String[] args) throws IOException {
        Excel excel = new Excel("e:\\" , "test.xlsx");
        ExcelUtils.VLOOKUP(excel, "test2", 'A', 1, 30, 'C', 'G', 1, 30, 'H');
        excel.updateExcel();
    }
}
