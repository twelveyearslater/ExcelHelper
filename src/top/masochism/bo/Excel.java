package top.masochism.bo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Excel {

    private String path;
    private String name;
    private File file;
    public Workbook workbook;

    public Excel(String path, String name) throws IOException {
        this.path = path;
        this.name = name;
        file = new File(path + name);
        InputStream is = new FileInputStream(file);
        if(name.indexOf("xlsx") > 0) {
            workbook = new XSSFWorkbook(is);
        }else{
            workbook = new HSSFWorkbook(is);
        }
    }

    public void updateExcel() throws IOException {
        OutputStream os = new FileOutputStream(file);
        workbook.write(os);
    }
}
