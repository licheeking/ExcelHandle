package com.intergraph;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class RowDelete {

    public static void main(String[] args) {
        delteRow();
    }

    public static void delteRow(){
        try{
            FileInputStream is = new FileInputStream("d://1.xls");
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            HSSFSheet sheet = workbook.getSheetAt(0);
            int ls=sheet.getLastRowNum();
            sheet.shiftRows(1, ls, -1);
            Row row=sheet.getRow(ls);
            sheet.removeRow(row);
            FileOutputStream os = new FileOutputStream("d://3.xls");
            workbook.write(os);
            is.close();
            os.close();
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
}
