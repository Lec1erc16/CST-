package org.example;
import service.ReadExcel;
import service.WriteExcel;

public class Main {
    public static void main(String[] args) throws Exception {

        String PATH = "D:\\CSTOA_Update_Load\\doc\\";//父目录的绝对路径
        String excelName = "excel.xlsx";
        ReadExcel readExcel = new ReadExcel();
        readExcel.generateJSON(PATH,excelName);
        //WriteExcel writeExcel = new WriteExcel();
        //writeExcel.writeExcel(PATH,excelName);
    }
}