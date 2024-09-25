package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


//录入名单后复制学号与密码，
public class WriteExcel {

    private String clipString;

    public WriteExcel() throws FileNotFoundException {
    }


    public void getClipboard() throws IOException, UnsupportedFlavorException {
        //获取系统剪切板
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        Transferable content = clipboard.getContents(null);//从系统剪切板中获取数据
        if (content.isDataFlavorSupported(DataFlavor.stringFlavor)) {//判断是否为文本类型
            String text = (String) content.getTransferData(DataFlavor.stringFlavor);//从数据中获取文本值
            if (text == null) {
                this.clipString = null;
            }
            this.clipString = text;
            // System.out.println(text);
        }

    }

    public void writeExcel(String PATH,String excelName) throws IOException, UnsupportedFlavorException {
        getClipboard();
        // System.out.println(clipString);

        //已知OA输出的账密是按照姓名排序的，这时候只需将录入Excel中按姓名升序排序，然后在最后一行插入密码即可

        FileInputStream fileInputStream=new FileInputStream(PATH+excelName);
        Workbook wb = new XSSFWorkbook(fileInputStream);
        Sheet sheet = wb.getSheetAt(0);

        int rowIndex = 1;//信息所在行
        int passwordColIndex=5;//根据实际调整所在列

        //匹配密码模式串
        Pattern pattern = Pattern.compile("\\S{22}[=][=]");
        Matcher matcher = pattern.matcher(clipString);
        while (matcher.find()) {
            Row row = sheet.getRow(rowIndex++);
            String pswd = clipString.substring(matcher.start(), matcher.end());
            Cell cell = row.createCell(passwordColIndex);
           cell.setCellValue(pswd);
            //System.out.println(clipString.substring(matcher.start(),matcher.end()));
        }
        OutputStream outputStream=new FileOutputStream(PATH+excelName);
        wb.write(outputStream);
        outputStream.close();
    }
}
