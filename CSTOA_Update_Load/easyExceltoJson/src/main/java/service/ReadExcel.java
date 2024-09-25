package service;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadExcel {

    public String PATH;
    public FileInputStream fileInputStream;
    public Workbook workbook;

    public void setPATH(String PATH, String excelname) throws Exception {
        this.PATH = PATH;
        fileInputStream = new FileInputStream(PATH + excelname);
        workbook = new XSSFWorkbook(fileInputStream);//读取的是.xlsx后缀的文件
    }

    public String getStringComment(int sheet, int row, int col) {
        Sheet sheet1 = workbook.getSheetAt(sheet);
        Row row1 = sheet1.getRow(row);
        Cell cell = row1.getCell(col);
        if (cell == null) return null;
        return cell.getStringCellValue();
    }

    public double getNumComment(int sheet, int row, int col) {
        Sheet sheet1 = workbook.getSheetAt(sheet);
        Row row1 = sheet1.getRow(row);
        Cell cell = row1.getCell(col);
        if (cell == null) return 0;
        return cell.getNumericCellValue();
    }

    public boolean isRowEmpty(int sheet, int row) {
        Sheet sheet1 = workbook.getSheetAt(sheet);
        if (sheet1.getRow(row) == null) return true;
        else return false;
    }

    /**
     * PATH 父目录
     * excelName 被读取Excel文件名
     * index 学号所在列
     */
    public void generateJSON(String PATH,String excelName) {

        try {
            File jsonPath = new File(PATH + "jsonName.json");
            if (!jsonPath.getParentFile().exists()) {
                // 如果父目录不存在，创建父目录
                jsonPath.getParentFile().mkdirs();
            }
            if (jsonPath.exists()) {
                // 如果json已存在,删除旧文件
                jsonPath.delete();
            }
            jsonPath.createNewFile();

            //构建json的string语句
            ReadExcel ReadExcel = new ReadExcel();
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.append("\n{\"members\": [");
            ReadExcel.setPATH(PATH, excelName);
            int rowCount = ReadExcel.workbook.getSheetAt(0).getPhysicalNumberOfRows();
            for (int i = 1; i <= rowCount; i++) {
                if (ReadExcel.isRowEmpty(0, i)) break;
                //以下根据实际excel表格结构进行调整
                int index = 0;
                String id = ReadExcel.getStringComment(0, i, index++);
                String name = ReadExcel.getStringComment(0, i, index++);
                String phone = ReadExcel.getStringComment(0, i, index++);
                String wxid = ReadExcel.getStringComment(0, i, index++);
                String qq = ReadExcel.getStringComment(0, i, index++);

                stringBuilder.append("{\"id\": \"" + id + "\",");
                stringBuilder.append("\"name\": \"" + name + "\",");
                stringBuilder.append("\"phone\": \"" + phone + "\",");
                stringBuilder.append("\"wxid\": \"" + id + "\",");
                if (i != ReadExcel.workbook.getSheetAt(0).getLastRowNum())
                    stringBuilder.append("\"qq\": \"" + qq + "\"},");
                else stringBuilder.append("\"qq\": \"" + qq + "\"}");

            }
            stringBuilder.append("]}");

            //将string写入json文件
            JSONObject object = JSONObject.parseObject(stringBuilder.toString());
            String pretty = JSON.toJSONString(object, SerializerFeature.PrettyFormat, SerializerFeature.WriteMapNullValue, SerializerFeature.WriteDateUseDateFormat);
            System.out.println(pretty);
            Writer write = new OutputStreamWriter(new FileOutputStream(jsonPath), "UTF-8");
            write.write(pretty);
            write.flush();
            write.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
