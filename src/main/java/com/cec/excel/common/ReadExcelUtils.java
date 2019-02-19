package com.cec.excel.common;

import java.io.*;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelUtils {
    private Workbook wb;
    private Sheet sheet;
    private Row row;

    public ReadExcelUtils(String filepath,String fileName){
        if(filepath==null){
            return;
        }
        String ext = fileName.substring(fileName.lastIndexOf("."));
        try {
            InputStream is = new FileInputStream(filepath);
            if(".xls".equals(ext)){
                wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(ext) || ".xlsm".equals(ext)){
                wb = new XSSFWorkbook(is);
            }else{
                wb = null;
            }
        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException :"+e);
        } catch (IOException e) {
            System.out.println("IOException :"+e);
        }
    }

    /**
     * 读取Excel表格表头的内容
     *
     * @param
     * @return String 表头内容的数组
     * @author luhk
     */
    public String[] readExcelTitle() throws Exception{
        if(wb==null){
            throw new Exception("Workbook对象为空！");
        }
        sheet = wb.getSheetAt(0);
        row = sheet.getRow(0);
        // 标题总列数
        int colNum = row.getPhysicalNumberOfCells();
        System.out.println("colNum:" + colNum);
        String[] title = new String[colNum];
        for (int i = 0; i < colNum; i++) {
            title[i] = row.getCell(i).getCellFormula();
        }
        return title;
    }

    /**
     * 创建一张存储数据库信息的sheet
     *
     * @param
     * @return String 表头内容的数组
     * @author luhk
     */
    public void createExcelForDatabase(String filePath, Map<String,String> param) throws Exception{
        if(wb==null){
            throw new Exception("Workbook对象为空！");
        }

        String database = param.get("database");
        String host = param.get("host");
        String databaseUser = param.get("databaseUser");
        String userPassword = param.get("userPassword");
        //创建Excel工作表对象
        Sheet sheet = wb.getSheet("数据库信息");
        if(sheet == null ){
            sheet = wb.createSheet("数据库信息");
        }
        //创建Excel工作表的行
        Row titleRow = sheet.createRow((short)0);
        //设置Excel工作表的值
        titleRow.createCell((short)0).setCellValue("数据库名称");
        titleRow.createCell((short)1).setCellValue("IP");
        titleRow.createCell((short)2).setCellValue("数据库用户");
        titleRow.createCell((short)3).setCellValue("用户密码");
        Row contentRow = sheet.createRow((short)1);
        contentRow.createCell((short)0).setCellValue(database);
        contentRow.createCell((short)1).setCellValue(host);
        contentRow.createCell((short)2).setCellValue(databaseUser);
        contentRow.createCell((short)3).setCellValue(userPassword);
        FileOutputStream os = new FileOutputStream(filePath);
        wb.write(os);
        os.flush();
        os.close();
    }

    /**
     * 读取Excel数据内容
     *
     * @param
     * @return Map 包含单元格数据内容的Map对象
     * @author luhk
     */
    public Map<Integer, Map<Integer,Object>> readExcelContent() throws Exception{
        if(wb==null){
            throw new Exception("Workbook对象为空！");
        }
        Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();
        //读取第几个sheet
        sheet = wb.getSheetAt(0);
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 0; i <= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
            while (j < colNum) {
                Object obj = getCellFormatValue(row.getCell(j));
                cellValue.put(j, obj);
                j++;
            }
            content.put(i, cellValue);
        }
        return content;
    }

    /**
     * 写入企业编码，申请时间
     *
     * @param
     * @return companyCode, applyDate
     * @author luhk
     */
    public void writeExcelContent(String filePath) throws Exception{
        if(wb==null){
            throw new Exception("Workbook对象为空！");
        }
        //读取第几个sheet
        sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(2);
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            throw new RuntimeException("请输入非数字企业名称！");
        }
        String companyName = getCellFormatValue(row.getCell(3)).toString();
        //创建样式
        CellStyle cellStyle = wb.createCellStyle();
        // 设置单元格水平方向对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置单元格垂直方向对齐方式
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Date currentTime = new Date();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        String applyDate = formatter.format(currentTime);
        Row companyRow = sheet.getRow(2);
        Row applyDateRow = sheet.getRow(3);
        Cell companyCell = companyRow.getCell((short)3);
        if (companyCell == null) {
            companyCell = companyRow.createCell((short)3);
        }
        companyCell.setCellValue(companyName.hashCode());
        companyCell.setCellStyle(cellStyle);
        Cell applyDateCell = applyDateRow.getCell((short)3);
        if (applyDateCell == null) {
            applyDateCell = applyDateRow.createCell((short)3);
        }
        applyDateCell.setCellValue(applyDate);
        applyDateCell.setCellStyle(cellStyle);
        FileOutputStream os = new FileOutputStream(filePath);
        wb.write(os);
        os.flush();
        os.close();
    }

    /**
     *
     * 根据Cell类型设置数据
     *
     * @param cell
     * @return
     * @author luhk
     */
    private Object getCellFormatValue(Cell cell) {
        Object cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:// 如果当前Cell的Type为NUMERIC
                    cellvalue = cell.getNumericCellValue();
                case Cell.CELL_TYPE_FORMULA: {
                    // 判断当前的cell是否为Date
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 如果是Date类型则，转化为Data格式
                        // data格式是带时分秒的：2013-7-10 0:00:00
                        // cellvalue = cell.getDateCellValue().toLocaleString();
                        // data格式是不带带时分秒的：2013-7-10
                        Date date = cell.getDateCellValue();
                        cellvalue = date;
                    } else {// 如果是纯数字

                        // 取得当前Cell的数值
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:// 如果当前Cell的Type为STRING
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:// 默认的Cell值
                    cellvalue = "";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;
    }

    /*public static void main(String[] args) {
        String userPassword = String.valueOf((int)((Math.random()*9+1)*100000));
        ReadExcelUtils excelReader = new ReadExcelUtils("/Users/luhk/工作/中电数据/表格制作/数据集选择医院锁定.xlsm","数据集选择医院锁定.xlsm");
        Map<String,String> param = new HashMap<String,String>();
        String defaultDatabase = "postgres";
        String database = "zdsj_test9";
        String host = "192.168.1.35";
        String username = "postgres";
        String password = "postgres";
        String databaseUser = "zdcs_user9";
        param.put("database",database);
        param.put("host",host);
        param.put("username",username);
        param.put("password",password);
        param.put("databaseUser",databaseUser);
        param.put("userPassword",userPassword);

        try{
            //数据库信息保存表格中
            excelReader.createExcelForDatabase("/Users/luhk/工作/中电数据/表格制作/数据集选择医院锁定.xlsm",param);
            excelReader.writeExcelContent("/Users/luhk/工作/中电数据/表格制作/数据集选择医院锁定.xlsm");
            String url = "jdbc:postgresql://" + host + ":5432/" + defaultDatabase;
            DBHelper dbHelper = new DBHelper(url,username,password);
            Statement statement = dbHelper.getStatement();
            //创建数据库
            statement.addBatch("create database " + database);
            //创建用户
            statement.addBatch("create user "+ databaseUser +" with password " + "'"+ userPassword + "'");
            //给用户授予连接权限
            statement.addBatch("grant connect on database "+database+" to "+databaseUser);
            statement.executeBatch();
            dbHelper.close();

            String url2 = "jdbc:postgresql://" + host + ":5432/" + database;
            DBHelper dbHelper2 = new DBHelper(url2,username,password);
            Statement statement2 = dbHelper2.getStatement();
            //创建模式
            statement2.addBatch("create schema " + databaseUser);
            //给用户授予usage权限
            statement2.addBatch("grant usage on schema "+databaseUser+" to " + databaseUser);
            //给用户授予select权限
            statement2.addBatch("alter default privileges in schema "+databaseUser+" grant select on tables to "+databaseUser);
            statement2.executeBatch();
            dbHelper2.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }*/
}