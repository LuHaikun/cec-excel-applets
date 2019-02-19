package com.cec.excel;

import com.cec.excel.common.DBHelper;
import com.cec.excel.common.ReadExcelUtils;
import java.io.FileNotFoundException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;

public class App 
{
    public static void main( String[] args )
    {
        if (args.length < 7) {
            throw new IllegalArgumentException("参数输入异常! ");
        }
        String defaultDatabase = "postgres";
        String newDatabase = args[0];
        String host = args[1];
        String port = args[2];
        String username = args[3];
        String password = args[4];
        String databaseUser = args[5];
        String filePath = args[6];
        String fileName = filePath.substring(filePath.lastIndexOf("/"));
        System.out.println("args = [" + newDatabase + ","+host+"," + port + ","+username+","+databaseUser+","+password+","+filePath+"]");
        String url = "jdbc:postgresql://" + host + ":" + port + "/";
        try {
            //读取excel文件
            String userPassword = String.valueOf((int)((Math.random()*9+1)*100000));
            ReadExcelUtils excelReader = new ReadExcelUtils(filePath,fileName);
            Map<String,String> param = new HashMap<String,String>();
            param.put("database",newDatabase);
            param.put("host",host);
            param.put("username",username);
            param.put("password",password);
            param.put("databaseUser",databaseUser);
            param.put("userPassword",userPassword);

            DBHelper defaultDbHelper = new DBHelper(url + defaultDatabase,username,password);
            //创建数据库、用户、授权
            Statement statement1 = defaultDbHelper.getStatement();
            //创建数据库
            statement1.addBatch("create database " + newDatabase);
            //创建用户
            statement1.addBatch("create user "+ databaseUser +" with password " + "'"+ userPassword + "'");
            //给用户授予connect权限
            statement1.addBatch("grant connect on database "+newDatabase+" to "+databaseUser);
            statement1.executeBatch();
            defaultDbHelper.close();
            DBHelper newDbHelper = new DBHelper(url + newDatabase,username,password);
            Statement statement2 = newDbHelper.getStatement();
            //创建模式
            statement2.addBatch("create schema " + databaseUser);
            //给用户授予usage权限
            statement2.addBatch("grant usage on schema "+databaseUser+" to " + databaseUser);
            //给用户授予select权限
            statement2.addBatch("alter default privileges in schema "+databaseUser+" grant select on tables to "+databaseUser);
            statement2.executeBatch();
            //数据库信息保存表格中
            excelReader.createExcelForDatabase(filePath,param);
            excelReader.writeExcelContent(filePath);
            newDbHelper.close();
            System.out.println("操作成功!");
        } catch (FileNotFoundException e) {
            System.out.println("未找到指定路径的文件!");
            e.printStackTrace();
        }catch (Exception e) {
            e.printStackTrace();
        }
    }
}
