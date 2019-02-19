package com.cec.excel.common;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;

/**
 * @Author: luhk
 * @Email lhk2014@163.com
 * @Date: 2019/1/16 3:52 PM
 * @Description: 数据库帮助类
 * @Created with cec-excel-applets
 * @Version: 1.0
 */
public class DBHelper {

    private static String driver = "org.postgresql.Driver";

    private Connection connection = null;

    public Connection getConnection() {
        return connection;
    }


    public Statement getStatement() {
        return statement;
    }


    private Statement statement = null;

    public DBHelper(String url,String user,String password) {
        try {
            Class.forName(driver);
            connection = DriverManager.getConnection(url, user, password);
            statement = connection.createStatement();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void close() {
        try {
            this.connection.close();
            this.statement.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
