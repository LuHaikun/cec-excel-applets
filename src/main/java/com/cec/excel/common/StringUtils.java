package com.cec.excel.common;

import java.util.regex.Pattern;

/**
 * @Author: luhk
 * @Email lhk2014@163.com
 * @Date: 2019/1/18 9:12 AM
 * @Description:
 * @Created with cec-excel-applets
 * @Version: 1.0
 */
public class StringUtils {
    public static boolean isIP(String str) {
        String regex = "[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}";
        Pattern pattern = Pattern.compile(regex);
        return pattern.matcher(str).matches();

    }

    public static boolean isNotEmpty(String str) {
        boolean flag = true;
        if(str!=null &&  str.length()>0){
            flag = true;
        }else {
            flag = false;
        }
        return flag;
    }
}
