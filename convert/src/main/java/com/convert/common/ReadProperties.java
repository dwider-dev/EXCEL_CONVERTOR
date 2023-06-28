package com.convert.common;

import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.util.Properties;

/**
 * 프로퍼티 설정 파일을 읽어들인다.
 *
 * 2023. 06. 27
 */
public class ReadProperties {
    private static String propFilePath = "resources/excel.properties";
    private static Properties props = new Properties();

    ReadProperties(){
    }

    public static String getProperty(String name) {
        String val = "";

        try {
            Reader reader = new FileReader(propFilePath);
            props.load(reader);

            val = props == null ? "" :props.getProperty(name);

            reader.close();
        }catch(IOException e) {
            return "";
        }

        return val;
    }
}
