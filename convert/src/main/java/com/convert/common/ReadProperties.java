package com.convert.common;

import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.util.Properties;

/**
 * 프로퍼티 설정 파일을 읽어들인다.
 *
 * 2023. 06. 27
 */
public class ReadProperties {
    private static String propFilePath = "excel.properties";
    private static Properties props = new Properties();

    private ReadProperties() {
    }

    public static String getProperty(String name) {
        String val = "";

        try {
            InputStream inputStream = ReadProperties.class.getClassLoader().getResourceAsStream(propFilePath);
            if (inputStream != null) {
                Reader reader = new java.io.InputStreamReader(inputStream, java.nio.charset.StandardCharsets.UTF_8);
                props.load(reader);
                val = props.getProperty(name);
                reader.close();
            }
        } catch (IOException e) {
            return "";
        }

        return val;
    }
}
