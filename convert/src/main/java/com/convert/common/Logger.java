package com.convert.common;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 로그를 기록한다
 */
public class Logger {
    // 로그 저장 경로
    private static String FILE_PATH;
    // 로그파일 저장시 날짜 포맷
    private static String FILE_DATE_FORMAT;
    // 끝 파일명
    private static String POST_FIX;
    // 첫 파일명
    private static String PRE_FIX;

    /**
     * Logging 레벨
     * <br>
     * <br>
     * DEBUG : Write debug, error, info
     * <br>
     * ERROR : Write debug, error
     * <br>
     * INFO  : Write info, error
     *
     */
    private static String LOG_MOD;

    private Class c;

    /**
     * Initialize Logger class
     *
     * @param
     */
    public Logger(){
        FILE_PATH        = FILE_PATH == null        ? ReadProperties.getProperty("FILE_PATH")        : FILE_PATH;
        FILE_DATE_FORMAT = FILE_DATE_FORMAT == null ? ReadProperties.getProperty("FILE_DATE_FORMAT") : FILE_DATE_FORMAT;
        POST_FIX         = POST_FIX == null         ? ReadProperties.getProperty("POST_FIX")         : POST_FIX;
        PRE_FIX          = PRE_FIX == null          ? ReadProperties.getProperty("PRE_FIX")          : PRE_FIX;

        LOG_MOD          = LOG_MOD == null          ? ReadProperties.getProperty("LOG_MOD")          : LOG_MOD;
    }

    /**
     * Logger 인스턴스를 반환한다.
     *
     * @param c
     * @return Logger
     */
    public Logger getLogger(Class c){
        this.c = c;

        return this;
    }

    /**
     * 메시지를 파일에 기록한다.
     *
     * @param msg
     */
    private static synchronized void fileWrite(String msg){
        // get Date ================================================================================
        SimpleDateFormat sdf = new SimpleDateFormat(FILE_DATE_FORMAT);
        Date date = new Date();
        String nowDate = sdf.format(date);
        // =========================================================================================

        try {
            File logFile = new File(FILE_PATH + "/" + PRE_FIX + nowDate + POST_FIX);
            FileWriter writer = new FileWriter(logFile, true);
            BufferedWriter br = new BufferedWriter(writer);
            br.write(msg + "\\n");

            br.close();
            writer.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Write DEBUG log message in the log file
     * <br>
     * **Only running with the DEBUG mode
     *
     * @param msg
     */
    public void debug(String msg){
        if(LOG_MOD.equals("INFO")){
            return;
        }

        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        Date date = new Date();
        String nowDate = sdf.format(date);

        String logMsg = nowDate + " [DEBUG] (" +  this.c.getName() + ") : " + msg;

        fileWrite(logMsg);
    }

    /**
     * Write INFO log message in the log file
     *
     * @param msg
     */
    public void info(String msg){
        if(LOG_MOD.equals("ERROR")){
            return;
        }

        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        Date date = new Date();
        String nowDate = sdf.format(date);

        String logMsg = nowDate + " [INFO] (" +  this.c.getName() + ") : " + msg;

        fileWrite(logMsg);
    }

    /**
     * Write Error log message in the log file
     *
     * @param msg
     */
    public void error(String msg){
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        Date date = new Date();
        String nowDate = sdf.format(date);

        String logMsg = nowDate + " [ERROR] (" +  this.c.getName() + ") : " + msg;

        fileWrite(logMsg);
    }

    /**
     * Write Error log message in the log file with Throwable
     *
     * @param msg
     * @param e
     */
    public void error(String msg, Throwable e){
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        Date date = new Date();
        String nowDate = sdf.format(date);

        String logMsg = nowDate + " [ERROR] (" +  this.c.getName() + ") : " + msg;

        fileWrite(logMsg);
        fileWrite(e.getStackTrace().toString());
    }
}
