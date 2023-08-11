package com.convert.excel;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel file read
 */
public class ExcelReader {
    private static Logger log = new Logger().getLogger(ExcelReader.class);

    // Excel 경로
    private static String INPUT_EXCEL_PATH;
    // Excel 확장자
    private static String EXCEL_POST_FIX;

    // Excel File list
    private static List<File> excelList;

    public ExcelReader(){
        System.out.println(ReadProperties.getProperty("INPUT_EXCEL_PATH"));
        INPUT_EXCEL_PATH = INPUT_EXCEL_PATH == null ? ReadProperties.getProperty("INPUT_EXCEL_PATH") : INPUT_EXCEL_PATH;
        EXCEL_POST_FIX   = EXCEL_POST_FIX   == null ? ReadProperties.getProperty("EXCEL_POST_FIX")   : EXCEL_POST_FIX;

        fileRead();
    }

    /**
     * List up Excel file list
     */
    private static void fileRead(){
        excelList = new ArrayList<File>();

        File path = new File(INPUT_EXCEL_PATH);

        if(!path.isDirectory()){
            log.debug("Excel 파일을 읽어들이기 위한 경로 설정이 잘못되었습니다. 엑셀파일 경로는 폴더로 설정해주세요.");

            return ;
        }

        // read all files
        File[] fileList = path.listFiles();
        for(File excel : fileList){
            if(excel.getName().indexOf(EXCEL_POST_FIX) > 0){
                excelList.add(excel);
                log.info("READ FILE [" + excelList.size() + "] : " + excel.getName());
            }
        }
    }

    /**
     * 특정 인덱스의 엑셀 FileReader 를 반환한다.
     *
     * @param index
     * @return
     * @throws FileNotFoundException
     */
    public static FileReader getFileReader(int index) throws FileNotFoundException {
        return excelList != null && excelList.size() > index ? new FileReader(excelList.get(index)) : null;
    }

    /**
     * 특정 인덱스의 File 을 반환한다.
     *
     * @param index
     * @return
     */
    public static File getFile(int index){
        return excelList != null && excelList.size() > index ? excelList.get(index) : null;
    }

    /**
     * 현재 읽어들인 파일의 갯수를 응답한다.
     *
     * @return
     */
    public static int size(){
        return excelList != null ? excelList.size() : 0;
    }

}
