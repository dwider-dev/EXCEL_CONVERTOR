package com.convert.main;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;
import com.convert.excel.ConvertDataSheet;
import com.convert.excel.ExcelReader;

import java.io.*;
import java.util.List;

public class ExcelConvertorMain {
    private static Logger log = new Logger().getLogger(ExcelConvertorMain.class);

    public static void main(String[] args) throws FileNotFoundException {

        ExcelReader excelReader = new ExcelReader();

        int fileCount = excelReader.size();
        if(fileCount == 0){
            log.debug("읽어들일 Excel 파일이 없습니다.");
            return;
        }
        for(int i =0; i< fileCount; i++){
            try{
                FileReader fileReader = ExcelReader.getFileReader(i);
                if (fileReader == null) {
                    log.debug("Excel 파일을 찾을 수 없습니다.");
                    continue;
                }

                File inputFile = ExcelReader.getFile(i);
                log.debug("변환 중인 파일: " + inputFile.getName());

                ConvertDataSheet convertDataSheet = new ConvertDataSheet(inputFile);
                convertDataSheet.convertRow(); // 변환 데이터 가져오



                fileReader.close();

            }catch(FileNotFoundException e){
                log.error("Excel 파일을 찾을 수 없습니다.");

            }catch(Exception e){
                log.error("Excel 파일 변환 중 오류가 발생했습니다.");
                e.printStackTrace();

            }
        }

    }
}
