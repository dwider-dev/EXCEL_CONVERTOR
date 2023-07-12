package com.convert.main;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;
import com.convert.excel.ConvertDataSheet;
import com.convert.excel.ExcelReader;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelConvertorMain {
    private static Logger log = new Logger().getLogger(ExcelConvertorMain.class);

    public static void main(String[] args){
        ExcelReader er = new ExcelReader();

        int fileLen = ExcelReader.size();
        List<ConvertDataSheet> excelList = new ArrayList<ConvertDataSheet>();

        for(int i = 0 ; i < fileLen ; i++){
            System.out.println("file Length ..  : " + fileLen);
            ConvertDataSheet excel = new ConvertDataSheet(ExcelReader.getFile(i));
            excelList.add(excel);
        }

        for(ConvertDataSheet excel : excelList){
            System.out.println("excel");
            excel.convertRow();
        }

    }
}
