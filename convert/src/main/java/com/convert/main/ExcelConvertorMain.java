package com.convert.main;

import com.convert.excel.ConvertDataSheet;
import com.convert.excel.ExcelReader;

import java.util.ArrayList;
import java.util.List;

public class ExcelConvertorMain {

    public static void main(String[] args){
        ExcelReader er = new ExcelReader();

        int fileLen = ExcelReader.size();
        List<ConvertDataSheet> excelList = new ArrayList<ConvertDataSheet>();

        for(int i = 0 ; i < fileLen ; i++){
            ConvertDataSheet excel = new ConvertDataSheet(ExcelReader.getFile(i));
            excelList.add(excel);
        }

        for(ConvertDataSheet excel : excelList){
            excel.convertRow();
        }

    }
}
