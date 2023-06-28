package com.convert.excel;

import com.convert.common.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 파일의 각 라인을 읽어들여 분석 로직에따라 변환한다.
 * 변환된 데이터를 보관한다.
 */
public class ConvertDataSheet {
    private static Logger log = new Logger().getLogger(ConvertDataSheet.class);

    private List<Object[]> rows;
    private String[] columnTitle;

    private int columnLength;
    private int rowLength;
    private XSSFSheet sheet;

    /**
     * 생성자 : 엑셀파일을 읽어 메모리에 저장한다.
     * @param workFile
     */
    public ConvertDataSheet(File workFile){
        columnLength  = 0;
        rowLength = 0;

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(workFile));
            sheet = workbook.getSheetAt(0);
            XSSFRow row;

            /*
            1. Row 갯수 조회

            2. Column 갯수 조회

            3. 첫 Row(Title) 별도 분리 및 기록

            4. 나머지 Row Data 입력

             */
            rowLength = sheet.getPhysicalNumberOfRows();
            if (rowLength > 0) {
                row = sheet.getRow(0);

                columnLength = row.getPhysicalNumberOfCells();
                columnTitle = new String[columnLength];

                for(int i = 0 ; i < columnLength ; i++){
                    XSSFCell cell = row.getCell(i);
                    columnTitle[i] = cell.getStringCellValue();
                }
            }
            else {
                return;
            }


            // Read row data

            // rows
            for(int i = 1 ; i < rowLength ; i ++){
                row = sheet.getRow(i);

                rows = new ArrayList<Object[]>();
                Object celObj[] = new Object[columnLength];

                // cols
                for(int j = 0 ; j < columnLength ; j++){
                    XSSFCell cell = row.getCell(j);
                    switch(cell.getCellType()) {   // 각셀의 데이터값을 가져올때 맞는 데이터형으로 변환한다.
                        case HSSFCell.CELL_TYPE_FORMULA:
                            String strValFormula = cell.getCellFormula();
                            celObj[j] = strValFormula;
                            break;
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            Integer intVal = Integer.parseInt(String.format("%d", cell.getNumericCellValue()));
                            celObj[j] = intVal;
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            String strVal = cell.getStringCellValue();
                            celObj[j] = strVal;
                            break;
                        case HSSFCell.CELL_TYPE_BLANK:
                            String strValBlank = "";
                            celObj[j] = strValBlank;
                            break;
                        case HSSFCell.CELL_TYPE_ERROR:
                            String strValError = String.valueOf(cell.getErrorCellValue());
                            celObj[j] = strValError;
                            break;
                        default:
                    }
                }
                // cols read end

                rows.add(celObj);
                // modify new data size
                rowLength = rows.size();
            }
        }catch (IOException e){
            log.error("읽기 파일에 문제가 있습니다. 다음 파일 : " + workFile.getName() , e);
        }
    }

    public int getRowLength(){
        return rowLength;
    }

    public Object[] convertRow(){
        Object[] row = new Object[columnLength];

        return row;
    }
}
