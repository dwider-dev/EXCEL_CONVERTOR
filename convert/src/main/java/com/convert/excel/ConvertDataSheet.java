package com.convert.excel;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;
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
    private List<Object[]> convertRows;
    private String[] columnTitle;

    private int columnLength;
    private int rowLength;
    private XSSFSheet sheet;

    /*
    설정 옵션 목록
    # $NONE : 참고할 데이터 없음. (원본 엑셀에서 가져오는 데이터가 없는 경우 사용)
    # $NUMBERING : 순서대로 번호를 매겨준다.
    # $STRING : 읽은 문자열 그대로 입력한다.
    # $DATE_YYMM : 원본 엑셀의 날짜필드를 YY/MM 형식으로 바꿔서 날짜를 입력해준다.
    # $DATE_YYYYMMDD : 원본 엑셀의 날짜필드를 YYYY-MM-DD 형식으로 바꿔서 날짜를 입력해준다.
    # $NOTE_CONVERT_CARNUM : 차량번호를 적요에서 찾아 입력해준다.
    # $NOTE_CONVERT_CARNAME : 차량명으로 예측되는 문자열을 찾아 입력해준다.
    # $NOTE_CONVERT_OIL : 주유비를 찾아 입력해준다.
    # $NOTE_CONVERT_INS : 보험료를 찾아 입력해준다.
    # $NOTE_CONVERT_PAR : 주차비를 찾아 입력해준다.
    # $NOTE_CONVERT_OIL_INS_PAR : 주유비, 보험료, 주차비 중에 어느하나라도 있으면 합산하여 입력해준다.
     */

    private final String OPT_NONE = "$NONE";
    private final String OPT_NUMBERING = "$NUMBERING";
    private final String OPT_STRING = "$STRING";
    private final String OPT_DATE_YYMM = "$DATE_YYMM";
    private final String OPT_DATE_YYYYMMDD = "$DATE_YYYYMMDD";
    private final String OPT_NOTE_CONVERT_CARNUM = "$NOTE_CONVERT_CARNUM";
    private final String OPT_NOTE_CONVERT_CARNAME = "$NOTE_CONVERT_CARNAME";
    private final String OPT_NOTE_CONVERT_OIL = "$NOTE_CONVERT_OIL";
    private final String OPT_NOTE_CONVERT_INS = "$NOTE_CONVERT_INS";
    private final String OPT_NOTE_CONVERT_PAR = "$NOTE_CONVERT_PAR";
    private final String OPT_NOTE_CONVERT_OIL_INS_PAR = "$NOTE_CONVERT_OIL_INS_PAR";

    /**
     * 생성자 : 엑셀파일을 읽어 메모리에 저장한다.
     * @param workFile
     */
    public ConvertDataSheet(File workFile){
        columnLength  = 0;
        rowLength = 0;

        List<Object[]> convertRows = new ArrayList<Object[]>();

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

    public void convertRow(){
        Object[] row = new Object[columnLength];

        String[] targets = ReadProperties.getProperty("DISTRIBUTE_TARGETS").split(",");
        String[] outputFileName = new String[targets.length];
        String[] stdValue = new String[targets.length];
        String[] columnData = new String[targets.length];

        // 타겟별 설정값 Load
        for(int i = 0 ; i < targets.length ; i++){
            outputFileName[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_FILE_NAME");
            stdValue[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_STD_VALUE");
            columnData[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_COLUMN_DATA");
        }


        // read row
        for(int i = 0 ; i < rowLength ; i++){
            row = rows.get(i);

            /*
             * 한 줄씩 분석
             *
             * STD_VALUE 의 분류기준에 해당하는 데이터 먼저 찾아낸다.
             */

            // read column
            for(int j = 0 ; j < columnLength ; j++){
                if(equalIndex(columnTitle, row, stdValue[j])){

                    log.debug("[" + j + "] " + stdValue[j] + " - 분류 - OUTPUT - " +  outputFileName[j]);
                    Object[] convertRow =  procRow(row, columnData[j], outputFileName[j]);

                    // 오류 발생 시 중단
                    if(convertRow == null){
                        return;
                    }

                    convertRows.add(convertRow);
                    log.debug("처리 완료 => " + convertRow);


                }
            }// column end

        }// row end

    }


    /**
     * 원본 엑셀의 컬럼명과 데이터가 분류기준에 부합하는 값인지 체크
     *
     *
     * @param column
     * @param row
     * @param value
     * @return
     */
    private boolean equalIndex(String[] column, Object[] row, String value){
        String strVal[] = value.split(":");

        for(int i = 0 ; i < columnLength ; i++){
            // 컬럼명 : 컬럼값 일치하는 Row 인지 검증
            if(column[i].equals(strVal[0]) && String.valueOf(row[i]).equals(strVal[1])){
                return true;
            }
        }


        return false;
    }

    /**
     * 분류된 대상의 Row 를 처리기준에 맞게 처리 후, 목표 파일에 저장
     *
     * @param row
     * @param convertFunc
     * @param outFile
     * @return
     */
    private Object[] procRow(Object[] row, String convertFunc, String outFile){
        String[] convertStr = convertFunc.split(":");
        boolean firstFlagNone = false;

        if(convertStr.length != 3){
            log.error("_OUTPUT_COLUMN_DATA 옵션 설정 오류");
            return null;
        }

        // 첫번째 필드는 NONE 여부만 체크
        if(convertStr[0].equals(OPT_NONE)){

        }

        return row;
    }


}
