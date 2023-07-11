package com.convert.excel;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
    private final String OPT_DATE_MMDD = "$DATE_MMDD";
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

    /**
     * 읽어들인 파일의 전체 Row 를 분석 및 변환 한다.
     */
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
            for(int j = 0 ; j < targets.length ; j++){
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
        String[] convertCols = convertFunc.split(",");
        Object[] procRow = new Object[convertCols.length];
        String[] outColumns = new String[convertCols.length];

        if(convertCols.length <= 0){
            log.error("_OUTPUT_COLUMN_DATA 옵션 설정 오류");
            return null;
        }

        /*
        * 변환 하고자하는 결과 파일의 컬럼 수 만큼 처리
        * 원본 엑셀 데이터의 특정 컬럼 값을 가져와 옵션에 따른 후 처리를 진행하여 Output 파일에 기록하기 위한 데이터 분석 및 처리 부분
        *
        * */
        int j = 0;
        for(String options : convertCols){
            String[] args = options.split(":");

            // 각 컬럼 설정 당 옵션은 3개로 구성되어있어야 한다.
            if(args.length != 3){
                log.error("_OUTPUT_COLUMN_DATA 옵션 설정 오류");
                return null;
            }

            // First option
            Object targetValue = null;
            for(int i = 0 ; i < columnLength ; i++) {
                // NONE 옵션의 경우 동작 없음
                if(args[0].equals(OPT_NONE)){
                    break;
                }

                if(args[0].equals(columnTitle[i])) {
                    targetValue = row[i];
                }
            }

            // Second option

            /*
            OPT_NUMBERING
            OPT_STRING
            OPT_DATE_YYMM
            OPT_DATE_YYYYMMDD
            OPT_NOTE_CONVERT_CARNUM
            OPT_NOTE_CONVERT_CARNAME
            OPT_NOTE_CONVERT_OIL
            OPT_NOTE_CONVERT_INS
            OPT_NOTE_CONVERT_PAR
            OPT_NOTE_CONVERT_OIL_INS_PAR
             */
            Object output = null;
            switch (args[1]) {
                case OPT_NUMBERING:
                    // 파일 기록 시 동작할 수 있도록 옵션을 그대로 입력
                    output = OPT_NUMBERING;
                    break;
                case OPT_STRING:
                    // 원본 엑셀에서 가져온 데이터를 그대로 입력한다.
                    output = targetValue;
                    break;
                case OPT_DATE_MMDD:
                    output = transDate(String.valueOf(targetValue), "MM/dd");
                    break;
                case OPT_DATE_YYYYMMDD:
                    output = transDate(String.valueOf(targetValue), "yyyy-MM-dd");
                    break;
                case OPT_NOTE_CONVERT_CARNUM:
                    output = findCarNum(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_CARNAME:
                    output = findCarName(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_OIL:
                    break;
                case OPT_NOTE_CONVERT_INS:
                    break;
                case OPT_NOTE_CONVERT_PAR:
                    break;
                case OPT_NOTE_CONVERT_OIL_INS_PAR:
                    break;
                default:
                    break;
            }

            procRow[j] = output;

            // Third option
            outColumns[j] = args[2];

            j++;


        } // End proc row


        // Write file
        writeFile(procRow, outColumns, new File(ReadProperties.getProperty("OUTPUT_EXCEL_PATH") + "/" + outFile));


        return null;
    }

    /**
     * 기능 옵션 : $DATE_MMDD, $DATE_YYYYMMDD
     * <br>
     * 원본 데이터의 날짜 형식을 지정한 양식으로 변경한다.
     *
     * @param dateStr
     * @param format
     * @return
     */
    private String transDate(String dateStr, String format){
        String formatString = dateStr;
        DateFormat dfInput = new SimpleDateFormat("yyyyMMdd");
        DateFormat df = new SimpleDateFormat(format);

        try {
            formatString = df.format(dfInput.parse(dateStr));
        } catch (ParseException e) {
            log.error("날짜 형식 변환 오류 발생 : " + dateStr);
        }

        return formatString;
    }

    /**
     * 기능 옵션 : $NOTE_CONVERT_CARNUM
     * <br>
     * 지정 필드에서 차량번호에 해당하는 값을 찾아준다.
     *
     * @param data
     * @return
     */
    private String findCarNum(String data) {
        String format = "(\\d{2,3}\\D\\d{4})";
        Pattern pattern = Pattern.compile(format);
        Matcher matcher = pattern.matcher(data);

        if(matcher.find()) {
            return matcher.group();
        }

        return "";
    }

    /**
     * 기능 옵션 : $NOTE_CONVERT_CARNAME
     * <br>
     * 지정 필드에서 차량명에 해당하는 값을 찾아준다.
     *
     * @param data
     * @return
     */
    private String findCarName(String data) {
        String tempStr = data;


        // 전화번호 제거
        Pattern pattern = Pattern.compile("(\\d{3}.\\d{4}.\\d{4})");
        Matcher matcher = pattern.matcher(tempStr);

        if(matcher.find()) {
            tempStr = matcher.replaceAll("");
        }

        // 차량번호 제거
        pattern = Pattern.compile("(\\d{2,3}\\D\\d{4})");
        matcher = pattern.matcher(tempStr);

        if(matcher.find()) {
            tempStr = matcher.replaceAll("");
        }

        // 추가금액 제거
        pattern = Pattern.compile("(.[주,보,주차,주유,보험]\\d+.\\d)");
        matcher = pattern.matcher(tempStr);

        if(matcher.find()) {
            tempStr = matcher.replaceAll("");
        }

        // 예약 문구 제거
        String[] expectStrs = {"출발", "도착", "반차", "핸들", "주차", "주유", "픽업", "청불", "보험", "경유", "전달"};

        for(String expect : expectStrs){

            if(tempStr.indexOf(expect) >= 0){
                String tempRm = tempStr.substring(tempStr.indexOf(expect));
                tempRm = tempRm.substring(0, tempRm.indexOf("/") >= 0 ? tempRm.indexOf("/") : tempRm.length());
                tempStr = tempStr.replaceAll(tempRm, "");
            }

        }

        tempStr = tempStr.replaceAll("/", "");

        return tempStr.trim();
    }

    /**
     * 파일을 생성하고 변환된 Row를 기록한다.
     *
     * @param row
     * @param outColumns
     * @param file
     */
    private void writeFile(Object[] row, String[] outColumns, File file){
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet outputSheet;

            if(workbook.getNumberOfSheets() <= 0){
                log.debug("New file : " + file.getName());
                outputSheet = workbook.createSheet();
            } else{
                outputSheet = workbook.getSheetAt(0);
            }

            int outFileRowCnt = outputSheet.getPhysicalNumberOfRows();

            // Write titles
            if(outFileRowCnt <= 0){

                // Write titles
                // Write to first row
                XSSFRow titleRow = outputSheet.createRow(0);

                for(int i = 0 ; i < outColumns.length ; i++){
                    Object cellData = outColumns[i];
                    XSSFCell cell = titleRow.createCell(i);

                    cell.setCellValue(String.valueOf(cellData));

                }
                outFileRowCnt++;

            }// End write titles

            // Write row data
            for(int i = 0 ; i < row.length ; i++){
                Object cellData = row[i];
                XSSFRow writeRow = outputSheet.createRow(outFileRowCnt);
                XSSFCell cell = writeRow.createCell(i);

                // Numbering 기능 사용시 값 치환
                if(cellData.equals(OPT_NUMBERING)){
                    cellData = outFileRowCnt;
                }

                if(cellData instanceof String){
                    cell.setCellValue(String.valueOf(cellData));
                } else if (cellData instanceof Integer) {
                    cell.setCellValue(Integer.parseInt(String.valueOf(cellData)));
                } else if (cellData instanceof Double) {
                    cell.setCellValue(Double.parseDouble(String.valueOf(cellData)));
                } else if (cellData instanceof Float) {
                    cell.setCellValue(Float.parseFloat(String.valueOf(cellData)));
                } else{
                    cell.setCellValue(String.valueOf(cellData));
                }

            }
            // End write row data

            workbook.write(new FileOutputStream(file));



        } catch (IOException e) {
            log.error("File 기록 중 오류 발생", e);
        }


    }



    public static void main(String args[]){
    }

}
