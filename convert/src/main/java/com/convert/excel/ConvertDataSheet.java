package com.convert.excel;

import com.convert.common.Logger;
import com.convert.common.ReadProperties;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel 파일의 각 라인을 읽어들여 분석 로직에따라 변환한다.
 * 변환된 데이터를 보관한다.
 */
public class ConvertDataSheet {
    private static Logger log = new Logger().getLogger(ConvertDataSheet.class);

    private ArrayList<Object[]> rows;
    private LinkedList<Object[]> convertRows;
    private ArrayList<Object[]> unProcessRows;
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
    # $NOTE_CONVERT_CARNUM_NAME : 차량명 + 차량번호를 적요에서 찾아 입력해준다.(/로 둘을 합친다)
    # $NOTE_CONVERT_OIL : 주유비를 찾아 입력해준다.
    # $NOTE_CONVERT_INS : 보험료를 찾아 입력해준다.
    # $NOTE_CONVERT_PAR : 주차비를 찾아 입력해준다.
    # $NOTE_CONVERT_OIL_INS_PAR : 주유비, 보험료, 주차비 중에 어느하나라도 있으면 합산하여 입력해준다.
     */

    private final String OPT_NONE = "$NONE";
    private final String OPT_NUMBERING = "$NUMBERING";
    private final String OPT_STRING = "$STRING";
    private final String OPT_ADRESS = "$ADDRESS";
    private final String OPT_CUST_NAME = "$CUST_NAME";
    private final String OPT_DRIVER_NAME = "$DRIVER_NAME";
    private final String OPT_DATE_MMDD = "$DATE_MMDD";
    private final String OPT_DATE_YYYYMMDD = "$DATE_YYYYMMDD";
    private final String OPT_NOTE_CONVERT_CARNUM = "$NOTE_CONVERT_CARNUM";
    private final String OPT_NOTE_CONVERT_CARNAME = "$NOTE_CONVERT_CARNAME";
    private final String OPT_NOTE_CONVERT_CARNUM_NAME = "$NOTE_CONVERT_CARNUM_NAME";
    private final String OPT_NOTE_CONVERT_OIL = "$NOTE_CONVERT_OIL";
    private final String OPT_NOTE_CONVERT_INS = "$NOTE_CONVERT_INS";
    private final String OPT_NOTE_CONVERT_PAR = "$NOTE_CONVERT_PAR";
    private final String OPT_NOTE_CONVERT_OIL_INS_PAR = "$NOTE_CONVERT_OIL_INS_PAR_CANCEL";
    private final String OPT_NOTE_CONVERT_WIP="$NOTE_CONVERT_WIP";
    private final String OPT_SUM_VAT="$SUM_VAT";
    private final String OPT_VAT="$VAT";

    /**
     * 생성자 : 엑셀파일을 읽어 메모리에 저장한다.
     * @param workFile
     */
    public ConvertDataSheet(File workFile){
        columnLength  = 0;
        rowLength = 0;

        convertRows = new LinkedList<Object[]>(); // 수정된 부분
        rows = new ArrayList<Object[]>();

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
                Object celObj[] = new Object[columnLength];

                // cols
                for(int j = 0 ; j < columnLength ; j++){
                    if (row == null) {

                    }else {
                        XSSFCell cell = row.getCell(j);
                        switch(cell.getCellType()) {   // 각셀의 데이터값을 가져올때 맞는 데이터형으로 변환한다.
                            case HSSFCell.CELL_TYPE_FORMULA:
                                String strValFormula = cell.getCellFormula();
                                celObj[j] = strValFormula;
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                double numericValue = cell.getNumericCellValue();
                                int intVal = (int) numericValue;
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
                }
                rows.add(celObj);
            }
            rowLength --;
        }catch (IOException e){
            log.error("읽기 파일에 문제가 있습니다. 다음 파일 : " + workFile.getName() , e);
        }
    }

    public int getRowLength(){
        return rowLength;
    }
    public ArrayList<Object[]> getUnprocessedRows() {
        return unProcessRows;
    }
    /**
     * 읽어들인 파일의 전체 Row 를 분석 및 변환 한다.
     */
    public void convertRow(){
        String[] targets = ReadProperties.getProperty("DISTRIBUTE_TARGETS").split(",");
        String[] outputFileName = new String[targets.length];
        String[] stdValue = new String[targets.length];
        String[] columnData = new String[targets.length];
        unProcessRows = new ArrayList<>();

        // 타겟별 설정값 Load
        for (int i = 0; i < targets.length; i++) {
            outputFileName[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_FILE_NAME");
            stdValue[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_STD_VALUE");
            columnData[i] = ReadProperties.getProperty(targets[i] + "_OUTPUT_COLUMN_DATA");
        }
        // read row
        for (int i = 0; i < rowLength; i++) {
            Object[] row = rows.get(i);

            /*
             * 한 줄씩 분석
             *
             * STD_VALUE 의 분류기준에 해당하는 데이터 먼저 찾아낸다.
             */

            // read column
            boolean processed = false; // 각 행의 처리 여부를 나타내는 플래그
            for (int j = 0; j < targets.length; j++) {
                if (equalIndex(columnTitle, row, stdValue[j])) {
                    log.debug("[" + j + "] " + stdValue[j] + " - 분류 - OUTPUT - " + outputFileName[j]);
                    Object[] convertRow = procRow(row, columnData[j], outputFileName[j]);

                    convertRows.add(convertRow);
                    log.debug("처리 완료 => " + convertRow);
                    processed = true;
                    break;
                }

            } // column end
            if (!processed) {
                unProcessRows.add(row);
            }
        } // row end
        writeUnprocessedRowsToFile();
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
    private boolean equalIndex(String[] column, Object[] row, String value) {
        if (value == null) {
            return false;
        }
        String[] strVal = value.split(":");
        String columnName = strVal[0];
        String columnValue = strVal[1];

        for (int i = 0; i < columnLength; i++) {
            if (column[i].equals(columnName)) {
                Object cellValue = row[i];
                if (cellValue != null && String.valueOf(cellValue).equals(columnValue)) {
                    return true;
                }
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
        Object[] procRow = new Object[convertCols.length];  // 크기를 설정하여 배열 초기화
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
                    targetValue = targetValue == null ? "" : targetValue;
                    output = targetValue;
                    break;
                case OPT_ADRESS:
                    output = transAddress(String.valueOf(targetValue));
                    break;
                case OPT_CUST_NAME:
                    output = transCustName(String.valueOf(targetValue));
                    break;
                case OPT_DRIVER_NAME:
                    output = transDriverName(String.valueOf(targetValue));
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
                case OPT_NOTE_CONVERT_CARNUM_NAME :
                    output = findCarName(String.valueOf(targetValue)) + "/" + findCarNum(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_WIP:
                    output = findWIP(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_OIL:
                    output = findOil(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_INS:
                    output = findIns(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_PAR:
                    output = findPark(String.valueOf(targetValue));
                    break;
                case OPT_NOTE_CONVERT_OIL_INS_PAR:
                    output = findOilInsPark(String.valueOf(targetValue));
                    break;
                case OPT_SUM_VAT:
                    output = sumVat(String.valueOf(targetValue));
                    break;
                case OPT_VAT:
                    output = getVat(String.valueOf(targetValue));
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
     * 기능 옵션 : $ADRESS
     * <br>
     * 원본 데이터에서 주소지만 출력한다 ( ']'앞의 시각 , 톨포등 주소가 아닌 문자열 제거)
     * @param data
     *
     * */
    private String transAddress(String data){

        String step1 = data.replace("즉후/탁/", "");
        String step2 = step1.replaceAll("@톨포\\]", "");

        // Step 3: Remove '@'
        String step3 = step2.replaceAll("@", "");

        // Step 4: Remove ']' and preceding characters
        String step4 = step3.replaceAll(".*?\\]", "");
/*

        String tempStr = data;
            String pattern = ".*?]";
            Pattern regex = Pattern.compile(pattern);
            Matcher matcher = regex.matcher(tempStr);

            tempStr = matcher.replaceAll("");

            //점, 쉼표, <>, 띄어쓰기, - 를 제외한 모든 특수기호 삭제
            String pattern2 = "[^\\p{L}\\p{N}\\s,.<>-]+";
            Pattern regex2 = Pattern.compile(pattern2);
            Matcher matcher2 = regex.matcher(tempStr);

            String output = matcher2.replaceAll("");*/
        return step4;
    }

    /**
     * 기능 옵션 : $CUST_NAME
     * <br>
     * 원본 데이터에서 '회사명/이름/직위' 로 되어있는 형태를 구분한다.
     * @param data
     *
     * */
    private String transCustName(String data){
        String tempStr = data;

        String pattern = "(?<=/)[가-힣]+";
        Pattern regex = Pattern.compile(pattern);
        Matcher matcher = regex.matcher(tempStr);

        while (matcher.find()) {
            String name = matcher.group();
            return name;
        }
        return "";
    }
    private String transDriverName(String data){
        // 괄호와 괄호 안의 값을 제거
        String pattern = "\\([^()]*\\)";
        Pattern regex = Pattern.compile(pattern);
        Matcher matcher = regex.matcher(data);
        while (matcher.find()) {
            data = data.replace(matcher.group(), "");
        }

        // 점('.')을 제거
        data = data.replace(".", "");

        return data.trim();
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
        DateFormat dfInput = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        DateFormat df = new SimpleDateFormat(format);

        try {
            Date date = dfInput.parse(dateStr);
            formatString = df.format(date);
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
/*
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
        String[] expectStrs = {"출발", "도착", "반차", "핸들", "주차", "주유", "픽업", "청불", "보험", "경유", "전달", "W", "w", "신청자","종료"};

        for (String expect : expectStrs) {
            int indexOfExpect = tempStr.indexOf(expect);

            while (indexOfExpect >= 0) {
                String tempRm = tempStr.substring(indexOfExpect);

                int nextSlashIndex = tempRm.indexOf("/");
                int nextBraceIndex = tempRm.indexOf("{");
                int endIndex = Math.min(nextSlashIndex >= 0 ? nextSlashIndex : tempRm.length(), nextBraceIndex >= 0 ? nextBraceIndex : tempRm.length());

                tempRm = tempRm.substring(0, endIndex);
                tempStr = tempStr.replaceAll("[^a-zA-Z0-9가-힣\\s]", "");
                tempStr = tempStr.replaceAll(tempRm, "");

                indexOfExpect = tempStr.indexOf(expect, indexOfExpect + 1);
            }
        }
        // WIP와 숫자 제거
*/
/*
        tempStr = tempStr.replaceAll("^/W.*?(/|$)", "");
*//*

        tempStr = tempStr.replaceAll("/", "");
*/

        Pattern carModelPattern = Pattern.compile("[가-힣a-zA-Z0-9]+"); // 차종을 추측할 수 있는 패턴
        Pattern vehicleNumberPattern = Pattern.compile("^[0-9가-힣]*[0-9]+[가-힣]*$"); // 차량 번호 패턴 (전체 문자열이 차량 번호와 일치해야 함)
        Matcher matcher = carModelPattern.matcher(data);

            while (matcher.find()) {
                String carModel = matcher.group();

                Matcher vehicleNumberMatcher = vehicleNumberPattern.matcher(carModel);
                if (!vehicleNumberMatcher.matches()) {
                    return carModel;
                }

        }
        return "";
    }
    private String findOil(String data){
        Pattern pattern = Pattern.compile("(주|보|주차|주유|보험)(\\d+\\.\\d+)");
        Matcher matcher = pattern.matcher(data);
        int intValue = 0;
        if (matcher.find()) {
            String label = matcher.group(1);
            double value = Double.parseDouble(matcher.group(2));

            intValue = (int) (value * 10000); // 값을 10000 곱해서 정수로 변환
        }
        return String.valueOf(intValue);
    }

    private String findIns(String data){
        Pattern pattern = Pattern.compile("(주|보|주차|주유|보험)(\\d+\\.\\d+)");
        Matcher matcher = pattern.matcher(data);
        int intValue = 0;
        if (matcher.find()) {
            String label = matcher.group(1);
            double value = Double.parseDouble(matcher.group(2));

            intValue = (int) (value * 10000); // 값을 10000 곱해서 정수로 변환
        }
        return String.valueOf(intValue);
    }

    private String findPark(String data){
        Pattern pattern = Pattern.compile("(주|보|주차|주유|보험)(\\d+\\.\\d+)");
        Matcher matcher = pattern.matcher(data);
        int intValue = 0;
        if (matcher.find()) {
            String label = matcher.group(1);
            double value = Double.parseDouble(matcher.group(2));

            intValue = (int) (value * 10000); // 값을 10000 곱해서 정수로 변환
        }
        return String.valueOf(intValue);
    }

    private String findOilInsPark(String data){
        String tempStr = data;
        Pattern pattern = Pattern.compile("/(주|보|주차|주유|보험)\\d+\\.\\d|취소비");
        Matcher matcher = pattern.matcher(tempStr);

        StringBuilder output = new StringBuilder();

        while (matcher.find()) {
            output.append(matcher.group()).append("/");
        }

        if (output.length() > 0) {
            output.deleteCharAt(output.length() - 1); // 마지막 '/' 제거
        }

        String result = output.toString();
        if (!result.isEmpty()) {
            return result;
        }
        return "";
    }

    private String sumVat(String data) {
        if (data != null && !data.equalsIgnoreCase("null")) {
            data = data.replace(",", ""); // 쉼표(,) 제거
            int tempStr = Integer.parseInt(data);
            tempStr = (int) (tempStr * 1.1);
            DecimalFormat decimalFormat = new DecimalFormat("#,###");
            String formattedResult = decimalFormat.format(tempStr);

            return formattedResult;
        }
        return "";
    }

    private String getVat(String data) {
        if (data != null && !data.equalsIgnoreCase("null")) {
            data = data.replace(",", ""); // 쉼표(,) 제거
            int tempStr = Integer.parseInt(data);
            tempStr = (int) (tempStr * 0.1);

            DecimalFormat decimalFormat = new DecimalFormat("#,###");
            String formattedResult = decimalFormat.format(tempStr);

            return formattedResult;
        }
        return "";
    }
    private static String findWIP(String data) {
        data = data.replaceAll("\\s", "");
        String pattern = "/[Ww]\\S*?(\\d+)(?:(?:\\s|:|$)|(?=/))";
        Pattern regex = Pattern.compile(pattern);
        Matcher matcher = regex.matcher(data);

        if (matcher.find()) {
            String matchStr = matcher.group(1);
            String extractedNumber = matchStr.replaceAll("[^\\d]", "");
            return extractedNumber;
        }
        return "";
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
            XSSFWorkbook workbook;
            XSSFSheet outputSheet;

            if (file.exists()) {
                // If the file already exists, open it and get the existing workbook and sheet
                FileInputStream fis = new FileInputStream(file);
                workbook = new XSSFWorkbook(fis);
                outputSheet = workbook.getSheetAt(0);
            } else {
                // If the file doesn't exist, create a new workbook and sheet
                workbook = new XSSFWorkbook();
                outputSheet = workbook.createSheet();
            }

            int outFileRowCnt = outputSheet.getPhysicalNumberOfRows();

            // Write titles
            if (outFileRowCnt <= 0) {
                // Write titles to the first row
                XSSFRow titleRow = outputSheet.createRow(0);
                for (int i = 0; i < outColumns.length; i++) {
                    Object cellData = outColumns[i];
                    XSSFCell cell = titleRow.createCell(i);
                    cell.setCellValue(String.valueOf(cellData));
                }
                outFileRowCnt++;
            }

            // Write row data
            XSSFRow writeRow = outputSheet.createRow(outFileRowCnt);
            for (int i = 0; i < row.length; i++) {
                Object cellData = row[i];
                XSSFCell cell = writeRow.createCell(i);

                // Numbering 기능 사용시 값 치환
                if (cellData.equals(OPT_NUMBERING)) {
                    cellData = outFileRowCnt;
                }

                if (cellData instanceof String) {
                    cell.setCellValue(String.valueOf(cellData));
                } else if (cellData instanceof Integer) {
                    cell.setCellValue(Integer.parseInt(String.valueOf(cellData)));
                } else if (cellData instanceof Double) {
                    cell.setCellValue(Double.parseDouble(String.valueOf(cellData)));
                } else if (cellData instanceof Float) {
                    cell.setCellValue(Float.parseFloat(String.valueOf(cellData)));
                } else {
                    cell.setCellValue(String.valueOf(cellData));
                }
            }
            outFileRowCnt++;

            // Write the workbook to the file
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();

        } catch (IOException e) {
            log.error("File 기록 중 오류 발생", e);
        }

    }

    /**
     * 처리되지 않은 엑셀 데이터를 파일로 저장한다.
     */
    private void writeUnprocessedRowsToFile() {
        try {
            String filePath = ReadProperties.getProperty("OUTPUT_EXCEL_PATH");
            String fileName = "미처리_내역.xlsx";
            File file = new File(filePath + "/" + fileName);

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();

            int rowIdx = 0;
            for (int i = 0; i < unProcessRows.size(); i++) {
                Object[] row = unProcessRows.get(i);
                XSSFRow excelRow = sheet.createRow(rowIdx++);
                for (int j = 0; j < columnLength; j++) {
                    XSSFCell cell = excelRow.createCell(j);
                    Object cellData = row[j];
                    if (cellData != null) {
                        cell.setCellValue(String.valueOf(cellData));
                    }
                }
            }
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();

        } catch (IOException e) {
            log.error("Failed to write unprocessed rows to file", e);
        }
    }


}
