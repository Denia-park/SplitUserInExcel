package org.example;

import com.opencsv.CSVWriter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.NumberFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Scanner;

public class Main {
    static final String PROGRAM_VERSION = "Version : 1.0 , UpdateDate : 22년 7월 10일";
    static boolean FILE_START_FLAG = true; // Title을 저장하기 위해서 사용한 변수
    static String fileNameCSV = getStringOfNowLocalDateTime();
    static final int RECEIVER_PHONE_NUMBER_CELL_INDEX = 7; //H [0부터 시작임.]

    public static void main(String[] args) {
        System.out.println("유저 나누기 프로그램을 시작합니다. [ " + PROGRAM_VERSION + " ]");

        Scanner sc = new Scanner(System.in); // 사용자로부터 데이터를 받기 위한 Scanner

        String path = System.getProperty("user.dir") + "\\"; //현재 작업 경로
        String fileName = "clothesOptionList.xlsx"; //파일명 설정

        XSSFSheet sheetDataFromExcel = readExcel(path, fileName); //엑셀 파일 Read
        if (sheetDataFromExcel == null) { //파일을 못 읽어오면 종료.
            System.out.println("파일을 찾지 못했으므로 프로그램을 종료 합니다.");

            System.out.println("Enter 를 치면 정상 종료됩니다.");
            sc.nextLine(); //프로그램 종료 전 Holding
            return; //프로그램 종료
        }

        //행 갯수 가져오기
        int rows = sheetDataFromExcel.getPhysicalNumberOfRows();

        XSSFRow row = sheetDataFromExcel.getRow(0); //Title Row 가져오기
        int cells = row.getPhysicalNumberOfCells(); //Title Cell 수 가져오기
        String[][] dataBufferArr = new String[2][cells]; //행을 읽어서 저장해둘 배열을 생성
        int currentSaveOrder = 0; //dataBufferArr 에서 몇번째 배열인지 알려줄 인자
        NumberFormat f = NumberFormat.getInstance(); //엑셀에서 NumberFormat이 나왔을때 저장할 수 있게 생성함
        f.setGroupingUsed(false);	//지수로 안나오게 설정

        //반드시 "행(row)"을 읽고 "열(cell)"을 읽어야함 ..
        //rowIndex = 0 => Title
        for(int rowIndex = 1 ; rowIndex < rows ; rowIndex++) {
            row = sheetDataFromExcel.getRow(rowIndex);

            for (int i = 0; i < cells; i++) {
                XSSFCell cell = row.getCell(i);
                dataBufferArr[currentSaveOrder][i] = readCell(cell,f);
            }

            //첫번째 행은 비교할게 없으므로 넘어간다.
            if(rowIndex == 1){
                if(dataBufferArr[currentSaveOrder][RECEIVER_PHONE_NUMBER_CELL_INDEX] == null){
                    throw new RuntimeException("휴대폰 번호 가 Null 입니다. ");
                }
            }
            //첫번째 행을 제외한 모든 행을 CSV로 저장
            else{
                writeDataToCSV(path, dataBufferArr, currentSaveOrder);
            }

            currentSaveOrder ^= 1; //dataBufferArr 에 저장할 순서 변경 0 -> 1 , 1 -> 0 : XOR을 사용했다.
        }

        //마지막 행 이므로 다시 한번 저장을 해줘야함.  그대로 csv에 저장하기.
        writeDataToCSV(path, dataBufferArr, currentSaveOrder);

        System.out.println("작업이 완료되었습니다.");

        System.out.println("Enter 를 치면 정상 종료됩니다.");
        sc.nextLine(); //프로그램 종료 전 Holding
    }

    private static String readCell(XSSFCell cell, NumberFormat f) {
        String tempValue = "zzz";
        if(cell != null){
            //타입 체크
            switch(cell.getCellType()) {
                case STRING:
                    tempValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    tempValue = f.format(cell.getNumericCellValue())+"";
                    break;
                case BLANK:
                    tempValue = "";
                    break;
                case ERROR:
                    tempValue = cell.getErrorCellValue()+"";
                    break;
            }
            return tempValue;
        }
        else
            throw new RuntimeException("Cell Read 중 NPE 발생함");
    }

    private static void writeDataToCSV(String path, String[][] dataBuffer, int currentSaveOrder) {
        File file = new File(path, fileNameCSV);
        try (
                FileOutputStream fos = new FileOutputStream(file,true);
                OutputStreamWriter osw = new OutputStreamWriter(fos, StandardCharsets.UTF_8);
                CSVWriter writer = new CSVWriter(osw)
        ) {
            if(FILE_START_FLAG){
                String[] title = {
                        "주문일시",
                        "상품명",
                        "옵션",
                        "수량",
                        "주문자 이름",
                        "주문자 연락처",
                        "수령자 이름",
                        "수령자 연락처"
                };
                writer.writeNext(title,false);
                FILE_START_FLAG = false;
            }

            int previousSaveOrder = currentSaveOrder ^ 1; //이전 값을 CSV에 저장한다
            String[] writeData = dataBuffer[previousSaveOrder];
            writer.writeNext(writeData,false);

            //휴대폰 번호가 다르면 저장할 행 사이에 null 값 넣기
            //후대폰 번호가 같으면 그냥 넣기
            if(isNotEqualWithPreviousPhoneNumber(dataBuffer, currentSaveOrder)){
                writer.writeNext(new String[] {null},false);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isNotEqualWithPreviousPhoneNumber(String[][] dataBuffer, int currentSaveOrder){
        int previousSaveOrder = currentSaveOrder ^ 1;

        if(dataBuffer[currentSaveOrder][RECEIVER_PHONE_NUMBER_CELL_INDEX] == null){
            throw new RuntimeException("휴대폰 번호 가 Null 입니다. ");
        }

        return !dataBuffer[currentSaveOrder][RECEIVER_PHONE_NUMBER_CELL_INDEX].equals(dataBuffer[previousSaveOrder][RECEIVER_PHONE_NUMBER_CELL_INDEX]);
    }

    public static XSSFSheet readExcel(String path, String fileName){
        try {
            FileInputStream file = new FileInputStream(path + fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            return workbook.getSheetAt(0); // 첫번째 시트만 사용
        } catch(IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static String getStringOfNowLocalDateTime() {
        // 현재 날짜/시간
        LocalDateTime now = LocalDateTime.now(); // 2021-06-17T06:43:21.419878100

        // 포맷팅
        String formatedNow = now.format(DateTimeFormatter.ofPattern("yyMMdd_HH_mm_ss")); // 220628_02_38_02

        return "Split User CSV_" + formatedNow + ".csv"; //Ex) CSV_220628_02_38_02

    }
}