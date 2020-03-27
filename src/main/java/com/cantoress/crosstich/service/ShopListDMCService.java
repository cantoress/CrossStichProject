package com.cantoress.crosstich.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

@Service
public class ShopListDMCService {

    private static final String PATH_TO_FILE = "D:/Downloads/Threads.xlsx";
    private static final String PATH_TO_FILE_TO_WRITE = "D:/Downloads/ThreadsToBuy.xlsx";

    private static void countThreadsToBuy() {

        try {
            FileInputStream excelFile = new FileInputStream(new File(PATH_TO_FILE));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet wholeDMCListSheet = workbook.getSheet("DMC pallette");
            Sheet myDMCListSheet = workbook.getSheet("DMC threads");

            Set<Integer> wholeDMCList = new TreeSet<>();
            Map<Integer, Double> myDMCMap = new HashMap<>();

            for (Row currentRow : wholeDMCListSheet) {
                if (currentRow.getCell(0) != null) {
                    wholeDMCList.add((int) currentRow.getCell(0).getNumericCellValue());
                }
            }

            for (Row currentRow : myDMCListSheet) {
                if (currentRow.getCell(0) != null && currentRow.getRowNum() != 0) {
                    myDMCMap.merge((int) currentRow.getCell(1).getNumericCellValue(),
                            currentRow.getCell(2).getNumericCellValue(),
                            Double::sum);
                }
            }

            List<Integer> absenceList = new ArrayList<>();
            List<Integer> halfEmptyList = new ArrayList<>();
            for (Integer threadNumber : wholeDMCList) {
                if (myDMCMap.containsKey(threadNumber)) {
                    if (myDMCMap.get(threadNumber) < 50) {
                        halfEmptyList.add(threadNumber);
                    }
                } else {
                    absenceList.add(threadNumber);
                }
            }

            XSSFWorkbook workbookToWrite = new XSSFWorkbook();
            int rowNum = 0;
            Sheet sheet = workbookToWrite.createSheet("Need to buy");
            for (Integer threadNumber : absenceList) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(threadNumber);
            }

            rowNum = 0;
            sheet = workbookToWrite.createSheet("Probably need to buy");
            for (Integer threadNumber : halfEmptyList) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(threadNumber);
            }

            FileOutputStream outputStream = new FileOutputStream(PATH_TO_FILE_TO_WRITE);
            workbookToWrite.write(outputStream);
            workbookToWrite.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void main(String[] args) {
        countThreadsToBuy();
    }

}

