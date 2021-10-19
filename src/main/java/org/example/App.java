package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class App {

    private static final String FILE_LOCATION = "G:\\My Drive\\Info\\CityGraph.xlsx";
    private static final String FILE_LOCATION2 = "howtodoinjava_demo.xlsx";
    private static List<int[]> elementsToModify =  new ArrayList<>();

    public static void main(String[] args) throws IOException {
        readExcelAndWriteElementsToModify(FILE_LOCATION, 0);
        System.out.println(elementsToModify);
        rewriteElementsAccordingToElementsToModifyList(FILE_LOCATION, 0);
    }

    private static void readExcel(String fileLocation, int sheetIndex) {
        try {
            FileInputStream file = new FileInputStream(fileLocation);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

            for (int rn = 0; rn <= sheet.getLastRowNum(); rn++) {
                Row row = sheet.getRow(rn);
                for (int cn = 0; cn < row.getLastCellNum(); cn++) {
                    Cell cell = row.getCell(cn);
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:

                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void readExcelAndWriteElementsToModify(String fileLocation, int sheetIndex) {
        try {
            FileInputStream file = new FileInputStream(fileLocation);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            elementsToModify = new ArrayList<>();
            for (int rn = 0; rn <= sheet.getLastRowNum(); rn++) {
                Row row = sheet.getRow(rn);
                for (int cn = 0; cn < row.getLastCellNum(); cn++) {
                    Cell cell = row.getCell(cn);
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            double numericCellValue = cell.getNumericCellValue();
                            if (numericCellValue > 500) {
                                elementsToModify.add(new int[]{rn, cn});
                            }
                            System.out.print(numericCellValue + "\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void rewriteElementsAccordingToElementsToModifyList(String fileLocation, int sheetIndex) {
        try {
            FileInputStream fileInputStream = new FileInputStream(fileLocation);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for(int[] address : elementsToModify) {
                int rowNum = address[0];
                int colNum = address[1];
                sheet.getRow(rowNum).getCell(colNum).setCellValue(0);
            }
            fileInputStream.close();
            FileOutputStream fileOutputStream = new FileOutputStream(fileLocation);
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}

