package com.zoetis.excelreader.app;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
public class ExcelContentHelper {
    @Value("${app.numberOfPage}")
    private String numberOfPage;

    public void getAllExcelFilesInDirectoryAndRead(String directoryPath) throws Exception {
        File directory = new File(directoryPath);

        System.out.println("==============================================================");
        System.out.println("File Path: " + directoryPath);
        System.out.println("==============================================================");

        // Count of Test Cases
        Integer countOfTestCases = 0;
        // Count Of Passed Tests Cases
        Integer countOfPassed = 0;
        // Count Of Failed Tests Cases
        Integer countOfFailed = 0;
        // Count Of Not implemented Tests Cases
        Integer countOfNotImplemented = 0;
        // Count Of Empty Tests Cases
        Integer countOfEmpty = 0;


        //get all the files from a directory
        File[] fList = directory.listFiles();
        for (File file : fList){
            String fileName = file.getName();

            String fileType = FilenameUtils.getExtension(fileName);

            if(fileType.equals("xls") || fileType.equals("xlsx")){
                Map<String, Integer> caseResult = ReadExcelFiles(directoryPath + fileName);

                if(caseResult.size() < 1){
                    throw new Exception(" There is no excel files in directory specified in application.properties! Please check File Path");
                }

                // Add counters for read file
                countOfTestCases += caseResult.get("Total");
                countOfPassed += caseResult.get("Passed");
                countOfFailed += caseResult.get("Failed");
                countOfNotImplemented += caseResult.get("Not");
                countOfEmpty += caseResult.get("Empty");
            }
        }

        // Create XLS File and add report there
        CreateExcelFile(directoryPath, countOfTestCases, countOfPassed, countOfFailed, countOfNotImplemented, countOfEmpty);
    }

    private Map<String, Integer> ReadExcelFiles(String filePath) throws IOException, FileNotFoundException {
        Map<String, Integer> caseResult = new HashMap<>();

        Integer numberOfPage = 0;
        Integer rowIndex = 0;

        FileInputStream stream = new FileInputStream(new File(filePath));

        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(numberOfPage);

        // Get Column with Test Cases Names
        Integer columnIndexWithLastTestResults = getExcelColumnIndexByRowIndexAndText(numberOfPage, rowIndex, "Result", filePath);
        Map<String, List<String>> caseColumnData = extractExcelContentByColumnIndex(numberOfPage, 1, columnIndexWithLastTestResults, filePath);

        // List of Test Cases Names
        List<String> list = new ArrayList<>();
        for (String data : caseColumnData.get("TestCases")) {
            list.add("Test Case Name: " + data + ";");
        }

        // Count of Test Cases
        Integer countOfTestCases = caseColumnData.get("TestCases").size();

        // Get CountOf Passed/Failed/Not implemented Test cases
        List<String> listOfPassed = new ArrayList<>();
        List<String> listOfFailed = new ArrayList<>();
        List<String> listOfNotImplemented = new ArrayList<>();
        List<String> listOfEmpty = new ArrayList<>();
        for (String data : caseColumnData.get("CasesResults")) {
            if (data.toLowerCase().contains("pas")) {
                listOfPassed.add(data);
            }
            if (data.toLowerCase().contains("fail")) {
                listOfFailed.add(data);
            }
            if (data.toLowerCase().contains("not")) {
                listOfNotImplemented.add(data);
            }
            if (data == "") {
                listOfEmpty.add(data);
            }
        }
        // Count Of Passed Tests Cases
        Integer countOfPassed = listOfPassed.size();

        // Count Of Failed Tests Cases
        Integer countOfFailed = listOfFailed.size();

        // Count Of Not implemented Tests Cases
        Integer countOfNotImplemented = listOfNotImplemented.size();

        // Count Of Empty Tests Cases
        Integer countOfEmpty = listOfNotImplemented.size();

        System.out.println("==============================================================");
        System.out.println("Total number of test cases: " + countOfTestCases);
        System.out.println("Passed: " + countOfPassed);
        System.out.println("Failed: " + countOfFailed);
        System.out.println("Not Implemented: " + countOfNotImplemented);
        System.out.println("Empty: " + countOfEmpty);
        System.out.println("==============================================================");

        caseResult.put("Total", countOfTestCases);
        caseResult.put("Passed", countOfPassed);
        caseResult.put("Failed", countOfFailed);
        caseResult.put("Not", countOfNotImplemented);
        caseResult.put("Empty", countOfEmpty);

        return caseResult;
    }

        private void CreateExcelFile(String filePath, int countOfTestCases, int countOfPassed, int countOfFailed, int countOfNotImplemented, int countOfEmpty) throws IOException, FileNotFoundException{
        String messageForOutPut = "Excel file with report has been generated!";

        String fileName = "Report.xlsx";
        String filePathWithReportName = filePath + fileName;

        XSSFWorkbook reportWorkbook = new XSSFWorkbook();
        XSSFSheet reportSheet = null;

        // Check If Report already exists. Use it. Clear all data.
        if(CheckIfReportExists(filePath, fileName)){
            FileInputStream stream = new FileInputStream(new File(filePathWithReportName));
            reportWorkbook = new XSSFWorkbook(stream);
            reportSheet = reportWorkbook.getSheetAt(Integer.parseInt(numberOfPage));
            Iterator<Row> rowIte =  reportSheet.rowIterator();

            while(rowIte.hasNext()){
                rowIte.next();
                rowIte.remove();
            }

            messageForOutPut = "Excel file with name = " + fileName + " was found. It was updated by new report!";
        }
        // If not exists - create it
        else
        {
            reportSheet = reportWorkbook.createSheet("FirstSheet");
        }

        //Set style for Excel
        CellStyle style = reportWorkbook.createCellStyle();
        Font font = reportWorkbook.createFont();
        font.setFontHeightInPoints((short)11);
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setBoldweight(HSSFFont.COLOR_NORMAL);
        font.setColor(HSSFColor.DARK_BLUE.index);

        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(font);


        reportSheet = reportWorkbook.getSheetAt(Integer.parseInt(numberOfPage));
        XSSFRow rowHead = reportSheet.createRow((short)0);

        XSSFCell cell0 = rowHead.createCell(0);
        cell0.setCellValue("File Directory");
        cell0.setCellStyle(style);

        XSSFCell cell1 = rowHead.createCell(1);
        cell1.setCellValue("Total number of test cases");
        cell1.setCellStyle(style);

        XSSFCell cell2 = rowHead.createCell(2);
        cell2.setCellValue("Passed");
        cell2.setCellStyle(style);

        XSSFCell cell3 = rowHead.createCell(3);
        cell3.setCellValue("Failed");
        cell3.setCellStyle(style);

        XSSFCell cell4 = rowHead.createCell(4);
        cell4.setCellValue("Not Implemented");
        cell4.setCellStyle(style);

        XSSFCell cell5 = rowHead.createCell(5);
        cell5.setCellValue("Empty");
        cell5.setCellStyle(style);

        XSSFRow row = reportSheet.createRow((short)1);
        row.createCell(0).setCellValue(filePath);
        row.createCell(1).setCellValue(countOfTestCases);
        row.createCell(2).setCellValue(countOfPassed);
        row.createCell(3).setCellValue(countOfFailed);
        row.createCell(4).setCellValue(countOfNotImplemented);
        row.createCell(5).setCellValue(countOfEmpty);

        reportSheet.autoSizeColumn(0);
        reportSheet.autoSizeColumn(1);
        reportSheet.autoSizeColumn(2);
        reportSheet.autoSizeColumn(3);
        reportSheet.autoSizeColumn(4);
        reportSheet.autoSizeColumn(5);

        FileOutputStream fileOut = new FileOutputStream(filePathWithReportName);
        reportWorkbook.write(fileOut);

        fileOut.close();

        System.out.println(messageForOutPut);
    }

    private boolean CheckIfReportExists(String directoryPath, String fileName){
        boolean result = false;

        File directory = new File(directoryPath);

        File[] fList = directory.listFiles();
        for (File file : fList) {
            if(file.getName().toLowerCase().equals(fileName.toLowerCase())){
                result = true;
                break;
            }
        }
        return result;
    }

    private Map<String, List<String>>  extractExcelContentByColumnIndex(int sheetIndex, int columnIndex, int testResultColumnIndex, String filePath) {
        Map<String, List<String>> result = new HashMap<>();
        ArrayList<String> columnDataTestCase = null;
        ArrayList<String> columnDataTestResult = null;
        int rowIndex = 0;
        try {
            columnDataTestCase = new ArrayList<>();
            columnDataTestResult = new ArrayList<>();

            FileInputStream stream = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() > 0) { //To filter column headings
                        if (cell.getColumnIndex() == columnIndex) {// To match column index
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    columnDataTestCase.add(cell.getNumericCellValue() + "");
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    String cellValue = cell.getStringCellValue();
                                    if(!cellValue.equals("ANIMAL_01_Create_animal"))
                                    {
                                        rowIndex = cell.getRowIndex();
                                        columnDataTestCase.add(cellValue);
                                        break;
                                    }
                            }
                        }

                        if (cell.getColumnIndex() == testResultColumnIndex && cell.getRowIndex() == rowIndex) {
                            String cellValue = cell.getStringCellValue();
                            if(!cellValue.equals("Passed / Failed ") && !cellValue.equals("Passed / Failed") && !cellValue.equals("Passed/Failed"))
                            {
                                columnDataTestResult.add(cellValue);
                            }
                        }
                    }
                }
            }
            stream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        result.put("TestCases", columnDataTestCase);
        result.put("CasesResults", columnDataTestResult);

        return result;
    }

    private Integer getExcelColumnIndexByRowIndexAndText(int sheetIndex, int rowIndex, String text, String filePath) {
        Integer columnIndex = 11;
        try {
            FileInputStream stream = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

            Row row = sheet.getRow(rowIndex);
            Iterator<Cell> cellIterator = row.cellIterator();
            Cell cell = cellIterator.next();
            if (row.getRowNum() > 0) { //To filter column headings
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        String cellValue = cell.getStringCellValue();
                        if (cellValue.toLowerCase().contains(text.toLowerCase())) {
                            columnIndex = cell.getColumnIndex();
                        }
                        break;
                }
            }
            stream.close();
            System.out.println("");
            System.out.println("");
            System.out.println("Column index with Test Result: " + columnIndex);
            System.out.println("");
        } catch (Exception e) {
            e.printStackTrace();
        }

        return columnIndex;
    }
}
