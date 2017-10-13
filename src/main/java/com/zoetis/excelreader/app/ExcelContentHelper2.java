package com.zoetis.excelreader.app;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
class ExcelContentHelper2 {
    @Value("${app.numberOfPage}")
    private String numberOfPage;

    void getAllExcelFilesInDirectoryAndRead(String directoryPath) throws Exception {
        List<TestCase> testCasesWithUnknownStatus = new ArrayList<>();

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
        // Count Of Blocked Tests Cases
        Integer countOfBlocked = 0;
        // Count Of Empty Tests Cases
        Integer countOfEmpty = 0;
        // Count Of Unknown Status Tests Cases
        Integer countOfUnknown = 0;

        Map<String, List<TestCase>> report = new HashMap<>();

        //File directory = new File(directoryPath);
        //File[] fList = directory.listFiles();

        // Get all the files from a directory
        List<File> listOfAllFilesFromDirectoryAndSubFolders = getAllFilesFromDirectoryAndSubdirectories(directoryPath);
        for (File file : listOfAllFilesFromDirectoryAndSubFolders) {
            String fileName = file.getName();
            String fileFullPath = file.getAbsolutePath();

            String fileType = FilenameUtils.getExtension(fileName);

            if (!fileName.contains("Report2017")) {
                if (fileType.equals("xls") || fileType.equals("xlsx")) {

                    System.out.println("---------------------------------------------------------------");
                    System.out.println("File Name: " + fileName);
                    System.out.println("---------------------------------------------------------------");

                    // Get All Test Cases info
                    List<TestCase> allTestCasesFromFile = ReadExcelFile(fileFullPath);

                    // Get Statistics by All Test Cases
                    Map<String, Integer> caseResult = getStatisticByAllTestCasesInOneFile(allTestCasesFromFile);

                    if (caseResult.size() < 1) {
                        throw new Exception(" There is no excel files in directory specified in application.properties! Please check File Path");
                    }

                    // Add counters for read file
                    countOfTestCases += caseResult.get("Total");
                    countOfPassed += caseResult.get("Passed");
                    countOfFailed += caseResult.get("Failed");
                    countOfNotImplemented += caseResult.get("Not");
                    countOfBlocked += caseResult.get("Blocked");
                    countOfEmpty += caseResult.get("Empty");
                    countOfUnknown += caseResult.get("Unknown");

                    report.put(fileName, allTestCasesFromFile);

                    // Get Test Cases with unknown result statuses
                    for (TestCase tCase : allTestCasesFromFile) {
                        String result = tCase.getTestResult();
                        if (result != null && !result.toLowerCase().contains("pas") && !result.toLowerCase().contains("fail") && !result.toLowerCase().contains("not") && !result.toLowerCase().contains("block") && !result.toLowerCase().contains("n/a") && !result.isEmpty()) {
                            testCasesWithUnknownStatus.add(tCase);
                        }
                    }
                }
            }
        }

        // Create XLS File and add report there
        CreateExcelFile(directoryPath, countOfTestCases, countOfPassed, countOfFailed, countOfNotImplemented, countOfBlocked, countOfEmpty, countOfUnknown, report, testCasesWithUnknownStatus);
    }

    private List<TestCase> ReadExcelFile(String filePath) throws Exception {
        Integer numberOfPageWithTestCase = Integer.parseInt(numberOfPage);
        Integer rowIndex = 0;

        // Get Column with Test Cases Names
        Integer columnIndexWithTestCase = getExcelColumnIndexByRowIndexAndText(numberOfPageWithTestCase, rowIndex, "Test Name", filePath);
        Integer columnIndexWithDescription = getExcelColumnIndexByRowIndexAndText(numberOfPageWithTestCase, rowIndex, "Description", filePath);
        Integer columnIndexWithRequirement = getExcelColumnIndexByRowIndexAndText(numberOfPageWithTestCase, rowIndex, "Req", filePath);
        Integer columnIndexWithLastTestResults = getExcelColumnIndexByRowIndexAndText(numberOfPageWithTestCase, rowIndex, "Result", filePath);

        List<TestCase> extractedTestCase = new ArrayList<>();
        // Check if File is Test Case - Read it. Else - ignor
        if (checkFileIsTestCase(filePath, numberOfPageWithTestCase)) {
            // Get List: Cells Values by each Test Case For: Test Name, Description, Requirement, Test Result(last)
            extractedTestCase = extractExcelFileContentByColumnIndexes(numberOfPageWithTestCase, columnIndexWithTestCase, columnIndexWithDescription, columnIndexWithRequirement, columnIndexWithLastTestResults, filePath);
        }
        // TODO : Create new Sheet in Test Case with report

        return extractedTestCase;
    }

    private List<TestCase>  extractExcelFileContentByColumnIndexes(int sheetIndex, int columnIndexWithTestCase, int columnIndexWithDescription, int columnIndexWithRequirement, int columnIndexWithTestResult, String filePath) throws IOException {
        List<TestCase> excelFileAllTestCases = new ArrayList<>();

        int rowIndex = 0;

        try {
            FileInputStream stream = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                TestCase testCase = new TestCase();

                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    if (row.getRowNum() > 0)
                    {
                        if(row.getRowNum() == 1 && cell.getStringCellValue().toLowerCase().equals("example") ){
                            break;
                        }
                        if (cell.getColumnIndex() == columnIndexWithTestCase)
                        {
                            String cellValue = cell.getStringCellValue();
                            if (!cellValue.isEmpty())
                            {
                                rowIndex = cell.getRowIndex();
                            }
                            testCase.setTestName(cellValue);
                        }
                        if (cell.getColumnIndex() == columnIndexWithDescription) {
                            int rowIndexOfCell = cell.getRowIndex();
                            if (rowIndexOfCell == rowIndex) {
                                String cellValue = cell.getStringCellValue();
                                if (!cellValue.isEmpty()) {
                                    testCase.setDescription(cellValue);
                                }
                            }
                        }
                        if (cell.getColumnIndex() == columnIndexWithRequirement) {
                            int rowIndexOfCell = cell.getRowIndex();
                            if (rowIndexOfCell == rowIndex) {
                                String cellValue = cell.getStringCellValue();
                                if (!cellValue.isEmpty()) {
                                    testCase.setRequirement(cellValue);
                                }
                            }
                        }
                        if (cell.getColumnIndex() == columnIndexWithTestResult) {
                            int rowIndexOfCell = cell.getRowIndex();
                            if (rowIndexOfCell == rowIndex) {
                                String cellValue = cell.getStringCellValue();
                                if (!cellValue.isEmpty()) {
                                    testCase.setTestResult(cellValue);
                                }
                            }
                        }
                    }
                }
                if(row.getRowNum() == rowIndex && rowIndex > 0) {
                    excelFileAllTestCases.add(testCase);
                }

            }

            stream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Remove from report Example Test Case
        List<TestCase> tempList = new ArrayList<>();
        for (TestCase tCase: excelFileAllTestCases) {
            if (tCase.getTestResult() != null && tCase.getTestResult().contains("Pas") && tCase.getTestResult().contains("Fail")) {
                tempList.add(tCase);
            }
        }
        if(tempList.size() > 0) {
            excelFileAllTestCases.remove(tempList.get(0));
        }

        return excelFileAllTestCases;
    }

    private Map<String, Integer> getStatisticByAllTestCasesInOneFile(List<TestCase> listOfAllTestCases){
        Map<String, Integer> caseResult = new HashMap<>();

        // Count of Test Cases
        Integer countOfTestCases = listOfAllTestCases.size();

        // Get CountOf Passed/Failed/Not implemented Test cases
        List<String> allResults = new ArrayList<>();
        for (TestCase tCase : listOfAllTestCases) {
            allResults.add(tCase.getTestResult());
        }
        List<String> listOfPassed = new ArrayList<>();
        List<String> listOfFailed = new ArrayList<>();
        List<String> listOfNotImplemented = new ArrayList<>();
        List<String> listOfBlocked = new ArrayList<>();
        List<String> listOfEmpty = new ArrayList<>();
        List<String> listOfUnknownStatus = new ArrayList<>();

        for (String data : allResults) {
            if (data == null) {
                listOfEmpty.add("");
            } else {
                if (data.toLowerCase().contains("pas") && data.toLowerCase().contains("fail")) {
                    countOfTestCases = countOfTestCases - 1;
                }
                if (data.toLowerCase().contains("pas") && !data.toLowerCase().contains("fail")) {
                    listOfPassed.add(data);
                }
                if (data.toLowerCase().contains("fail") && !data.toLowerCase().contains("pas")) {
                    listOfFailed.add(data);
                }
                if (data.toLowerCase().contains("not") || data.toLowerCase().contains("n/a")) {
                    listOfNotImplemented.add(data);
                }
                if (data.toLowerCase().contains("block")) {
                    listOfBlocked.add(data);
                }
                if (!data.toLowerCase().contains("not") && !data.toLowerCase().contains("n/a") && !data.toLowerCase().contains("fail") && !data.toLowerCase().contains("pas")&& !data.toLowerCase().contains("block")) {
                    listOfUnknownStatus.add(data);
                }
                if (data.isEmpty()) {
                    listOfEmpty.add(data);
                }
            }
        }
        // Count Of Passed Tests Cases Results
        Integer countOfPassed = listOfPassed.size();

        // Count Of Failed Tests Cases Results
        Integer countOfFailed = listOfFailed.size();

        // Count Of Not implemented Tests Cases Results
        Integer countOfNotImplemented = listOfNotImplemented.size();

        // Count Of Not implemented Tests Cases Results
        Integer countOfBlocked = listOfBlocked.size();

        // Count Of Empty Tests Cases Results
        Integer countOfEmpty = listOfEmpty.size();

        // Count Of UnKnown type of Tests Cases Results
        Integer countOfUnknown = listOfUnknownStatus.size();

        System.out.println("Total number of test cases: " + countOfTestCases);
        System.out.println("Passed: " + countOfPassed);
        System.out.println("Failed: " + countOfFailed);
        System.out.println("Not Implemented: " + countOfNotImplemented);
        System.out.println("Blocked: " + countOfBlocked);
        System.out.println("Empty: " + countOfEmpty);
        System.out.println("Unknown: " + countOfUnknown);
        System.out.println("==============================================================");

        caseResult.put("Total", countOfTestCases);
        caseResult.put("Passed", countOfPassed);
        caseResult.put("Failed", countOfFailed);
        caseResult.put("Not", countOfNotImplemented);
        caseResult.put("Blocked", countOfNotImplemented);
        caseResult.put("Empty", countOfEmpty);
        caseResult.put("Unknown", countOfUnknown);

        return caseResult;
    }

    private Integer getExcelColumnIndexByRowIndexAndText(int sheetIndex, int rowIndex, String text, String filePath) {
        Integer columnIndex = 11;
        try {
            FileInputStream stream = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

            Row row = sheet.getRow(rowIndex);
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        String cellValue = cell.getStringCellValue();
                        if (text.equals("Description")) {
                            if (cellValue.toLowerCase().equals(text.toLowerCase())) {
                                columnIndex = cell.getColumnIndex();
                                break;
                            }
                        }
                        else {
                            if (cellValue.toLowerCase().contains(text.toLowerCase())) {
                                columnIndex = cell.getColumnIndex();
                                break;
                            }
                        }
                }
            }
            stream.close();
            System.out.println("Column index with text like: " + text + " has index: " + columnIndex);
        }
        catch (Exception e) {
            e.printStackTrace();
        }

        return columnIndex;
    }

    private void CreateExcelFile(String filePath, int countOfTestCases, int countOfPassed, int countOfFailed, int countOfNotImplemented, int countOfBlocked, int countOfEmpty, int countOfUnknown, Map<String, List<TestCase>> report, List<TestCase> listOfTestCasesWithUnknownStatuses) throws IOException {
        String messageForOutPut = "Excel file with report has been generated!";

        DateTime dateTime = DateTime.now();
        String neededData = dateTime.toString().split("\\+")[0];
        String fileName = "Report" + neededData + ".xlsx";
        String filePathWithReportName = filePath + fileName.replaceAll("[:]", "-").replaceFirst("[.]", "-");

        XSSFWorkbook reportWorkbook = new XSSFWorkbook();

        // Statistics Sheet
        reportWorkbook.createSheet("StatisticsSheet");

        //Set style for Excel
        CellStyle style = reportWorkbook.createCellStyle();
        Font font = reportWorkbook.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setBoldweight(HSSFFont.COLOR_NORMAL);
        font.setColor(HSSFColor.DARK_BLUE.index);
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(font);

        XSSFSheet statisticsSheet = reportWorkbook.getSheetAt((short) 0);
        XSSFRow rowHead = statisticsSheet.createRow((short) 0);

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
        cell5.setCellValue("Blocked");
        cell5.setCellStyle(style);

        XSSFCell cell6 = rowHead.createCell(6);
        cell6.setCellValue("Empty");
        cell6.setCellStyle(style);

        XSSFCell cell7 = rowHead.createCell(7);
        cell7.setCellValue("UnKnown");
        cell7.setCellStyle(style);

        XSSFCell cell8 = rowHead.createCell(8);
        cell8.setCellValue("UnKnown Results");
        cell8.setCellStyle(style);

        XSSFRow row = statisticsSheet.createRow((short) 1);

        int tempCounter = 8;
        if(listOfTestCasesWithUnknownStatuses.size() > 0) {
            for (int i = 0; i < listOfTestCasesWithUnknownStatuses.size(); i++) {
                TestCase tCase = listOfTestCasesWithUnknownStatuses.get(i);

                String unknownResult = tCase.getTestResult();

                XSSFRow rowUnknownStatus = statisticsSheet.getRow((short) i + 1);
                if (i > 0) {
                    rowUnknownStatus = statisticsSheet.createRow((short) i + 1);
                }
                XSSFCell tempCell = rowUnknownStatus.createCell(tempCounter);
                tempCell.setCellValue(unknownResult);

                Font font2 = reportWorkbook.createFont();
                font2.setFontHeightInPoints((short) 9);
                font2.setFontName(HSSFFont.FONT_ARIAL);
                font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

                CellStyle style2 = reportWorkbook.createCellStyle();
                style.setFont(font2);
                style2.setFillPattern(CellStyle.NO_FILL);
                tempCell.setCellStyle(style2);
                statisticsSheet.autoSizeColumn(tempCounter);
            }
        }
        style.setFont(font);

        row.createCell(0).setCellValue(filePath);
        row.createCell(1).setCellValue(countOfTestCases);
        row.createCell(2).setCellValue(countOfPassed);
        row.createCell(3).setCellValue(countOfFailed);
        row.createCell(4).setCellValue(countOfNotImplemented);
        row.createCell(5).setCellValue(countOfBlocked);
        row.createCell(6).setCellValue(countOfEmpty);
        row.createCell(7).setCellValue(countOfUnknown);

        statisticsSheet.autoSizeColumn(0);
        statisticsSheet.autoSizeColumn(1);
        statisticsSheet.autoSizeColumn(2);
        statisticsSheet.autoSizeColumn(3);
        statisticsSheet.autoSizeColumn(4);
        statisticsSheet.autoSizeColumn(5);
        statisticsSheet.autoSizeColumn(6);
        statisticsSheet.autoSizeColumn(7);

        // Report Sheet
        reportWorkbook.createSheet("ReportSheet");
        XSSFSheet reportSheet = reportWorkbook.getSheetAt((short) 1);

        // Create Head of Report
        XSSFRow rowHead2 = reportSheet.createRow((short) 0);

        XSSFCell cell00 = rowHead2.createCell(0);
        cell00.setCellValue("File Name");
        cell00.setCellStyle(style);

        XSSFCell cell01 = rowHead2.createCell(1);
        cell01.setCellValue("Test Name");
        cell01.setCellStyle(style);

        XSSFCell cell02 = rowHead2.createCell(2);
        cell02.setCellValue("Description");
        style.setWrapText(true);
        cell02.setCellStyle(style);
        style.setWrapText(false);

        XSSFCell cell03 = rowHead2.createCell(3);
        cell03.setCellValue("Req #");
        cell03.setCellStyle(style);

        XSSFCell cell04 = rowHead2.createCell(4);
        cell04.setCellValue("Result");
        cell04.setCellStyle(style);

        int rowIterator = 0;
        for (String repKey : report.keySet()) {
            List<TestCase> allTestCasesFromFile = report.get(repKey);

            for (int j = 0; j < allTestCasesFromFile.size(); j++) {
                TestCase tCase = allTestCasesFromFile.get(j);

                XSSFRow tempRaw = reportSheet.createRow((short) rowIterator + 1);
                //if(j < 1) {
                    //tempRaw.createCell(0).setCellValue(repKey);
                //}
                tempRaw.createCell(0).setCellValue(repKey);
                tempRaw.createCell(1).setCellValue(tCase.getTestName());
                tempRaw.createCell(2).setCellValue(tCase.getDescription());
                tempRaw.createCell(3).setCellValue(tCase.getRequirement());
                tempRaw.createCell(4).setCellValue(tCase.getTestResult());
                rowIterator ++;
            }
        }

        // Do auto size for all columns
        reportSheet.autoSizeColumn(0);
        reportSheet.autoSizeColumn(1);
        reportSheet.autoSizeColumn(2);
        reportSheet.autoSizeColumn(3);
        reportSheet.autoSizeColumn(4);

        // Save changes
        FileOutputStream fileOut = new FileOutputStream(filePathWithReportName);
        reportWorkbook.write(fileOut);

        fileOut.close();

        System.out.println(messageForOutPut);
    }

    private boolean CheckIfReportExists(String directoryPath, String fileName){
        boolean result = false;

        File directory = new File(directoryPath);

        File[] fList = directory.listFiles();
        assert fList != null;
        for (File file : fList) {
            if(file.getName().toLowerCase().equals(fileName.toLowerCase())){
                result = true;
                break;
            }
        }
        return result;
    }

    private boolean checkFileIsTestCase(String filePath, int sheetIndex) throws Exception {
        Boolean result = true;

        FileInputStream stream = new FileInputStream(new File(filePath));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

        XSSFRow firstRow = sheet.getRow(0);

        // Check 3 Necessary fields
        List<String> neededCellValues = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            XSSFCell tempCell = firstRow.getCell(i);
            if (tempCell != null && tempCell.getCellType() == Cell.CELL_TYPE_STRING) {
                String value = tempCell.getStringCellValue();
                if (value.toLowerCase().equals("test name") || value.toLowerCase().equals("step id") || value.toLowerCase().equals("description")) {
                    neededCellValues.add(value);
                }
            }
        }

        // Check size must be = 3
        if(neededCellValues.size() != 3){
            result = false;
        }

        return result;
    }

    private List<File> getAllFilesFromDirectoryAndSubdirectories(String directoryPath) throws Exception {
        File directory = new File(directoryPath);

        List<File> tempListList = new ArrayList<>();
        List<File> resultList = new ArrayList<>();

        // get all the files from a directory
        File[] fList = directory.listFiles();
        tempListList.addAll(Arrays.asList(fList));
        for (File file : fList) {
            if (file.isFile()) {

            } else if (file.isDirectory()) {
                tempListList.addAll(getAllFilesFromDirectoryAndSubdirectories(file.getAbsolutePath()));
            }
        }
        for (File tempFile : tempListList) {
            if (tempFile.isFile()) {
                resultList.add(tempFile);
            }
        }

        return resultList;
    }
}
