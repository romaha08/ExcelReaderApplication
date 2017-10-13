package com.zoetis.excelreader.app;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.Properties;

@Component
public class Executer {
    @Value("${app.filePath}")
    private String filePath;

    @Autowired
    private ExcelContentHelper excelContentHelper;
    @Autowired
    private ExcelContentHelper2 excelContentHelper2;

    public void start() throws Exception {
        //excelContentHelper.getAllExcelFilesInDirectoryAndRead(filePath);
        excelContentHelper2.getAllExcelFilesInDirectoryAndRead(filePath);
    }
}
