package com.zoetis.excelreader.app;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class PropertyReader {
    public static Properties load(String file) throws AssertionError{
        Properties props = new Properties();
        try {
            props.load(new FileInputStream(file));
        } catch (Exception e) {
            try {
                ClassLoader loader = Thread.currentThread().getContextClassLoader();
                InputStream resourceStream = loader.getResourceAsStream(file);
                props.load(resourceStream);

            } catch (FileNotFoundException e1) {
                throw new AssertionError("File with locator's information not found: " + e.toString());
            } catch (IOException e1) {
                throw new AssertionError("IO error while trying to reach locator's information file: " + e.toString());
            }
        }
        return props;
    }
}
