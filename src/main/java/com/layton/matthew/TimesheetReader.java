package com.layton.matthew;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;

public class TimesheetReader {

    public XWPFDocument readTimeSheet(String filePath) {
        try {
            ClassLoader classLoader = getClass().getClassLoader();
            File file = new File(classLoader.getResource(filePath).getFile());
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            fis.close();
            return document;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
