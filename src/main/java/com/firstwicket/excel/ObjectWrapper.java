package com.firstwicket.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

/**
 * Created by andrewssamuel on 5/13/16.
 */
class ObjectWrapper implements Serializable {

    public Class<?> aClass;
    public List<String> headerInfo;
    public Map<String, Class<?>> header;
    public Row headerRow;

    public Row getHeaderRow() {
        return headerRow;
    }

    public void setHeaderRow(Row headerRow) {
        this.headerRow = headerRow;
    }

    public Class<?> getaClass() {
        return aClass;
    }

    public void setaClass(Class<?> aClass) {
        this.aClass = aClass;
    }

    public List<String> getHeaderInfo() {
        return headerInfo;
    }

    public void setHeaderInfo(List<String> headerInfo) {
        this.headerInfo = headerInfo;
    }

    public Map<String, Class<?>> getHeader() {
        return header;
    }

    public void setHeader(Map<String, Class<?>> header) {
        this.header = header;
    }
}
