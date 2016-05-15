package com.firstwicket.excel;

import java.util.*;

/**
 * Created by andrewssamuel on 5/14/16.
 */
public interface XcelConverter {

    /* To parse the excel into java objects */
    List excelParser(String path);

    /* To get the Json String of the lists */
    String toJson(List list);

    /* Instanciate XcelConverter to invoke excelParser*/
    static XcelConverter newInstance(){

        XcelConverter xcelConverter = new XcelConverterImpl();
        return xcelConverter;

    }


}
