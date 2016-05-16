package com.firstwicket.excel;


import org.junit.*;
import java.io.*;

import static org.junit.Assert.assertThat;
import static org.hamcrest.CoreMatchers.instanceOf;


/**
 * Created by andrewssamuel on 5/14/16.
 */

public class XcelConverterTests {


   @Test
    public void verifyInstance(){
       XcelConverter xcelConverter = XcelConverter.newInstance();
       assertThat(xcelConverter, instanceOf(XcelConverter.class));
    }

    @Test
    public void verifyJsonString(){
        XcelConverter xcelConverter = XcelConverter.newInstance();
        Assert.assertEquals(xcelConverter.toJson(xcelConverter.excelParserByFile("Movie.xlsx")), "[[{\"Movie\":\"Money Monster\",\"Release_date\":\"42503.0\"},{\"Movie\":\"Mother's Day\",\"Release_date\":\"42489.0\"}]]");
    }

    @Test
    public void verifyJsonStringByInputStream(){
        XcelConverter xcelConverter = XcelConverter.newInstance();
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(new File("Movie.xlsx"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Assert.assertEquals(xcelConverter.toJson(xcelConverter.excelParser(fileInputStream)), "[[{\"Movie\":\"Money Monster\",\"Release_date\":\"42503.0\"},{\"Movie\":\"Mother's Day\",\"Release_date\":\"42489.0\"}]]");
    }

    @Test
    public void verifyJsonStringForExcelWithMultipleSheets(){
        XcelConverter xcelConverter = XcelConverter.newInstance();
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(new File("Movie_Multiple.xlsx"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        Assert.assertEquals(xcelConverter.toJson(xcelConverter.excelParser(fileInputStream)), "[[{\"Rdate\":\"42489.0\",\"Film\":\"Mother's Day\"}],[{\"Movie\":\"Money Monster\",\"Release_date\":\"42503.0\"}]]");

    }


}
