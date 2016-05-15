package com.firstwicket.excel;

import com.google.gson.*;
import org.junit.*;
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
        Assert.assertEquals(xcelConverter.toJson(xcelConverter.excelParser("Movie.xlsx")), "[[{\"Movie\":\"Money Monster\",\"Release_date\":\"42503.0\"},{\"Movie\":\"Mother's Day\",\"Release_date\":\"42489.0\"}]]");
    }


}
