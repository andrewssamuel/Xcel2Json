package com.firstwicket.excel;

import com.google.gson.*;
import javassist.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.*;
import java.util.*;
import java.util.stream.*;

/**
 * Created by andrewssamuel on 5/11/16.
 */
class XcelConverterImpl implements XcelConverter {

    InputStream inputStream = null;
    Workbook workbook = null;
    Class<?> clazz = null;
    List<List> mList = new ArrayList<List>();


    public List<List> excelParserByFile(String xcelPath){

        try {

            inputStream = new FileInputStream(new File(xcelPath));

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        return excelParser(inputStream);

    }

    public List<List> excelParser(InputStream inputStream) {



        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        int i = workbook.getNumberOfSheets();
        List<Sheet> sheets = new ArrayList<Sheet>();

        IntStream.range(0, i).parallel().forEach(
                sheet -> sheets.add(workbook.getSheetAt(sheet))
        );


        for (Sheet sheet : sheets) {

            List<Object> list = new ArrayList<Object>();
            ObjectWrapper objectWrapper = null;

            String sheetName = sheet.getSheetName().replace(' ', '_');


            for (Iterator<Row> rowIterator = sheet.rowIterator(); rowIterator.hasNext(); ) {

                Row row = rowIterator.next();

                if (row.getRowNum() == 0) {

                    try {
                        objectWrapper = generateNewClass(row, sheetName);
                    } catch (NotFoundException e) {
                        e.printStackTrace();
                    } catch (CannotCompileException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    } catch (InstantiationException e) {
                        e.printStackTrace();
                    }

                } else {

                    Object obj = null;
                    try {
                        obj = objectWrapper.getaClass().newInstance();
                    } catch (InstantiationException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }

                    for (Iterator<Cell> cellsIterator = row.cellIterator(); cellsIterator.hasNext(); ) {

                        Cell cell = cellsIterator.next();

                        try {
                            populatePOJO(cell, objectWrapper.getHeaderInfo(), objectWrapper.getaClass(), obj);
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        } catch (InstantiationException e) {
                            e.printStackTrace();
                        } catch (NoSuchMethodException e) {
                            e.printStackTrace();
                        } catch (InvocationTargetException e) {
                            e.printStackTrace();
                        }

                    }

                    list.add(obj);

                }

            }

            mList.add(list);
        }

        return mList;

    }

    private Class<?> getClass4CellType(Cell cell) {
        Class c = null;

        switch (cell.getCellType()) {
            case 0:
                c = Integer.class;
                break;
            case 1:
                c = String.class;
                break;
            case 2:
                c = String.class;
                break;
            case 3:
                c = String.class;
                break;
            case 4:
                c = Boolean.class;
                break;
            case 5:
                c = Byte.class;
                break;

        }

        return c;
    }


    private Object populatePOJO(Cell cell, List<String> headerInfo, Class<?> clazz, Object obj) throws IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {


        switch (cell.getCellType()) {

            case 0:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, Double.toString(cell.getNumericCellValue()));
                break;
            case 1:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, cell.getStringCellValue());
                break;
            case 2:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, cell.getCellFormula());
                break;
            case 3:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, "");
                break;
            case 4:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, Boolean.toString(cell.getBooleanCellValue()));
                break;
            case 5:
                clazz.getMethod("set" + headerInfo.get(cell.getColumnIndex()), String.class).invoke(obj, Byte.toString(cell.getErrorCellValue()));
                break;
        }

        return obj;

    }

    private ObjectWrapper generateNewClass(Row row, String sheetName) throws NotFoundException, CannotCompileException, IllegalAccessException, InstantiationException {

        Map<String, Class<?>> header = new HashMap<String, Class<?>>();
        List<String> headerInfo = new ArrayList<String>();
        ObjectWrapper objectWrapper = new ObjectWrapper();


        if (row.getRowNum() == 0) {

            for (Iterator<Cell> cellsIterator = row.cellIterator(); cellsIterator.hasNext(); ) {
                Cell cell = cellsIterator.next();
                String headerValue = toCamelCase(cell.getStringCellValue().replace("'", "").replace(' ', '_'));

                header.put(headerValue, String.class);
                headerInfo.add(headerValue);

            }

            Random rand = new Random();
            Integer randInteger = new Integer(rand.nextInt(50) + 1);

            clazz = PojoGenerator.generate(
                    "com.firstwicket.Excel$" +randInteger.toString() + sheetName, header);


        }


        objectWrapper.setaClass(clazz);
        objectWrapper.setHeader(header);
        objectWrapper.setHeaderInfo(headerInfo);

        return objectWrapper;


    }

    private static String toCamelCase(final String init) {
        if (init == null)
            return null;

        final StringBuilder ret = new StringBuilder(init.length());

        for (final String word : init.split(" ")) {
            if (!word.isEmpty()) {
                ret.append(word.substring(0, 1).toUpperCase());
                ret.append(word.substring(1).toLowerCase());
            }
            if (!(ret.length() == init.length()))
                ret.append(" ");
        }

        return ret.toString();
    }

    public String toJson(List list) {

        Gson gson = new GsonBuilder().disableHtmlEscaping().create();
        return  gson.toJson(list);
    }

}


