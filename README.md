# Xcel2Json
Convert Any Excel into Json

This is a tool that converts Excel files into JSON Stirng/Java Objects.

After you have included this Jar file as a project dependencies, Please call the API as below.

  XcelConverter xcelConverter = XcelConverter.newInstance();
  String Json = xcelConverter.toJson(xcelConverter.excelParser("Movie.xlsx")); 
