# Xcel2Json
Convert Any Excel into Json

This is a tool that converts Excel files into JSON Stirng/Java Objects.

Maven Dependencies

<dependency>
    <groupId>com.firstwicket</groupId>
    <artifactId>Xcel2Json</artifactId>
    <version>1.0.1</version>
</dependency>

After you have included this Jar file as a project dependencies, Please call the API as below.

  XcelConverter xcelConverter = XcelConverter.newInstance();
  String Json = xcelConverter.toJson(xcelConverter.excelParserByFile("Movie.xlsx"));


Note: Please make sure that excel sheet header row should not contains any special characters. Spaces and Underscores are allowed.
