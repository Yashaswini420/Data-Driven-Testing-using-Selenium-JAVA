package com.example;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingdata {
    public static void main(String[] args) throws IOException{
    
        FileOutputStream file= new FileOutputStream("C:\\Users\\YashaswiniR\\Desktop\\Project Selenium_Playwright\\DataDriven Selenium\\Testdata\\writingdatafetch.xlsx");

        XSSFWorkbook workbook=new XSSFWorkbook();

        XSSFSheet sheet=workbook.createSheet("Data");

        XSSFRow row1=sheet.createRow(0);
            row1.createCell(0).setCellValue("Company");
            row1.createCell(1).setCellValue("Project");
            row1.createCell(2).setCellValue("name");

        XSSFRow row2=sheet.createRow(1);
            row2.createCell(0).setCellValue("ABC");
            row2.createCell(1).setCellValue("Lama");
            row2.createCell(2).setCellValue("Yashaswini");

        XSSFRow row3=sheet.createRow(2);
            row3.createCell(0).setCellValue("Telse");
            row3.createCell(1).setCellValue("Santa");
            row3.createCell(2).setCellValue("Yaana");

        XSSFRow row4=sheet.createRow(3);
            row4.createCell(0).setCellValue("Motherland");
            row4.createCell(1).setCellValue("Claude");
            row4.createCell(2).setCellValue("Henry");

        XSSFRow row5=sheet.createRow(4);
            row5.createCell(0).setCellValue("Carryland");
            row5.createCell(1).setCellValue("Jumanji");
            row5.createCell(2).setCellValue("Rajni");

        workbook.write(file);

        System.out.println("Data has been created");
        workbook.close();
        file.close();
    }
}
