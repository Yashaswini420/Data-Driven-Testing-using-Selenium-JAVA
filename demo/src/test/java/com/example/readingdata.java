package com.example;


import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readingdata {
//excel--> workbook -->Sheet --> Rows--> cells

    public static void main(String[] args) throws IOException{
        FileInputStream file = new FileInputStream("C:\\Users\\YashaswiniR\\Desktop\\Project Selenium_Playwright\\DataDriven Selenium\\Testdata\\AllTestData.xlsx");

        XSSFWorkbook workbook=new XSSFWorkbook(file);

        XSSFSheet sheet= workbook.getSheet("Sheet1");
        int totalrows= sheet.getLastRowNum();
        int totalcells= sheet.getRow(1).getLastCellNum();

        System.out.println("the number of rows: "+totalrows);
        System.out.println("the number of cell :"+totalcells);

        //extract the data from the sheet 
        //Fetch the data from Row then from Cell
        //Row counts start from :0 & cell count start from 1 but in java cell also count from :0

        for(int row=0; row<=totalrows;row++){
            XSSFRow currentrow=sheet.getRow(row);

            for(int cell=0;cell<totalcells;cell++){

                XSSFCell celldata = currentrow.getCell(cell);
                System.out.print(celldata.toString()+"\t");
            }
            System.out.println();
        }
        workbook.close();
        file.close();
    }
}