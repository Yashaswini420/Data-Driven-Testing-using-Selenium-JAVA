package com.example;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingdynamicdata {
    public static void main(String[] args) throws IOException{

        FileOutputStream file=new FileOutputStream("C:\\Users\\YashaswiniR\\Desktop\\Project Selenium_Playwright\\DataDriven Selenium\\Testdata\\Dynamic.xlsx");

        XSSFWorkbook workbook=new XSSFWorkbook();

        XSSFSheet sheet= workbook.createSheet("Data");

        Scanner sc=new Scanner(System.in);

        System.out.println("How many rows needed");
        int numofrow=sc.nextInt();

        System.out.println("How many cells needed");
        int numofcell=sc.nextInt();

        for(int r=0; r <=numofrow; r++){
             XSSFRow currentRow=sheet.createRow(r);

            for(int c=0; c<numofcell ; c++){
                XSSFCell currentCell=currentRow.createCell(c);
                //in this current cell user need to give the data
                
                currentCell.setCellValue(sc.next());

            }
        }
        // If u want to setvalue to random cell or random row

        //XSSFRow row=sheet.createRow(indexvalue:3)
        //XSSFCell cell =row.createCell(indexvalue:5)
        //cell.setValue("The value")

        workbook.write(file);
        workbook.close();
        file.close();
    }

    
}
