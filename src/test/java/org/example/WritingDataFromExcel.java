package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WritingDataFromExcel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheetNew");

        XSSFRow row1= sheet.createRow(0);
        row1.createCell(0).setCellValue("Java");
        row1.createCell(1).setCellValue("19");
        row1.createCell(2).setCellValue("Automation");

        XSSFRow row2= sheet.createRow(1);
        row1.createCell(0).setCellValue("python");
        row1.createCell(1).setCellValue(3);
        row1.createCell(2).setCellValue("Automation");

        XSSFRow row3= sheet.createRow(2);
        row1.createCell(0).setCellValue("C#");
        row1.createCell(1).setCellValue("5");
        row1.createCell(2).setCellValue("Automation");

        FileOutputStream file=new FileOutputStream("//Users//krithikj//IdeaProjects//seleniumwebdriver//testdata//DummyData2.xlsx");
        workbook.write(file);
        workbook.close();
        file.close();
        System.out.println("File is created");
    }
}