package com.caglayan.excelgenerator;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Set;
import java.util.TreeMap;

public class ExcelGenerator {
    private static final String EXCEL_FILE = "C:\\OMDB\\myexcel.xlsx";

    public static void main(String[] args) throws IOException {
        generateExcel();
    }

    private static void generateExcel() throws IOException {
        TreeMap<String, ExcelRow> data = new TreeMap<>();
        data.put("1", new ExcelRow(1, "Caglayan", "2021-12-19", 12.3f));
        data.put("2", new ExcelRow(2, "Ali", "2020-12-19", 178.8f));
        data.put("3", new ExcelRow(3, "Veli", "2019-12-19", 23f));
        data.put("4", new ExcelRow(4, "Selami", "2021-10-19", 48.33465f));
        data.put("5", new ExcelRow(5, "Ayşe", "2021-12-15", 767f));
        data.put("6", new ExcelRow(6, "Fatma", "2021-08-25", 1313.48f));

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("MySheet");
        byte[] rgb= DefaultIndexedColorMap.getDefaultRGB(IndexedColors.CORNFLOWER_BLUE.getIndex());
        sheet.setTabColor(new XSSFColor(rgb,null));
        sheet.setColumnWidth(3, 12*256);
        sheet.autoSizeColumn(2);

        XSSFRow row;
        int rowCount = 0;
        String[] header = {"Id", "Isim", "Kayit Tarihi", "Height"};

        sheet.createRow(rowCount++);
        row = sheet.createRow(rowCount++);
        int colCount = 0;
        row.createCell(colCount++);

        for(String string : header){
            Cell cell = row.createCell(colCount++);
            cell.setCellValue(string);
        }

        Set<String> keys = data.keySet();
        for (String key: keys) {
            row = sheet.createRow(rowCount++);
            ExcelRow rowData = data.get(key);
            short cellCount = 0;
            row.createCell(cellCount++);
            Cell cell1 = row.createCell(cellCount++);
            cell1.setCellValue(rowData.getNumber());
            Cell cell2 = row.createCell(cellCount++);
            cell2.setCellValue(rowData.getName());
            Cell cell3 = row.createCell(cellCount++);
            cell3.setCellValue(rowData.getDate());
            Cell cell4 = row.createCell(cellCount++);
            cell4.setCellValue(rowData.getHeight());

        }

        FileOutputStream fos = new FileOutputStream(new File(ExcelGenerator.EXCEL_FILE));
        workbook.write(fos);
        workbook.close();
        fos.close();
    }

    private static class ExcelRow{
        private int number;
        private String name;
        private LocalDate date;
        private float height;

        public ExcelRow(int number, String name, String date, float height) {
            this.number = number;
            this.name = name;
            this.date = LocalDate.parse(date);
            this.height = height;
        }

        public int getNumber() {
            return number;
        }

        public String getName() {
            return name;
        }

        public LocalDate getDate() {
            return date;
        }

        public float getHeight() {
            return height;
        }
    }
}
