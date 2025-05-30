package com.cts.fastexcel.function;

import org.apache.poi.ss.usermodel.DateUtil;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.ReadingOptions;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.stream.Stream;

public class TestReadExcel {
    public static void main(String[] args) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String excelPath = "/Users/a2107739/Downloads/Sample_Test_ReadExcel.xlsx";
        try (InputStream inputStream = Files.newInputStream(Paths.get(excelPath))) {
            ReadableWorkbook wb = new ReadableWorkbook(inputStream, new ReadingOptions(true, false));
            Sheet sheet = wb.getFirstSheet();
            int startRow = 1;
            int batchSize = 10;
            int totalRows = 0;

            try (Stream<Row> rows = sheet.openStream().skip(startRow - 1).limit(batchSize)) {
                rows.forEach(row -> {
                    row.forEach(c -> {
                        switch (c.getType()) {
                            case NUMBER:
                                if (c.getDataFormatId() != null && DateUtil.isADateFormat(c.getDataFormatId(), c.getDataFormatString())) {
                                    System.out.println("Row: " + row.getRowNum() + " Col: " + c.getColumnIndex() + " Value: " + c.asDate().format(formatter));
                                    break;
                                }
                                System.out.println("Row: " + row.getRowNum() + " Col: " + c.getColumnIndex() + " Value: " + c.asNumber().toString());
                                break;
                            case FORMULA:
                                System.out.println("Row: " + row.getRowNum() + " Col: " + c.getColumnIndex() + " Value: " + c.getRawValue());
                                System.out.println("Formula is: " + c.getFormula());
                                break;
                            default:
                                System.out.println("Row: " + row.getRowNum() + " Col: " + c.getColumnIndex() + " Value: " + c.getText());
                                break;
                        }
                    });
                });
            }
        } catch (IOException e) {
            System.out.println(Arrays.toString(e.getStackTrace()));
            throw new RuntimeException(e);
        }
    }
}

