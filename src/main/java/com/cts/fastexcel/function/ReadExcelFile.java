package com.cts.fastexcel.function;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;

import com.appiancorp.ps.plugins.typetransformer.AppianList;
import com.appiancorp.ps.plugins.typetransformer.AppianObject;
import com.appiancorp.ps.plugins.typetransformer.AppianTypeFactory;
import com.appiancorp.suiteapi.common.paging.DataSubset;
import com.appiancorp.suiteapi.common.paging.PagingInfo;
import com.appiancorp.suiteapi.common.paging.TypedValueDataSubset;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.exceptions.InvalidContentException;
import com.appiancorp.suiteapi.expression.annotations.Category;
import com.appiancorp.suiteapi.expression.annotations.Function;
import com.appiancorp.suiteapi.expression.annotations.Parameter;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.type.*;
import com.appiancorp.type.AppianTypeLong;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.dhatim.fastexcel.reader.*;

@Category("fastExcelFunctions")
public class ReadExcelFile {
    private static final Logger LOG = LogManager.getLogger(ReadExcelFile.class);

    @Function
    @Type(namespace = Type.APPIAN_NAMESPACE, name = DataSubset.LOCAL_PART)
    public static DataSubset readUsingFastExcel(
            TypeService ts,
            ContentService cs,
            @Parameter(required = true) @DocumentDataType Long excelDocument,
            @Parameter(required = true) int sheetNumber,
            @Parameter(required = true) PagingInfo pagingInfo,
            @Parameter(required = false) String password
    ) throws InvalidContentException {
        List<TypedValue> identifiers = new ArrayList<>();
        List<TypedValue> result = new ArrayList<>();
        AppianTypeFactory tf = AppianTypeFactory.newInstance(ts);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        //DataFormatter formatter = new DataFormatter();

        List<List<String>> resultRows = new ArrayList<>();
        String excelPath = cs.getInternalFilename(excelDocument);

        int startRow = pagingInfo.getStartIndex();
        int batchSize = pagingInfo.getBatchSize();

        //String excelPath = "/Users/a2107739/Downloads/Disbursement25-03-06_1740.xlsx";
        try (InputStream inputStream = Files.newInputStream(Paths.get(excelPath))) {
            ReadableWorkbook wb = new ReadableWorkbook(inputStream, new ReadingOptions(true, false));
            Sheet sheet = wb.getFirstSheet();

            int totalRows = 0;

            try (Stream<Row> rows = sheet.openStream().skip(startRow - 1).limit(batchSize)) {
                rows.forEach(r -> {
                    resultRows.add(getRowData(r, formatter));
                    identifiers.add(tf.toTypedValue(tf.createLong((long) r.getRowNum() + 1)));
                });
            }
            for (List<String> row : resultRows) {
                result.add(toTypedValue(row, tf));
                totalRows++;
            }
            return new TypedValueDataSubset(pagingInfo, totalRows, result, identifiers);
        } catch (IOException e) {
            System.out.println(Arrays.toString(e.getStackTrace()));
            throw new RuntimeException(e);
        }
    }


    private static List<String> getRowData(Row row, DateTimeFormatter formatter) {
        List<String> rowData = new ArrayList<>();
        row.forEach(c -> {
            switch (c.getType()) {
                case NUMBER:
                    if (c.getDataFormatId() != null && DateUtil.isADateFormat(c.getDataFormatId(), c.getDataFormatString())) {
                        rowData.add(c.asDate().format(formatter));
                        break;
                    }
                    rowData.add(c.asNumber().toString());
                    System.out.println("Cell Num value is: " + c.asNumber());
                    break;
                case FORMULA:
                    rowData.add(c.asNumber().toString());
                    System.out.println("Formula is: " + c.getFormula());
                    System.out.println("Formula value is: " + c.asString());
                    break;
                default:
                    rowData.add(c.getText());
                    break;
            }
        });
        return rowData;
    }

    private static TypedValue toTypedValue(List<String> row, AppianTypeFactory tf) {

        AppianObject rowResult = (AppianObject) tf.createElement(AppianType.DICTIONARY);
        AppianList values = tf.createList(AppianType.LIST_OF_STRING);

        for (String item : row) {
            values.add(tf.createString(item));
        }

        rowResult.put("values", values);
        return tf.toTypedValue(rowResult);
    }
}
