package com.cts.fastexcel.smartservice;

import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;
import com.appiancorp.suiteapi.security.external.SecureCredentialsStore;
import com.cts.fastexcel.datatype.SQLSheetData;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Ignore;
import org.junit.Test;
import org.mockito.Mock;

import javax.naming.Context;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;

public class SQLToExcelWriterTest {

    @Mock
    private SmartServiceContext smartServiceCtx;

    @Mock
    private ContentService contentService;

    @Mock
    private Context context;

    @Mock
    private SecureCredentialsStore scs;

    private final WriteSQLToExcel writeSQLToExcel = new WriteSQLToExcel(contentService, context);
    private static final Logger LOG = LogManager.getLogger(SQLToExcelWriterTest.class);

    // update db connection info
    private static final String URL = "jdbc:oracle:thin:@//localhost:1521/db";
    private static final String USER = "username";
    private static final String PASSWORD = "password";

    @Test
    public void testSQLToExcel() throws Exception {
        SQLSheetData[] sqlSheetDataList = new SQLSheetData[4];

        SQLSheetData data1 = new SQLSheetData();
        data1.setSheetName("Sheet1");
        data1.setSqlQuery("SELECT * from TABLE");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Sheet2");
        data2.setSqlQuery("SELECT * from TABLE1");

        SQLSheetData data3 = new SQLSheetData();
        data3.setSheetName("Sheet3");
        data3.setSqlQuery("SELECT * from TABLE2");

        SQLSheetData data4 = new SQLSheetData();
        data4.setSheetName("Sheet4");
        data4.setSqlQuery("SELECT * from TABLE3");

        sqlSheetDataList[0] = data1;
        sqlSheetDataList[1] = data2;
        sqlSheetDataList[2] = data3;
        sqlSheetDataList[3] = data4;

        Path tempFile = Files.createFile(Paths.get("./ExportData_Test.xlsx"));
        Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
        String headerColor = "d3d3d3";
        writeSQLToExcel.exportSQLToWorksheet(sqlSheetDataList, tempFile, connection, headerColor);
    }

    @Test
    public void testSQLToProtectedExcel() throws Exception {
        SQLSheetData[] sqlSheetDataList = new SQLSheetData[4];

        SQLSheetData data1 = new SQLSheetData();
        data1.setSheetName("Sheet1");
        data1.setSqlQuery("SELECT * from TABLE");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Sheet2");
        data2.setSqlQuery("SELECT * from TABLE1");

        SQLSheetData data3 = new SQLSheetData();
        data3.setSheetName("Sheet3");
        data3.setSqlQuery("SELECT * from TABLE2");

        SQLSheetData data4 = new SQLSheetData();
        data4.setSheetName("Sheet4");
        data4.setSqlQuery("SELECT * from TABLE3");

        sqlSheetDataList[0] = data1;
        sqlSheetDataList[1] = data2;
        sqlSheetDataList[2] = data3;
        sqlSheetDataList[3] = data4;

        Path tempFile = Files.createFile(Paths.get("./ExportDataProtected_Test.xlsx"));
        Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
        String headerColor = "#FFFFFF";
        String password = "1234qwer";
        writeSQLToExcel.exportSQLToProtectedWorksheet(sqlSheetDataList, tempFile, connection, headerColor, password);
    }
}
