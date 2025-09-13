package com.cts.fastexcel.smartservice;

import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;
import com.appiancorp.suiteapi.security.external.SecureCredentialsStore;
import com.cts.fastexcel.datatype.SQLSheetData;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Test;
import org.mockito.Mock;

import javax.naming.Context;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;

public class MySQLToExcelWriterTest {

    @Mock
    private SmartServiceContext smartServiceCtx;

    @Mock
    private ContentService contentService;

    @Mock
    private Context context;

    @Mock
    private SecureCredentialsStore scs;

    private final WriteSQLToExcel writeSQLToExcel = new WriteSQLToExcel(contentService, context);
    private static final Logger LOG = LogManager.getLogger(MySQLToExcelWriterTest.class);

    // update db connection info
    private static final String URL = "jdbc:mysql://localhost:3306/appian?allowPublicKeyRetrieval=true&useSSL=false";
    private static final String USER = "root";
    private static final String PASSWORD = "password";

    @Test
    public void testSQLToExcel() throws Exception {
        SQLSheetData[] exportDataList = new SQLSheetData[2];

        SQLSheetData data1 = new SQLSheetData();
        data1.setSheetName("Application");
        data1.setSqlQuery("select * from table1");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Latest");
        data2.setSqlQuery("SELECT * from table2");


        exportDataList[0] = data1;
        exportDataList[1] = data2;

        Path tempFile = Files.createFile(Paths.get("./ExportData_MySQLTest.xlsx"));
        Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
        String headerColor = "";
        writeSQLToExcel.exportSQLToWorksheet(exportDataList, tempFile, connection, headerColor);
    }

    @Test
    public void testSQLToProtectedExcel() throws Exception {
        SQLSheetData[] exportDataList = new SQLSheetData[2];

        SQLSheetData data1 = new SQLSheetData();
        data1.setSheetName("Application");
        data1.setSqlQuery("select * from table1");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Latest");
        data2.setSqlQuery("SELECT * from table2");


        exportDataList[0] = data1;
        exportDataList[1] = data2;

        Path tempFile = Files.createFile(Paths.get("./ExportProtected_MySQLTest.xlsx"));
        Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
        String headerColor = "d3d3d3";
        String passsword = "!@#$QWER";
        writeSQLToExcel.exportSQLToProtectedWorksheet(exportDataList, tempFile, connection, headerColor, passsword);
    }

}
