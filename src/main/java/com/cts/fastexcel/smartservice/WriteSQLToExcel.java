package com.cts.fastexcel.smartservice;

import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.ContentUploadOutputStream;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.*;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.dhatim.fastexcel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import com.cts.fastexcel.datatype.SQLSheetData;

import javax.naming.Context;
import javax.sql.DataSource;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.sql.*;

@PaletteInfo(paletteCategory = "Fast Excel Tools", palette = "Write SQL to Excel")
@Order({"SQLSheetDataList", "jndiName", "SaveInFolder", "NewDocumentName", "NewDocumentDesc", "DateFormat", "DateTimeFormat", "HeaderColor", "ExcelPassword"})
public class WriteSQLToExcel extends AppianSmartService {
    private static final Logger LOG = LogManager.getLogger(WriteSQLToExcel.class);
    private final ContentService contentService;
    private final Context context;
    private SQLSheetData[] sqlSheetDataList;
    private String jndiName;
    private Long saveInFolder;
    private String newDocumentName;
    private String newDocumentDesc;
    private String dateFormat;
    private String dateTimeFormat;
    private String headerColor;
    private String excelPassword;
    private Long newDocumentCreated;
    private String errorMessage;
    private static final String EXPORT_DATA_INPUT = "ExportDataList";

    public WriteSQLToExcel(ContentService contentService, Context context) {
        super();
        this.contentService = contentService;
        this.context = context;
    }

    @Input(required = Required.ALWAYS)
    @Name("SQLSheetDataList")
    public void setExportDataList(SQLSheetData[] sqlSheetDataList) {
        this.sqlSheetDataList = sqlSheetDataList;
    }

    @Input(required = Required.ALWAYS)
    @Name("jndiName")
    public void setJndiName(String jndiName) {
        this.jndiName = jndiName;
    }

    @Input(required = Required.ALWAYS)
    @Name("SaveInFolder")
    @FolderDataType
    public void setSaveInFolder(Long saveInFolder) {
        this.saveInFolder = saveInFolder;
    }

    @Input(required = Required.ALWAYS)
    @Name("NewDocumentName")
    public void setNewDocumentName(String newDocumentName) {
        this.newDocumentName = newDocumentName;
    }

    @Input(required = Required.OPTIONAL)
    @Name("NewDocumentDesc")
    public void setNewDocumentDesc(String newDocumentDesc) {
        this.newDocumentDesc = newDocumentDesc;
    }

    @Input(required = Required.OPTIONAL)
    @Name("DateFormat")
    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    @Input(required = Required.OPTIONAL)
    @Name("DateTimeFormat")
    public void setDateTimeFormat(String dateTimeFormat) {
        this.dateTimeFormat = dateTimeFormat;
    }

    @Input(required = Required.OPTIONAL)
    @Name("HeaderColor")
    public void setHeaderColor(String headerColor) {
        this.headerColor = headerColor;
    }

    @Input(required = Required.OPTIONAL)
    @Name("ExcelPassword")
    public void setExcelPassword(String excelPassword) {
        this.excelPassword = excelPassword;
    }

    @Name("NewDocumentCreated")
    @DocumentDataType
    public Long getNewDocumentCreated() {
        return newDocumentCreated;
    }

    @Name("ErrorMessage")
    public String getErrorMessage() {
        return errorMessage;
    }

    @Override
    public void validate(MessageContainer messages) {
        if (jndiName == null || jndiName.isEmpty()) {
            messages.addError("jndiName", "jndiName.missing");
        }
        if (newDocumentName == null || newDocumentName.isEmpty()) {
            messages.addError("NewDocumentName", "newDocumentName.missing");
        }
        if (saveInFolder == null) {
            messages.addError("SaveInFolder", "saveInFolder.missing");
        }
        if (sqlSheetDataList == null || sqlSheetDataList.length == 0) {
            messages.addError(EXPORT_DATA_INPUT, "exportDataList.missing");
        }
        if (sqlSheetDataList != null && sqlSheetDataList.length > 20) {
            messages.addError(EXPORT_DATA_INPUT, "exportDataList.maximum");
        }
        //Using separate method to reduce Cognitive Complexity
        validateExportDataList(messages);
    }

    @Override
    public void run() throws SmartServiceException {
        File tempFile = null;
        try {
            tempFile = File.createTempFile("ExportData", ".xlsx");
            LOG.debug("JNDI name is {}", jndiName);
            LOG.debug("Context String Name : {}", context);
            DataSource dataSource = (DataSource) context.lookup(jndiName); // nosemgrep datasourse read from appian const
            try (Connection connection = dataSource.getConnection()) {
                if (excelPassword != null && !excelPassword.isEmpty()) {
                    exportSQLToProtectedWorksheet(sqlSheetDataList, tempFile.toPath(), connection, headerColor, excelPassword);
                } else {
                    exportSQLToWorksheet(sqlSheetDataList, tempFile.toPath(), connection, headerColor);
                }
                createAppianDocument(tempFile);
            }

        } catch (Exception e) {
            errorMessage = e.getMessage();
            LOG.error("Error generating Excel from SQL", e);
        } finally {
            if (tempFile != null) {
                boolean result = tempFile.delete();
                if (result) {
                    LOG.debug("Temp File Deleted Successfully");
                } else {
                    LOG.debug("Failed to delete temp file");
                }
            }
        }
    }

    public void exportSQLToWorksheet(SQLSheetData[] sqlSheetDataList, Path tempFile, Connection connection, String headerColor) throws SQLException, IOException {
        try (OutputStream os = Files.newOutputStream(tempFile);
             Workbook wb = new Workbook(os, "ExportData", "1.0")) {

            processWorksheet(sqlSheetDataList, connection, wb, headerColor);
            wb.finish();
            LOG.info("Successfully generated Excel");
        }
    }

    public void exportSQLToProtectedWorksheet(SQLSheetData[] sqlSheetDataList, Path tempFile, Connection connection, String headerColor, String password) throws SQLException, IOException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream(); POIFSFileSystem fs = new POIFSFileSystem();
             Workbook wb = new Workbook(bos, "ExportData", "1.0")) {

            processWorksheet(sqlSheetDataList, connection, wb, headerColor);
            wb.finish();
            byte[] bytes = bos.toByteArray();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);

            Encryptor enc = info.getEncryptor();
            enc.confirmPassword(password);

            // Read in an existing OOXML file and write to encrypted output stream
            try (OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes)); OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            } catch (GeneralSecurityException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
            // Write out the encrypted version
            try (FileOutputStream fos = new FileOutputStream(String.valueOf(tempFile))) {
                fs.writeFilesystem(fos);
            }
            LOG.info("Successfully generated protected Excel");
        }
    }

    private void processWorksheet(SQLSheetData[] sqlSheetDataList, Connection connection, Workbook wb, String headerColor) throws SQLException, IOException {
        String excelDateFormat = (this.dateFormat == null || this.dateFormat.isEmpty()) ? "dd-mm-yyyy" : this.dateFormat;
        LOG.debug("Set Excel Date format: {}", excelDateFormat);
        String excelDateTimeFormat = (this.dateTimeFormat == null || this.dateTimeFormat.isEmpty()) ? "dd-mm-yyyy hh:mm:ss" : this.dateTimeFormat;
        LOG.debug("Set Excel Timestamp format: {}", excelDateTimeFormat);

        for (SQLSheetData sqlSheetDataItem : sqlSheetDataList) {

            String sheetName = sqlSheetDataItem.getSheetName();
            LOG.debug("Start Processing for Sheet: {}", sheetName);

            try (PreparedStatement preparedStatement = connection.prepareStatement(sqlSheetDataItem.getSqlQuery());
                 ResultSet rs = preparedStatement.executeQuery()) {
                preparedStatement.setFetchSize(100);
                LOG.debug("SQL FetchSize set to 100");

                //Create New Worksheet
                try (Worksheet ws = wb.newWorksheet(sqlSheetDataItem.getSheetName())) {

                    //Get ResultSet Metadata and Column Count
                    ResultSetMetaData rsmd = rs.getMetaData();
                    int colCount = rsmd.getColumnCount();
                    System.out.println("Start Row is: " + sqlSheetDataItem.getRowOffset());
                    processHeaders(ws, colCount, rsmd, headerColor, sqlSheetDataItem.getRowOffset(), sqlSheetDataItem.getColumnOffset());
                    processRows(rs, ws, colCount, rsmd, excelDateFormat, excelDateTimeFormat, sqlSheetDataItem.getRowOffset(), sqlSheetDataItem.getColumnOffset());

                    ws.finish();
                    LOG.debug(" Sheet: {}. Flushed Sheet to File", sheetName);
                    LOG.info("Completed processing for Sheet: {}", sheetName);
                }
            } catch (SQLException e) {
                throw new SQLException("Error processing request for Sheet: " + sheetName + " Error: " + e.getMessage(), e);
            }
        }
    }

    private void processHeaders(Worksheet ws, int colCount, ResultSetMetaData rsmd, String headerColor, int startRow, int colOffset) throws SQLException {
        //Process Headers from ResultSet. Add Stying for Headers in Excel
        for (int i = 0; i < colCount; i++) {
            ws.value(startRow, colOffset + i, rsmd.getColumnName(i + 1));
            if (headerColor != null && !headerColor.isEmpty()) {
                ws.style(startRow, i)
                        .bold()
                        .fillColor(headerColor.replace("#", ""))
                        .borderColor(BorderSide.RIGHT, Color.BLACK)
                        .borderStyle(BorderSide.RIGHT, BorderStyle.THIN)
                        .borderColor(BorderSide.LEFT, Color.BLACK)
                        .borderStyle(BorderSide.LEFT, BorderStyle.THIN)
                        .borderColor(BorderSide.TOP, Color.BLACK)
                        .borderStyle(BorderSide.TOP, BorderStyle.THIN)
                        .borderColor(BorderSide.BOTTOM, Color.BLACK)
                        .borderStyle(BorderSide.BOTTOM, BorderStyle.THIN)
                        .set();
            }
        }
    }

    private void processRows(ResultSet rs, Worksheet ws, int colCount, ResultSetMetaData rsmd, String excelDateFormat, String excelDateTimeFormat, int startRow, int colOffset) throws SQLException, IOException {
        // Process ResultSet Rows based on Column Type
        while (rs.next()) {
            int row = startRow + rs.getRow();
            int col;
            int sheetCol;
            for (int i = 0; i < colCount; i++) {
                col = i + 1;
                sheetCol = colOffset + 1;
                int columnType = rsmd.getColumnType(col);
                switch (columnType) {
                    case Types.BIT:
                    case Types.BOOLEAN:
                        ws.value(row, sheetCol, (rs.getBoolean(col) ? "Yes" : "No"));
                        break;
                    case Types.DATE:
                        ws.value(row, sheetCol, rs.getDate(col));
                        ws.style(row, sheetCol).format(excelDateFormat).set();
                        break;
                    case Types.TIMESTAMP:
                        ws.value(row, sheetCol, rs.getTimestamp(col));
                        ws.style(row, sheetCol).format(excelDateTimeFormat).set();
                        break;
                    case Types.DOUBLE:
                    case Types.FLOAT:
                    case Types.DECIMAL:
                    case Types.REAL:
                        ws.value(row, sheetCol, 123);
                        break;
                    case Types.INTEGER:
                    case Types.BIGINT:
                    case Types.SMALLINT:
                    case Types.TINYINT:
                        ws.value(row, sheetCol, rs.getLong(col));
                        break;
                    default:
                        ws.value(row, sheetCol, rs.getString(col));
                        break;
                }
            }
            if (row % 100 == 0) {
                ws.flush();
                LOG.debug("Flushed Next 100 rows to File");
            }
        }
    }

    public void createAppianDocument(File sourceFile) throws IOException {
        try {
            Document doc = new Document();
            doc.setName(newDocumentName);
            doc.setDescription(newDocumentDesc);
            doc.setExtension("xlsx");
            doc.setParent(saveInFolder);

            //Copy file from temp location to appian document using the ContentUploadOutputStream
            try(FileInputStream inputStream = new FileInputStream(sourceFile);
            ContentUploadOutputStream contentUploadOutputStream = contentService.uploadDocument(doc, ContentConstants.UNIQUE_NONE);) {
                byte[] buffer = new byte[1024];
                int length;
                while ((length = inputStream.read(buffer)) > 0) {
                    contentUploadOutputStream.write(buffer, 0, length);
                }
                newDocumentCreated = contentUploadOutputStream.getContentId();
            }
            LOG.info("Successfully copied file to Appian Folder");
        } catch (Exception e) {
            throw new IOException(e);
        }
    }

    private void validateExportDataList(MessageContainer messages) {
        for (SQLSheetData sqlSheetData : sqlSheetDataList) {
            if (sqlSheetData.getSheetName() == null || sqlSheetData.getSheetName().isEmpty()) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sheetName.missing");
            }
            if (sqlSheetData.getSqlQuery() == null || sqlSheetData.getSqlQuery().isEmpty()) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sqlQuery.missing");
            }
            if (!sqlSheetData.getSqlQuery().toUpperCase().startsWith("SELECT")) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sqlQuery.startSyntax");
            }
            if (sqlSheetData.getSqlQuery().endsWith(";")) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sqlQuery.endSyntax");
            }
            if (sqlSheetData.getRowOffset() <= 0) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sqlQuery.endSyntax");
            }
            if (sqlSheetData.getColumnOffset() <= 0) {
                messages.addError(EXPORT_DATA_INPUT, "exportDataList.sqlQuery.endSyntax");
            }
        }
    }
}
