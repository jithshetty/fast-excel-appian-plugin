package com.cts.fastexcel.datatype;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlRootElement(name="FE_SQLSheetData", namespace="urn:com.cts:types:FE_SQLSheetData")
@XmlType(name="FE_SQLSheetData", namespace="urn:com.cts:types:FE_SQLSheetData", propOrder={"sheetName", "sqlQuery", "rowOffset", "columnOffset"})
public class SQLSheetData {
    private String sheetName;
    private String sqlQuery;
    private int rowOffset;
    private int columnOffset;

    @XmlElement
    public String getSheetName() {
        return sheetName;
    }
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    @XmlElement
    public String getSqlQuery() {
        return sqlQuery;
    }
    public void setSqlQuery(String sqlQuery) {
        this.sqlQuery = sqlQuery;
    }

    @XmlElement
    public int getRowOffset() {
        return rowOffset;
    }
    public void setRowOffset(int rowOffset) {
        this.rowOffset = rowOffset;
    }

    @XmlElement
    public int getColumnOffset() {
        return columnOffset;
    }
    public void setColumnOffset(int columnOffset) {
        this.columnOffset = columnOffset;
    }
}
