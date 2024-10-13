package com.cts.fastexcel.datatype;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlRootElement(name="FE_SQLSheetData", namespace="urn:com.cts:types:FE_SQLSheetData")
@XmlType(name="FE_SQLSheetData", namespace="urn:com.cts:types:FE_SQLSheetData", propOrder={"sheetName", "sqlQuery"})
public class SQLSheetData {
    private String sheetName;
    private String sqlQuery;

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
}
