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

    private static final String URL = "jdbc:oracle:thin:@//localhost:1521/OSGDEV";
    private static final String USER = "appianbiz";
    private static final String PASSWORD = "aPPian01_";

    @Test
    public void testSQLToExcel() throws Exception {
        SQLSheetData[] sqlSheetDataList = new SQLSheetData[4];

        SQLSheetData data1 = new SQLSheetData();
        data1.setSheetName("Application");
        data1.setSqlQuery("SELECT REFERENCE_ID  AS  \"Reference ID\",APPLICANT_TYPE  AS \"Applicant Type\",AGENCY AS  \"Agency\",GRANT_SCHEME  AS \"Grant Scheme\",PROJ_TYPE_NAME AS \"Project Type\",IS_AUTO_DISBURSEMENT AS \"Auto Disbursement\",IS_FINAL_CLAIM AS \"Final Claim\",TRANCHE_NUMBER AS \"Tranche Number\",FIRST_SUBMISSION_DT AS \"Submission Date\",CURRENT_DISBURSEMENT_STATUS As \"Current Submission Status\",APPROVED_GRANT_AMOUNT AS \"Approved Grant Amount (S$)\",DISBURSED_AMOUNT AS \"Amount Disbursed To-Date (S$)\",RECOVERED_AMOUNT_TO_DATE AS \"Amount Recovered To-Date (S$)\",BALANCE_AMOUNT AS \"Balance Grant Amount (S$)\",TRANCHE_AMOUNT AS \"Tranche Amount (S$)\",CURRENT_TRANCHE_RECOVERED_AMOUNT AS \"Recovered Amount (S$)\",OWNER AS \"Owner\",DEPARTMENT AS \"Owner Department (if any)\",OWNER_ASSESSMENT_REMARK AS \"Owner Assessment Remark\",CERTIFYING_OFFICER AS \"Certifying Officer\",DCO_ASSESSMENT_REMARK AS \"Certifying Officer Remarks\",APPROVING_OFFICER AS \"Approver\",DAO_ASSESSMENT_REMARK AS \"Approver Assessment Remark\",APPROVED_DT AS \"Approved Date\",PC_SEND_OUT_DT AS \"Payment Confirmation Send Out Date\",SALUTATION AS \"Salutation\",APPLICANT_NAME AS \"Applicant Name\",NRIC AS \"NRIC\",RESIDENT_STATUS AS \"Resident Status\",RACE AS \"Race\",GENDER As \"Gender\",DOB AS \"Date of Birth\",APPLICANT_AGE AS \"Applicant Age\",CONTACT_NUMBER AS \"Contact Number\",EMAIL AS \"Email\",RESIDENTIAL_ADDRESS AS \"Residential Address\",OCCUPATION AS \"Occupation\",DESIGNATION AS \"Designation\",ORGANISATION AS \"Organisation\",ORGANISATION_ADDRESS AS \"Organisation Address\",ORG_CONTACT_NUMBER AS \"Organisation Contact Number\",UEN AS \"UEN\",ORGANISATION_TYPE AS \"Organisation Type\",REGISTERED_AS AS \"Registered As\",REGISTRATION_DATE AS \"Registration Date\",CHARITY_CHECK AS \"Registered Charity\",RAISE_MEMBER_CHECK \"Raise Member\",IPC_CHECK AS \"Institution of Public Character (IPC)\",IPC_EXPIRY_DATE AS \"IPC Expiry Date\",ORGANISATION_WEBSITE AS \"Organisation Website\",ORGANISATION_PAST_WORK AS \"Past Achievements/Projects/Work\",ORGANISATION_DESCRIPTION AS \"Organisation Vision/Description\",ADDRESSEE_SALUTATION AS \"LOA Addressee Salutation\",ADDRESSEE_NAME AS \"LOA Addressee Name\",ADDRESSEE_DESIGNATION AS \"LOA Addressee Designation\",ADDRESSEE_EMAIL AS \"LOA Addressee Email\",IS_MAIN_CONTACT AS \"Main Contact\",MAIN_CONTACT_SALUTATION AS \"Main Contact  Salutation\",MAIN_CONTACT_NAME AS \"Main Contact Name\",MAIN_CONTACT_OCCUPATION AS \"Main Contact Role\",MAIN_CONTACT_ORG AS \"Main Contact Organisation\",MAIN_CONTACT_PHN_NO AS \"Main Contact Phone Number\",MAIN_CONTACT_EMAIL AS \"Main Contact Email\",MAILING_ADDRESS AS \"Mailing Address\",ALTERNATE_CONTACT_SALUTATION AS \"Alternate Contact  Salutation\",ALTERNATE_CONTACT_NAME AS \"Alternate Contact Name\",ALTERNATE_CONTACT_OCCUPATION AS \"Alternate Contact Role\",ALTERNATE_CONTACT_PHN_NO AS \"Alternate Contact Phone Number\",ALTERNATE_CONTACT_EMAIL AS \"Alternate Contact Email\",BREACH_STATUS AS \"Presence on Due Diligence List during Assessment\",TO_CHAR(LATEST_ASSESSMENT_DT, 'DD-MM-YYYY HH24:MI:SS') AS \"Latest Assessment Date\",PAYMENT_DATE AS \"Payment Date\",PAYMENT_STATUS AS \"Payment Status\",CATEGORY AS \"Project Category\",TITLE AS \"Title\",DESCRIPTION AS \"Description\",START_DATE AS \"Start Date\",END_DATE AS \"End Date\",VENUE AS \"Venue\",PROJECTED_BUDGET AS \"Projected Budget\",REQUESTED_GRANT_AMOUNT AS \"Requested Amount\",AIC_CLAIM_PERIOD_FROM AS \"Claim Period From\",AIC_CLAIM_PERIOD_TO AS \"Claim Period To\",AIC_CLAIM_AMOUNT AS \"Claim Amount\",IS_MIGRATED  AS  \"Is Migrated\" FROM V_GAM_EXPORT_DISBURSEMENTS WHERE PROJ_TYPE_CODE IN ('AIC_OPEN_GRANT','PROJECT','AIC_CLOSED_GRANT') AND  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-22 12:00:00 AM' AND '15-Jul-24 11:59:59 PM' ORDER BY \"Submission Date\"");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Latest");
        data2.setSqlQuery("SELECT REFERENCE_ID  AS  \"Reference ID\",APPLICANT_TYPE  AS \"Applicant Type\",AGENCY AS  \"Agency\",GRANT_SCHEME  AS \"Grant Scheme\",PROJ_TYPE_NAME AS \"Project Type\",IS_AUTO_DISBURSEMENT AS \"Auto Disbursement\",IS_FINAL_CLAIM AS \"Final Claim\",TRANCHE_NUMBER AS \"Tranche Number\",FIRST_SUBMISSION_DT AS \"Submission Date\",CURRENT_DISBURSEMENT_STATUS As \"Current Submission Status\",APPROVED_GRANT_AMOUNT AS \"Approved Grant Amount (S$)\",DISBURSED_AMOUNT AS \"Amount Disbursed To-Date (S$)\",RECOVERED_AMOUNT_TO_DATE AS \"Amount Recovered To-Date (S$)\",BALANCE_AMOUNT AS \"Balance Grant Amount (S$)\",TRANCHE_AMOUNT AS \"Tranche Amount (S$)\",CURRENT_TRANCHE_RECOVERED_AMOUNT AS \"Recovered Amount (S$)\",OWNER AS \"Owner\",DEPARTMENT AS \"Owner Department (if any)\",OWNER_ASSESSMENT_REMARK AS \"Owner Assessment Remark\",CERTIFYING_OFFICER AS \"Certifying Officer\",DCO_ASSESSMENT_REMARK AS \"Certifying Officer Remarks\",APPROVING_OFFICER AS \"Approver\",DAO_ASSESSMENT_REMARK AS \"Approver Assessment Remark\",APPROVED_DT AS \"Approved Date\",PC_SEND_OUT_DT AS \"Payment Confirmation Send Out Date\",SALUTATION AS \"Salutation\",APPLICANT_NAME AS \"Applicant Name\",NRIC AS \"NRIC\",RESIDENT_STATUS AS \"Resident Status\",RACE AS \"Race\",GENDER As \"Gender\",DOB AS \"Date of Birth\",APPLICANT_AGE AS \"Applicant Age\",CONTACT_NUMBER AS \"Contact Number\",EMAIL AS \"Email\",RESIDENTIAL_ADDRESS AS \"Residential Address\",OCCUPATION AS \"Occupation\",DESIGNATION AS \"Designation\",ORGANISATION AS \"Organisation\",ORGANISATION_ADDRESS AS \"Organisation Address\",ORG_CONTACT_NUMBER AS \"Organisation Contact Number\",UEN AS \"UEN\",ORGANISATION_TYPE AS \"Organisation Type\",REGISTERED_AS AS \"Registered As\",REGISTRATION_DATE AS \"Registration Date\",CHARITY_CHECK AS \"Registered Charity\",RAISE_MEMBER_CHECK \"Raise Member\",IPC_CHECK AS \"Institution of Public Character (IPC)\",IPC_EXPIRY_DATE AS \"IPC Expiry Date\",ORGANISATION_WEBSITE AS \"Organisation Website\",ORGANISATION_PAST_WORK AS \"Past Achievements/Projects/Work\",ORGANISATION_DESCRIPTION AS \"Organisation Vision/Description\",ADDRESSEE_SALUTATION AS \"LOA Addressee Salutation\",ADDRESSEE_NAME AS \"LOA Addressee Name\",ADDRESSEE_DESIGNATION AS \"LOA Addressee Designation\",ADDRESSEE_EMAIL AS \"LOA Addressee Email\",IS_MAIN_CONTACT AS \"Main Contact\",MAIN_CONTACT_SALUTATION AS \"Main Contact  Salutation\",MAIN_CONTACT_NAME AS \"Main Contact Name\",MAIN_CONTACT_OCCUPATION AS \"Main Contact Role\",MAIN_CONTACT_ORG AS \"Main Contact Organisation\",MAIN_CONTACT_PHN_NO AS \"Main Contact Phone Number\",MAIN_CONTACT_EMAIL AS \"Main Contact Email\",MAILING_ADDRESS AS \"Mailing Address\",ALTERNATE_CONTACT_SALUTATION AS \"Alternate Contact  Salutation\",ALTERNATE_CONTACT_NAME AS \"Alternate Contact Name\",ALTERNATE_CONTACT_OCCUPATION AS \"Alternate Contact Role\",ALTERNATE_CONTACT_PHN_NO AS \"Alternate Contact Phone Number\",ALTERNATE_CONTACT_EMAIL AS \"Alternate Contact Email\",BREACH_STATUS AS \"Presence on Due Diligence List during Assessment\",TO_CHAR(LATEST_ASSESSMENT_DT, 'DD-MM-YYYY HH24:MI:SS') AS \"Latest Assessment Date\",PAYMENT_DATE AS \"Payment Date\",PAYMENT_STATUS AS \"Payment Status\",CATEGORY AS \"Project Category\",TITLE AS \"Title\",DESCRIPTION AS \"Description\",START_DATE AS \"Start Date\",END_DATE AS \"End Date\",VENUE AS \"Venue\",PROJECTED_BUDGET AS \"Projected Budget\",REQUESTED_GRANT_AMOUNT AS \"Requested Amount\",AIC_CLAIM_PERIOD_FROM AS \"Claim Period From\",AIC_CLAIM_PERIOD_TO AS \"Claim Period To\",AIC_CLAIM_AMOUNT AS \"Claim Amount\",IS_MIGRATED  AS  \"Is Migrated\" FROM V_GAM_EXPORT_DISBURSEMENTS WHERE PROJ_TYPE_CODE IN ('AIC_OPEN_GRANT','PROJECT','AIC_CLOSED_GRANT') AND  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM' AND REFERENCE_ID IN (SELECT REFERENCE_ID FROM V_MGP_LATEST_GRANT) ORDER BY \"Submission Date\"");

        SQLSheetData data3 = new SQLSheetData();
        data3.setSheetName("Actions");
        data3.setSqlQuery("SELECT REFERENCE_ID AS \"Reference ID\",REF_NAME AS \"User  Name\",REQUEST_TYPE AS \"Request Type\",UPDATED_BY AS \"Updated By\",to_char(ACT_UPDATED_DT,'DD-MM-YYYY HH24:MI:SS') AS \"Updated  Date\",REMARKS AS \"Remarks\" FROM V_MGP_EXPORT_DISB_ACTION_LOGS WHERE  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM'");

        SQLSheetData data4 = new SQLSheetData();
        data4.setSheetName("Assessments");
        data4.setSqlQuery("SELECT REFERENCE_ID AS \"Reference ID\",USER_NAME AS \"User Name\",TRANCHE_AMOUNT AS \"Tranche Amount (S$)\",to_char(ACTION_TAKEN_ON,'DD-MM-YYYY HH24:MI:SS') AS \"Action Taken On\",REMARKS AS \"Remarks\" FROM V_MGP_EXPORT_DISB_ASSESSMENTS WHERE  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM'");

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
        data1.setSheetName("Tranches");
        data1.setSqlQuery("select * from OSG_DISBURSEMENT_TRANCHE_VERSIONS");

        SQLSheetData data2 = new SQLSheetData();
        data2.setSheetName("Latest");
        data2.setSqlQuery("SELECT REFERENCE_ID  AS  \"Reference ID\",APPLICANT_TYPE  AS \"Applicant Type\",AGENCY AS  \"Agency\",GRANT_SCHEME  AS \"Grant Scheme\",PROJ_TYPE_NAME AS \"Project Type\",IS_AUTO_DISBURSEMENT AS \"Auto Disbursement\",IS_FINAL_CLAIM AS \"Final Claim\",TRANCHE_NUMBER AS \"Tranche Number\",FIRST_SUBMISSION_DT AS \"Submission Date\",CURRENT_DISBURSEMENT_STATUS As \"Current Submission Status\",APPROVED_GRANT_AMOUNT AS \"Approved Grant Amount (S$)\",DISBURSED_AMOUNT AS \"Amount Disbursed To-Date (S$)\",RECOVERED_AMOUNT_TO_DATE AS \"Amount Recovered To-Date (S$)\",BALANCE_AMOUNT AS \"Balance Grant Amount (S$)\",TRANCHE_AMOUNT AS \"Tranche Amount (S$)\",CURRENT_TRANCHE_RECOVERED_AMOUNT AS \"Recovered Amount (S$)\",OWNER AS \"Owner\",DEPARTMENT AS \"Owner Department (if any)\",OWNER_ASSESSMENT_REMARK AS \"Owner Assessment Remark\",CERTIFYING_OFFICER AS \"Certifying Officer\",DCO_ASSESSMENT_REMARK AS \"Certifying Officer Remarks\",APPROVING_OFFICER AS \"Approver\",DAO_ASSESSMENT_REMARK AS \"Approver Assessment Remark\",APPROVED_DT AS \"Approved Date\",PC_SEND_OUT_DT AS \"Payment Confirmation Send Out Date\",SALUTATION AS \"Salutation\",APPLICANT_NAME AS \"Applicant Name\",NRIC AS \"NRIC\",RESIDENT_STATUS AS \"Resident Status\",RACE AS \"Race\",GENDER As \"Gender\",DOB AS \"Date of Birth\",APPLICANT_AGE AS \"Applicant Age\",CONTACT_NUMBER AS \"Contact Number\",EMAIL AS \"Email\",RESIDENTIAL_ADDRESS AS \"Residential Address\",OCCUPATION AS \"Occupation\",DESIGNATION AS \"Designation\",ORGANISATION AS \"Organisation\",ORGANISATION_ADDRESS AS \"Organisation Address\",ORG_CONTACT_NUMBER AS \"Organisation Contact Number\",UEN AS \"UEN\",ORGANISATION_TYPE AS \"Organisation Type\",REGISTERED_AS AS \"Registered As\",REGISTRATION_DATE AS \"Registration Date\",CHARITY_CHECK AS \"Registered Charity\",RAISE_MEMBER_CHECK \"Raise Member\",IPC_CHECK AS \"Institution of Public Character (IPC)\",IPC_EXPIRY_DATE AS \"IPC Expiry Date\",ORGANISATION_WEBSITE AS \"Organisation Website\",ORGANISATION_PAST_WORK AS \"Past Achievements/Projects/Work\",ORGANISATION_DESCRIPTION AS \"Organisation Vision/Description\",ADDRESSEE_SALUTATION AS \"LOA Addressee Salutation\",ADDRESSEE_NAME AS \"LOA Addressee Name\",ADDRESSEE_DESIGNATION AS \"LOA Addressee Designation\",ADDRESSEE_EMAIL AS \"LOA Addressee Email\",IS_MAIN_CONTACT AS \"Main Contact\",MAIN_CONTACT_SALUTATION AS \"Main Contact  Salutation\",MAIN_CONTACT_NAME AS \"Main Contact Name\",MAIN_CONTACT_OCCUPATION AS \"Main Contact Role\",MAIN_CONTACT_ORG AS \"Main Contact Organisation\",MAIN_CONTACT_PHN_NO AS \"Main Contact Phone Number\",MAIN_CONTACT_EMAIL AS \"Main Contact Email\",MAILING_ADDRESS AS \"Mailing Address\",ALTERNATE_CONTACT_SALUTATION AS \"Alternate Contact  Salutation\",ALTERNATE_CONTACT_NAME AS \"Alternate Contact Name\",ALTERNATE_CONTACT_OCCUPATION AS \"Alternate Contact Role\",ALTERNATE_CONTACT_PHN_NO AS \"Alternate Contact Phone Number\",ALTERNATE_CONTACT_EMAIL AS \"Alternate Contact Email\",BREACH_STATUS AS \"Presence on Due Diligence List during Assessment\",TO_CHAR(LATEST_ASSESSMENT_DT, 'DD-MM-YYYY HH24:MI:SS') AS \"Latest Assessment Date\",PAYMENT_DATE AS \"Payment Date\",PAYMENT_STATUS AS \"Payment Status\",CATEGORY AS \"Project Category\",TITLE AS \"Title\",DESCRIPTION AS \"Description\",START_DATE AS \"Start Date\",END_DATE AS \"End Date\",VENUE AS \"Venue\",PROJECTED_BUDGET AS \"Projected Budget\",REQUESTED_GRANT_AMOUNT AS \"Requested Amount\",AIC_CLAIM_PERIOD_FROM AS \"Claim Period From\",AIC_CLAIM_PERIOD_TO AS \"Claim Period To\",AIC_CLAIM_AMOUNT AS \"Claim Amount\",IS_MIGRATED  AS  \"Is Migrated\" FROM V_GAM_EXPORT_DISBURSEMENTS WHERE PROJ_TYPE_CODE IN ('AIC_OPEN_GRANT','PROJECT','AIC_CLOSED_GRANT') AND  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM' AND REFERENCE_ID IN (SELECT REFERENCE_ID FROM V_MGP_LATEST_GRANT) ORDER BY \"Submission Date\"");

        SQLSheetData data3 = new SQLSheetData();
        data3.setSheetName("Actions");
        data3.setSqlQuery("SELECT REFERENCE_ID AS \"Reference ID\",REF_NAME AS \"User  Name\",REQUEST_TYPE AS \"Request Type\",UPDATED_BY AS \"Updated By\",to_char(ACT_UPDATED_DT,'DD-MM-YYYY HH24:MI:SS') AS \"Updated  Date\",REMARKS AS \"Remarks\" FROM V_MGP_EXPORT_DISB_ACTION_LOGS WHERE  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM'");

        SQLSheetData data4 = new SQLSheetData();
        data4.setSheetName("Assessments");
        data4.setSqlQuery("SELECT REFERENCE_ID AS \"Reference ID\",USER_NAME AS \"User Name\",TRANCHE_AMOUNT AS \"Tranche Amount (S$)\",to_char(ACTION_TAKEN_ON,'DD-MM-YYYY HH24:MI:SS') AS \"Action Taken On\",REMARKS AS \"Remarks\" FROM V_MGP_EXPORT_DISB_ASSESSMENTS WHERE  AGENCY_ID IN (1181 ) AND  SCHEME_ID IN (5441 , 5042 , 6301 , 5521 , 4841 , 5181 , 5361 , 5063 , 5081 , 5121 , 5161 , 5201 ) AND  FIRST_SUBMISSION_DT BETWEEN '01-Jan-24 12:00:00 AM' AND '01-Jul-24 11:59:59 PM'");

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
