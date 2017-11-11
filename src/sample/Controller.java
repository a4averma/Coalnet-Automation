package sample;


import javafx.fxml.FXML;
import javafx.scene.control.*;

import java.io.*;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;


public class Controller {

    DateFormat dateFormat = new SimpleDateFormat("dd-MM-YYYY");
    Date today = new Date();
    String date = dateFormat.format(today);

    @FXML
    private TextArea logger;

    @FXML
    private Button start;

    @FXML
    private MenuItem about;

    @FXML
    public void initialize() {
        logger.appendText("Welcome to Coal-Net BTS Automation \n");
        start.setOnAction(event -> {
            generateExcel();
            upload();
        });
        about.setOnAction(event -> {
            Alert alert = new Alert(Alert.AlertType.INFORMATION, "My Project Submission for ECL Internship under Mr. Sudhesh Kumar.\nsiradityaverma@gmail.com", ButtonType.OK);
            alert.showAndWait();
            if (alert.getResult() == ButtonType.OK) {
                alert.close();
            }

        });

    }

    public void generateExcel() {
        try {
            Class.forName("oracle.jdbc.driver.OracleDriver");

            Connection con = DriverManager.getConnection(
                    "jdbc:oracle:thin:@172.18.3.251:1522:eclprod",
                    "username",
                    "password");

            Statement stmt = con.createStatement();
            ResultSet rs = stmt.executeQuery("select \n" +
                    "       A.BILL_NO as Bill_track_no,\n" +
                    "       A.BILL_DATE,\n" +
                    "       A.WORK_ORD_NO,\n" +
                    "       A.WORK_ORD_DATE,\n" +
                    "       B.vendor_NAME AS PARTY_NAME,\n" +
                    "       B.Vendor_Code AS PARTY_ID,\n" +
                    "       B.Vendor_Telephone_O1 as PARTY_PHONE,\n" +
                    "       B.Vendor_Email as PARTY_EMAIL, \n" +
                    "       A.party_bill_no,\n" +
                    "       A.PARTY_BILL_DATE,\n" +
                    "       A.PARTY_BILL_AMOUNT,\n" +
                    "       A.PARTY_BILL_AMOUNT as NET_AMOUNT,\n" +
                    "       A.REMARKS as NARRATION,\n" +
                    "       d.approval_flag,\n" +
                    "       d.approved_date,\n" +
                    "       A.PAYMENT_STATUS,\n" +
                    "       A.Payment_Date,\n" +
                    "       d.rec_dept,\n" +
                    "       c.rec_dept as current_dept,\n" +
                    "       c.remarks\n" +
                    "       \n" +
                    " FROM t_fis_dept_bill_rcpt_hdr a, m_mms_vendor b,t_fis_dept_bill_rcpt_dtl c, \n" +
                    " (select bill_id, rec_dept, approval_flag,approved_date from t_fis_dept_bill_rcpt_dtl where line_no=1 order by bill_id) d\n" +
                    " where a.payee_id = b.vendor_ID(+)\n" +
                    " AND   a.bill_id=c.bill_id(+) \n" +
                    " and a.bill_id=d.bill_id(+)\n" +
                    " and c.line_no=(select max(line_no) from t_fis_dept_bill_rcpt_dtl where bill_id=c.bill_id)\n" +
                    " and A.PAYEE_TYPE_ID=1\n" +
                    " and a.delete_flag='N'\n" +
                    " \n" +
                    " UNION\n" +
                    " \n" +
                    " select \n" +
                    "       A.BILL_NO as Bill_track_no,\n" +
                    "       A.BILL_DATE,\n" +
                    "       A.WORK_ORD_NO,\n" +
                    "       A.WORK_ORD_DATE,\n" +
                    "       B.party_NAME AS PARTY_NAME,\n" +
                    "       B.party_code as party_id,\n" +
                    "       B.phone_number as party_phone,\n" +
                    "       B.email as party_email, \n" +
                    "       A.party_bill_no,\n" +
                    "       A.PARTY_BILL_DATE,\n" +
                    "       A.PARTY_BILL_AMOUNT,\n" +
                    "       A.PARTY_BILL_AMOUNT as NET_AMOUNT,\n" +
                    "       A.REMARKS as NARRATION,\n" +
                    "       d.approval_flag,\n" +
                    "       d.approved_date,\n" +
                    "       A.PAYMENT_STATUS,\n" +
                    "       A.Payment_Date,\n" +
                    "       d.rec_dept,\n" +
                    "       c.rec_dept as current_dept,\n" +
                    "       c.remarks\n" +
                    "       \n" +
                    " FROM t_fis_dept_bill_rcpt_hdr a, m_fis_party b,t_fis_dept_bill_rcpt_dtl c, \n" +
                    " (select bill_id, rec_dept, approval_flag,approved_date from t_fis_dept_bill_rcpt_dtl where line_no=1 order by bill_id) d\n" +
                    " where a.payee_id = b.party_ID(+)\n" +
                    " AND   a.bill_id=c.bill_id (+)\n" +
                    " and a.bill_id=d.bill_id(+)\n" +
                    " and c.line_no=(select max(line_no) from t_fis_dept_bill_rcpt_dtl where bill_id=c.bill_id)\n" +
                    " and A.PAYEE_TYPE_ID=3\n" +
                    " and a.delete_flag='N'");

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("db");
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell;
            cell = row.createCell(0);
            cell.setCellValue("BILL_TRACK_NO");
            cell = row.createCell(1);
            cell.setCellValue("BILL_DATE");
            cell = row.createCell(2);
            cell.setCellValue("WORK_ORD_NO");
            cell = row.createCell(3);
            cell.setCellValue("WORK_ORD_DATE");
            cell = row.createCell(4);
            cell.setCellValue("PARTY_NAME");
            cell = row.createCell(5);
            cell.setCellValue("PARTY_ID");
            cell = row.createCell(6);
            cell.setCellValue("PARTY_PHONE");
            cell = row.createCell(7);
            cell.setCellValue("PARTY_EMAIL");
            cell = row.createCell(8);
            cell.setCellValue("PARTY_BILL_NO");
            cell = row.createCell(9);
            cell.setCellValue("PARTY_BILL_DATE");
            cell = row.createCell(10);
            cell.setCellValue("PARTY_BILL_AMOUNT");
            cell = row.createCell(11);
            cell.setCellValue("NET_AMOUNT");
            cell = row.createCell(12);
            cell.setCellValue("NARRATION");
            cell = row.createCell(13);
            cell.setCellValue("APPROVAL_FLAG");
            cell = row.createCell(14);
            cell.setCellValue("APPROVED_DATE");
            cell = row.createCell(15);
            cell.setCellValue("PAYMENT_STATUS");
            cell = row.createCell(16);
            cell.setCellValue("PAYMENT_DATE");
            cell = row.createCell(17);
            cell.setCellValue("REC_DEPT");
            cell = row.createCell(18);
            cell.setCellValue("CURRENT_DEPT");
            cell = row.createCell(19);
            cell.setCellValue("REMARKS");
            XSSFCreationHelper createHelper = workbook.getCreationHelper();
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("dd-MM-yyyy"));


            int i = 1;
            while (rs.next()) {
                row = sheet.createRow(i);
                cell = row.createCell(0);
                cell.setCellValue(rs.getString(1));
                cell = row.createCell(1);
                cell.setCellValue(rs.getDate(2));
                cell.setCellStyle(cellStyle);
                cell = row.createCell(2);
                cell.setCellValue(rs.getString(3));
                cell = row.createCell(3);
                if (rs.getDate(4) != null) {
                    cell.setCellValue(rs.getDate(4));
                    cell.setCellStyle(cellStyle);
                } else {

                }

                cell = row.createCell(4);
                cell.setCellValue(rs.getString(5));
                cell = row.createCell(5);
                cell.setCellValue(rs.getString(6));
                cell = row.createCell(6);
                cell.setCellValue(rs.getString(7));
                cell = row.createCell(7);
                cell.setCellValue(rs.getString(8));
                cell = row.createCell(8);
                cell.setCellValue(rs.getString(9));
                cell = row.createCell(9);
                if (rs.getDate(10) != null) {
                    cell.setCellValue(rs.getDate(10));
                    cell.setCellStyle(cellStyle);
                } else {

                }

                cell = row.createCell(10);
                cell.setCellValue(rs.getString(11));
                cell = row.createCell(11);
                cell.setCellValue(rs.getString(12));
                cell = row.createCell(12);
                cell.setCellValue(rs.getString(13));
                cell = row.createCell(13);
                cell.setCellValue(rs.getString(14));
                cell = row.createCell(14);
                if (rs.getDate(15) != null) {
                    cell.setCellValue(rs.getDate(15));
                    cell.setCellStyle(cellStyle);
                } else {

                }
                cell = row.createCell(15);
                cell.setCellValue(rs.getString(16));
                cell = row.createCell(16);
                cell.setCellValue(rs.getDate(17));
                cell = row.createCell(17);
                cell.setCellValue(rs.getString(18));
                cell = row.createCell(18);
                cell.setCellValue(rs.getString(19));
                cell = row.createCell(19);
                cell.setCellValue(rs.getString(20));
                i++;
            }



            FileOutputStream out = new FileOutputStream(
                    new File("BTS " + date +".xlsx"));
            workbook.write(out);
            out.close();
            logger.appendText(
                    "BTS " + date + ".xlsx written successfully \n");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void upload() {
        String url = "https://secureloginecl.co.in/bts/chk_user.php";
        System.setProperty("http.proxyHost", "172.18.3.33");
        System.setProperty("http.proxyPort", "3128");
        try {
            org.jsoup.Connection.Response response = Jsoup.connect(url).data("username", "admin", "password", "password").method(org.jsoup.Connection.Method.POST).execute();
            Map<String, String> loginCookies = response.cookies();

            String uploadUrl = "https://secureloginecl.co.in/bts/bill_info_update.php";
            File excelFile = new File("BTS " + date + ".xlsx");
            FileInputStream file = new FileInputStream(excelFile);

            Document document = Jsoup.connect(uploadUrl).cookies(loginCookies).data("file1", "BTS " + date + ".xlsx", file).post();

            Elements element = document.select(".colm h4");
            boolean append = true;
            FileHandler handler = new FileHandler("default.log", append);
            Logger log = Logger.getLogger("com.easterncoalfields.coalnet");
            log.addHandler(handler);
            log.info(element.text());
            logger.appendText(date + ": " + element.text());

        } catch (Exception e) {
            exception(e);
        }
    }

    public void exception(Exception e) {
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        e.printStackTrace(pw);
        String exception = sw.toString();
        logger.appendText(exception);
    }
}



