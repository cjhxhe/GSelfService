package com.gang.gselfservice.dao;

import com.gang.gselfservice.bo.DailyRateBO;
import com.gang.gselfservice.config.ConnectionPool;

import java.sql.*;

public class RateDAO {

    public static void createTable() throws SQLException {
        Connection conn = null;
        Statement stmt = null;
        try {
            conn = ConnectionPool.getInstance().getConnection();
            DatabaseMetaData meta = conn.getMetaData();
            ResultSet rsTables = meta.getTables(null, null, "RATE_INFO",
                    new String[]{"TABLE"});
            if (!rsTables.next()) {
                stmt = conn.createStatement();
                stmt.execute("CREATE TABLE `RATE_INFO` (\n" +
                        "  `id` INT NOT NULL AUTO_INCREMENT,\n" +
                        "  `date` VARCHAR(32) ,\n" +
                        "  `rate_name` VARCHAR(64) ,\n" +
                        "  `rate_role` VARCHAR(32) ,\n" +
                        "  `phone` VARCHAR(32) ,\n" +
                        "  `biz_name` VARCHAR(64) ,\n" +
                        "  `social_credit_code` VARCHAR(64) ,\n" +
                        "  `tax_payer` VARCHAR(64) ,\n" +
                        "  `send_time` VARCHAR(32) ,\n" +
                        "  `comment` VARCHAR(256) ,\n" +
                        "  `rate_result` VARCHAR(256) ,\n" +
                        "  `rate_time` VARCHAR(32) ,\n" +
                        "  `source` VARCHAR(64) ,\n" +
                        "  `tax_authority` VARCHAR(64) ,\n" +
                        "  PRIMARY KEY (`id`))");
            }
            rsTables.close();
        } finally {
            releaseConnection(conn, stmt, null);
        }
    }

    public static void deleteByDate(String date) throws SQLException {
        Connection conn = null;
        PreparedStatement stmt = null;
        try {
            conn = ConnectionPool.getInstance().getConnection();
            stmt = conn.prepareStatement("delete from RATE_INFO where date = ?");
            stmt.setString(1, date);
            stmt.execute();
        } finally {
            releaseConnection(conn, stmt, null);
        }
    }

    public static void addInfo(DailyRateBO dailyRateBO) throws SQLException {
        Connection conn = null;
        PreparedStatement stmt = null;
        try {
            conn = ConnectionPool.getInstance().getConnection();

            stmt = conn
                    .prepareStatement("INSERT INTO RATE_INFO(date, rate_name, rate_role, phone, biz_name, social_credit_code, tax_payer, send_time, comment, rate_result, rate_time, source, tax_authority) " +
                            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)");
            int i = 1;
            stmt.setString(i++, dailyRateBO.getDate());
            stmt.setString(i++, dailyRateBO.getRateName());
            stmt.setString(i++, dailyRateBO.getRateRole());
            stmt.setString(i++, dailyRateBO.getPhone());
            stmt.setString(i++, dailyRateBO.getBizName());
            stmt.setString(i++, dailyRateBO.getSocialCreditCode());
            stmt.setString(i++, dailyRateBO.getTaxPayer());
            stmt.setString(i++, dailyRateBO.getSendTime());
            stmt.setString(i++, dailyRateBO.getComment());
            stmt.setString(i++, dailyRateBO.getRateResult().getResult());
            stmt.setString(i++, dailyRateBO.getRateTime());
            stmt.setString(i++, dailyRateBO.getSource().getTitle());
            stmt.setString(i++, dailyRateBO.getTaxAuthority());
            stmt.execute();
        } finally {
            releaseConnection(conn, stmt, null);
        }
    }

    private static void releaseConnection(Connection conn, Statement stmt,
                                          ResultSet rs) throws SQLException {
        if (rs != null) {
            rs.close();
        }
        if (stmt != null) {
            stmt.close();
        }
        if (conn != null) {
            conn.close();
        }
    }
}
