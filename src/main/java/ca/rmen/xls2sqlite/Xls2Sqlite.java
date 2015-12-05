/*
 * ----------------------------------------------------------------------------
 * "THE WINE-WARE LICENSE" Version 1.0:
 * Authors: Carmen Alvarez.
 * As long as you retain this notice you can do whatever you want with this stuff.
 * If we meet some day, and you think this stuff is worth it, you can buy me a
 * glass of wine in return.
 *
 * THE AUTHORS OF THIS FILE ARE NOT RESPONSIBLE FOR LOSS OF LIFE, LIMBS, SELF-ESTEEM,
 * MONEY, RELATIONSHIPS, OR GENERAL MENTAL OR PHYSICAL HEALTH CAUSED BY THE
 * CONTENTS OF THIS FILE OR ANYTHING ELSE.
 * ----------------------------------------------------------------------------
 */

package ca.rmen.xls2sqlite;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Create an SQLite file from an Excel file.
 */
public class Xls2Sqlite {

    private static final int BATCH_INSERT_SIZE = 100;

    /**
     * Reads the given Excel file and creates an SQLite file at the given path containing the data from the Excel file.
     * One table is created for each Excel sheet.  The first row of the Excel file must contain the column names.
     * Empty rows are ignored.
     * Column names beginning with # are ignored.
     * This is a very simple implementation, with the following limitations:
     * <ul>
     * <li>No primary keys or foreign keys are created.</li>
     * <li>No autoincrement or unique indexes are created.</li>
     * <li>No NOT NULL constraints (or any constraints for that matter) are created.</li>
     * <li>All fields in the Excel file are treated as strings.</li>
     * <li>All columns in the SQLite file are created as TEXT.</li>
     * <li>If the SQLite file exists already, it is deleted before creating a new one.</li>
     * </ul>
     *
     * @param xlsPath    the path to the Excel file to read.
     * @param sqlitePath the path to the SQLite file to create.
     * @throws IOException
     * @throws BiffException
     * @throws SQLException
     */
    public static void xls2sqlite(String xlsPath, String sqlitePath) throws IOException, BiffException, SQLException, ClassNotFoundException {
        File sqliteFile = new File(sqlitePath);

        if (sqliteFile.isFile()) sqliteFile.delete();

        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setEncoding("iso-8859-1");
        Workbook wb = Workbook.getWorkbook(new File(xlsPath), wbSettings);
        Connection connection = getDbConnection(sqlitePath);
        try {
            Sheet[] sheets = wb.getSheets();
            for (Sheet sheet : sheets) {
                processSheet(connection, sheet);
            }
        } finally {
            wb.close();
        }
    }

    private static void processSheet(Connection connection, Sheet sheet) throws SQLException {
        String sheetName = normalizeForDb(sheet.getName());
        int colCount = sheet.getColumns();
        int rowCount = sheet.getRows();
        if (rowCount < 2 || colCount < 1) {
            System.err.println("Ignoring empty sheet " + sheet.getName());
            return;
        }
        List<String> columnDefinitions = new ArrayList<>();
        List<Integer> ignoredColumns = new ArrayList<>();
        List<String> columnNames = new ArrayList<>();
        List<String> insertArgs = new ArrayList<>();
        for (int col = 0; col < colCount; col++) {
            Cell header = sheet.getCell(col, 0);
            if (header.getContents().startsWith("#")) {
                ignoredColumns.add(col);
                continue;
            }
            String columnName = normalizeForDb(header.getContents());
            columnDefinitions.add(columnName + " TEXT");
            columnNames.add(columnName);
            insertArgs.add("?");
        }

        String sqlCreateTable = String.format("CREATE TABLE %s (%s)", sheetName, String.join(", ", columnDefinitions));
        Statement createTableStatement = connection.createStatement();

        createTableStatement.execute(sqlCreateTable);
        createTableStatement.close();
        System.out.println("executed " + sqlCreateTable);

        String sqlInsertRow = String.format("INSERT INTO %s (%s) VALUES (%s)", sheetName,
                String.join(", ", columnNames),
                String.join(", ", insertArgs));
        PreparedStatement insertStatement = connection.prepareStatement(sqlInsertRow);

        int insertedCount = 0;
        for (int row = 1; row < rowCount; row++) {
            List<String> insertValues = new ArrayList<>();

            boolean isRowEmpty = true;
            for (int col = 0; col < colCount; col++) {
                if (ignoredColumns.contains(col)) continue;
                String value = sheet.getCell(col, row).getContents();
                if (value != null && !value.trim().isEmpty()) isRowEmpty = false;
                insertValues.add(value);
            }

            // Skip empty rows
            if (isRowEmpty) continue;

            for (int i = 0; i < insertValues.size(); i++) {
                insertStatement.setString(i + 1, insertValues.get(i));
            }
            insertStatement.addBatch();
            if (row % BATCH_INSERT_SIZE == 0) {
                int[] result = insertStatement.executeBatch();
                insertedCount += result.length;
                System.out.println(insertedCount + " rows inserted into " + sheetName);
            }
        }
        int[] result = insertStatement.executeBatch();
        insertedCount += result.length;
        System.out.println(insertedCount + " rows inserted into " + sheetName);
        insertStatement.close();
    }

    /**
     * @return a version of the given string that can be used in db column or table names.
     */
    private static String normalizeForDb(String string) {
        string = string.toLowerCase();
        return string.replaceAll("[^a-z][^a-z]*", "_");
    }

    private static Connection getDbConnection(String dbPath) throws SQLException, ClassNotFoundException {
        Class.forName("org.sqlite.JDBC");
        return DriverManager.getConnection("jdbc:sqlite:" + dbPath);
    }
}
