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

import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;

public class Main {
    public static void main(String[] args) throws IOException, BiffException, SQLException, ClassNotFoundException {
        if (args.length != 2) usage();
        String xlsPath = args[0];
        String sqlitePath = args[1];
        Xls2Sqlite.xls2sqlite(xlsPath, sqlitePath);
    }

    private static void usage() {
        System.err.println(getProgramName() + " </path/to/file.xls> </path/to/file.db>");
        System.exit(-1);
    }

    /**
     * Try to return a string like "java -jar /path/to/app.jar" to be used within the usage.
     */
    private static String getProgramName() {
        try {
            File file = new File(Main.class.getProtectionDomain().getCodeSource().getLocation().getPath());
            if (file.isFile()) {
                String relativeFile = new File(".").toURI().relativize(file.toURI()).getPath();
                return "java -jar " + relativeFile;
            }
        } catch (Throwable ignored) {
        }
        return "java " + Main.class.getName();
    }
}
