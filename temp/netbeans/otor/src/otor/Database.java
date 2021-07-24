/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package otor;

import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author El-Wattaneya
 */
public class Database {

    public static Connection conn;
    public static Statement stmt;
    public static String lastQuery;
    public static ResultSet lastResult;
    public static String lastUpdate;
    public static String lastValues[][];
    public static int lastRowsUpdated = 0;
    public static boolean showMassages = true;
    public static boolean intialized = false;

    public static boolean InitDatabase(String url ,boolean show) {
        try {
            conn = DriverManager.getConnection(url);
            stmt = conn.createStatement();

        } catch (SQLException ex) {
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
        intialized = true;
        showMassages = show;
        System.out.print((showMassages) ? "Database stated successfuly...\n" : "");
        return true;
    }

    public static int executeQuery(String q) {
        try {
            lastResult = stmt.executeQuery(q);
            lastQuery = q;
        } catch (SQLException ex) {
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
            System.out.print((showMassages) ? "error while quering...\n" : "");
            return 0;
        }

        System.out.print((showMassages) ? "successful Query...\n" : "");
        return organizeResult(lastResult, lastQuery);
    }

    public static int executeUpdate(String q) {
        try {
            lastRowsUpdated = stmt.executeUpdate(q);
            lastUpdate = q;
        } catch (SQLException ex) {
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
            System.out.print((showMassages) ? "error while updating...\n" : "");
            return -1;
        }

        System.out.print((showMassages) ? "successful Update...\n" : "");
        return lastRowsUpdated;
    }

    public static int organizeResult(ResultSet rs, String q) {

        int counter = 0;
        int records = 0;
        try {
//            System.out.println((showMassages)? (lastQuery.toUpperCase().replace(lastQuery.substring(lastQuery.indexOf(" "), lastQuery.indexOf("FROM"))," COUNT(*) ")) + "\n":"");
//            ResultSet r =  stmt.executeQuery(lastQuery.toUpperCase().replace(lastQuery.substring(lastQuery.indexOf(" "), lastQuery.indexOf("FROM"))," COUNT(*) "));
            ResultSet r = conn.createStatement().executeQuery(lastQuery);
            while (r.next()) {
                records++;
            }
            System.out.print((showMassages) ? "Rowcount: " + records + " \n" : "");
            r.close();
        } catch (SQLException ex) {
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
            System.out.print((showMassages) ? "error while rowcount...\n" : "");
            return -1;
        }
        String params[] = q.toUpperCase().replaceFirst("SELECT", "").substring(0, q.indexOf("FROM") - 6).split(",");
        String values[][] = new String[records + 1][params.length];
        values[counter] = params;

        System.out.print((showMassages) ? (q.toUpperCase().replaceFirst("SELECT", "").substring(0, q.indexOf("FROM") - 6)) + "\n" : "");
        System.out.print((showMassages) ? params.length + "\n" : "");

        try {
            while (rs.next()) {
                counter++;
                values[counter] = new String[params.length];

                for (int j = 0; j < params.length; j++) {
                    System.out.print((showMassages) ? "counter: " + counter + " - j: " + j + " >> " + params[j].trim() + "\n" : "");
                    values[counter][j] = rs.getString(params[j].trim());
                }

            }
        } catch (SQLException ex) {
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
            System.out.print((showMassages) ? "error while parsing...\n" : "");
            return -1;
        }
        lastValues = values;
        return counter;
    }

    public static void printTable(String t[][]) {


        for (int i = 0; i < t.length; i++) {
            System.out.print("|");
            for (int j = 0; j < t[0].length; j++) {
                int temp = (int) (t[1][j].length() + 10);
                System.out.print(t[i][j] + ((t[i][j].length() >= temp) ? "" : "                                                         ".substring(0, temp - t[i][j].length())) + "|");
            }
            System.out.println("");
            printLine(t[0].length, t[0][0].length());
        }

    }

    public static void printLine(int counter, int len) {
        for (int i = 0; i < counter * len * 4; i++) {
            System.out.print("-");
        }
        System.out.println("");
    }

    public static void showLastResults() {
        System.out.println("Last Query: " + lastQuery);
        System.out.println("result: ");
        printTable(lastValues);
        System.out.println("----------------");
        System.out.println("Last Update: " + lastUpdate);
        System.out.println("Num of Updated Rows: " + lastRowsUpdated);
    }

    public static void showLastQuery() {
        System.out.println("Last Query: " + lastQuery);
        System.out.println("result: ");
        printTable(lastValues);
    }

    public static void showLastUpdate() {
        System.out.println("Last Update: " + lastUpdate);
        System.out.println("Num of Updated Rows: " + lastRowsUpdated);
    }

    public static void closeDatabase(boolean commit) {
        try {
            conn.commit();
            stmt.close();
            conn.close();
        } catch (SQLException ex) {
            System.out.print((showMassages) ? "error while closing...\n" : "");
            Logger.getLogger(Database.class.getName()).log(Level.SEVERE, null, ex);
        }

        lastQuery = "";
        lastUpdate = "";
        lastResult = null;
        lastValues = null;
        lastRowsUpdated = 0;
        showMassages = true;
        intialized = false;
    }
}
