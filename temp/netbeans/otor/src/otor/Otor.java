/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package otor;


import java.sql.*;

/**
 *
 * @author El-Wattaneya
 */
public class Otor {

    public static void main(String[] args) throws SQLException, ClassNotFoundException {

       Database.InitDatabase("jdbc:odbc:otor32" , true);

       Database.executeQuery("SELECT ID , GROUP_CODE , CODE FROM LOOKUP ORDER BY ID DESC");
       Database.showLastQuery();
       
       Database.executeUpdate("INSERT INTO LOOKUP (ID, GROUP_CODE ,CODE) VALUES ( '504' , 'الطبيب' ,'DOCTOR' )");
       Database.showLastUpdate();
       
       Database.closeDatabase(true);
    }
}
