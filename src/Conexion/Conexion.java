/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Conexion;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

/**
 *
 * @author clopez
 */
public class Conexion {

    static Connection con = null;
    static Connection con2 = null;
    static Connection con3 = null;
    public static String query;
    public static Statement stat;
    public static ResultSet rs;

    public static Connection conectar() {

        String connectionUrl = "jdbc:sqlserver://192.168.10.1:1433;"
                + "databaseName=NodumEP;user=sa;password=EPpal2003;";

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            con = DriverManager.getConnection(connectionUrl);
        } catch (ClassNotFoundException | SQLException e) {
            System.out.println("error " + e);
        }
        return con;
    }

    public static void cerrarConexion() {
        try {
            stat.close();
            con.close();
        } catch (SQLException e) {
            System.out.println("Error en cerrar la base de datos" + e.toString());
        }
    }

}
