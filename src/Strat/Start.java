/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Strat;

import Conexion.CronTask1;
import java.io.IOException;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import org.quartz.SchedulerException;

/**
 *
 * @author clopez
 */
public class Start {

    static SimpleDateFormat formato = new SimpleDateFormat("mm");
    static boolean correr = true;

    public static void main(String[] args) throws SQLException, IOException, SchedulerException, InterruptedException {
        CronTask1 javaPoiUtils = new CronTask1();

        while (correr) {
            if (formato.format(new Date()).equals("59")) {
                System.out.println("Ha iniciado la tarea para enviar reporte");
                correr = false;
                javaPoiUtils.ejecutarCron();
                break;
            }
            Thread.sleep(1000);
        }

////        javaPoiUtils.ejecutarCron();
//
//        javaPoiUtils.generarArchivoExcel();
    }

}
