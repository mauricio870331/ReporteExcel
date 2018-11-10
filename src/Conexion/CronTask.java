/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Conexion;

import Model.Data;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Locale;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.quartz.Job;
import static org.quartz.JobBuilder.newJob;
import org.quartz.JobDetail;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.quartz.Scheduler;
import org.quartz.SchedulerException;
import org.quartz.SchedulerFactory;
import static org.quartz.SimpleScheduleBuilder.simpleSchedule;
import org.quartz.Trigger;
import static org.quartz.TriggerBuilder.newTrigger;

public class CronTask implements Job {

    static Connection cn;
    static Connection con3 = null;
    static PreparedStatement pstm;
    static ResultSet rs;
    SimpleDateFormat formato = new SimpleDateFormat("yyyy-MM-dd");
    static CallableStatement cstmt;
    SchedulerFactory schedFact = new org.quartz.impl.StdSchedulerFactory();
    Scheduler sched;

    @Override
    public void execute(JobExecutionContext jec) throws JobExecutionException {
        try {
            System.out.println("Generando excel");
            generarArchivoExcel();
        } catch (SQLException | IOException ex) {
            System.out.println("error");
        }
    }

    public void ejecutarCron() throws SchedulerException {
        sched = schedFact.getScheduler();
        sched.start();
        JobDetail job = newJob(CronTask.class).withIdentity("myJob", "group1").build();
        Trigger trigger = newTrigger()
                .withIdentity("myTrigger", "group1")
                .startNow()
                .withSchedule(simpleSchedule()
                        .withIntervalInMinutes(60)
                        .repeatForever())
                .build();
        sched.scheduleJob(job, trigger);
    }

    public void detenerCron() throws SchedulerException {
        sched.shutdown();
        sched = null;
    }

    public void generarArchivoExcel() throws SQLException, FileNotFoundException, IOException {

        try {

            ArrayList<Data> listDatos = new ArrayList();
            cn = Conexion.conectar();
            String query = "select * from [dbo].[DatosExcel]('" + formato.format(new Date()) + " 00:00:00','" + formato.format(new Date()) + " 23:59:59',0,1,-1)";
            System.out.println("query " + query);
            pstm = cn.prepareStatement(query);
            rs = pstm.executeQuery();
            while (rs.next()) {
                listDatos.add(new Data(rs.getString(1), rs.getInt(2), rs.getDouble(3), rs.getDouble(4), rs.getInt(5)));
                System.out.println(rs.getString(1) + " " + rs.getInt(2) + " " + rs.getDouble(3) + " " + rs.getDouble(4) + " " + rs.getDouble(5));
            }
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("Datos Agencias");
            XSSFRow row;
            XSSFCell cell;
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("AGENCIA");
            cell = row.createCell(1);
            cell.setCellValue("PASAJES");
            cell = row.createCell(2);
            cell.setCellValue("IMPORTEBASE");
            cell = row.createCell(3);
            cell.setCellValue("DESCUENTOS");
            cell = row.createCell(4);
            cell.setCellValue("FINAL");
            for (int r = 0; r < listDatos.size(); r++) {
                row = sheet.createRow(r + 1);
                for (int c = 0; c < 5; c++) {

                    if (c == 0) {
                        cell = row.createCell(c);
                        cell.setCellValue(listDatos.get(r).getAgencia());
                    }
                    if (c == 1) {
                        cell = row.createCell(c);
                        cell.setCellValue(listDatos.get(r).getPasajes());
                    }
                    if (c == 2) {
                        cell = row.createCell(c);
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        cell.setCellValue((double)listDatos.get(r).getImporteBase());

                    }
                    if (c == 3) {
                        cell = row.createCell(c);
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        cell.setCellValue((double) listDatos.get(r).getDescuentos());
                    }
                    if (c == 4) {
                        cell = row.createCell(c);
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        cell.setCellValue((double) listDatos.get(r).getVfinal());

                    }
                }
            }
            //Escribimos el resultado en fichero
//            String nombreFichero = "Z:/DatosExcel.xlsx";//unidad de red al 10.1
            String nombreFichero = "C:/DatosExcel.xlsx";
            File file = new File(nombreFichero);
            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workBook.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
//Envia el archivo desde sqlserver
//            cstmt = cn.prepareCall("{call EnviarEmailConAdjuntos (?,?,?,?,?)}");
//            cstmt.setString(1, "Email Sistemas");
//            cstmt.setString(2, "desarrollo1@expresopalmira.com.co");
//            cstmt.setString(3, "Prueba reporte");
//            cstmt.setString(4, "revisar si llego");
//            cstmt.setString(5, "\\\\192.168.10.1\\ExcelParaEmail\\DatosExcel.xlsx");
//            cstmt.execute();

        } catch (SQLException e) {
            System.out.println("error " + e);
        } finally {
            cn.close();
            pstm.close();
        }
    }

}
