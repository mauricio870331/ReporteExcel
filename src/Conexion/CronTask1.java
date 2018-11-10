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
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.hssf.record.formula.functions.Match;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
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

public class CronTask1 implements Job {

    static Connection cn;
    static Connection con3 = null;
    static PreparedStatement pstm;
    static ResultSet rs, rs2;
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
        JobDetail job = newJob(CronTask1.class).withIdentity("myJob", "group1").build();
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
//        Locale.setDefault(new Locale("ES", "COL"));
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(new Date());
        calendar.add(Calendar.YEAR, -1);
//        System.out.println(formato.format(calendar.getTime())); // Devuelve el objeto Date con los nuevos días añadidos
        try {
            ArrayList<Data> Actual = new ArrayList();
            ArrayList<Data> Anterior = new ArrayList();
            cn = Conexion.conectar();

            String query = "select * from [dbo].[DatosExcel]('" + formato.format(new Date()) + " 00:00:00','" + formato.format(new Date()) + " 23:59:59',0,1,-1) "
                    + "where  Nombre not in ('MONITOREO','GERENCIA','SOACHA VARIANTE') order by 1";
            System.out.println("query " + query);
            pstm = cn.prepareStatement(query);
            rs = pstm.executeQuery();
            while (rs.next()) {
                Actual.add(new Data(rs.getString(1), rs.getInt(2), rs.getDouble(3), rs.getDouble(4), rs.getInt(5)));
                Anterior.add(new Data(rs.getString(1), 0, 0.0, 0.0, 0));
            }
            //anterior
            String query2 = "select * from [dbo].[DatosExcel]('" + formato.format(calendar.getTime()) + " 00:00:00','" + formato.format(calendar.getTime()) + " 23:59:59',0,1,-1) "
                    + "where  Nombre not in ('MONITOREO','GERENCIA','SOACHA VARIANTE') order by 1";
            System.out.println("query " + query2);
            pstm = cn.prepareStatement(query2);
            rs2 = pstm.executeQuery();
            while (rs2.next()) {
                for (Data data : Anterior) {
                    if (data.getAgencia().equals(rs2.getString(1))) {
//                        tring agencia, int pasajes, Double importeBase, Double descuentos, int Vfinal
                        data.setPasajes(rs2.getInt(2));
                        data.setImporteBase(rs2.getDouble(3));
                        data.setDescuentos(rs2.getDouble(4));
                        data.setVfinal(rs2.getInt(5));
                    }
                }
            }

            //un año atras
            //---------------------------------
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Datos Agencias");
            XSSFFont font = workbook.createFont();
            font.setFontHeightInPoints((short) 9);
            font.setFontName("Arial");
            font.setItalic(false);
            font.setBold(true);
            font.setColor(HSSFColor.BLACK.index);
            XSSFCellStyle styleCurrencyFormat = null;
            styleCurrencyFormat = workbook.createCellStyle();
            styleCurrencyFormat.setFont(font);
            XSSFRow row = null;
            XSSFCell cell;
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("Comparativa ventas año " + formato.format(new Date()) + " | " + formato.format(calendar.getTime()));
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));

            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue("AGENCIA");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(1);
            cell.setCellValue("PAS18");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(2);
            cell.setCellValue("PAS17");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(3);
            cell.setCellValue("TOT18 $");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(4);
            cell.setCellValue("TOT17 $");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(5);
            cell.setCellValue("DIFERENCIA $");
            cell.setCellStyle(styleCurrencyFormat);
            cell = row.createCell(6);
            cell.setCellValue("PORCENTAJE %");
            cell.setCellStyle(styleCurrencyFormat);
            int anchoColums = 2000;
            sheet.setColumnWidth(0, 5000);
            sheet.setColumnWidth(1, anchoColums);
            sheet.setColumnWidth(2, anchoColums);
            sheet.setColumnWidth(3, 4000);
            sheet.setColumnWidth(4, 4000);
            sheet.setColumnWidth(5, 4000);
            sheet.setColumnWidth(6, 4000);

            int i = 2;
            int totalTiquetes = 0, totalTiquetes2 = 0;
            double Vfinal = 0, Vfinal2 = 0, dif = 0, porcent = 0;
            DecimalFormat formateador = new DecimalFormat("###,###.##");
            System.out.println("actual " + Actual.size() + " Anterior " + Anterior.size());

            for (int j = 0; j < Actual.size(); j++) {
                row = sheet.createRow(i);
                cell = row.createCell(0);
                cell.setCellValue(Actual.get(j).getAgencia());

                cell = row.createCell(1);
                cell.setCellValue(Actual.get(j).getPasajes());
                totalTiquetes += Actual.get(j).getPasajes();

                cell = row.createCell(2);
                cell.setCellValue(Anterior.get(j).getPasajes());
                totalTiquetes2 += Anterior.get(j).getPasajes();

                cell = row.createCell(3);
                cell.setCellValue("$ " + formateador.format(Actual.get(j).getVfinal()));
                Vfinal += Actual.get(j).getVfinal();

                cell = row.createCell(4);
                cell.setCellValue("$ " + formateador.format(Anterior.get(j).getVfinal()));
                Vfinal2 += Anterior.get(j).getVfinal();

                cell = row.createCell(5);
                cell.setCellValue("$ " + formateador.format(Actual.get(j).getVfinal() - Anterior.get(j).getVfinal()));
                dif += Actual.get(j).getVfinal() - Anterior.get(j).getVfinal();

                cell = row.createCell(6);
                if (Actual.get(j).getVfinal() > 0 && Anterior.get(j).getVfinal() > 0) {
                    cell.setCellValue(Math.ceil(new Double(Actual.get(j).getVfinal()) / new Double(Anterior.get(j).getVfinal()) * 100) + " %");
                } else {
                    cell.setCellValue(0 + " %");
                }

//                porcent += ((new Double(Actual.get(j).getVfinal()) / new Double(Anterior.get(j).getVfinal())) * 100);
                //System.out.println("Actual.get(j).getVfinal() - Anterior.get(j).getVfinal()" + new Float((Actual.get(j).getVfinal() / Anterior.get(j).getVfinal())));
                i++;
            }

            // Create SUM formula
            styleCurrencyFormat = null;
            styleCurrencyFormat = workbook.createCellStyle();
            styleCurrencyFormat.setFont(font);
            System.out.println("i = " + i);
            row = sheet.createRow(i + 1);

            cell = row.createCell(0);
            cell.setCellValue("Total = ");
            cell.setCellStyle(styleCurrencyFormat);

            cell = row.createCell(1);
            cell.setCellValue("" + totalTiquetes);

            cell = row.createCell(2);
            cell.setCellValue("" + totalTiquetes2);

            cell = row.createCell(3);
            cell.setCellValue("$ " + formateador.format(Vfinal));

            cell = row.createCell(4);
            cell.setCellValue("$ " + formateador.format(Vfinal2));

            cell = row.createCell(5);
            cell.setCellValue("$ " + formateador.format(dif));

            cell = row.createCell(6);
            if (Vfinal > 0 && Vfinal2 > 0) {
                cell.setCellValue(Math.ceil(Vfinal / Vfinal2 * 100) + " %");
            } else {
                cell.setCellValue(0 + " %");
            }

            //Escribimos el resultado en fichero
//            
            try {
//                 String nombreFichero = "Z:/DatosExcel.xlsx";//unidad de red al 10.1
                String nombreFichero = "Z:/DatosExcel.xlsx";
                FileOutputStream out = new FileOutputStream(new File(nombreFichero));
                workbook.write(out);
                out.close();
                System.out.println("Se creo el archivo correctamente..");
            } catch (IOException e) {
                System.out.println("error " + e.getMessage());
            }

//Envia el archivo desde sqlserver
            //si se desea agregar un correo mas , al final de esta linea se puede hacer
            String correos = "desarrollo1@expresopalmira.com.co;jose.manzi@expresopalmira.com.co;cristina.ocana@expresopalmira.com.co";
            cstmt = cn.prepareCall("{call EnviarEmailConAdjuntos (?,?,?,?,?)}");
            cstmt.setString(1, "Email Sistemas");
            cstmt.setString(2, correos);
            cstmt.setString(3, "Reporte Venta Diaria");
            cstmt.setString(4, "Se adjunta el reporte de venta diaria Expreso Palmira");
            cstmt.setString(5, "\\\\192.168.10.1\\ExcelParaEmail\\DatosExcel.xlsx");
            cstmt.execute();
        } catch (SQLException e) {
            System.out.println("error " + e.getMessage());
        } finally {
            cn.close();
            pstm.close();
        }
    }

}
