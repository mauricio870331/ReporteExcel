/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Model;

/**
 *
 * @author clopez
 */
public class Data {
    private String agencia;
    private int pasajes;
    private Double importeBase;
    private Double descuentos;
    private int Vfinal;

    public Data(String agencia, int pasajes, Double importeBase, Double descuentos, int Vfinal) {
        this.agencia = agencia;
        this.pasajes = pasajes;
        this.importeBase = importeBase;
        this.descuentos = descuentos;
        this.Vfinal = Vfinal;
    }

    public int getVfinal() {
        return Vfinal;
    }

    public void setVfinal(int Vfinal) {
        this.Vfinal = Vfinal;
    }

    public String getAgencia() {
        return agencia;
    }

    public void setAgencia(String agencia) {
        this.agencia = agencia;
    }

    public int getPasajes() {
        return pasajes;
    }

    public void setPasajes(int pasajes) {
        this.pasajes = pasajes;
    }

    public Double getImporteBase() {
        return importeBase;
    }

    public void setImporteBase(Double importeBase) {
        this.importeBase = importeBase;
    }

    public Double getDescuentos() {
        return descuentos;
    }

    public void setDescuentos(Double descuentos) {
        this.descuentos = descuentos;
    }

    @Override
    public String toString() {
        return "Data{" + "agencia=" + agencia + ", pasajes=" + pasajes + ", importeBase=" + importeBase + ", descuentos=" + descuentos + ", Vfinal=" + Vfinal + '}';
    }
    
    
    
}
