/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package practicasi2poi;

/**
 *
 * @author David Gonzalez Alvarez
 * @author Santiago Martinez Lopez
 * @version 1.0
 */
public class Trabajador {

    private String nifnie;
    private String nombre;
    private String primerApellido;
    private String segundoApellido;
    private String empresa;
    private String categoria;

    public Trabajador() {

    }

    public Trabajador(String nifnie, String nombre, String primerApellido, String segundoApellido, String empresa, String categoria) {
        this.nifnie = nifnie;
        this.nombre = nombre;
        this.primerApellido = primerApellido;
        this.segundoApellido = segundoApellido;
        this.empresa = empresa;
        this.categoria = categoria;
    }

    public String getNifnie() {
        return nifnie;
    }

    public void setNifnie(String nifnie) {
        this.nifnie = nifnie;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getPrimerApellido() {
        return primerApellido;
    }

    public void setPrimerApellido(String primerApellido) {
        this.primerApellido = primerApellido;
    }

    public String getSegundoApellido() {
        return segundoApellido;
    }

    public void setSegundoApellido(String segundoApellido) {
        this.segundoApellido = segundoApellido;
    }

    public String getEmpresa() {
        return empresa;
    }

    public void setEmpresa(String empresa) {
        this.empresa = empresa;
    }

    public String getCategoria() {
        return categoria;
    }

    public void setCategoria(String categoria) {
        this.categoria = categoria;
    }
}
