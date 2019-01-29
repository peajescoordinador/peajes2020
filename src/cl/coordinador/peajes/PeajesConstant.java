/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package cl.coordinador.peajes;

/**
 *
 * @author www.flconsulting.cl
 */
public class PeajesConstant {

    /**
     * Representacion abreviada de los meses del año:
     * <li>0=Ene</li>
     * <li>1=Feb</li>
     * <li>2=Mar</li>
     * <li>3=Abr</li>
     * <li>4=May</li>
     * <li>5=Jun</li>
     * <li>6=Jul</li>
     * <li>7=Ago</li>
     * <li>8=Sep</li>
     * <li>9=Oct</li>
     * <li>10=Nov</li>
     * <li>11=Dic</li>
     */
    public static final String[] MESES = {"Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"};

    /**
     * Total numero de meses del año (12)
     */
    public static final int NUMERO_MESES = 12;

    /**
     * File path separator as defined by Operating System
     */
    public static final String SLASH = java.io.File.separator;

    /**
     * Dimension inicial por defecto de los arrays
     */
    public static final int INIT_SIZE_ARRAY = 1500;

    /**
     * Valor usado para detectar compresiones sospechosas.
     * <br>Valor por defecto es 1%. Usar 0% para permitir cualquier planilla
     * Excel
     * <br>TODO: move to config file. Workaround: Value zero will prevent zip
     * bomb exception in poi but it may become vulnerable to malicious data
     * <br>Note: There is a detected bug in poi library when dealing with large
     * complex equations
     */
    public static final int MAX_COMPRESSION_RATIO = 0;

}
