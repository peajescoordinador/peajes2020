/*
 * Copyright 2019 Coordinador Electrico Nacional
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cl.coordinador.peajes;

/**
 * Clase estatica para despliegue de constantes y enums para el calculo de
 * peajes
 *
 * @author Frank Leanez at www.flconsulting.cl
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
    
    /**
     * Caracteres que definen el separador de la llave del archivo de
     * propiedades de configuracion
     */
    protected static final String KEY_SEPARATOR = "::";

    /**
     * Tipo de datos para usar en el archivo de configuracion
     */
    public enum DataType {

        /**
         * Datos tipo String (para archivo de opciones)
         */
        STRING,

        /**
         * Datos tipo DOUBLE (para archivo de opciones)
         */
        DOUBLE,

        /**
         * Datos tipo INTEGER (para archivo de opciones)
         */
        INTEGER,

        /**
         * Datos tipo BOOLEAN (para archivo de opciones)
         */
        BOOLEAN;
    }
    
    /**
     * Tipo de usos de memoria
     */
    public enum UsoMemoria {
        /**
         * Auto: Decide la applicacion
         */
        Auto, 
        /**
         * Minimo: Uso de disco
         */
        Min, 
        /**
         * Maximo: Todo en memoria (por ahora)
         */
        Max;
    }
    
    /**
     * Horizontes de calculos para liquidaciones: Mensual o Anual
     */
    public enum HorizonteCalculo {

        /**
         * Mensual
         */
        Mensual,

        /**
         * Anual
         */
        Anual
    }
    
}
