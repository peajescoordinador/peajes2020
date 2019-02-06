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

import static cl.coordinador.peajes.PeajesConstant.INIT_SIZE_ARRAY;
import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.MAX_COMPRESSION_RATIO;
import static cl.coordinador.peajes.PeajesConstant.MESES;
import static cl.coordinador.peajes.PeajesConstant.SLASH;
import java.io.File;

/**
 * 
 * @author aramos
 */
public class Peajes {

    public static void calculaPeajes(File DirEntrada, File DirSalida, int Ano) {

        String DirBaseEnt = DirEntrada.toString();
        String DirBaseSal = DirSalida.toString();
        String libroEntrada = DirBaseEnt + SLASH + "Ent" + Ano + ".xlsx";
        System.out.println(libroEntrada);
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);
        /**********
         * lee VATT
         **********/
//        String LibroAVI= DirBaseEnt + SLASH +"AVI_COMA.xls"; //Deprecated
        //Lee.leeEscribeArchivoVATT(libroEntrada,LibroAVI, Ano);

        double[][] Aux = new double[INIT_SIZE_ARRAY][NUMERO_MESES];
        String[] TxtTemp1 = new String[INIT_SIZE_ARRAY];
        String[] TxtTemp2 = new String[INIT_SIZE_ARRAY];
        int numLineasVATT = Lee.leeVATT(libroEntrada, TxtTemp1, TxtTemp2, Aux);
        double[][] VATT = new double[numLineasVATT][NUMERO_MESES];
        String[] nomLinVATT = new String[numLineasVATT];
        String[] nomProp = new String[numLineasVATT];
        String[] TxtTemp3 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp3[i] = "";
        }
        int numTx = 0;
        for (int i = 0; i < numLineasVATT; i++) {
            nomLinVATT[i] = TxtTemp1[i];
            nomProp[i] = TxtTemp2[i];
            int t = Calc.Buscar(nomProp[i], TxtTemp3);
            if (t == -1) {
                TxtTemp3[numTx] = nomProp[i];
                numTx++;
            }
            for (int j = 0; j < NUMERO_MESES; j++) {
                VATT[i][j] = Aux[i][j];
            }
        }
        String[] TxtTemp4 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp4[i] = "";
        }
        int numLinTx = 0;
        for (int i = 0; i < numLineasVATT; i++) {
            int l = Calc.Buscar(nomLinVATT[i] + "#" + nomProp[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinVATT[i] + "#" + nomProp[i];
                numLinTx++;
            }
        }
        String[] nombreTx = new String[numTx];
        for (int i = 0; i < numTx; i++) {
            nombreTx[i] = TxtTemp3[i];
        }
        String[] nomLinTx = new String[numLinTx];
        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
        }

        /************
         * lee Líneas
         *************/
        TxtTemp1 = new String[INIT_SIZE_ARRAY];
        int numLineas = Lee.leeDeflin(libroEntrada, TxtTemp1, Aux);
        double[][] paramLineas = new double[numLineas][10];
        String[] nomLin = new String[numLineas];
        for (int i = 0; i < numLineas; i++) {
            for (int j = 0; j <= 8; j++) {
                paramLineas[i][j] = Aux[i][j];
            }
            nomLin[i] = TxtTemp1[i];
        }
         /*************
         * lee Indices
         *************/
        double[] dolar = new double[NUMERO_MESES];
        double[] interes = new double[NUMERO_MESES];
       // Lee.leeEscribeIndices(libroEntrada, LibroAVI, Ano);
        Lee.leeIndices(libroEntrada, dolar, interes);

        /**********************
         * lee Líneas Troncales
         **********************/
        TxtTemp1 = new String[INIT_SIZE_ARRAY];
        TxtTemp2 = new String[INIT_SIZE_ARRAY];
        int[] intAux1 = new int[INIT_SIZE_ARRAY];
        int[] numLineasIT=new int[1];
        double[][] ITEAux = new double[INIT_SIZE_ARRAY][NUMERO_MESES];
        double[][] ITEGAux = new double[INIT_SIZE_ARRAY][NUMERO_MESES];
        double[][] ITERAux = new double[INIT_SIZE_ARRAY][NUMERO_MESES];
        double[][] ITPAux = new double[INIT_SIZE_ARRAY][NUMERO_MESES];
        int numLinT = Lee.leeLintronIT2(libroEntrada, TxtTemp1, TxtTemp2,
                nomLin, intAux1, ITEAux,ITEGAux,ITERAux, ITPAux,numLineasIT);
        String[] nomLinIT = new String[numLinT];
        String[] nomLinIT_Tx = new String[numLinT];
        int[] indiceLintron = new int[numLinT];
        double[][] ITEi = new double[numLinT][NUMERO_MESES];
        double[][] ITEGi = new double[numLinT][NUMERO_MESES];
        double[][] ITERi = new double[numLinT][NUMERO_MESES];
        double[][] ITPi = new double[numLinT][NUMERO_MESES];
        double[][] ITEo = new double[numLinT][NUMERO_MESES];
        double[][] ITEGo = new double[numLinT][NUMERO_MESES];
        double[][] ITERo = new double[numLinT][NUMERO_MESES];
        double[][] ITPo = new double[numLinT][NUMERO_MESES];
        double[][] ITo = new double[numLinT][NUMERO_MESES];


        for (int i = 0; i < numLinT; i++) {
            nomLinIT[i] = TxtTemp1[i];
            nomLinIT_Tx[i] = TxtTemp2[i];
            indiceLintron[i] = intAux1[i];
            for (int m = 0; m < NUMERO_MESES; m++) {
                ITEi[i][m] = ITEAux[i][m]* (1 + interes[m]);
                ITEGi[i][m] = ITEGAux[i][m]* (1 + interes[m]);
                ITERi[i][m] = ITERAux[i][m]* (1 + interes[m]);
                ITPi[i][m] = ITPAux[i][m]* (1 + interes[m]);
            }
        }
       

        /******************************************
         * Calcula Peajes
         ******************************************/
        double[][] peajeLin = new double[numLinTx][NUMERO_MESES];
        double[][] VATTLin = new double[numLinTx][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
                int l = Calc.Buscar(nomLinTx[i], nomLinIT_Tx);
                if(l==-1){
                    System.out.println("Error!!!");
                    System.out.println("El tramo - "+nomLinTx[i]+" - de la hoja 'VATT' en archivo AVI_COMA.xls no posee una instalación Toncal con IT asociado en la hoja 'lintron'");
                    System.out.println("Corregir y ejecutar nuevamente el botón Peajes");
                    System.out.println();
                }
                else {
                    for (int m = 0; m < NUMERO_MESES; m++) {
                    peajeLin[i][m] += -(ITEi[l][m]);
                    peajeLin[i][m] += -(ITPi[l][m]);
                    ITEo[i][m]=ITEi[l][m];
                    ITEGo[i][m]=ITEGi[l][m];
                    ITERo[i][m]=ITERi[l][m];
                    ITPo[i][m]=ITPi[l][m];
                    ITo[i][m]=ITEo[i][m] + ITPo[i][m];
                }
            }
        }
        for (int i = 0; i < numLineasVATT; i++) {
            int l = Calc.Buscar(nomLinVATT[i] + "#" + nomProp[i], nomLinTx);
                for (int m = 0; m < NUMERO_MESES; m++) {
                    peajeLin[l][m] += VATT[i][m] * dolar[m] * (1 + interes[m]) * 1000;
                    VATTLin[l][m]+= VATT[i][m] * dolar[m] * (1 + interes[m]) * 1000;
                    //peajeLin[l][m] += VATT[i][m] * (1 + interes[m]) * 1000;
                    //VATTLin[l][m]+= VATT[i][m] * (1 + interes[m]) * 1000;
                }
        }
        //System.out.println(numLinTx+" "+numLinT);

        /*
         * Escritura de Resultados
         * =======================
         */
        String libroSalidaGXLS = DirBaseSal + SLASH + "Peaje" + Ano + ".xlsx";
        Escribe.crearLibro(libroSalidaGXLS);
        Escribe.creaH1F_2d_long(
                "Peajes [$]", peajeLin,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "Peajes", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía Asignable a Generadores[$]", ITEGo,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "ITEG", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía Asignable a Retiro[$]", ITERo,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "ITER", "#,###,##0;[Red]-#,###,##0;\"-\"");
        
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía [$]", ITEo,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "ITE", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Potencia [$]", ITPo,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "ITP", "#,###,##0;[Red]-#,###,##0;\"-\"");
        //modificar por formulas
        
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario [$]", ITo,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "IT", "#,###,##0;[Red]-#,###,##0;\"-\"");
        // \modificar por formulas
               
        Escribe.creaH1F_2d_long(
                "VATT [$]", VATTLin,
                "Línea", nomLinTx,
                "Mes", MESES,
                libroSalidaGXLS, "VATT", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_verifProrr(peajeLin,
                numLineasIT[0], nomLinTx,
                libroEntrada, "verProrr","0.000%;[Red]-0.000%;\"-\"",0);
        System.out.println("Peajes Calculados");
        System.out.println();
    }
}
