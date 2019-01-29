package cl.coordinador.peajes;


import java.io.*;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author aramos
 */
public class Peajes {

    private static String slash = File.separator;
    private static final int numMeses = 12;
    static String[] nomLinIT_Tx;

    public static void calculaPeajes(File DirEntrada, File DirSalida, int Ano) {

        String DirBaseEnt = DirEntrada.toString();
        String DirBaseSal = DirSalida.toString();
        String[] nomMes = {"Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul",
            "Ago", "Sep", "Oct", "Nov", "Dic"};
        String libroEntrada = DirBaseEnt + slash + "Ent" + Ano + ".xlsx";
        System.out.println(libroEntrada);
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(PeajesCDEC.MAX_COMPRESSION_RATIO);
        /**********
         * lee VATT
         **********/
        String LibroAVI= DirBaseEnt + slash +"AVI_COMA.xls";
        //Lee.leeEscribeArchivoVATT(libroEntrada,LibroAVI, Ano);


        double[][] Aux = new double[1500][numMeses];
        String[] TxtTemp1 = new String[1500];
        String[] TxtTemp2 = new String[1500];
        int numLineasVATT = Lee.leeVATT(libroEntrada, TxtTemp1, TxtTemp2, Aux);
        double[][] VATT = new double[numLineasVATT][numMeses];
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
            for (int j = 0; j < numMeses; j++) {
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
        TxtTemp1 = new String[1500];
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
        double[] dolar = new double[numMeses];
        double[] interes = new double[numMeses];
       // Lee.leeEscribeIndices(libroEntrada, LibroAVI, Ano);
        Lee.leeIndices(libroEntrada, dolar, interes);

        /**********************
         * lee Líneas Troncales
         **********************/
        TxtTemp1 = new String[1500];
        TxtTemp2 = new String[1500];
        int[] intAux1 = new int[1500];
        int[] numLineasIT=new int[1];
        double[][] ITEAux = new double[1500][numMeses];
        double[][] ITEGAux = new double[1500][numMeses];
        double[][] ITERAux = new double[1500][numMeses];
        double[][] ITPAux = new double[1500][numMeses];
        int numLinT = Lee.leeLintronIT2(libroEntrada, TxtTemp1, TxtTemp2,
                nomLin, intAux1, ITEAux,ITEGAux,ITERAux, ITPAux,numLineasIT);
        String[] nomLinIT = new String[numLinT];
        nomLinIT_Tx = new String[numLinT];
        int[] indiceLintron = new int[numLinT];
        double[][] ITEi = new double[numLinT][numMeses];
        double[][] ITEGi = new double[numLinT][numMeses];
        double[][] ITERi = new double[numLinT][numMeses];
        double[][] ITPi = new double[numLinT][numMeses];
        double[][] ITEo = new double[numLinT][numMeses];
        double[][] ITEGo = new double[numLinT][numMeses];
        double[][] ITERo = new double[numLinT][numMeses];
        double[][] ITPo = new double[numLinT][numMeses];
        double[][] ITo = new double[numLinT][numMeses];


        for (int i = 0; i < numLinT; i++) {
            nomLinIT[i] = TxtTemp1[i];
            nomLinIT_Tx[i] = TxtTemp2[i];
            indiceLintron[i] = intAux1[i];
            for (int m = 0; m < numMeses; m++) {
                ITEi[i][m] = ITEAux[i][m]* (1 + interes[m]);
                ITEGi[i][m] = ITEGAux[i][m]* (1 + interes[m]);
                ITERi[i][m] = ITERAux[i][m]* (1 + interes[m]);
                ITPi[i][m] = ITPAux[i][m]* (1 + interes[m]);
            }
        }
       

        /******************************************
         * Calcula Peajes
         ******************************************/
        double[][] peajeLin = new double[numLinTx][numMeses];
        double[][] VATTLin = new double[numLinTx][numMeses];
        for (int i = 0; i < numLinTx; i++) {
                int l = Calc.Buscar(nomLinTx[i], nomLinIT_Tx);
                if(l==-1){
                    System.out.println("Error!!!");
                    System.out.println("El tramo - "+nomLinTx[i]+" - de la hoja 'VATT' en archivo AVI_COMA.xls no posee una instalación Toncal con IT asociado en la hoja 'lintron'");
                    System.out.println("Corregir y ejecutar nuevamente el botón Peajes");
                    System.out.println();
                }
                else {
                    for (int m = 0; m < numMeses; m++) {
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
                for (int m = 0; m < numMeses; m++) {
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
        String libroSalidaGXLS = DirBaseSal + slash + "Peaje" + Ano + ".xlsx";
        Escribe.crearLibro(libroSalidaGXLS);
        Escribe.creaH1F_2d_long(
                "Peajes [$]", peajeLin,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "Peajes", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía Asignable a Generadores[$]", ITEGo,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "ITEG", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía Asignable a Retiro[$]", ITERo,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "ITER", "#,###,##0;[Red]-#,###,##0;\"-\"");
        
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Energía [$]", ITEo,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "ITE", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario Potencia [$]", ITPo,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "ITP", "#,###,##0;[Red]-#,###,##0;\"-\"");
        //modificar por formulas
        
        Escribe.creaH1F_2d_long(
                "Ingreso Tarifario [$]", ITo,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "IT", "#,###,##0;[Red]-#,###,##0;\"-\"");
        // \modificar por formulas
               
        Escribe.creaH1F_2d_long(
                "VATT [$]", VATTLin,
                "Línea", nomLinTx,
                "Mes", nomMes,
                libroSalidaGXLS, "VATT", "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_verifProrr(peajeLin,
                numLineasIT[0], nomLinTx,
                libroEntrada, "verProrr","0.000%;[Red]-0.000%;\"-\"",0);
        System.out.println("Peajes Calculados");
        System.out.println();
    }
}
