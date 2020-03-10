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

import static cl.coordinador.peajes.PeajesConstant.MAX_COMPRESSION_RATIO;
import static cl.coordinador.peajes.PeajesConstant.MESES;
import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.SLASH;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author vtoro
 */
public class PeajesIny {

    private static final boolean USE_MEMORY_READER = true; //swtich para usar nuevo API lectura poi
    private static final boolean USE_MEMORY_WRITER = true; //swtich para usar nuevo API escritura poi
    
    static double peajeEmpTxO[][][]; //TODO: Encapsulate
    private static double prorrMesGO[][][];
    private static double peajeLinGO[][][];
    private static double ItLinGO[][][];
    static double peajeCenTxO[][][]; //TODO: Encapsulate
    private static double peajeCenO[][];
    static double peajeEmpGO[][]; //TODO: Encapsulate
    static float MGNCO[]; //TODO: Encapsulate
    static double[] PotNetaO; //TODO: Encapsulate
    static float[] FAERO; //TODO: Encapsulate
    static float[] CETO; //TODO: Encapsulate
    static float[] Tabla1O; //TODO: Encapsulate
    static double[][] GenPromMesCenO; //TODO: Encapsulate
    static double[][] facPagoO; //TODO: Encapsulate
    static double[][] peajeGenO ; //TODO: Encapsulate
    private static double[][] ExcTotCenO;
    static double[][] AjusMGNCTotO; //TODO: Encapsulate
    static double[][] PagoTotCenO; //TODO: Encapsulate
    private static double[][][] ExcenCenO;
    static double[][][] AjusMGNCTxO; //TODO: Encapsulate
    static double[][][] PagoTotCenTxO; //TODO: Encapsulate
    static double[][][] AjusMGNCEmpTxO; //TODO: Encapsulate
    static double[][][] PagoEmpTxO; //TODO: Encapsulate
    static double[][][] ItEmpTxO; //TODO: Encapsulate
    static double[][] PagoAnualEmpGTxO; //TODO: Encapsulate
    static double[][] PeajeAnualEmpGTxO; //TODO: Encapsulate
    static double[][] AjusMGNCEmpO; //TODO: Encapsulate
    static double[][] PagoEmpO; //TODO: Encapsulate
    static double[][] ItEmpO; //TODO: Encapsulate
    static double[] PagoAnualEmpGO; //TODO: Encapsulate
    static double[] PeajeAnualEmpGO; //TODO: Encapsulate
    static String[] nomGen; //TODO: Encapsulate
    static String[] nombreTx; //TODO: Encapsulate
    static String[] nomEmpGO; //TODO: Encapsulate
    static String[] nomLinTx; //TODO: Encapsulate
    static String[] nomGenO; //TODO: Encapsulate
    static String[] nomMGNCO; //TODO: Encapsulate
    private static String DirBaseSal;
    private static  String[] nomLinIT;
    private static String[] nomLineasN;
    private static  int[] zonaLinIT;
    private static  int[] zonaLinPe;
    static  int[] zonaLinTx; //TODO: Encapsulate
    static double[][][] prorrMesGenTx; //TODO: Encapsulate
    static double[][] prorrMesGenTxTot; //TODO: Encapsulate
    static double[][] peajeAnualMGNCTxO; //TODO: Encapsulate
    static double[] peajeAnualMGNCO; //TODO: Encapsulate
    static double[] PotNetaMGNCO; //TODO: Encapsulate
    static double[][] ExcenAnualMGNCTxO; //TODO: Encapsulate
    static double[] ExcenAnualMGNCO; //TODO: Encapsulate
    static double[] facPagoMGNCO; //TODO: Encapsulate
    static double[][]  AjusMGNCAnualEmpTxO; //TODO: Encapsulate
    static double[]  AjusMGNCAnualEmpO; //TODO: Encapsulate
    static double[][][] ExcenMGNCTxO; //TODO: Encapsulate
    static double[][] ExcenMGNCO; //TODO: Encapsulate
    private static boolean cargandoInfo=false;
    private static boolean calcPagos=false;
    private static boolean EscribirPagos=false;

    public static void calculaPeajesIny(File DirEntrada, File DirSalida, int Ano, boolean LiquidacionReliquidacion) {
        calculaPeajesIny(PeajesConstant.HorizonteCalculo.Anual, DirEntrada, DirSalida, Ano, 0, LiquidacionReliquidacion);
    }
    
    public static void calculaPeajesIny(PeajesConstant.HorizonteCalculo horizon, File DirEntrada, File DirSalida, int Ano, int Mes, boolean LiquidacionReliquidacion) {
        String DirBaseEnt = DirEntrada.toString();
        DirBaseSal = DirSalida.toString();
        DecimalFormat DosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();
        cargandoInfo=true;
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);
        
        //Lee opciones de configuracion:
        int maxLeeLINEA = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de LINEAS leer en planillas Ent",PeajesConstant.DataType.INTEGER));
        int maxLeeGEN = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de CENTRALES leer en planillas Ent",PeajesConstant.DataType.INTEGER));
        int maxLeeZONA = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de ZONAS a leer en planillas Ent",PeajesConstant.DataType.INTEGER));
        
        // Libro peajes
        String libroEntradaPeajes = DirBaseSal + SLASH + "Peaje" + Ano + ".xlsx";
        XSSFWorkbook wb_Peajes;
        try {
            wb_Peajes = new XSSFWorkbook(new java.io.FileInputStream( libroEntradaPeajes ));
        } catch (IOException e) {
            System.out.println("No se pudo conectar con planilla entrada " + libroEntradaPeajes);
            System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
            return;
        }

        /************
         * lee Peajes e IT
         ************/
        double[][] longAux = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxIT = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxITR = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxVATT = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxITP = new double[maxLeeLINEA][NUMERO_MESES];
        String[] TxtTemp = new String[maxLeeLINEA];
        String[] TxtTempIT = new String[maxLeeLINEA];
        int numLinea;
        int numLineaVATT;
        int numLineaIT;
        int numLineaITR;
        int numLineaITP;
        if (USE_MEMORY_READER) {
            numLinea = Lee.leePeajes(wb_Peajes, TxtTemp, longAux);
            numLineaIT = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxIT, "ITEG");
            numLineaITR = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxITR, "ITER");
            numLineaVATT = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxVATT, "VATT");
            numLineaITP = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxITP, "ITP");
        } else {
            numLinea = Lee.leePeajes(libroEntradaPeajes, TxtTemp, longAux);
            numLineaIT = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxIT, "ITEG");
            numLineaITR = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxITR, "ITER");
            numLineaVATT = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxVATT, "VATT");
            numLineaITP = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxITP, "ITP");
        }
        nomLineasN = new String[numLinea];
        double[][] PeajeN = new double[numLinea][12];
        double[][] VATTN = new double[numLineaVATT][12];
        double[][] ITP_N = new double[numLineaITP][12];
        double[][] IT_N = new double[numLineaIT][12];
        double[][] IT_NR = new double[numLineaITR][12];
        double[] PeajeNMes = new double[12];
        double[] VATTNMes = new double[12];
        double[] ITP_N_Mes = new double[12];
        double[] IT_N_Mes = new double[12];
        for (int i = 0; i < numLinea; i++) {
            nomLineasN[i] = TxtTemp[i];
            for (int j = 0; j < 12; j++) {
                PeajeN[i][j] = longAux[i][j];
                PeajeNMes[j]+=PeajeN[i][j];
                VATTN[i][j]= longAuxVATT[i][j];
                VATTNMes[j]+=VATTN[i][j];
                IT_N[i][j] = longAuxIT[i][j];
                IT_N_Mes[j]+=IT_N[i][j];
                ITP_N[i][j] = longAuxITP[i][j];
                ITP_N_Mes[j]+=ITP_N[i][j];
                IT_NR[i][j] = longAuxITR[i][j];
            }
        }

        // Libro Ent
        String libroEntrada = DirBaseEnt + SLASH + "Ent" + Ano + ".xlsx";
        XSSFWorkbook wb_Ent;
        try {
            wb_Ent = new XSSFWorkbook(new java.io.FileInputStream( libroEntrada ));
        } catch (IOException e) {
            System.out.println("No se pudo conectar con planilla entrada " + libroEntrada);
            System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
            return;
        }

        /**********
         * lee VATT
         **********/
        double[][] Aux = new double[maxLeeLINEA][NUMERO_MESES];
        String[] TxtTemp1 = new String[maxLeeLINEA];
        String[] TxtTemp2 = new String[maxLeeLINEA];
        int numLineasVATT ;
        if (USE_MEMORY_READER) {
            numLineasVATT = Lee.leeVATT(wb_Ent, TxtTemp1, TxtTemp2, Aux);
        } else {
            numLineasVATT = Lee.leeVATT(libroEntrada, TxtTemp1, TxtTemp2, Aux);
        }
        String[] nomProp = new String[numLineasVATT];
        String[] TxtTemp3 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp3[i] = "";
        }
        int numTx = 0; //numero de transmisores
        for (int i = 0; i < numLineasVATT; i++) {
//            nomLinVATT[i] = TxtTemp1[i];
            nomProp[i] = TxtTemp2[i];
            int t = Calc.Buscar(nomProp[i], TxtTemp3);
            if (t == -1) {
                TxtTemp3[numTx] = nomProp[i];
                numTx++;
            }
        }
        /*String[] TxtTemp4 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp4[i] = "";
        }
        numLinTx = 0;
        for (int i = 0; i < numLineasVATT; i++) {
            int l = Calc.Buscar(nomLinVATT[i] + "#" + nomProp[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinVATT[i] + "#" + nomProp[i];
                numLinTx++;
            }
        }
         *
         */
        nombreTx = new String[numTx];
        System.arraycopy(TxtTemp3, 0, nombreTx, 0, numTx); /*
        nomLinTx = new String[numLinTx];
        for (int i = 0; i < numLinTx; i++) {
        nomLinTx[i] = TxtTemp4[i];
        }
        // Ordena
        int[] nt = Calc.OrdenarBurbujaStr(nomLinTx);
        String[] nomLinTxO = new String[numLinTx];
        for (int i = 0; i < numLinTx; i++) {
        nomLinTxO[i] = nomLinTx[nt[i]];
        //System.out.println(nomLinTxO[i]);
        }
         *
         */

        /***************
         * lee Centrales
         ***************/
        TxtTemp1 = new String[maxLeeGEN];
        double PotNetaTot=0;
        float[] Temp1 = new float[maxLeeGEN];
        float[] Temp2= new float[maxLeeGEN];
        float[] Temp3= new float[maxLeeGEN];
        float[] Temp4= new float[maxLeeGEN];
        float[] Temp5= new float[maxLeeGEN];
        int numCen;
        if (USE_MEMORY_READER) {
            numCen = Lee.leeCentrales(wb_Ent, TxtTemp1,Temp1,Temp2,Temp3,Temp4,Temp5);
        } else {
            numCen = Lee.leeCentrales(libroEntrada, TxtTemp1,Temp1,Temp2,Temp3,Temp4,Temp5);
        }
        nomGen = new String[numCen];
        double[] PotNeta = new double[numCen];
        float[] MGNC = new float[numCen];
        float[] FAER = new float[numCen];
        float[] CET = new float[numCen];
        float[] Tabla1 = new float[numCen];
        int numMGNC=0;
//        int[] indMGNC=new int[numCen];

        for (int i = 0; i < numCen; i++) {
            nomGen[i] = TxtTemp1[i];
            PotNeta[i] = Temp1[i];
            PotNetaTot+=PotNeta[i];
            MGNC[i] = Temp2[i];
            FAER[i] = Temp3[i];
            CET[i] = Temp4[i];
            Tabla1[i] = Temp5[i];
            //if(MGNC[i]==1){
//                indMGNC[numMGNC]=i;
                numMGNC++;
            //}
        }


        TxtTemp1 = new String[numCen];
        for (int i = 0; i < numCen; i++) {
            TxtTemp1[i] = "";
        }
        int numEmpG = 0;
        for (int j = 0; j < numCen; j++) {
            String[] tmp = nomGen[j].split("#");
            int l = Calc.Buscar(tmp[0], TxtTemp1);
            if (l == -1) {
                TxtTemp1[numEmpG] = tmp[0];
                numEmpG++;
            }
        }
       
        String[] nomEmp = new String[numEmpG];
        System.arraycopy(TxtTemp1, 0, nomEmp, 0, numEmpG);

        /************
         * lee Lineas
         *************/
        TxtTemp1 = new String[maxLeeLINEA];
        Aux = new double[maxLeeLINEA][NUMERO_MESES];
        int numLineas;
        if (USE_MEMORY_READER) {
            numLineas = Lee.leeDeflin(wb_Ent, TxtTemp1, Aux);
        } else {
            numLineas = Lee.leeDeflin(libroEntrada, TxtTemp1, Aux);
        }
        double[][] paramLineas = new double[numLineas][10];
        String[] nomLin = new String[numLineas];
        for (int i = 0; i < numLineas; i++) {
            for (int j = 0; j <= 8; j++) {
                paramLineas[i][j] = Aux[i][j];
            }
            nomLin[i] = TxtTemp1[i];
        }

        /**********************
         * lee Lineas Troncales
         **********************/
        TxtTemp1 = new String[maxLeeLINEA];
        TxtTemp2 = new String[maxLeeLINEA];
        int[] intAux1 = new int[maxLeeZONA];
        int[][] intAux2 = new int[maxLeeZONA][NUMERO_MESES];
        int numLinIT;
        if (USE_MEMORY_READER) {
            numLinIT = Lee.leeLintron(wb_Ent, TxtTemp1, nomLin, TxtTemp2,intAux1, intAux2);
        } else {
            numLinIT = Lee.leeLintron(libroEntrada, TxtTemp1, nomLin, TxtTemp2,intAux1, intAux2);
        }
        String TxtTemp4[]=new String[numLinIT];
        nomLinIT = new String[numLinIT];
        zonaLinIT= new int[numLinIT];        
        int[] indZonaLinIT=new int[numLinIT];
//        String[] nomZonaLinIT=new String[numLinIT];
        String[] propietario = new String[numLinIT];
        for (int i = 0; i < numLinIT; i++) {
            TxtTemp4[i]="";
            nomLinIT[i] = TxtTemp1[i];
            zonaLinIT[i]=intAux2[i][0];
            propietario[i]=TxtTemp2[i];
            if(zonaLinIT[i]==1){
            indZonaLinIT[i]=0;
//            nomZonaLinIT[i]="N";
            }
            else if(zonaLinIT[i]==0){
            indZonaLinIT[i]=1;
//            nomZonaLinIT[i]="A";
            }
            else if(zonaLinIT[i]==-1){
            indZonaLinIT[i]=2;
//            nomZonaLinIT[i]="S";
            }
        }
        int[] indZonaLinPe=new int[numLinea];
        zonaLinPe= new int[numLinea];
        for (int i = 0; i < numLinea; i++) {
         String[] tmp = nomLineasN[i].split("#");
            int l = Calc.Buscar(tmp[0], nomLinIT);
            if(l==-1){
             System.out.println("Error!!!");
             System.out.println("Línea Troncal - "+tmp[0]+" - en archivo Peaje"+Ano+".xls no se encuentra en la hoja 'lintron' del archivo Ent"+Ano+".xlsx");
             System.out.println("Debe asegurarse que las Líneas del archivo AVI_COMA.xls se encuentren en la hoja 'lintron' y ejecutar el botón Peajes");
            }
            else{
            zonaLinPe[i]=zonaLinIT[l];
            indZonaLinPe[i]=indZonaLinIT[l];
            }
        }

//         TxtTemp3=new String[numLinIT];


        int numLinTx = 0;
        for (int i = 0; i <numLinIT; i++) {
            int l = Calc.Buscar(nomLinIT[i] + "#" + propietario[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinIT[i] + "#" + propietario[i];
                TxtTemp2[numLinTx]=propietario[i];
                numLinTx++;
            }
        }
        nomLinTx = new String[numLinTx];//solo registros unicos Linea-Transmisor de hoja lintron
//        String[] nomPropTx = new String[numLinTx];

        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
            //System.out.println(nomLinTx[i]);
//            nomPropTx[i]=TxtTemp2[i];
        }
//        int[] indZonaLinTx=new int[numLinTx];
        zonaLinTx= new int[numLinTx];
         for (int i = 0; i < numLinTx; i++) {
         String[] tmp = nomLinTx[i].split("#");
            int l2 = Calc.Buscar(tmp[0], nomLinIT);
            zonaLinTx[i]=zonaLinIT[l2];
//            indZonaLinTx[i]=indZonaLinIT[l2];
         }

        
        /*****************************
         * lee Prorratas de Generacion e Inyeccion Centrales
         *****************************/
        prorrMesGenTx = new double[numLinTx][numCen][NUMERO_MESES];
        double[][] GenerMensual = new double[numCen][NUMERO_MESES];
        
        //Lee prorratas desde csv (defecto):
        long tInicioLecturaProrratas = System.currentTimeMillis();
        System.out.println("Inicio lectura prorratas.."); //TEMP!
        File f_prorratasGCSV;
        if (horizon == PeajesConstant.HorizonteCalculo.Anual) {
            f_prorratasGCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_PRORRATAGEN + Ano + ".csv");
        } else {
            f_prorratasGCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_PRORRATAGEN + Ano + MESES[Mes] + ".csv");
        }
        File f_GMesCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_GMES + Ano + ".csv");
        if (f_prorratasGCSV.exists() && f_GMesCSV.exists()) {
            System.out.println("Leyendo archivos csv de prorratas y generacion mensual..");
            try {
                int nReadP = Lee.leeProrratasCSV(f_prorratasGCSV.getAbsolutePath(), prorrMesGenTx, horizon);
            } catch (IOException e) {
                System.out.println("No se pudo conectar con archivo prorratas " + f_prorratasGCSV.getAbsolutePath());
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
            try {
                int nReadGx = Lee.leeGeneracionMesCSV(f_GMesCSV.getAbsolutePath(), GenerMensual, horizon);
            } catch (IOException e) {
                System.out.println("No se pudo conectar con archivo generacion mensual " + f_GMesCSV.getAbsolutePath());
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
        } else {
            // Si no, intenta buscar libro Prorrata Excel
            System.out.println("Leyendo prorratas y generacion mensual desde Excel..");
            String libroEntradaP = DirBaseSal + SLASH + "Prorrata" + Ano + ".xlsx";
            XSSFWorkbook wb_Prorratas;
            try {
                wb_Prorratas = new XSSFWorkbook(new java.io.FileInputStream(libroEntradaP));
                if (USE_MEMORY_READER) {
                    Lee.leeProrratasGxExcel(wb_Prorratas, prorrMesGenTx);
                } else {
                    Lee.leeProrratasGx(libroEntradaP, prorrMesGenTx);
                }
                if (USE_MEMORY_READER) {
                    Lee.leeGeneracionMes(wb_Prorratas, GenerMensual);
                } else {
                    Lee.leeGeneracionMes(libroEntradaP, GenerMensual);
                }
                wb_Prorratas.close();
            } catch (IOException e) {
                System.out.println("No se pudo conectar con planilla prorratas " + libroEntradaP);
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
        }
        long tFinLecturaProrratas = System.currentTimeMillis();
        System.out.println("Finalizado lectura prorratas. Tiempo: " + (tFinLecturaProrratas - tInicioLecturaProrratas) / 1000 + " seg"); //TEMP!
        
        //Prorratas agregadas por linea:
        prorrMesGenTxTot = new double[numLinTx][NUMERO_MESES];
        for (int l = 0 ; l < numLinTx; l++){
            for (int m = 0 ; m < NUMERO_MESES; m ++ ){
                for (int c = 0 ; c < numCen; c ++ ){
                prorrMesGenTxTot [l][m]+=prorrMesGenTx[l][c][m];
                }
            }
        }
        
        //Inyeccion Centrales agregadas:
        double[] GeneTotMesProm= new double[NUMERO_MESES];
        double[][] GenPromMesCen= new double[numCen][NUMERO_MESES];
        double[] GenAnoxCen= new double[numCen];
        int [][] MesesAct=new int[numCen][NUMERO_MESES];
        int [] numMesesAct=new int[numCen];

        /*LEE MESES CENTRALES ACTIVO*/
        for (int i=0;i<numCen;i++){
            for(int m=0; m<NUMERO_MESES;m++){
                MesesAct[i][m]+=1;  //AGREGAR RUTINA Q LEE MANT CEN 
            }
        }
        
        
        for (int i = 0; i < numCen; i++) {
            for (int m = 0; m < NUMERO_MESES; m++) {
                MesesAct[i][m] = 0;
                if (GenerMensual[i][m] != 0) {
                    GenAnoxCen[i] += GenerMensual[i][m];
                    MesesAct[i][m] = 1;
                    numMesesAct[i] += 1;
                }
            }
            for (int m = 0; m < NUMERO_MESES; m++) {
                GenPromMesCen[i][m] = 0;
                if (MesesAct[i][m] == 1) {
                    GenPromMesCen[i][m] = GenAnoxCen[i] / numMesesAct[i];
                    GeneTotMesProm[m] += GenPromMesCen[i][m];
                }
            }
        }
        
        
        
        long tFinalLectura = System.currentTimeMillis();
        long tInicioCalculo = System.currentTimeMillis();
        cargandoInfo=false;
        calcPagos=true;

        /******************************************
         * Calcula Pagos por Inyeccion de Centrales
         ******************************************/
        double[][][] peajeLinCen = new double[numLinTx][numCen][NUMERO_MESES];
        double[][][]  ItLinCen = new double[numLinTx][numCen][NUMERO_MESES];
        double[][][][] peajeGenTxZona = new double[numCen][numTx][3][NUMERO_MESES];
        double[][][] peajeGenTx = new double[numCen][numTx][NUMERO_MESES];
        double[][][] peajeGenTxExcen = new double[numCen][numTx][NUMERO_MESES];
        double[][][] ItGenTxExcen = new double[numCen][numTx][NUMERO_MESES];
        double[][][] ItGenTx = new double[numCen][numTx][NUMERO_MESES];
        double[][] peajeGen = new double[numCen][NUMERO_MESES];
        double[] pagoInyMesLin = new double[NUMERO_MESES];
        
        
        
        
        for (int l = 0; l < numLinea; l++) {
            String[] tmp = nomLineasN[l].split("#");
            int l2 = Calc.Buscar(nomLineasN[l], nomLinTx);
            //System.out.println(nomLineasN[l]+" "+l2);
            for (int j = 0; j < numCen; j++) {
                for (int m = 0; m < NUMERO_MESES; m++) {
                    if (!LiquidacionReliquidacion) { //si es reliquidacion, divide IT entre dentro de AIC y fuera de AIC y se asigna por separado
                        peajeLinCen[l][j][m] = prorrMesGenTxTot[l2][m] == 0 ? 0: VATTN[l][m] * prorrMesGenTx[l2][j][m] - ITP_N[l][m] * prorrMesGenTx[l2][j][m] - IT_N[l][m]*prorrMesGenTx[l2][j][m]/prorrMesGenTxTot[l2][m];
                        ItLinCen[l][j][m]    = prorrMesGenTxTot[l2][m] == 0 ? 0:                                       ITP_N[l][m] * prorrMesGenTx[l2][j][m] + IT_N[l][m]*prorrMesGenTx[l2][j][m]/prorrMesGenTxTot[l2][m]; 
                        if (prorrMesGenTxTot[l2][m] == 0 && IT_N[l][m] != 0){
                            System.err.println(nomLineasN[l] + " " + m + " " + "IT:" + IT_N[l][m] + " " +ITP_N[l][m] );
                        }
                    }
                    else { //si es liquidacion, IT se asigna igual que VATT
                        peajeLinCen[l][j][m] = prorrMesGenTxTot[l2][m] == 0 ? 0: ((VATTN[l][m] - ITP_N[l][m] - IT_N[l][m] - IT_NR[l][m]) * prorrMesGenTx[l2][j][m]);
                        ItLinCen[l][j][m]    = prorrMesGenTxTot[l2][m] == 0 ? 0: ((              IT_N[l][m]  + ITP_N[l][m] + IT_NR[l][m])* prorrMesGenTx[l2][j][m]);
                    }
                    
                    
                    int t = Calc.Buscar(tmp[1], nombreTx);
                    peajeGenTx[j][t][m] += peajeLinCen[l][j][m];
                    ItGenTx[j][t][m] += ItLinCen[l][j][m];
                    if (LiquidacionReliquidacion) { //si es liquidacion, todo los peajes son expectuables, no importa el signo
                        peajeGenTxExcen[j][t][m] +=  peajeLinCen[l][j][m];
                        ItGenTxExcen[j][t][m]    +=  ItLinCen[l][j][m];
                    }
                    else { //si es reliquidacion, el peaje exceptuable son solo los positivos, o sea que vatt > IT
                            if (peajeLinCen[l][j][m] > 0) 
                                peajeGenTxExcen[j][t][m] +=  peajeLinCen[l][j][m];
                                ItGenTxExcen[j][t][m]    +=  ItLinCen[l][j][m];
                            }
                    peajeGenTxZona[j][t][indZonaLinPe[l]][m]+=peajeLinCen[l][j][m];
                    peajeGen[j][m] += peajeLinCen[l][j][m];
                    pagoInyMesLin[m]+=peajeLinCen[l][j][m];
                    //prorrMesGenTx[l][j][m]=prorrMesG[l2][j][m];
                }
            }
        }
        
        
        
        
         /******************************************
         * Calcula Exencion de Centrales
         ******************************************/
        double[][] facPago=new double[numCen][NUMERO_MESES];
        double[] CapConjExcep=new double[NUMERO_MESES];
        double[][][] ExcenCenTx=new double[numCen][numTx][NUMERO_MESES];
        double[][][] ExcenCenItTx=new double[numCen][numTx][NUMERO_MESES];
        double[][] ExcTotCen=new double[numCen][NUMERO_MESES];
        double[][] ExcTotItCen=new double[numCen][NUMERO_MESES];
        double[] InyTotMGC=new double[numCen];
        double[][] ExcTotTx=new double[numTx][NUMERO_MESES];
        double[][] ExcTotItTx=new double[numTx][NUMERO_MESES];
        double[][][] ExcenCenPeajLin = new double[numCen][numLinTx][NUMERO_MESES]; 
        double[][] ExcTotPeajLin = new double [numLinTx][NUMERO_MESES];
        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                    //if(MGNC[i]==1){
                      //  if(PotNeta[i]<9){
                        //    facPago[i][m]=0;
                        //}
                        //else{
                            //facPago[i][m]=1-(20-PotNeta[i])/11;
                            facPago[i][m]=MGNC[i]; //MGNC es factor de exención
                      //  }
                    //}
                    //else{
                        //facPago[i][m]=1;
                        InyTotMGC[m]+=GenPromMesCen[i][m];
                    //}
                    //CapConjExcep[m]+=PotNeta[i]*(1-facPago[i][m])/PotNetaTot;
                    CapConjExcep[m]+=PotNeta[i]*(facPago[i][m])/PotNetaTot;
            }
        }
         double[] FCorrec=new double[NUMERO_MESES];

        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                 //FCorrec[m]=Math.min(1,1/CapConjExcep[m]*0.05);
                 for(int t=0;t<numTx;t++){
                    //ExcenCenTx[i][t][m]=peajeGenTxExcen[i][t][m]*(1-facPago[i][m]);
                    ExcenCenTx[i][t][m]=peajeGenTxExcen[i][t][m]*(facPago[i][m]);//*FCorrec[m];
                    ExcTotCen[i][m]+=ExcenCenTx[i][t][m];
                    ExcTotTx[t][m]+=ExcenCenTx[i][t][m];
                    //IT
                    //ExcenCenItTx[i][t][m]=ItGenTxExcen[i][t][m]*(1-facPago[i][m]);
                    ExcenCenItTx[i][t][m]=ItGenTxExcen[i][t][m]*(facPago[i][m]);//*FCorrec[m];
                    ExcTotItCen[i][m]+=ExcenCenItTx[i][t][m];
                    ExcTotItTx[t][m]+=ExcenCenItTx[i][t][m];
                }
            }
            
        }
        
        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                 FCorrec[m]=Math.min(1,1/CapConjExcep[m]*0.05);
                 for(int l=0;l<numLinTx;l++){
                    //ExcenCenItLin[numCen][numLinTx][numMeses]
                    //ExcenCenPeajLin[i][l][m] = peajeLinCen[l][i][m]*(1-facPago[i][m]);//*FCorrec[m];
                    ExcenCenPeajLin[i][l][m] = peajeLinCen[l][i][m]*(facPago[i][m]);
                    ExcTotPeajLin[l][m]+=ExcenCenPeajLin[i][l][m];
                }
            }
        }
        
        
        /******************************************
         * Calcula Ajuste por MGNC de Centrales
         ******************************************/
        
        double[][][] AjusMGNCTx=new double[numCen][numTx][NUMERO_MESES];
        double[][][] AjusItMGNCTx=new double[numCen][numTx][NUMERO_MESES];
        
        double[][] AjusMGNCTot=new double[numCen][NUMERO_MESES];
        double[][][] AjusPeajMGNCLin = new double[numCen][numLinTx][NUMERO_MESES];
        
        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                for(int t=0;t<numTx;t++){
                    //if(MGNC[i]==1){
                        AjusMGNCTx[i][t][m]=-ExcenCenTx[i][t][m];
                        AjusItMGNCTx[i][t][m]=-ExcenCenItTx[i][t][m];
                    //}
                    //else{
                      // AjusMGNCTx[i][t][m]= GenPromMesCen[i][m]/InyTotMGC[m]*ExcTotTx[t][m];
                       //AjusItMGNCTx[i][t][m]= GenPromMesCen[i][m]/InyTotMGC[m]*ExcTotItTx[t][m];
                    //}
                    AjusMGNCTot[i][m]+=AjusMGNCTx[i][t][m];
                    
                }
            }
        }
        
        
        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                for(int l=0;l<numLinTx;l++){
                    //if(MGNC[i]==1){
                        AjusPeajMGNCLin[i][l][m]=-ExcenCenPeajLin[i][l][m];
                    //}
                    //else{
                      //  AjusPeajMGNCLin[i][l][m]= GenPromMesCen[i][m]/InyTotMGC[m]*ExcTotPeajLin[l][m];
                  //  }
                    
                }
            }
        }
                
                
         /******************************************
         * Separa Datos de Centrales MGNC
         ******************************************/
        double[][][] peajeMGNCTx=new double[numMGNC][numTx][NUMERO_MESES];
        double[][] peajeMGNC=new double[numMGNC][NUMERO_MESES];
        double[][] peajeAnualMGNCTx=new double[numMGNC][numTx];
        double[] peajeAnualMGNC=new double[numMGNC];
        String nomMGNC[]=new String[numMGNC];
        double[] PotNetaMGNC = new double[numMGNC];
        
        int aux=0;
        double[][][] ExcenMGNCTx=new double[numMGNC][numTx][NUMERO_MESES];
        double[][] ExcenMGNC=new double[numMGNC][NUMERO_MESES];
        double[][] ExcenAnualMGNCTx=new double[numMGNC][numTx];
        double[] ExcenAnualMGNC=new double[numMGNC];
        double[] facPagoMGNC=new double[numMGNC];


         for(int i=0;i<numCen;i++){
            // if(MGNC[i]==1){
                 PotNetaMGNC[aux]=PotNeta[i];
                 nomMGNC[aux]=nomGen[i];
                 facPagoMGNC[aux]=facPago[i][11];//factor de Pago de Diciembre
                 for(int m=0;m<NUMERO_MESES;m++){
                     for(int t=0;t<numTx;t++){
                     peajeMGNCTx[aux][t][m]=peajeGenTx[i][t][m];
                     peajeMGNC[aux][m]+=peajeGenTx[i][t][m];
                     peajeAnualMGNCTx[aux][t]+=peajeGenTx[i][t][m];
                     peajeAnualMGNC[aux]+=peajeGenTx[i][t][m];
                     ExcenMGNCTx[aux][t][m]=ExcenCenTx[i][t][m];
                     ExcenMGNC[aux][m]+=ExcenCenTx[i][t][m];
                     ExcenAnualMGNCTx[aux][t]+=ExcenCenTx[i][t][m];
                     ExcenAnualMGNC[aux]+=ExcenCenTx[i][t][m];
               
                     }
               }
                 aux++;
            //}
         }
        
        /******************************************
         * Calcula Pago Total por central
         ******************************************/
        double[][][] PagoTotCenTx=new double[numCen][numTx][NUMERO_MESES];
        double[][][] ItTotCenTx=new double[numCen][numTx][NUMERO_MESES];
        double[][] PagoTotCen=new double[numCen][NUMERO_MESES];
        double[][][] ProrrPeajEmpLin = new double[numEmpG][numLinTx][NUMERO_MESES];
        for(int i=0;i<numCen;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                for(int t=0;t<numTx;t++){
                    PagoTotCenTx[i][t][m]=peajeGenTx[i][t][m]+ AjusMGNCTx[i][t][m];
                    ItTotCenTx[i][t][m]=ItGenTx[i][t][m]+ AjusItMGNCTx[i][t][m];
                    PagoTotCen[i][m]+=PagoTotCenTx[i][t][m];
                    }
                }
            }
        

        
        
        
        for(int i=0;i<numCen;i++){
            String[] tmp = nomGen[i].split("#");
            int emp = Calc.Buscar(tmp[0], nomEmp);
            for(int m=0;m<NUMERO_MESES;m++){
                for( int l=0 ; l<numLinTx ; l++ ){
                //ItTotCenLin[i][l][m]=ItLinCen[l][i][m]+ AjusItMGNCLin[i][l][m];
                    if  (PeajeN[l][m] == 0) {
                        ProrrPeajEmpLin[emp][l][m] += 0;
                    }
                    else {
                        ProrrPeajEmpLin[emp][l][m]+=(peajeLinCen[l][i][m]+ AjusPeajMGNCLin[i][l][m])/PeajeN[l][m];
                    }
                }
            }
        }
        
        
        try {
            FileWriter writer = new FileWriter(DirBaseSal + SLASH + "prorratas_pago_iny.csv");
            writer.append("Central");
            writer.append(',');
            writer.append("Linea");
            writer.append(',');
            writer.append("Mes");
            writer.append(',');
            writer.append("Prorrata");
            writer.append('\n');
            for (int m = 0; m < NUMERO_MESES; m++) {
                for (int i = 0; i < numEmpG; i++) {
                    for (int t = 0; t < numLinTx; t++) {
                        writer.append(nomEmp[i]);
                        writer.append(',');
                        writer.append(nomLineasN[t]);
                        writer.append(',');
                        writer.append(Float.toString(m + 1));
                        writer.append(',');
                        writer.append(Double.toString(ProrrPeajEmpLin[i][t][m]));
                        writer.append('\n');
                    }
                }
            }
            writer.flush();
            writer.close();
        } catch (IOException e) {
            System.out.println("No se pudo escribir con exito prorratas_pago_iny.csv");
            e.printStackTrace(System.out);
        }
          
          
        /******************************************
         * Calcula Pagos por empresa
         ******************************************/
        double[][][] peajeEmpGTx = new double[numEmpG][numTx][NUMERO_MESES];
        double[][] peajeEmpG = new double[numEmpG][NUMERO_MESES];
        double[][][] AjusMGNCEmpTx=new double[numEmpG][numTx][NUMERO_MESES];
        double[][][] PagoEmpGTx=new double[numEmpG][numTx][NUMERO_MESES];
        double[][][] ItEmpGTx=new double[numEmpG][numTx][NUMERO_MESES];
        double[][] PagoAnualEmpGTx=new double[numEmpG][numTx];
        double[][] PeajeAnualEmpGTx=new double[numEmpG][numTx];
        double[][] AjusMGNCEmp=new double[numEmpG][NUMERO_MESES];
        double[][] PagoEmp=new double[numEmpG][NUMERO_MESES];
        double[][] ItEmp=new double[numEmpG][NUMERO_MESES];
        double[] PagoAnualEmpG=new double[numEmpG];
        double[] PeajeAnualEmpG=new double[numEmpG];
        double[] PagoInyMes=new double[NUMERO_MESES];
        for (int j = 0; j < numCen; j++) {
            String[] tmp = nomGen[j].split("#");
            int l = Calc.Buscar(tmp[0], nomEmp);
            for (int m = 0; m < NUMERO_MESES; m++) {
                AjusMGNCEmp[l][m]+=AjusMGNCTot[j][m];
                for (int t = 0; t < numTx; t++) {
                    peajeEmpGTx[l][t][m] += peajeGenTx[j][t][m];
                    peajeEmpG[l][m] += peajeGenTx[j][t][m];
                    AjusMGNCEmpTx[l][t][m] += AjusMGNCTx[j][t][m];
                    PagoEmpGTx[l][t][m]+=PagoTotCenTx[j][t][m];
                    ItEmpGTx[l][t][m]+=ItTotCenTx[j][t][m];
                    PagoAnualEmpGTx[l][t]+=PagoTotCenTx[j][t][m];
                    PagoEmp[l][m]+=PagoTotCenTx[j][t][m];
                    ItEmp[l][m]+=ItTotCenTx[j][t][m];
                    PagoAnualEmpG[l]+=PagoTotCenTx[j][t][m];
                    PeajeAnualEmpG[l]+=peajeGenTx[j][t][m];
                    PeajeAnualEmpGTx[l][t]+=peajeGenTx[j][t][m];
                    PagoInyMes[m]+=PagoTotCenTx[j][t][m];
                    
                }
            }
        }



        // Ordena los archivos de salida de Inyeccion por empresas
        int[] ng = Calc.OrdenarBurbujaStr(nomGen);
       nomGenO = new String[numCen];
        for (int i = 0; i < numCen; i++) {
            nomGenO[i] = nomGen[ng[i]];
        }
        int []nmgnc = Calc.OrdenarBurbujaStr(nomMGNC);
        nomMGNCO = new String[numMGNC];
        for (int i = 0; i < numMGNC; i++) {
            nomMGNCO[i] = nomMGNC[nmgnc[i]];
        }
        // -------------------------------------------------------------------
        prorrMesGO = new double[numLinTx][numCen][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCen; j++) {
                System.arraycopy(prorrMesGenTx[i][ng[j]], 0, prorrMesGO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        peajeLinGO = new double[numLinTx][numCen][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCen; j++) {
                System.arraycopy(peajeLinCen[i][ng[j]], 0, peajeLinGO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        ItLinGO = new double[numLinTx][numCen][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCen; j++) {
                System.arraycopy(ItLinCen[i][ng[j]], 0, ItLinGO[i][j], 0, NUMERO_MESES);
            }
        }// -------------------------------------------------------------------
        peajeCenTxO = new double[numCen][numTx][NUMERO_MESES];
        for (int i = 0; i < numCen; i++) {
            for (int j = 0; j < numTx; j++) {
                System.arraycopy(peajeGenTx[ng[i]][j], 0, peajeCenTxO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        peajeCenO = new double[numCen][NUMERO_MESES];
        for (int i = 0; i < numCen; i++) {
            System.arraycopy(peajeGen[ng[i]], 0, peajeCenO[i], 0, NUMERO_MESES);
        }
        // -------------------------------------------------------------------

        int []ne = Calc.OrdenarBurbujaStr(nomEmp);
        nomEmpGO = new String[numEmpG];
        for (int i = 0; i < numEmpG; i++) {
            nomEmpGO[i] = nomEmp[ne[i]];
        }
        // -------------------------------------------------------------------
        peajeEmpTxO = new double[numEmpG][numTx][NUMERO_MESES];
        for (int i = 0; i < numEmpG; i++) {
            for (int j = 0; j < numTx; j++) {
                System.arraycopy(peajeEmpGTx[ne[i]][j], 0, peajeEmpTxO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        peajeEmpGO = new double[numEmpG][NUMERO_MESES];
        for (int i = 0; i < numEmpG; i++) {
            System.arraycopy(peajeEmpG[ne[i]], 0, peajeEmpGO[i], 0, NUMERO_MESES);
        }
        // -------------------------------------------------------------------
        ExcenMGNCTxO=new double[numMGNC][numTx][NUMERO_MESES];
        ExcenMGNCO=new double[numMGNC][NUMERO_MESES];
        for (int i = 0; i < numMGNC; i++) {
            for (int j = 0; j < numTx; j++) {
            for (int m = 0; m < NUMERO_MESES; m++) {
               ExcenMGNCTxO[i][j][m]=ExcenMGNCTx[nmgnc[i]][j][m];
               ExcenMGNCO[i][m]=ExcenMGNC[nmgnc[i]][m];
            }
            }
        }
        // -------------------------------------------------------------------
        MGNCO = new float[numCen];
        PotNetaO = new double[numCen];
        FAERO = new float[numCen];
        CETO= new float[numCen];
        Tabla1O= new float[numCen];       
        GenPromMesCenO= new double[numCen][NUMERO_MESES];
        facPagoO=new double[numCen][NUMERO_MESES];
        peajeGenO = new double[numCen][NUMERO_MESES];
        ExcTotCenO=new double[numCen][NUMERO_MESES];
        AjusMGNCTotO=new double[numCen][NUMERO_MESES];
        PagoTotCenO=new double[numCen][NUMERO_MESES];
        ExcenCenO=new double[numCen][numTx][NUMERO_MESES];
        AjusMGNCTxO=new double[numCen][numTx][NUMERO_MESES];
        PagoTotCenTxO=new double[numCen][numTx][NUMERO_MESES];
        double[][][][] peajeGenTxZonaO = new double[numCen][numTx][3][NUMERO_MESES];


        for (int i = 0; i < numCen; i++) {
            MGNCO[i]=MGNC[ng[i]];
            PotNetaO[i]=PotNeta[ng[i]];
            FAERO[i]=FAER[ng[i]];
            CETO[i]=CET[ng[i]];
            Tabla1O[i]=Tabla1[ng[i]];
            for (int k = 0; k < NUMERO_MESES; k++) {
                GenPromMesCenO[i][k]=GenPromMesCen[ng[i]][k];
                facPagoO[i][k]=facPago[ng[i]][k];
                peajeGenO[i][k]=peajeGen[ng[i]][k];
                ExcTotCenO[i][k]=ExcTotCen[ng[i]][k];
                AjusMGNCTotO[i][k]=AjusMGNCTot[ng[i]][k];
                PagoTotCenO[i][k]=PagoTotCen[ng[i]][k];
                for (int j = 0; j < numTx; j++) {
                    ExcenCenO[i][j][k] = ExcenCenTx[ng[i]][j][k];
                    AjusMGNCTxO[i][j][k] = AjusMGNCTx[ng[i]][j][k];
                    PagoTotCenTxO[i][j][k] = PagoTotCenTx[ng[i]][j][k];
                    for(int z=0;z<3;z++) peajeGenTxZonaO[i][j][z][k] =peajeGenTxZona[ng[i]][j][z][k];
                }
            }
        }
        // -------------------------------------------------------------------
        AjusMGNCEmpTxO= new double[numEmpG][numTx][NUMERO_MESES];
        PagoEmpTxO=new double[numEmpG][numTx][NUMERO_MESES];
        ItEmpTxO=new double[numEmpG][numTx][NUMERO_MESES];
        PagoAnualEmpGTxO=new double[numEmpG][numTx];
        PeajeAnualEmpGTxO=new double[numEmpG][numTx];
        AjusMGNCEmpO=new double[numEmpG][NUMERO_MESES];
        PagoEmpO=new double[numEmpG][NUMERO_MESES];
        ItEmpO=new double[numEmpG][NUMERO_MESES];
        PagoAnualEmpGO=new double[numEmpG];
        PeajeAnualEmpGO=new double[numEmpG];
        AjusMGNCAnualEmpTxO=new double[numEmpG][numTx];
        AjusMGNCAnualEmpO=new double[numEmpG];

        for (int i = 0; i < numEmpG; i++) {
            PagoAnualEmpGO[i]= PagoAnualEmpG[ne[i]];
             PeajeAnualEmpGO[i]= PeajeAnualEmpG[ne[i]];
                for (int k = 0; k < NUMERO_MESES; k++) {
                     AjusMGNCEmpO[i][k] = AjusMGNCEmp[ne[i]][k];
                     PagoEmpO[i][k] = PagoEmp[ne[i]][k];
                     ItEmpO[i][k] = ItEmp[ne[i]][k];
                     for (int j = 0; j < numTx; j++) {
                     AjusMGNCEmpTxO[i][j][k] = AjusMGNCEmpTx[ne[i]][j][k];
                     AjusMGNCAnualEmpTxO[i][j]+=AjusMGNCEmpTx[ne[i]][j][k];
                     AjusMGNCAnualEmpO[i]+=AjusMGNCEmpTx[ne[i]][j][k];
                     PagoEmpTxO[i][j][k] = PagoEmpGTx[ne[i]][j][k];
                     ItEmpTxO[i][j][k] = ItEmpGTx[ne[i]][j][k];
                     PagoAnualEmpGTxO[i][j]=PagoAnualEmpGTx[ne[i]][j];
                     PeajeAnualEmpGTxO[i][j]=PeajeAnualEmpGTx[ne[i]][j];
                }
            }
        }

        double[][][] peajeMGNCTxO=new double[numMGNC][numTx][NUMERO_MESES];
        double[][] peajeMGNCO=new double[numMGNC][NUMERO_MESES];
        peajeAnualMGNCTxO=new double[numMGNC][numTx];
        peajeAnualMGNCO=new double[numMGNC];
        PotNetaMGNCO=new double[numMGNC];
        ExcenAnualMGNCTxO=new double[numMGNC][numTx];
        ExcenAnualMGNCO=new double[numMGNC];
        facPagoMGNCO=new double[numMGNC];
        

        for(int i=0;i<numMGNC;i++){
            for(int m=0;m<NUMERO_MESES;m++){
                for(int t=0;t<numTx;t++){
                     peajeMGNCTxO[i][t][m]= peajeMGNCTx[nmgnc[i]][t][m];
                     peajeMGNCO[i][m]=peajeMGNC[nmgnc[i]][m];
                     peajeAnualMGNCTxO[i][t]=peajeAnualMGNCTx[nmgnc[i]][t];
                     peajeAnualMGNCO[i]=peajeAnualMGNC[nmgnc[i]];
                     PotNetaMGNCO[i]=PotNetaMGNC[nmgnc[i]];
                     ExcenAnualMGNCTxO[i][t]=ExcenAnualMGNCTx[nmgnc[i]][t];
                     ExcenAnualMGNCO[i]=ExcenAnualMGNC[nmgnc[i]];
                     facPagoMGNCO[i]=facPagoMGNC[nmgnc[i]];
                     }
                }
            }
        calcPagos=false;
        EscribirPagos=true;
        long tFinalCalculo = System.currentTimeMillis();


        /*
         * Escritura de Resultados Anuales:
         * =======================
         */
        long tInicioEscritura = System.currentTimeMillis();
        System.out.println("Escribiendo resultados a archivos de salida");
        if (horizon == PeajesConstant.HorizonteCalculo.Anual) {
            String sEscribeXLS = PeajesCDEC.getOptionValue("Imprime pagos a Excel", PeajesConstant.DataType.BOOLEAN);
            boolean bEscribeXLS = Boolean.parseBoolean(sEscribeXLS);
            if (bEscribeXLS) {
                String libroSalidaGXLS = DirBaseSal + SLASH + "PagoIny" + Ano + ".xlsx";
                if (!USE_MEMORY_WRITER) {
                    Escribe.crearLibro(libroSalidaGXLS);
                    Escribe.creaH2F_3d_long(
                            "Pago de Peaje por Línea y Central [$]", peajeLinGO,
                            "Línea", nomLineasN,
                            "Central", nomGenO,
                            "Factor de Exención", MGNCO,
                            "Mes", MESES,
                            libroSalidaGXLS, "PjeCenLin",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    for (int m = 0; m < NUMERO_MESES; m++) {
                        Escribe.creaPIny(m,
                                "Pago Peaje por Empresa y Transmisor [$]", peajeEmpTxO,
                                AjusMGNCEmpTxO, PagoEmpTxO,
                                peajeEmpGO, AjusMGNCEmpO, PagoEmpO,
                                "Empresa", nomEmpGO,
                                "Transmisor", nombreTx,
                                libroSalidaGXLS, MESES[m],
                                "#,###,##0;[Red]-#,###,##0;\"-\"");

                        Escribe.creaDetallePIny(m,
                                "Detalle de Pagos por Central [$]", peajeGenTxZonaO, peajeCenTxO, ExcenCenO,
                                AjusMGNCTxO, PagoTotCenTxO,
                                peajeGenO, ExcTotCenO, AjusMGNCTotO, PagoTotCenO,
                                CapConjExcep, FCorrec,
                                "Central", nomGenO,
                                "Transmisor", nombreTx,
                                "Factor Exención", MGNCO,
                                //"PNeta", PotNetaO,
                                //"Inyeccion Mensual", GenPromMesCenO,
                                //"Factor",facPagoO ,
                                libroSalidaGXLS, MESES[m],
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                    }
                    Escribe.crea_verificaIny(
                            "Verifica Pagos de Inyección", libroEntrada,
                            "Mes", MESES,
                            "Calculo", PagoInyMes,
                            "Prorrata Línea", pagoInyMesLin,
                            "Diferencia",
                            "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.crea_verificaCalcPeajes(
                            "Verifica cálculo de Peajes", libroEntrada,
                            "Mes", MESES,
                            "Peajes", PeajeNMes,
                            "Pago Ret", "Pago Iny", "Diferencia",
                            "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                } else {

                    try {
                        XSSFWorkbook wb_salida = Escribe.crearLibroVacio(libroSalidaGXLS);
                        Escribe.creaH2F_3d_long(
                                "Pago de Peaje por Línea y Central [$]", peajeLinGO,
                                "Línea", nomLineasN,
                                "Central", nomGenO,
                                "Factor de Exención", MGNCO,
                                "Mes", MESES,
                                wb_salida, "PjeCenLin",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        for (int m = 0; m < NUMERO_MESES; m++) {
                            Escribe.creaPIny(m,
                                    "Pago Peaje por Empresa y Transmisor [$]", peajeEmpTxO,
                                    AjusMGNCEmpTxO, PagoEmpTxO,
                                    peajeEmpGO, AjusMGNCEmpO, PagoEmpO,
                                    "Empresa", nomEmpGO,
                                    "Transmisor", nombreTx,
                                    wb_salida, MESES[m],
                                    "#,###,##0;[Red]-#,###,##0;\"-\"");
                            Escribe.creaDetallePIny(m,
                                    "Detalle de Pagos por Central [$]", peajeGenTxZonaO, peajeCenTxO, ExcenCenO,
                                    AjusMGNCTxO, PagoTotCenTxO,
                                    peajeGenO, ExcTotCenO, AjusMGNCTotO, PagoTotCenO,
                                    CapConjExcep, FCorrec,
                                    "Central", nomGenO,
                                    "Transmisor", nombreTx,
                                    "Factor Exención", MGNCO,
                                    //"PNeta", PotNetaO,
                                    //"Inyeccion Mensual", GenPromMesCenO,
                                    //"Factor",facPagoO ,
                                    wb_salida, MESES[m],
                                    "#,###,##0;[Red]-#,###,##0;\"-\"");
                        }
                        Escribe.crea_verificaIny(
                                "Verifica Pagos de Inyección", wb_Ent,
                                "Mes", MESES,
                                "Calculo", PagoInyMes,
                                "Prorrata Línea", pagoInyMesLin,
                                "Diferencia",
                                "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.crea_verificaCalcPeajes(
                                "Verifica cálculo de Peajes", wb_Ent,
                                "Mes", MESES,
                                "Peajes", PeajeNMes,
                                "Pago Ret", "Pago Iny", "Diferencia",
                                "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.guardaLibroDisco(wb_salida, libroSalidaGXLS);
                        Escribe.guardaLibroDisco(wb_Ent, libroEntrada);
                        wb_Peajes.close();
                        wb_Ent.close();
                        wb_salida.close();
                    } catch (IOException e) {
                        System.out.println("Error al escribir resultados de pagos inyecciones a archivo " + libroSalidaGXLS);
                        System.out.println(e.getMessage());
                        e.printStackTrace(System.err);
                    }
                }
            }

            //Escribe archivos csv de salida:
            String sEscribeCSV = PeajesCDEC.getOptionValue("Imprime pagos a csv", PeajesConstant.DataType.BOOLEAN);
            boolean bEscribeCSV = Boolean.parseBoolean(sEscribeCSV);
            if (bEscribeCSV) {
                String libroSalidaGCSV = DirBaseSal + SLASH + "PagoIny" + Ano + ".csv";
                BufferedWriter writerCSV = null;
                try {
                    writerCSV = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(libroSalidaGCSV), StandardCharsets.ISO_8859_1));
                    String sLineText;
                    //Escribimos el header:
                    writerCSV.write("Tramo,Transmisor,Central,Empresa Generación,Mes,Factor de Exención,FAER,CET,Tabla 1,Pago de Peaje [$]");
                    writerCSV.newLine();
                    //Escribimos los datos:
                    for (int l = 0; l < numLinTx; l++) {
                        String[] sTramoTransmisor = nomLineasN[l].split("#");
                        assert (sTramoTransmisor.length == 2) : "Como se formaron estos nombres?";
                        for (int g = 0; g < numCen; g++) {
                            String[] sEmpresaGen = nomGenO[g].split("#");
                            assert (sEmpresaGen.length == 2) : "Como se formaron estos nombres?";
                            for (int m = 0; m < NUMERO_MESES; m++) {
                                sLineText = "";
                                for (String s : sTramoTransmisor) {
                                    sLineText += s + ",";
                                }
                                sLineText += sEmpresaGen[1] + ",";
                                sLineText += sEmpresaGen[0] + ",";
                                sLineText += MESES[m] + ",";
                                sLineText += MGNCO[g] + ",";
                                sLineText += FAERO[g] + ",";
                                sLineText += CETO[g] + ",";
                                sLineText += Tabla1O[g] + ",";
                                sLineText += peajeLinGO[l][g][m];
                                writerCSV.write(sLineText);
                                writerCSV.newLine();
                            }
                        }
                    }
                    System.out.println("Finalizado escritura de resultados PagoIny.csv");
                } catch (IOException e) {
                    System.out.println("WARNING: No se pudo escribir PagoIny.csv. Error: " + e.getMessage());
                    e.printStackTrace(System.out);
                } finally {
                    if (writerCSV != null) {
                        try {
                            writerCSV.close();
                        } catch (IOException e) {
                            System.out.println("No se pudo cerrar conexion con PagoIny.csv. Error: " + e.getMessage());
                            e.printStackTrace(System.out);
                        }
                    }
                }
            }
        }
        
        /*
         * Escritura de Resultados Mensuales:
         * =======================
         */
        if (horizon == PeajesConstant.HorizonteCalculo.Mensual) {
            LiquiMesIny(MESES[Mes], Ano);
        }
        long tFinalEscritura = System.currentTimeMillis();
        EscribirPagos=true;
        
        System.out.println("Pagos de Inyección Calculados");
        System.out.println("Tiempo Adquisicion de datos     : "+DosDecimales.format((tFinalLectura-tInicioLectura)/(1000.0))+" seg");
        System.out.println("Tiempo Cálculo                  : "+DosDecimales.format((tFinalCalculo-tInicioCalculo)/(1000.0))+" seg");
        System.out.println("Tiempo Escritura de Resultados  : "+DosDecimales.format((tFinalEscritura-tInicioEscritura)/(1000.0))+" seg");
        System.out.println();
    }

    public static void LiquiMesIny(String mes, int Ano) {
        int m = 0;
        for (int i = 0; i < NUMERO_MESES; i++) {
            if (mes.equals(MESES[i])) {
                m = i;
            }
        }
        String libroSalidaGXLSMes = DirBaseSal + SLASH + "PagoIny" + MESES[m] + ".xlsx";
        if (!USE_MEMORY_WRITER) {
            Escribe.crearLibro(libroSalidaGXLSMes);
            Escribe.creaLiquidacionMesIny(m,
                    "Pago de Peajes de Inyección", peajeCenTxO,
                    AjusMGNCTxO, PagoTotCenTxO,
                    peajeGenO, AjusMGNCTotO, PagoTotCenO,
                    "Central", nomGenO,
                    "Transmisor", nombreTx,
                    "MGNC", MGNCO,
                    "PNeta", PotNetaO,
                    "Inyeccion [GWh]", GenPromMesCenO,
                    "Factor", facPagoO,
                    libroSalidaGXLSMes, MESES[m], Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaProrrataMes(m,
                    "Participación de Inyecciones [%]", prorrMesGenTx, "Participación " + MESES[m],
                    "Cliente", nomGen,
                    "Línea", nomLinTx,
                    "AIC", zonaLinTx,
                    libroSalidaGXLSMes, "PartIny" + MESES[m],
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaProrrataMes_long(m,
                    "Pagos por Inyección " + MESES[m] + " [$]",
                    peajeLinGO,
                    "Pago " + MESES[m],
                    "Central",
                    nomGenO,
                    "Línea",
                    nomLineasN,
                    "AIC",
                    zonaLinPe,
                    "Pago IT",
                    ItLinGO,
                    libroSalidaGXLSMes,
                    "PagoxLinea",
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
        } else {

            try {
                XSSFWorkbook wb_salida = Escribe.crearLibroVacio(libroSalidaGXLSMes);
                Escribe.creaLiquidacionMesIny(m,
                        "Pago de Peajes de Inyección", peajeCenTxO,
                        AjusMGNCTxO, PagoTotCenTxO,
                        peajeGenO, AjusMGNCTotO, PagoTotCenO,
                        "Central", nomGenO,
                        "Transmisor", nombreTx,
                        "MGNC", MGNCO,
                        "PNeta", PotNetaO,
                        "Inyeccion [GWh]", GenPromMesCenO,
                        "Factor", facPagoO,
                        wb_salida, MESES[m], Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaProrrataMes(m,
                        "Participación de Inyecciones [%]", prorrMesGenTx, "Participación " + MESES[m],
                        "Cliente", nomGen,
                        "Línea", nomLinTx,
                        "AIC", zonaLinTx,
                        wb_salida, "PartIny" + MESES[m],
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaProrrataMes_long(m,
                        "Pagos por Inyección " + MESES[m] + " [$]",
                        peajeLinGO,
                        "Pago " + MESES[m],
                        "Central",
                        nomGenO,
                        "Línea",
                        nomLineasN,
                        "AIC",
                        zonaLinPe,
                        "Pago IT",
                        ItLinGO,
                        wb_salida,
                        "PagoxLinea",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.guardaLibroDisco(wb_salida, libroSalidaGXLSMes);
                wb_salida.close();
            } catch (IOException e) {
                System.out.println("Error al escribir resultados de pagos inyecciones a archivo " + libroSalidaGXLSMes);
                System.out.println(e.getMessage());
                e.printStackTrace(System.err);
            }

        }

        System.out.println("Archivo Pago de Inyección Mensual creado");
        System.out.println();

    }

    public static boolean cargando() {
        return cargandoInfo;
    }

    public static boolean calculando() {
        return calcPagos;
    }

    public static boolean escribiendo() {
        return calcPagos;
    }

    public static void Comenzar(final File DirIn, final File DirOut, final int AnoAEvaluar, final boolean LiquidacionReliquidacion) {
        javax.swing.SwingWorker worker = new javax.swing.SwingWorker() {

            @Override
            protected Object doInBackground() throws Exception {
                calculaPeajesIny(DirIn, DirOut, AnoAEvaluar, LiquidacionReliquidacion);
                return null;
            }

        };
        worker.execute();
    }

    

}
