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

import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.MAX_COMPRESSION_RATIO;
import static cl.coordinador.peajes.PeajesConstant.MESES;
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
public class PeajesRet {
    
    private static final boolean USE_MEMORY_READER = true; //swtich para usar nuevo API lectura poi
    private static final boolean USE_MEMORY_WRITER = true; //swtich para usar nuevo API escritura poi

    private static String DirBaseSal;
    static double[][][]  RetEmpSinAjuTxO; //TODO: Encapsulate
    static double[][] RetEmpSinAjuO; //TODO: Encapsulate
    private static double[][][] AjusEmpCTxO;
    static double[][] AjusAnualEmpCTxO; //TODO: Encapsulate
    static double[][] TotRetEmpO; //TODO: Encapsulate
    static double[][] TotItRetEmpO; //TODO: Encapsulate
    private static double[] TotAnualRetEmpO;
    static double[][][] TotRetEmpTxO; //TODO: Encapsulate
    static double[][][] TotRetItEmpTxO; //TODO: Encapsulate
    private static double[][][] TotRetEmpTxRE2288O;
    private static double[][] TotAnualRetEmpTxO;
    private static double[][] TotAnualRetEmpTxRE2288O;
    private static double[][] AjusEmpCO;
    static double[] AjusAnualEmpCO; //TODO: Encapsulate
    static String[] nomEmpO; //TODO: Encapsulate
    static int numEmp; //TODO: Encapsulate
    static String[] nombreTx; //TODO: Encapsulate
    static String[] nomCli; //TODO: Encapsulate
    static String[] nomLinIT; //TODO: Encapsulate
    static double[][][] prorrMesC; //TODO: Encapsulate
    private static double[][] prorrMesCTot;
    static String[] nomLinTx; //TODO: Encapsulate
    private static int[] zonaLin;
    static int[] zonaLinTx; //TODO: Encapsulate
    private static double[][][] PUO;
    private static String[] nomBarO;
    private static String[] nomCliO;
    private static double[][][] peajeClienTxNOExenO;
    private static String[] nombreCliNOExenO;
    static String[] nombreClientesExenO; //TODO: Encapsulate
    private static double[][][] peajeClienTxExenO;
    private static double[][][] AjusClienExenCenTxO ;
    private static String[] nomCenO ;
    private static double[] GenAnoxCenO;
    private static double[][] GenPromMesCenO;
    private static String[] nomEmpC;
    private static int numEmpC;
    private static String[] nomEmp;
    static int numTx; //TODO: Encapsulate
    private static int[] ne;
    private static int[] ne2288;
    private static String[] nomEmpCO;
    static String[] nomSumiRM88O; //TODO: Encapsulate
    static double[][] TotAnualPjeRetEmpTxO; //TODO: Encapsulate
    static double[][] TotConRe2288AnualPjeRetEmpTxO; //TODO: Encapsulate
    static double[][] TotAnualPjeRetEmpTxRE2288O; //TODO: Encapsulate
    static double[] TotAnualPjeRetEmpO; //TODO: Encapsulate
    static double[] TotConRe2288AnualPjeRetEmpO; //TODO: Encapsulate
    static double[] TotAnualPjeRetEmpRE2288O; //TODO: Encapsulate
    static  double[][] pjeAnualClienTxExenO; //TODO: Encapsulate
    static double[] pjeAnualClienExenO; //TODO: Encapsulate
    static int numClienExentos; //TODO: Encapsulate
    private static double[][][][] pjeEmpDxTx ;
    private static  String[] nomDx;
    private static String[] nomSumi;
    private static String[] nomSumiRM88;
    private static double[][][] facDx;
    private static double[][] proEfirme;
    static double[][][] TotRetEmpTxRE2288OO; //TODO: Encapsulate
    static double[][] TotRetEmpRE2288O;

    public static void calculaPeajesRet(File DirEntrada, File DirSalida, int Ano, boolean LiquidacionReliquidacion) {
        calculaPeajesRet(PeajesConstant.HorizonteCalculo.Anual, DirEntrada, DirSalida, Ano, 0, LiquidacionReliquidacion);
    }
    
    public static void calculaPeajesRet(PeajesConstant.HorizonteCalculo horizon, File DirEntrada, File DirSalida, int Ano, int Mes, boolean LiquidacionReliquidacion) {
        String DirBaseEnt = DirEntrada.toString();
        DirBaseSal = DirSalida.toString();
        DecimalFormat DosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);

        //Lee opciones de configuracion:
        int maxLeeLINEA = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de LINEAS leer en planillas Ent",PeajesConstant.DataType.INTEGER));
        int maxLeeGEN = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de CENTRALES leer en planillas Ent",PeajesConstant.DataType.INTEGER));
        int maxLeeCLIENTE = Integer.parseInt(PeajesCDEC.getOptionValue("Maximo numero de CLIENTES leer en planillas Ent",PeajesConstant.DataType.INTEGER));
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
        double[][] longAuxVATT = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxIT = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxITG = new double[maxLeeLINEA][NUMERO_MESES];
        double[][] longAuxITP = new double[maxLeeLINEA][NUMERO_MESES];
        String[] TxtTemp = new String[maxLeeLINEA];
        String[] TxtTempIT = new String[maxLeeLINEA];
        int numLinea;
        int numLineaVATT;
        int numLineaIT;
        int numLineaITG;
        int numLineaITP;
        if (USE_MEMORY_READER) {
            numLinea = Lee.leePeajes(wb_Peajes, TxtTemp, longAux);
            numLineaVATT = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxVATT, "VATT");
            numLineaIT = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxIT, "ITER");
            numLineaITG = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxITG, "ITEG");
            numLineaITP = Lee.leeIT(wb_Peajes, TxtTempIT, longAuxITP, "ITP");
        } else {
            numLinea = Lee.leePeajes(libroEntradaPeajes, TxtTemp, longAux);
            numLineaVATT = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxVATT, "VATT");
            numLineaIT = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxIT, "ITER");
            numLineaITG = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxITG, "ITEG");
            numLineaITP = Lee.leeIT(libroEntradaPeajes, TxtTempIT, longAuxITP, "ITP");
        }
        String[] nomLineasN = new String[numLinea];
        double[][] PeajeN = new double[numLinea][12];
        double[][] VATTN = new double[numLineaVATT][12];
        double[][] IT_N = new double[numLineaIT][12];
        double[][] IT_NG = new double[numLineaITG][12];     
        double[][] ITP_N = new double[numLineaITP][12];
        double[] PeajeNMes = new double[NUMERO_MESES];
        double[] VATTNMes = new double[NUMERO_MESES];
        double[] IT_N_Mes = new double[12];
        double[] ITP_N_Mes = new double[12];
        for (int i = 0; i < numLinea; i++) {
            nomLineasN[i] = TxtTemp[i];
            for (int j = 0; j < 12; j++) {
                PeajeN[i][j] = longAux[i][j];
                PeajeNMes[j]+=PeajeN[i][j];
                VATTN[i][j] = longAuxVATT[i][j];
                VATTNMes[j]+=VATTN[i][j];
                IT_N[i][j] = longAuxIT[i][j];
                IT_NG[i][j] = longAuxITG[i][j];
                IT_N_Mes[j]+=IT_N[i][j];
                ITP_N[i][j] = longAuxITP[i][j];
                ITP_N_Mes[j]+=ITP_N[i][j];        
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
        int numLineasVATT;
        if (USE_MEMORY_READER) {
            numLineasVATT = Lee.leeVATT(wb_Ent, TxtTemp1, TxtTemp2, Aux);
        } else {
            numLineasVATT = Lee.leeVATT(libroEntrada, TxtTemp1, TxtTemp2, Aux);
        }
//        String[] nomLinVATT = new String[numLineasVATT];
        String[] nomProp = new String[numLineasVATT];
        String[] TxtTemp3 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp3[i] = "";
        }
        numTx = 0;
        for (int i = 0; i < numLineasVATT; i++) {
//            nomLinVATT[i] = TxtTemp1[i];
            nomProp[i] = TxtTemp2[i];
            int t = Calc.Buscar(nomProp[i], TxtTemp3);
            if (t == -1) {
                TxtTemp3[numTx] = nomProp[i];
                numTx++;
            }
        }
       /* String[] TxtTemp4 = new String[numLineasVATT];
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
        *
        */
        nombreTx = new String[numTx];
        System.arraycopy(TxtTemp3, 0, nombreTx, 0, numTx); /*nomLinTx = new String[numLinTx];
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

        /**************
         * lee Clientes
         **************/
        String[] Exen = new String[maxLeeCLIENTE];
        TxtTemp1 = new String[maxLeeCLIENTE];
        int numCli;
        if (USE_MEMORY_READER) {
            numCli = Lee.leeClientes(wb_Ent, TxtTemp1, Exen);
        } else {
            numCli = Lee.leeClientes(libroEntrada, TxtTemp1, Exen);
        }
        
        nomCli = new String[numCli];
        System.arraycopy(TxtTemp1, 0, nomCli, 0, numCli);
        String[] TxtTemp0 = new String[numCli];
        for (int i = 0; i < numCli; i++) {
            TxtTemp0[i] = "";
        }
        numEmpC = 0;
        for (int j = 0; j < numCli; j++) {
            String[] tmp2 = nomCli[j].split("#")[1].split("#");
            int l = Calc.Buscar(tmp2[0], TxtTemp0);

            if (l == -1) {
                TxtTemp0[numEmpC] = tmp2[0];
                numEmpC++;
            }

        }
        nomEmpC = new String[numEmpC];
        System.arraycopy(TxtTemp0, 0, nomEmpC, 0, numEmpC);

        /***********Extrae Barra de Clientes**********************************/

        String[] tmpo;
        String[] BPeajeC=new String[numCli];

        TxtTemp2 = new String[numCli];
        for (int i = 0; i < numCli; i++) {
        TxtTemp2[i] = "";
        }
        int numBarC = 0;
        for (int j = 0; j < numCli; j++) {
            tmpo=nomCli[j].split("#");
            BPeajeC[j]= tmpo[2];
            int l = Calc.Buscar(tmpo[2], TxtTemp2);
            if (l == -1) {
                TxtTemp2[numBarC] = tmpo[2];
                numBarC++;
            }
        }
        String[] nomBar = new String[numBarC];
        System.arraycopy(TxtTemp2, 0, nomBar, 0, numBarC);


        //Busca clientes exentos
        int[] indiceClienExen = new int[numCli];
        double[] CondiClienExe = new double[numCli];
        int[] indiceClienNOExen = new int[numCli];

        numClienExentos=0;
        int numClienNOExentos=0;
          for (int j=0; j<numCli; j++) {
            double exento= Double.parseDouble(Exen[j]);
            CondiClienExe [j]=exento;
            if(CondiClienExe [j]==0){
            indiceClienExen[numClienExentos]=j;//contiene indice de clientes exentos
            numClienExentos++;
            }
            if(CondiClienExe [j]==-1){
            indiceClienNOExen[numClienNOExentos]=j;//contiene indice de clientes NO exentos
            numClienNOExentos++;
            }

            }


        String[] nombreClientesExen= new String[numClienExentos];
        for (int j=0; j<numClienExentos; j++) {
        nombreClientesExen[j]=nomCli[indiceClienExen[j]];
        }

        String[] nombreCliNOExen= new String[numClienNOExentos];
        for (int j=0; j<numClienNOExentos; j++) {
        nombreCliNOExen[j]=nomCli[indiceClienNOExen[j]];
        }

        //Extrae barra de Clientes NO exentos
        int numBarCNoEx = 0;
        for (int j = 0; j < numCli; j++) {
            tmpo=nomCli[indiceClienNOExen[j]].split("#");
          //  BPeajeC[j]= tmpo[2];
            int l = Calc.Buscar(tmpo[2], TxtTemp2);
            if (l == -1) {
                TxtTemp2[numBarCNoEx] = tmpo[2];
                numBarCNoEx++;
            }

        }
        String[] nomBarNoEx = new String[numBarCNoEx];
        System.arraycopy(TxtTemp2, 0, nomBarNoEx, 0, numBarCNoEx);

        /***************
         * Lee Centrales
         ***************/
        TxtTemp1 = new String[maxLeeGEN];
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
        String[] nomCen = new String[numCen];
        System.arraycopy(TxtTemp1, 0, nomCen, 0, numCen);
        TxtTemp1 = new String[numCen];
        for (int i=0; i<numCen; i++) {
            TxtTemp1[i] = "";
        }
        numEmp = 0;
        for (int j=0; j<numCen; j++) {
            String[] tmp = nomCen[j].split("#");
            int l = Calc.Buscar(tmp[0], TxtTemp1);
            if (l==-1) {
                TxtTemp1[numEmp] = tmp[0];
                numEmp++;
            }
        }
        nomEmp = new String[numEmp];
        System.arraycopy(TxtTemp1, 0, nomEmp, 0, numEmp);

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
        nomProp = new String[numLinIT];
        nomLinIT = new String[numLinIT];
        zonaLin= new int[numLinIT];
        String TxtTemp4[]=new String[numLinIT];
        for (int i = 0; i < numLinIT; i++) {
            nomLinIT[i] = TxtTemp1[i];
            zonaLin[i]=intAux2[i][0];
            nomProp[i]=TxtTemp2[i];
            TxtTemp4[i]="";
        }
         int[] TxtTemp5=new int[numLinIT];

        int numLinTx = 0;
        for (int i = 0; i <numLinIT; i++) {
            int l = Calc.Buscar(nomLinIT[i] + "#" + nomProp[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinIT[i] + "#" + nomProp[i];
                TxtTemp2[numLinTx]=nomProp[i];
                TxtTemp5[numLinTx]=zonaLin[i];
                numLinTx++;
            }
        }
        nomLinTx = new String[numLinTx];//solo registros inico Linea-Transmisor de hoja lintron
//        String[] nomPropTx = new String[numLinTx];
        zonaLinTx= new int[numLinTx];

        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
//            nomPropTx[i]=TxtTemp2[i];
            zonaLinTx[i]=TxtTemp5[i];
        }
         /**************
         * lee Porrratas Distribuidoras
         **************/

        String[] TxtTe = new String[300];
        String[] TxtTe1 = new String[300];
        String[] TxtTe2 = new String[300];
        double[][][] facDxaux= new double[100][100][NUMERO_MESES];
        double[][] prorrataEfirmeAux = new double[NUMERO_MESES][300];
        int num[];
        int sumRM88;
        if (USE_MEMORY_READER) {
            num = Lee.leeDistribuidoras(wb_Ent, TxtTe, TxtTe1, facDxaux);
            sumRM88 = Lee.leeProrrataEfirme(wb_Ent, TxtTe2, prorrataEfirmeAux);
        } else {
            num = Lee.leeDistribuidoras(libroEntrada, TxtTe, TxtTe1, facDxaux);
            sumRM88 = Lee.leeProrrataEfirme(libroEntrada, TxtTe2, prorrataEfirmeAux);
        }
        
        int numSumi=num[1];
        int numDx=num[0];

        nomDx= new String[numDx];
        nomSumi= new String[numSumi];
        nomSumiRM88 =  new String[sumRM88];

        facDx= new double[numDx][numSumi][NUMERO_MESES];
        proEfirme= new double[sumRM88][NUMERO_MESES];

        System.arraycopy(TxtTe, 0, nomDx, 0, numDx);
        System.arraycopy(TxtTe1, 0, nomSumi, 0, numSumi);
        System.arraycopy(TxtTe2, 0, nomSumiRM88, 0, sumRM88);
        
        for(int i=0;i<numDx;i++)
          for(int j=0;j<numSumi;j++)
              System.arraycopy(facDxaux[i][j], 0, facDx[i][j], 0, NUMERO_MESES);

        
        for(int j=0;j<sumRM88;j++)
            for(int k=0;k<NUMERO_MESES;k++)
                proEfirme[j][k]=prorrataEfirmeAux[k][j];

        // Libro Prorrata
        

          /***************
         * Lee Prorratas, Inyeccion Centrales y Consumos Mensuales
         ***************/
        prorrMesC = new double[numLinTx][numCli][NUMERO_MESES];//las prorratas se encuentran en el orden nomLinTx (rgistros unicos hoja lintron)
        prorrMesCTot = new double[numLinTx][NUMERO_MESES];
        double[][] GenerMensual = new double[numCen][NUMERO_MESES];
        double[] GeneTotMesProm= new double[NUMERO_MESES];
        double[][] GenPromMesCen= new double[numCen][NUMERO_MESES];
        double[] GenAnoxCen= new double[numCen];
        int [][] MesesAct=new int[numCen][NUMERO_MESES];
        int [] numMesesAct=new int[numCen];
        double[][] CMesCli = new double[numCli][NUMERO_MESES];
        double[][][] CUE = new double[numCli][3][NUMERO_MESES];
        double[] ECUAnual = new double[2];
        
        //Lee prorratas desde csv (defecto):
        long tInicioLecturaProrratas = System.currentTimeMillis();
        
        File f_prorratasCCSV;
        if (horizon == PeajesConstant.HorizonteCalculo.Anual) {
            f_prorratasCCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_PRORRATACONSUMO + Ano + ".csv");
        } else {
            f_prorratasCCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_PRORRATACONSUMO + Ano + MESES[Mes] + ".csv");
        }
        File f_GMesCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_GMES + Ano + ".csv");
        File f_CMesCSV = new File(DirBaseSal + SLASH + PeajesConstant.PREFIJO_CMES + Ano + ".csv");
        
        if (f_prorratasCCSV.exists() && f_GMesCSV.exists() && f_CMesCSV.exists()) {
            System.out.println("Leyendo archivos csv de prorratas, generacion y consumo mensual..");
            try {
                int nReadP = Lee.leeProrratasCSV(f_prorratasCCSV.getAbsolutePath(), prorrMesC, horizon);
            } catch (IOException e) {
                System.out.println("No se pudo conectar con archivo prorratas " + f_prorratasCCSV.getAbsolutePath());
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
            try {
                int nReadCx = Lee.leeGeneracionMesCSV(f_GMesCSV.getAbsolutePath(), GenerMensual, horizon);
            } catch (IOException e) {
                System.out.println("No se pudo conectar con archivo generacion mensual " + f_GMesCSV.getAbsolutePath());
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
            try {
                int nReadCx = Lee.leeConsumoMesCSV(f_CMesCSV.getAbsolutePath(), CMesCli, horizon);
            } catch (IOException e) {
                System.out.println("No se pudo conectar con archivo consumo mensual " + f_CMesCSV.getAbsolutePath());
                System.out.println("Verifique la ruta y vuelva a intentar. Error: " + e.getMessage());
                return;
            }
        } else {
            // Si no, intenta buscar libro Prorrata:
            System.out.println("Leyendo prorratas, generacion y consumo mensual desde Excel..");
            String libroEntradaP = DirBaseSal + SLASH + "Prorrata" + Ano + ".xlsx";
            XSSFWorkbook wb_Prorratas;
            try {
                wb_Prorratas = new XSSFWorkbook(new java.io.FileInputStream(libroEntradaP));
                if (USE_MEMORY_READER) {
                    Lee.leeProrratasConsumoExcel(wb_Prorratas, prorrMesC);
                } else {
                    Lee.leeProrratasC(libroEntradaP, prorrMesC);
                }
                if (USE_MEMORY_READER) {
                    Lee.leeGeneracionMes(wb_Prorratas, GenerMensual);
                } else {
                    Lee.leeGeneracionMes(libroEntradaP, GenerMensual);
                }
                if (USE_MEMORY_READER) {
                    Lee.leeConsumoMes(wb_Prorratas, CMesCli, CUE);
                } else {
                    Lee.leeConsumoMes(libroEntradaP, CMesCli, CUE);
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
        
        for (int i=0;i<numCen;i++){
            for(int m=0; m<NUMERO_MESES;m++){
            MesesAct[i][m]=0;
                if(GenerMensual[i][m]!=0){
                   GenAnoxCen[i]+=GenerMensual[i][m];
                   MesesAct[i][m]=1;
                   numMesesAct[i]+=1;
                }
            }

            for(int m=0; m<NUMERO_MESES;m++){
            GenPromMesCen[i][m]=0;
                if(MesesAct[i][m]==1){
                    GenPromMesCen[i][m]=GenAnoxCen[i]/numMesesAct[i];
                    GeneTotMesProm[m]+= GenPromMesCen[i][m];
                }
            }
        }
        
        for (int l = 0 ; l < numLinTx; l++){
            for (int m = 0 ; m < NUMERO_MESES; m ++ ){
                for (int c = 0 ; c < numCli; c ++ ){
                prorrMesCTot [l][m]+=prorrMesC[l][c][m];
                }
            }
        }
        
        
        long tFinalLectura = System.currentTimeMillis();
        long tInicioCalculo = System.currentTimeMillis();
        
        
        /****************************
         * Calcula Pagos por Consumos
         ****************************/
        
        
        double[][][] peajeLinCon = new double[numLinTx][numCli][NUMERO_MESES];
        double[][][] ItLinCli = new double[numLinTx][numCli][NUMERO_MESES];
        double[][][] peajeCliTx = new double[numCli][numTx][NUMERO_MESES];
        double[][][] ItCliTx = new double[numCli][numTx][NUMERO_MESES];
        double[][] peajeCli = new double[numCli][NUMERO_MESES];
        double[][] ItCli = new double[numCli][NUMERO_MESES];

        for (int l = 0; l < numLinTx; l++) {
            //System.out.println(nomLinTxO[l]);
            String[] tmp = nomLineasN[l].split("#");
            int l2 = Calc.Buscar( nomLineasN[l], nomLinTx);//antes buscaba nomLinTxO con nomLinTx lineas de peajes ordenadas que era la salida de prorratas pero no estaba ordenada
            if(l2==-1){//ahora la variable nomLinTx son los registros unicos en orden de lintron (asi estaba salida de prorratas)
             System.out.println();
             System.out.println("Línea Trocal - "+nomLineasN[l]+" - en archivo Peaje"+Ano+".xls no se encuentra en la hoja 'lintron' del archivo Ent"+Ano+".xlsx");
             System.out.println("Debe asegurarse que las Líneas en el archivo AVI_COMA.xls sean las mismas de la hoja 'lintron' y ejecutar el botón Peajes");
            }
            else{
                //System.out.println(l+" "+nomLineasN[l]+" "+l2+" "+PeajeN[l][1]+" "+nomLinTx[l2]+prorrMesC[l2][15][1]);
                for (int j = 0; j < numCli; j++) {
                    for (int m = 0; m < NUMERO_MESES; m++) {
                        if (LiquidacionReliquidacion) {
                            peajeLinCon[l][j][m] = prorrMesCTot[l2][m] == 0 ? 0: ((VATTN[l][m] - ITP_N[l][m] - IT_N[l][m] - IT_NG[l][m])*prorrMesC[l2][j][m]);
                            ItLinCli[l][j][m]    = prorrMesCTot[l2][m] == 0 ? 0: ((              IT_N[l][m] + ITP_N[l][m] + IT_NG[l][m])*prorrMesC[l2][j][m]);
                        }
                        else {
                        peajeLinCon[l][j][m] = prorrMesCTot[l2][m] == 0 ? 0:VATTN[l][m]*prorrMesC[l2][j][m] - ITP_N[l][m]*prorrMesC[l2][j][m] - IT_N[l][m]*prorrMesC[l2][j][m]/prorrMesCTot[l2][m];
                            if( prorrMesCTot[l2][m] == 0 && IT_N[l][m] != 0 ) {
                                System.err.println(nomLineasN[l] + " " + m + " " + " IT: " + IT_N[l][m] + " " + ITP_N[l][m] );
                            }
                        ItLinCli[l][j][m] = prorrMesCTot[l2][m] == 0 ? 0 : IT_N[l][m]*prorrMesC[l2][j][m]/prorrMesCTot[l2][m] + ITP_N[l][m]*prorrMesC[l2][j][m];
                        
                        }
                        int t = Calc.Buscar(tmp[1], nombreTx);
                        peajeCliTx[j][t][m] += peajeLinCon[l][j][m]; //j cliente l linea m mes
                        ItCliTx[j][t][m] +=ItLinCli[l][j][m];
                        peajeCli[j][m] += peajeLinCon[l][j][m];
                        ItCli[j][m] += ItLinCli[l][j][m];
                    }
                }
            }
        }
        
        
       /*****************************
       * Calculo de Peajes Unitarios y Cargo Unitario
       *****************************/
        double[][][] PagoBarra= new double[numBarC][numTx][NUMERO_MESES];
        double[][][] ItBarra= new double[numBarC][numTx][NUMERO_MESES];
        double[][] PagoAnoBarra= new double[numBarC][numTx];
        double[][] ConsBarra= new double[numBarC][NUMERO_MESES];
        double[][][] ECUbarra= new double[numBarC][3][12];
        double[][][] PU= new double[numBarC][numTx][NUMERO_MESES];
        double[][][] ITU= new double[numBarC][numTx][NUMERO_MESES];
        

        for(int i=0;i<numBarC;i++){
            for (int m = 0; m < NUMERO_MESES; m++){
            ECUbarra[i][0][m]=0;
            ECUbarra[i][1][m]=0;
            ECUbarra[i][2][m]=0;
            }
        }
        for(int i=0;i<numCli;i++){
            if(CondiClienExe[i]==-1){
             int l = Calc.Buscar(BPeajeC[i], nomBar);// falta poner solo las barras de clientes No exentos
               for(int m=0;m<NUMERO_MESES;m++){
                    ECUbarra[l][0][m]+=CUE[i][0][m];
                    ECUbarra[l][1][m]+=CUE[i][1][m];
                    ECUbarra[l][2][m]+=CUE[i][2][m];
                    ECUAnual[0]+=CUE[i][0][m];
                    ECUAnual[1]+=CUE[i][1][m];
               }
              for(int m=0;m<NUMERO_MESES;m++){
                  ConsBarra[l][m]+=CMesCli[i][m];
                  for(int t=0; t<numTx;t++){
                      PagoBarra[l][t][m]+=peajeCliTx[i][t][m];
                      ItBarra[l][t][m]+=ItCliTx[i][t][m];
              }
             }
            }
        }

        for(int b=0; b<numBarC;b++){
            for(int t=0; t<numTx; t++){
                for (int m=0; m<NUMERO_MESES; m++){
                    PagoAnoBarra[b][t]+=PagoBarra[b][t][m];
                    if(ConsBarra[b][m]==0){
                    PU[b][t][m]=0;
                    ITU[b][t][m]=0;
                    }
                    else{
                    PU[b][t][m]=PagoBarra[b][t][m]/ConsBarra[b][m];
                    ITU[b][t][m]=ItBarra[b][t][m]/ConsBarra[b][m];
                    }
                }
            }
        }
       /*****************************
       * Calculo de CUE
       *****************************/
      double[][] SumaCUE= new double[numBarC][NUMERO_MESES];
      double[][][] ProrrCU= new double[numBarC][2][NUMERO_MESES];
      double[][] ProrrCUano= new double[numBarC][2];
      double[][] ECUbarraAno= new double[numBarC][3];
      double[][][] PagoCU= new double[numBarC][numTx][2];
      double[] PagoCUAnual= new double[2];

      for(int b=0; b<numBarC;b++){
          for(int m=0;m<NUMERO_MESES;m++){
          SumaCUE[b][m]=ECUbarra[b][0][m]+ ECUbarra[b][1][m]+ ECUbarra[b][2][m];
          ECUbarraAno[b][0] += ECUbarra[b][0][m];
          ECUbarraAno[b][1] += ECUbarra[b][1][m];
          ECUbarraAno[b][2] += ECUbarra[b][0][m]+ ECUbarra[b][1][m]+ ECUbarra[b][2][m];
          }
      }
       for(int b=0; b<numBarC;b++){
           for(int m=0;m<NUMERO_MESES;m++){
              if(SumaCUE[b][m]==0){
                    ProrrCU[b][0][m]=ProrrCU[b][1][m]=0;
                }
                else{
                 ProrrCU[b][0][m]= ECUbarra[b][0][m]/SumaCUE[b][m];
                 ProrrCU[b][1][m]= ECUbarra[b][1][m]/SumaCUE[b][m];
                 ProrrCUano[b][0]+= ECUbarra[b][0][m]/ECUbarraAno[b][2];
                 ProrrCUano[b][1]+= ECUbarra[b][1][m]/ECUbarraAno[b][2];       

                }
                 for(int t=0; t<numTx;t++){
                     PagoCU[b][t][0]+=PagoBarra[b][t][m]*ProrrCU[b][0][m];
                     PagoCU[b][t][1]+=PagoBarra[b][t][m]*ProrrCU[b][1][m];

                }
           }
       }
       
       for(int b=0; b<numBarC;b++){
            for(int t=0; t<numTx;t++){
                     PagoCUAnual[0]+=PagoCU[b][t][0];
                     PagoCUAnual[1]+=PagoCU[b][t][1];
            }
       }
       
       
       for(int b=0; b<numBarC;b++){
           
              if(ECUbarraAno[b][2]==0){
                    ProrrCUano[b][0]=ProrrCUano[b][1]=0;
                }
                else{
                 ProrrCUano[b][0]= ECUbarraAno[b][0]/ECUbarraAno[b][2];
                 ProrrCUano[b][1]= ECUbarraAno[b][1]/ECUbarraAno[b][2];       

                }
       }
       

       //ProrrCUCons[][]*peajeCliTx[i][t][m];
       
       
          /*****************************
         * Pagos Clientes NO exentos
         *****************************/
        double[][][] pjeCliTxNOExen = new double[numClienNOExentos][numTx][NUMERO_MESES];
        double[][] pjeCliNOExen = new double[numClienNOExentos][NUMERO_MESES];
        double[][][] ItCliTxNOExen = new double[numClienNOExentos][numTx][NUMERO_MESES];
        double[][] ItCliNOExen = new double[numClienNOExentos][NUMERO_MESES];
        double[][] TotMesTxCliNOExen = new double[numTx][NUMERO_MESES];
        String Barra[];

                for (int i=0; i<numClienNOExentos;i++){
                    Barra=nombreCliNOExen[i].split("#");
                    int l2 = Calc.Buscar(Barra[2], nomBar);
                    for (int t=0; t<numTx; t++) {
                        for (int m=0; m<NUMERO_MESES; m++) {
                            
     /* ITU por cmes*/                       
                               pjeCliTxNOExen[i][t][m]=PU[l2][t][m]*CMesCli[indiceClienNOExen[i]][m];
                               pjeCliNOExen[i][m]+=pjeCliTxNOExen[i][t][m];
                               ItCliTxNOExen[i][t][m]=ITU[l2][t][m]*CMesCli[indiceClienNOExen[i]][m];
                               ItCliNOExen[i][m]+=ItCliTxNOExen[i][t][m];
                               TotMesTxCliNOExen[t][m]+=pjeCliTxNOExen[i][t][m];
                           }
                     }
                }

          /*****************************
         * Pagos por Empresa Clientes No Exentos
         *****************************/
        double[][][] pjeEmpSinAjuTx = new double[numEmp][numTx][NUMERO_MESES];// el nombre que posee en peajes ret es: pjeEmpCNoExTx
        double[][] pjeEmpSinAju = new double[numEmp][NUMERO_MESES];
        double[][][] ItEmpSinAjuTx = new double[numEmp][numTx][NUMERO_MESES];
        double[][] ItEmpSinAju = new double[numEmp][NUMERO_MESES];
        double[][][] pjeEmpSinAjuTxRE2288 = new double[numEmp][numTx][NUMERO_MESES];// el nombre que posee en peajes ret es: pjeEmpCNoExTx
        double[][] pjeEmpSinAjuRE2288 = new double[numEmp][NUMERO_MESES];
        double[][][] ItEmpSinAjuTxRE2288 = new double[numEmp][numTx][NUMERO_MESES];
        double[][] ItEmpSinAjuRE2288 = new double[numEmp][NUMERO_MESES];

          /*****************************
         * Calcula pago por empresa y agrega pago de Distribuidoras a los Suministradores
         *****************************/

        pjeEmpDxTx = new double[numDx][numSumi][numTx][NUMERO_MESES];// el nombre que posee en peajes ret es: pjeEmpCNoExTx

        for (int j = 0; j < numClienNOExentos; j++) {
            String[] tmp = nombreCliNOExen[j].split("#");//Busca la Empresa del Cliente
            int l = Calc.Buscar(tmp[1], nomEmp);
            if(l!=-1){
                 for (int t=0; t<numTx; t++) {
                        for (int m=0; m<NUMERO_MESES; m++) {
                            pjeEmpSinAjuTx[l][t][m]+=pjeCliTxNOExen[j][t][m];
                            pjeEmpSinAju[l][m]+=pjeCliTxNOExen[j][t][m];
                            ItEmpSinAjuTx[l][t][m]+=ItCliTxNOExen[j][t][m];
                            ItEmpSinAju[l][m]+=ItCliTxNOExen[j][t][m];
                            
                        }
                 }
            }
            else{//si no este como empresa de generacion Busca si es una Distribuidora
                int l1 = Calc.Buscar(tmp[1], nomDx);
                 if(l1!=-1){
                    for(int i=0;i<numSumi;i++){
                        int l2= Calc.Buscar(nomSumi[i], nomEmp);//asigna la distribuidara a los suministradores
                        if(l2!=-1){
                            for (int t=0; t<numTx; t++) {
                                if (nomSumi[i].equals("RE2288")){
                                    for (int m=0; m<NUMERO_MESES; m++) {
                                        for(int s=0;s<sumRM88;s++){
                                            pjeEmpSinAjuTxRE2288[s][t][m]+= pjeCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            pjeEmpSinAjuRE2288[s][m]+=pjeCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            ItEmpSinAjuTxRE2288[s][t][m]+= ItCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            ItEmpSinAjuRE2288[s][m]+=ItCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                        }
                                    }    
                                }
                                else {
                                    for (int m=0; m<NUMERO_MESES; m++) {
                                        pjeEmpSinAjuTx[l2][t][m]+= pjeCliTxNOExen[j][t][m]*facDx[l1][i][m];
                                        pjeEmpSinAju[l2][m]+=pjeCliTxNOExen[j][t][m]*facDx[l1][i][m];
                                        ItEmpSinAjuTx[l2][t][m]+= ItCliTxNOExen[j][t][m]*facDx[l1][i][m];
                                        ItEmpSinAju[l2][m]+=ItCliTxNOExen[j][t][m]*facDx[l1][i][m];
                                        pjeEmpDxTx[l1][i][t][m]+=pjeCliTxNOExen[j][t][m]*facDx[l1][i][m];
                                    }   
                                }
                            }
                        }
                        else {                    
                            System.out.println("El suministrador "+nomSumi[i]+" "+j+" "+numTx+" "+" no se encuentra como empresa generadora en 'centrales'");
                        }
                    
                    }   
                }
                if(l1==-1){//si no encuentra la empresa
                    System.out.println("El Suministrador "+tmp[1]+" en 'clientes'"+" "+j+" no esté asignado como Distribuidora o como empresa de Generación");
                }
            }
        }
        
        
        //Calcula prorrata para cuadro IT 
        double[][][] ProrrataRetCenLin1 = new double[numLinTx][numEmp][NUMERO_MESES];
        double[][][] ProrrataRetCenLin2 = new double[numLinTx][numEmp][NUMERO_MESES];
        double[][][] ProrrataRetCenLin3 = new double[numLinTx][numEmp][NUMERO_MESES];

        for (int j = 0; j < numClienNOExentos; j++) {
            int c;
            String[] tmp = nombreCliNOExen[j].split("#");//Busca la Empresa del Cliente
            int l = Calc.Buscar(tmp[1], nomEmp);
            if(l!=-1){
                 for (int t=0; t<numLinTx; t++) {
                        for (int m=0; m<NUMERO_MESES; m++) {
                           c = l;
                           //ProrrataRetCenLin[numLinTx][numEmp][numMeses];
                           //prorrMesC[numLinTx][numCli][numMeses]
                           ProrrataRetCenLin1[t][c][m] += prorrMesC[t][j][m];
                        }
                 }
            }
            else{//si no esté como empresa de generación Busca si es una Distribuidora
                int l1 = Calc.Buscar(tmp[1], nomDx);
                 if(l1!=-1){
                    for(int i=0;i<numSumi;i++){
                        int l2= Calc.Buscar(nomSumi[i], nomEmp);//asigna la distribuidara a los suministradores
                        if(l2!=-1){
                            for (int t=0; t<numLinTx; t++) {
                                if (nomSumi[i].equals("RE2288")){
                                    for (int m=0; m<NUMERO_MESES; m++) {
                                        for(int s=0;s<sumRM88;s++){
                                            c = Calc.Buscar(nomSumiRM88[s], nomEmp);
                                            if ( c == -1) {
                                                System.out.println("El suministrador "+nomSumi[i]+" "+j+" "+numTx+" "+" no se encuentra como empresa generadora en 'centrales'");
                                            //c = s;
                                            }
                                            else {
                                                ProrrataRetCenLin3[t][c][m] += prorrMesC[t][j][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            }
                                            
                                        }
                                    }    
                                }
                                else {
                                    for (int m=0; m<NUMERO_MESES; m++) {
                                       c = l2;
                                       ProrrataRetCenLin2[t][c][m] += prorrMesC[t][j][m]*facDx[l1][i][m];
                                    }   
                                }
                            }
                        }
                        else {                    
                            System.out.println("El suministrador "+nomSumi[i]+" "+j+" "+numTx+" "+" no se encuentra como empresa generadora en 'centrales'");
                        }
                    
                    }   
                }
                if(l1==-1){//si no encuentra la empresa
                    System.out.println("El Suministrador "+tmp[1]+" en 'clientes'"+" "+j+" no esté asignado como Distribuidora o como empresa de Generación");
                }
            }
        }
 
        //imprime prorratas
        try {
            FileWriter writer = new FileWriter(DirBaseSal + SLASH + "prorratas_pago_ret.csv");
            writer.append("Central");
            writer.append(',');
            writer.append("Linea");
            writer.append(',');
            writer.append("Mes");
            writer.append(',');
            writer.append("Prorrata1");
            writer.append(',');
            writer.append("Prorrata2");
            writer.append(',');
            writer.append("Prorrata3");
            writer.append('\n');
            for (int m = 0; m < NUMERO_MESES; m++) {
                for (int i = 0; i < numEmp; i++) {
                    for (int t = 0; t < numLinTx; t++) {
                        writer.append(nomEmp[i]);
                        writer.append(',');
                        writer.append(nomLinTx[t]);
                        writer.append(',');
                        writer.append(Float.toString(m + 1));
                        writer.append(',');
                        writer.append(Double.toString(ProrrataRetCenLin1[t][i][m]));
                        writer.append(',');
                        writer.append(Double.toString(ProrrataRetCenLin2[t][i][m]));
                        writer.append(',');
                        writer.append(Double.toString(ProrrataRetCenLin3[t][i][m]));
                        writer.append('\n');
                    }
                }
            }
            writer.flush();
            writer.close();
        } catch (IOException e) {
            System.out.println("No se pudo escribir con exito prorratas_pago_ret.csv");
            e.printStackTrace(System.out);
        }
        
          /*****************************
         * Pagos Clientes exentos
         *****************************/
        double[][][] peajeClienTxExen = new double[numCli][numTx][NUMERO_MESES];
        double[][] peajeClienExen = new double[numCli][NUMERO_MESES];
        double[][] pjeAnualClienTxExen = new double[numCli][numTx];
        double[] pjeAnualClienExen = new double[numCli];
        double[][] TotMesTxClienExen = new double[numTx][NUMERO_MESES];

         for (int i=0; i<numClienExentos;i++){
                    for (int m=0; m<NUMERO_MESES; m++) {
                          pjeAnualClienExen[i]+=peajeCli[indiceClienExen[i]][m];
                          for (int t=0; t<numTx; t++) {
                                peajeClienTxExen[i][t][m]=peajeCliTx[indiceClienExen[i]][t][m];
                                pjeAnualClienTxExen[i][t]+=peajeCliTx[indiceClienExen[i]][t][m];
                                peajeClienExen[i][m]=peajeCli[indiceClienExen[i]][m];
                                TotMesTxClienExen[t][m]+=peajeClienTxExen[i][t][m];
                           }
                     }
                }
         /*****************************
         * Calcula Ajuste por Clientes exentos
         *****************************/
       double[][] FactAjusClienExenCen = new double[numCen][NUMERO_MESES];
       double[][][] AjusClienExenCenTx = new double[numCen][numTx][NUMERO_MESES];
       for(int j=0;j<numCen;j++){
       for(int m=0;m<NUMERO_MESES;m++){
         FactAjusClienExenCen[j][m]=GenPromMesCen[j][m]/GeneTotMesProm[m];
          for(int t=0;t<numTx;t++){
             AjusClienExenCenTx[j][t][m]=TotMesTxClienExen[t][m]*FactAjusClienExenCen[j][m];
          }
       }
       }
         /*****************************
         * Asigna Ajuste de Exentos a las empresas
         *****************************/
       double[][][] pagoEmpAjusTx= new double[numEmp][numTx][NUMERO_MESES];
       double[][] pagoEmpAjus= new double[numEmp][NUMERO_MESES];
            for (int j = 0; j < numCen; j++) {
            String[] tmp = nomCen[j].split("#");
            int l = Calc.Buscar(tmp[0], nomEmp);
                for (int m = 0; m < NUMERO_MESES; m++) {
                    for (int t = 0; t < numTx; t++) {
                    pagoEmpAjusTx[l][t][m] += AjusClienExenCenTx[j][t][m];
                    pagoEmpAjus[l][m] += AjusClienExenCenTx[j][t][m];
                }
            }
        }

         /*****************************
         * Pagos Totales por Empresa
         *****************************/

        //pago por empresa de generacion por clientes no exentos
        double[][] TotAnualPjeRetEmpGTx = new double[numEmp][numTx];
        double[] TotAnualPjeRetEmpG = new double[numEmp];
        double[][] TotAnualItRetEmpGTx = new double[numEmp][numTx];
        double[] TotAnualItRetEmpG = new double[numEmp];
        double[][] TotAnualPjeRetEmpGTxRE2288 = new double[numEmp][numTx];
        double[] TotAnualPjeRetEmpGRE2288 = new double[numEmp];
        

        //Pago Total (No exento + Ajustex Exento)
        double[][][] TotRetEmpTx = new double[numEmp][numTx][NUMERO_MESES];
        double[][] TotRetEmp = new double[numEmp][NUMERO_MESES];
        double[][][] TotItRetEmpTx = new double[numEmp][numTx][NUMERO_MESES];
        double[][] TotItRetEmp = new double[numEmp][NUMERO_MESES];
        double[][][] TotRetEmpTxRE2288 = new double[numEmp][numTx][NUMERO_MESES];
        double[][] TotRetEmpRE2288 = new double[numEmp][NUMERO_MESES];
        
        double[][][] TotItRetEmpTxRE2288= new double[numEmp][numTx][NUMERO_MESES];
        double[][] TotItRetEmpRE2288= new double[numEmp][NUMERO_MESES];
        double[][] TotAnualItRetEmpGTxRE2288= new double[numEmp][numTx];
        double[] TotAnualItRetEmpGRE2288= new double[numEmp];
        
         for (int j = 0; j < numEmp; j++) {
            for (int m = 0; m < NUMERO_MESES; m++) {
                for (int t = 0; t < numTx; t++) {
                    TotRetEmpTx[j][t][m] = pjeEmpSinAjuTx[j][t][m]+ pagoEmpAjusTx[j][t][m];
                    TotRetEmp[j][m]+=pjeEmpSinAjuTx[j][t][m]+ pagoEmpAjusTx[j][t][m];
                    TotAnualPjeRetEmpGTx[j][t]+=pjeEmpSinAjuTx[j][t][m]+ pagoEmpAjusTx[j][t][m];
                    TotAnualPjeRetEmpG[j]+=pjeEmpSinAjuTx[j][t][m]+ pagoEmpAjusTx[j][t][m];
                    
                    TotItRetEmpTx[j][t][m] = ItEmpSinAjuTx[j][t][m];
                    TotItRetEmp[j][m]+=ItEmpSinAjuTx[j][t][m];
                    TotAnualItRetEmpGTx[j][t]+=ItEmpSinAjuTx[j][t][m];
                    TotAnualItRetEmpG[j]+=ItEmpSinAjuTx[j][t][m];
                 }
            }
        }

         for (int j = 0; j < sumRM88; j++) {
            for (int m = 0; m < NUMERO_MESES; m++) {
                for (int t = 0; t < numTx; t++) {
                    TotRetEmpTxRE2288[j][t][m] = pjeEmpSinAjuTxRE2288[j][t][m];
                    TotRetEmpRE2288[j][m]+=pjeEmpSinAjuTxRE2288[j][t][m];
                    TotAnualPjeRetEmpGTxRE2288[j][t]+=pjeEmpSinAjuTxRE2288[j][t][m];
                    TotAnualPjeRetEmpGRE2288[j]+=pjeEmpSinAjuTxRE2288[j][t][m];
                     
                    TotItRetEmpTxRE2288[j][t][m] =ItEmpSinAjuTxRE2288[j][t][m];
                    TotItRetEmpRE2288[j][m]+= ItEmpSinAjuTxRE2288[j][t][m];
                    TotAnualItRetEmpGTxRE2288[j][t]+= ItEmpSinAjuTxRE2288[j][t][m];
                    TotAnualItRetEmpGRE2288[j]+= ItEmpSinAjuTxRE2288[j][t][m];
                    
                    
                }
            }
         }    
 //agregar tablas de pago finales    


/*****************************************************************/

        // Ordena los archivos de salida de Retiros por empresas
        int[] nc = Calc.OrdenarBurbujaStr(nomCli);
        nomCliO = new String[numCli];
        for (int i = 0; i < numCli; i++) {
            nomCliO[i] = nomCli[nc[i]];
        }

          // Ordena los archivos de salida de Retiros para clientes exentos
        int[] ncExen = Calc.OrdenarBurbujaStr(nombreClientesExen);
        nombreClientesExenO = new String[numClienExentos];
        for (int i=0; i < numClienExentos; i++)
            nombreClientesExenO[i] = nombreClientesExen[ncExen[i]];

          // Ordena los archivos de salida de Retiros para clientes NO exentos
        int[] ncNOExen = Calc.OrdenarBurbujaStr(nombreCliNOExen);
        nombreCliNOExenO = new String[numClienNOExentos];
        for (int i=0; i < numClienNOExentos; i++)
            nombreCliNOExenO[i] = nombreCliNOExen[ncNOExen[i]];


        // Ordena los archivos de salida Ajuste por empresas
        int[] n = Calc.OrdenarBurbujaStr(nomCen);
        nomCenO = new String[numCen];
        for (int i=0; i < numCen; i++){
            nomCenO[i] = nomCen[n[i]];
        }

        // Ordena los archivos de salida Ajuste por empresas
        int[] nb = Calc.OrdenarBurbujaStr(nomBar);
        nomBarO = new String[numBarC];
        for (int i=0; i < numBarC; i++){
            nomBarO[i] = nomBar[nb[i]];
        }
         // Ordena los archivos de salida para PU (deprecated)


        // -------------------------------------------------------------------
        double[][][] prorrMesCO = new double[numLinTx][numCli][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCli; j++) {
                System.arraycopy(prorrMesC[i][nc[j]], 0, prorrMesCO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        double[][][] peajeLinCO = new double[numLinTx][numCli][NUMERO_MESES];
        double[] SumMensualPjeLin = new double[NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCli; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    peajeLinCO[i][j][k] = peajeLinCon[i][nc[j]][k];
                    SumMensualPjeLin[k]+=peajeLinCO[i][j][k];
                }
            }
        }
        // -------------------------------------------------------------------
        peajeClienTxNOExenO = new double[numClienNOExentos][numTx][NUMERO_MESES];
        for (int i = 0; i < numClienNOExentos; i++) {
            for (int j = 0; j < numTx; j++) {
                System.arraycopy(pjeCliTxNOExen[ncNOExen[i]][j], 0, peajeClienTxNOExenO[i][j], 0, NUMERO_MESES);
            }
        }
        // -------------------------------------------------------------------
        double[][] peajeClienNOExenO = new double[numClienNOExentos][NUMERO_MESES];
        for (int i = 0; i < numClienNOExentos; i++) {
            System.arraycopy(pjeCliNOExen[ncNOExen[i]], 0, peajeClienNOExenO[i], 0, NUMERO_MESES);
        }
        // -------------------------------------------------------------------
        nc = Calc.OrdenarBurbujaStr(nomEmpC);
        nomEmpCO = new String[numEmpC];
        for (int i = 0; i < numEmpC; i++) {
            nomEmpCO[i] = nomEmpC[nc[i]];
        }
          // -------------------------------------------------------------------
        ne = Calc.OrdenarBurbujaStr(nomEmp);
        nomEmpO = new String[numEmp];
        for (int i = 0; i < numEmp; i++) {
            nomEmpO[i] = nomEmp[ne[i]];
        }
        
        // ---------------------------------------------------------------------
        double[][] peajeEmpCO = new double[numEmpC][NUMERO_MESES];
        for (int i = 0; i < numEmpC; i++) {
            System.arraycopy(pjeEmpSinAju[nc[i]], 0, peajeEmpCO[i], 0, NUMERO_MESES);
        }
          // ---------------------------------------------------------------------
        peajeClienTxExenO = new double[numClienExentos][numTx][NUMERO_MESES];
        pjeAnualClienTxExenO = new double[numClienExentos][numTx];
        pjeAnualClienExenO= new double[numClienExentos];
        for (int i=0; i < numClienExentos; i++){
        for (int j=0; j < numTx; j++){
        for (int k=0; k < NUMERO_MESES; k++){
            peajeClienTxExenO[i][j][k] = peajeClienTxExen[ncExen[i]][j][k];
            pjeAnualClienTxExenO[i][j]=pjeAnualClienTxExen[ncExen[i]][j];
            pjeAnualClienExenO[i]=pjeAnualClienExen[ncExen[i]];
        }
        }
        }
          // -------------------------------------------------------------------
        double[][] peajeClienExenO = new double[numClienExentos][NUMERO_MESES];
        for (int i = 0; i < numClienExentos; i++) {
            System.arraycopy(peajeClienExen[ncExen[i]], 0, peajeClienExenO[i], 0, NUMERO_MESES);
        }
          // ---------------------------------------------------------------------
       AjusClienExenCenTxO = new double[numCen][numTx][NUMERO_MESES];
       GenAnoxCenO=new double[numCen];//ajuste
       GenPromMesCenO=new double[numCen][NUMERO_MESES];
        for (int j=0; j < numCen; j++)
        for (int t=0; t < numTx; t++)
        for (int m=0; m < NUMERO_MESES; m++){
           AjusClienExenCenTxO[j][t][m]= AjusClienExenCenTx[n[j]][t][m];
           GenAnoxCenO[j]=GenAnoxCen[n[j]];//ajuste
           GenPromMesCenO[j][m]=GenPromMesCen[n[j]][m];
        }
        // ---------------------------------------------------------------------
        AjusEmpCTxO = new double[numEmp][numTx][NUMERO_MESES];
        AjusAnualEmpCTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    AjusEmpCTxO[i][j][k] = pagoEmpAjusTx[ne[i]][j][k];
                     AjusAnualEmpCTxO[i][j]+=AjusEmpCTxO[i][j][k];
                }
            }
        }
        // ---------------------------------------------------------------------
        AjusEmpCO = new double[numEmp][NUMERO_MESES];
        AjusAnualEmpCO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < NUMERO_MESES; j++) {
                AjusEmpCO[i][j] = pagoEmpAjus[ne[i]][j];
                AjusAnualEmpCO[i]+= AjusEmpCO[i][j];
            }
        }
          // ---------------------------------------------------------------------
        TotRetEmpTxO = new double[numEmp][numTx][NUMERO_MESES];
        TotAnualRetEmpTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotRetEmpTxO[i][j][k] = TotRetEmpTx[ne[i]][j][k];//Esto es lo que estaria malo
                     TotAnualRetEmpTxO[i][j]+=TotRetEmpTxO[i][j][k];
                }
            }
        }
        
        
        ne2288 = Calc.OrdenarBurbujaStr(nomSumiRM88);
        nomSumiRM88O = new String[sumRM88];
        for (int i = 0; i < sumRM88; i++) {
            nomSumiRM88O[i] = nomSumiRM88[ne2288[i]];
        }
        
        
        TotRetEmpTxRE2288O = new double[sumRM88][numTx][NUMERO_MESES];
        TotAnualRetEmpTxRE2288O = new double[sumRM88][numTx];
        double[][][] TotItRetEmpTxRE2288O = new double[sumRM88][numTx][NUMERO_MESES];
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotRetEmpTxRE2288O[i][j][k] = TotRetEmpTxRE2288[ne2288[i]][j][k];
                    TotItRetEmpTxRE2288O[i][j][k] = TotItRetEmpTxRE2288[ne2288[i]][j][k];
                     TotAnualRetEmpTxRE2288O[i][j]+=TotRetEmpTxRE2288O[i][j][k];
                }
            }
        }
        
        
        
        
        TotRetItEmpTxO = new double[numEmp][numTx][NUMERO_MESES];
        //TotAnualRetEmpTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotRetItEmpTxO[i][j][k] = TotItRetEmpTx[ne[i]][j][k];
                     //TotAnualRetEmpTxO[i][j]+=TotRetEmpTxO[i][j][k];
                }
            }
        }
        
        
        
                  // ---------------------------------------------------------------------
       RetEmpSinAjuTxO = new double[numEmp][numTx][NUMERO_MESES];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                System.arraycopy(pjeEmpSinAjuTx[ne[i]][j], 0, RetEmpSinAjuTxO[i][j], 0, NUMERO_MESES);
            }
        }
        // ---------------------------------------------------------------------
        TotRetEmpO = new double[numEmp][NUMERO_MESES];
        TotItRetEmpO = new double[numEmp][NUMERO_MESES];
        double[] TotMensualRetEmp=new double[NUMERO_MESES];
        
        TotAnualRetEmpO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < NUMERO_MESES; j++) {
                TotRetEmpO[i][j] = TotRetEmp[ne[i]][j];
                TotItRetEmpO[i][j] = TotItRetEmp[ne[i]][j];
                TotAnualRetEmpO[i]+=TotRetEmp[ne[i]][j];
                TotMensualRetEmp[j]+=TotRetEmp[ne[i]][j];
            }
        }
        // ---------------------------------------------------------------------
        RetEmpSinAjuO = new double[numEmp][NUMERO_MESES];
        for (int i = 0; i < numEmp; i++) {
            System.arraycopy(pjeEmpSinAju[ne[i]], 0, RetEmpSinAjuO[i], 0, NUMERO_MESES);
        }
        // ---------------------------------------------------------------------
        double[][][] TotPjeRetEmpTxO = new double[numEmp][numTx][NUMERO_MESES];
        TotAnualPjeRetEmpTxO = new double[numEmp][numTx];
        TotAnualPjeRetEmpO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotAnualPjeRetEmpTxO[i][j]=TotAnualPjeRetEmpGTx[ne[i]][j]; //sumar
                    TotAnualPjeRetEmpO[i]=TotAnualPjeRetEmpG[ne[i]]; //sumar
                    TotPjeRetEmpTxO[i][j][k] = pjeEmpSinAjuTx[ne[i]][j][k];
                }
            }
        }
        
        TotAnualPjeRetEmpRE2288O = new double[numEmp];
        TotAnualPjeRetEmpTxRE2288O = new double[sumRM88][numTx];
        double[] TotMensualPjeRetEmpRE2288O = new double[NUMERO_MESES];
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotAnualPjeRetEmpTxRE2288O[i][j]=TotAnualPjeRetEmpGTxRE2288[ne2288[i]][j]; //sumar
                    TotAnualPjeRetEmpRE2288O[i]=TotAnualPjeRetEmpGRE2288[ne2288[i]]; //sumar
                    TotMensualPjeRetEmpRE2288O[k]+=pjeEmpSinAjuTxRE2288[i][j][k];
                    //TotPjeRetEmpTxO[i][j][k] = pjeEmpSinAjuTx[ne[i]][j][k];
                }
            }
        }
        
        int[] re2288toNomsumi = new int[numEmp];
        for (int i =0; i<nomEmpO.length;i++){
             re2288toNomsumi[i] = Calc.Buscar(nomEmpO[i], nomSumiRM88O);
        }
        //for (int i = 0; i < numEmp; i++) {
        //    System.out.println(i + " " + re2288toNomsumi[i]);
        //}
        
        TotConRe2288AnualPjeRetEmpTxO = new double[numEmp][numTx];
        TotConRe2288AnualPjeRetEmpO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                //System.out.println(i + " " + re2288toNomsumi[i]);
                TotConRe2288AnualPjeRetEmpTxO[i][j] = re2288toNomsumi[i] == -1? TotAnualPjeRetEmpTxO[i][j]: TotAnualPjeRetEmpTxO[i][j] + TotAnualPjeRetEmpTxRE2288O[re2288toNomsumi[i]][j];
            }    
            //System.out.println(i + "-" + re2288toNomsumi[i]);
            TotConRe2288AnualPjeRetEmpO[i] = re2288toNomsumi[i] == -1? TotAnualPjeRetEmpO[i]: TotAnualPjeRetEmpO[i] + TotAnualPjeRetEmpRE2288O[re2288toNomsumi[i]];
        }
                
        for (int j = 0; j < NUMERO_MESES; j++) {
            TotMensualRetEmp[j]+=TotMensualPjeRetEmpRE2288O[j];
            
        }
        double[][][] TotRetEmpTxRE2288O = new double[sumRM88][numTx][NUMERO_MESES];
        double[][] TotAnualPjeRetEmpGTxRE2288O = new double[numEmp][numTx];
        
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                System.arraycopy(TotRetEmpTxRE2288[ne2288[i]][j], 0, TotRetEmpTxRE2288O[i][j], 0, NUMERO_MESES);
                TotAnualPjeRetEmpGTxRE2288O[i][j]=TotAnualPjeRetEmpGTxRE2288[ne2288[i]][j];
            }
        }
        
        
        TotRetEmpTxRE2288OO = new double[numEmp][numTx][NUMERO_MESES];
        TotRetEmpRE2288O=new double[numEmp][NUMERO_MESES];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    TotRetEmpTxO[i][j][k]= re2288toNomsumi[i] == -1? TotRetEmpTxO[i][j][k]:TotRetEmpTxO[i][j][k]+TotRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];//Esto es lo que estaria malo
                    TotRetEmpO[i][k]+=re2288toNomsumi[i] == -1?0:TotRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];
                    TotRetEmpTxRE2288OO[i][j][k] += re2288toNomsumi[i] == -1?0:TotRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];
                    TotRetEmpRE2288O[i][k]+= re2288toNomsumi[i] == -1?0:TotRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];
                    TotRetItEmpTxO[i][j][k]+= re2288toNomsumi[i] == -1?0:TotItRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];
                    TotItRetEmpO[i][k]+=  re2288toNomsumi[i] == -1?0:TotItRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];
                }
                TotAnualRetEmpTxO[i][j]= re2288toNomsumi[i] == -1?TotAnualRetEmpTxO[i][j]:TotAnualRetEmpTxO[i][j]+TotAnualPjeRetEmpGTxRE2288O[re2288toNomsumi[i]][j];
            }
        }
                /*TotItRetEmpTxRE2288O[j][t][m]
                TotRetItEmpTxO;
                 
                
                /* TotItRetEmpO;*/
        
        
          // --------------------------------------------------------------------
      PUO = new double[numBarC][numTx][NUMERO_MESES];
        for (int j=0; j < numBarC; j++)
        for (int t=0; t < numTx; t++)
            System.arraycopy(PU[nb[j]][t], 0, PUO[j][t], 0, NUMERO_MESES);
        // --------------------------------------------------------------------
       double[][] ProrrCUO= new double[numBarC][2];
       double[][][] PagoCUO= new double[numBarC][numTx][2];
       double[][] ECUbarraO= new double[numBarC][2];
       double[][] PagoAnoBarraO= new double[numBarC][numTx];

        for (int j=0; j < numBarC; j++){
            for (int k=0; k < 2; k++){
                for (int t=0; t < numTx; t++){
                     PagoCUO[j][t][k]=PagoCU[nb[j]][t][k];
                     PagoAnoBarraO[j][t]= PagoAnoBarra[nb[j]][t];
                 //for (int m=0; m < numMeses; m++){    
                     ProrrCUO[j][k]=ProrrCUano[nb[j]][k];
                     ECUbarraO[j][k]= ECUbarraAno[nb[j]][k];
                 //}
          }
        }
        }
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
                String libroSalidaCXLS = DirBaseSal + SLASH + "PagoRet" + Ano + ".xlsx";
                if (!USE_MEMORY_WRITER) {
                    Escribe.crearLibro(libroSalidaCXLS);
                    Escribe.creaH2F_3d2_long(
                            "Pago de Peaje por Línea y Cliente [$]", peajeLinCO,
                            "Línea", nomLineasN,
                            "Cliente", nomCliO,
                            "Mes", MESES,
                            libroSalidaCXLS, "PjeClienLin",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaH2F_3d2_long(
                            "Pago Peaje por Cliente y Transmisor (Clientes No Exentos) [$]", peajeClienTxNOExenO,
                            "Cliente", nombreCliNOExenO,
                            "Transmisor", nombreTx,
                            "Mes", MESES,
                            libroSalidaCXLS, "PjeClienTx",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    if (numClienExentos != 0) {
                        Escribe.creaH2F_3d2_long(
                                "Pago Peaje de Cliente Exento y Transmisor[$]", peajeClienTxExenO,
                                "Cliente", nombreClientesExenO,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                libroSalidaCXLS, "PjeClienTxExen", "#,###,##0;[Red]-#,###,##0;\"-\"");
                    }
                    Escribe.creaH3F_3d_double(
                            "Pago por Ajuste de Retiros Exentos por Central y Transmisor [$]", AjusClienExenCenTxO,
                            "Central", nomCenO,
                            "Transmisor", nombreTx,
                            "Mes", MESES,
                            "Inyeccion Anual", GenAnoxCenO,
                            libroSalidaCXLS, "AjusExenTx", "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaH1F_2d_double(
                            "Ajuste por Empresa [$]", AjusEmpCO,
                            "Empresa", nomEmpO,
                            "Mes", MESES,
                            libroSalidaCXLS, "AjusEmp",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaH3F_4d_double(
                            "Pago por Contratos con Distribuidoras [$]", pjeEmpDxTx,
                            "Sumnistrador", nomSumi,
                            "Distrubuidora", nomDx,
                            "Mes", MESES,
                            "Transmisor", nombreTx,
                            libroSalidaCXLS, "PagosDx", "#,###,##0;[Red]-#,###,##0;\"-\"");

                    Escribe.creaH2F_3d_double(
                            "Pagos de Peaje de Retiro RE2288 por Empresa y Transmisor [$] (Incluye ajuste por Exentos)", TotRetEmpTxRE2288O,
                            "Empresa", nomSumiRM88O,
                            "Transmisor", nombreTx,
                            "Mes", MESES,
                            libroSalidaCXLS, "PagosRE2288",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaH2F_3d_double(
                            "Pagos de Peaje de Retiro por Empresa y Transmisor [$] (Incluye ajuste por Exentos)", TotRetEmpTxO,
                            "Empresa", nomEmpO,
                            "Transmisor", nombreTx,
                            "Mes", MESES,
                            libroSalidaCXLS, "PagoEmpTx",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.crea_SalidaCU(
                            "Cargo Unitario [$/MWh]",
                            "Barra", nomBarO,
                            "Transmisor", nombreTx,
                            "Consumo", "Consumo CU2", "Consumo CU30", ECUbarraO,
                            "Prorrata", "Prorrata CU2", "Prorrata CU30", ProrrCUO,
                            "Pagos", "Pago CU2", "Pago CU30", PagoCUO,
                            PagoAnoBarraO,
                            libroSalidaCXLS, "CargoUnitario",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaH2F_3d_double(
                            "Pago Unitario [$/MWh]", PUO,
                            "Barra", nomBarO,
                            "Transmisor", nombreTx,
                            "Mes", MESES,
                            libroSalidaCXLS, "PUnit",
                            "#,###,##0.###;[Red]-#,###,##0.###;\"-\"");
                    Escribe.crea_verificaRet(
                            "Verifica Pagos de Retiro", libroEntrada,
                            "CUE", "CUE2", "CUE30", "Pago", "Consumo",
                            PagoCUAnual, ECUAnual,
                            "Mes", MESES,
                            "Calculo", TotMensualRetEmp,
                            "Prorrata Línea", SumMensualPjeLin,
                            "Diferencia",
                            "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.crea_verificaCalcPeajes(
                            "Verifica Cálculo de Peajes", libroEntrada,
                            "Mes", MESES,
                            "Peajes", PeajeNMes,
                            "Pago Ret", "Pago Iny", "Diferencia",
                            "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                } else {

                    try {
                        XSSFWorkbook wb_salida = Escribe.crearLibroVacio(libroSalidaCXLS);
                        Escribe.creaH2F_3d2_long(
                                "Pago de Peaje por Línea y Cliente [$]", peajeLinCO,
                                "Línea", nomLineasN,
                                "Cliente", nomCliO,
                                "Mes", MESES,
                                wb_salida, "PjeClienLin",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH2F_3d2_long(
                                "Pago Peaje por Cliente y Transmisor (Clientes No Exentos) [$]", peajeClienTxNOExenO,
                                "Cliente", nombreCliNOExenO,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                wb_salida, "PjeClienTx",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        if (numClienExentos != 0) {
                            Escribe.creaH2F_3d2_long(
                                    "Pago Peaje de Cliente Exento y Transmisor[$]", peajeClienTxExenO,
                                    "Cliente", nombreClientesExenO,
                                    "Transmisor", nombreTx,
                                    "Mes", MESES,
                                    wb_salida, "PjeClienTxExen", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        }
                        Escribe.creaH3F_3d_double(
                                "Pago por Ajuste de Retiros Exentos por Central y Transmisor [$]", AjusClienExenCenTxO,
                                "Central", nomCenO,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                "Inyeccion Anual", GenAnoxCenO,
                                wb_salida, "AjusExenTx", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH1F_2d_double(
                                "Ajuste por Empresa [$]", AjusEmpCO,
                                "Empresa", nomEmpO,
                                "Mes", MESES,
                                wb_salida, "AjusEmp",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH3F_4d_double(
                                "Pago por Contratos con Distribuidoras [$]", pjeEmpDxTx,
                                "Sumnistrador", nomSumi,
                                "Distrubuidora", nomDx,
                                "Mes", MESES,
                                "Transmisor", nombreTx,
                                wb_salida, "PagosDx", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH2F_3d_double(
                                "Pagos de Peaje de Retiro RE2288 por Empresa y Transmisor [$] (Incluye ajuste por Exentos)", TotRetEmpTxRE2288O,
                                "Empresa", nomSumiRM88O,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                wb_salida, "PagosRE2288",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH2F_3d_double(
                                "Pagos de Peaje de Retiro por Empresa y Transmisor [$] (Incluye ajuste por Exentos)", TotRetEmpTxO,
                                "Empresa", nomEmpO,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                wb_salida, "PagoEmpTx",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.crea_SalidaCU(
                                "Cargo Unitario [$/MWh]",
                                "Barra", nomBarO,
                                "Transmisor", nombreTx,
                                "Consumo", "Consumo CU2", "Consumo CU30", ECUbarraO,
                                "Prorrata", "Prorrata CU2", "Prorrata CU30", ProrrCUO,
                                "Pagos", "Pago CU2", "Pago CU30", PagoCUO,
                                PagoAnoBarraO,
                                wb_salida, "CargoUnitario",
                                "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.creaH2F_3d_double(
                                "Pago Unitario [$/MWh]", PUO,
                                "Barra", nomBarO,
                                "Transmisor", nombreTx,
                                "Mes", MESES,
                                wb_salida, "PUnit",
                                "#,###,##0.###;[Red]-#,###,##0.###;\"-\"");
                        Escribe.crea_verificaRet(
                                "Verifica Pagos de Retiro", wb_Ent,
                                "CUE", "CUE2", "CUE30", "Pago", "Consumo",
                                PagoCUAnual, ECUAnual,
                                "Mes", MESES,
                                "Calculo", TotMensualRetEmp,
                                "Prorrata Línea", SumMensualPjeLin,
                                "Diferencia",
                                "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.crea_verificaCalcPeajes(
                                "Verifica Cálculo de Peajes", wb_Ent,
                                "Mes", MESES,
                                "Peajes", PeajeNMes,
                                "Pago Ret", "Pago Iny", "Diferencia",
                                "verifica", "#,###,##0;[Red]-#,###,##0;\"-\"");
                        Escribe.guardaLibroDisco(wb_salida, libroSalidaCXLS);
                        Escribe.guardaLibroDisco(wb_Ent, libroEntrada);
                        wb_Peajes.close();
                        wb_Ent.close();
                        wb_salida.close();
                    } catch (IOException e) {
                        System.out.println("Error al escribir resultados de pagos retiros a archivo " + libroSalidaCXLS);
                        System.out.println(e.getMessage());
                        e.printStackTrace(System.err);
                    }
                }
            }

            //Escribe archivos csv de salida:
            String sEscribeCSV = PeajesCDEC.getOptionValue("Imprime pagos a csv", PeajesConstant.DataType.BOOLEAN);
            boolean bEscribeCSV = Boolean.parseBoolean(sEscribeCSV);
            if (bEscribeCSV) {

                //Escribe reporte PjeClienLin:
                String libroSalidaCLCSV = DirBaseSal + SLASH + "PjeClienLin" + Ano + ".csv";
                BufferedWriter writerCSV = null;
                try {
                    writerCSV = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(libroSalidaCLCSV), StandardCharsets.ISO_8859_1));
                    String sLineText;
                    //Escribimos el header:
                    writerCSV.write("Tramo,Transmisor,Cliente,Suministrador,Barra,Mes,PagoxLinea[$]");
    //                writerCSV.write("Tramo,Transmisor,Cliente,Suministrador,Barra,Mes,Pago(todos)xLinea[$],Pago(No-Exentos)xLinea[$]");
                    writerCSV.newLine();
                    //Escribimos los datos:
                    for (int l=0; l<numLinTx; l++) {
                        String[] sTramoTransmisor = nomLineasN[l].split("#");
                        assert (sTramoTransmisor.length == 2) : "Como se formaron estos nombres?";
                        for (int c=0; c<numCli; c++) {
                            String[] sClienteSuministradorBarra = nomCliO[c].split("#");
                            assert (sClienteSuministradorBarra.length == 3) : "Como se formaron estos nombres?";
                            for (int m=0; m<NUMERO_MESES; m++) {
                                sLineText = "";
                                for (String s : sTramoTransmisor) {
                                    sLineText += s + ","; //Tramo,Transmisor
                                }
                                for (String s : sClienteSuministradorBarra) {
                                    sLineText += s + ","; //Cliente,Suministrador,Barra
                                }
                                sLineText += MESES[m] + ","; //mes
                                sLineText += peajeLinCO[l][c][m] + ","; //pagostodos

                                writerCSV.write(sLineText);
                                writerCSV.newLine();
                            }
                        }
                    }
                    System.out.println("Finalizado escritura de resultados PjeClienLin.csv");
                } catch (IOException e) {
                    System.out.println("WARNING: No se pudo escribir PjeClienLin.csv. Error: " + e.getMessage());
                    e.printStackTrace(System.out);
                } finally {
                    if (writerCSV != null) {
                        try {
                            writerCSV.close();
                        } catch (IOException e) {
                            System.out.println("No se pudo cerrar conexion con PjeClienLin.csv. Error: " + e.getMessage());
                            e.printStackTrace(System.out);
                        }
                    }
                }

                //Escribe reporte PjeClienTx:
                String libroSalidaCTCSV = DirBaseSal + SLASH + "PjeClienTx" + Ano + ".csv";
                try {
                    writerCSV = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(libroSalidaCTCSV), StandardCharsets.ISO_8859_1));
                    String sLineText;
                    //Escribimos el header:
                    writerCSV.write("Cliente,Suministrador,Barra,Transmisor,Mes,PagoUnitario,Energia");
    //                writerCSV.write("Tramo,Transmisor,Cliente,Suministrador,Barra,Mes,Pago(todos)xLinea[$],Pago(No-Exentos)xLinea[$]");
                    writerCSV.newLine();
                    //Escribimos los datos:
                    for (int c = 0; c < numCli; c++) {
                        String[] sClienteSuministradorBarra = nomCli[c].split("#");
                        assert (sClienteSuministradorBarra.length == 3) : "Como se formaron estos nombres?";
                        int nPosBarraC = Calc.Buscar(sClienteSuministradorBarra[2], nomBar);
                        for (int t = 0; t < numTx; t++) {
                            for (int m = 0; m < NUMERO_MESES; m++) {
                                sLineText = "";
                                for (String s : sClienteSuministradorBarra) {
                                    sLineText += s + ","; //Cliente,Suministrador,Barra
                                }
                                sLineText += nombreTx[t] + ","; //mes
                                sLineText += MESES[m] + ","; //mes
                                if (CondiClienExe[c] == -1) {
                                    sLineText += PU[nPosBarraC][t][m] + ","; //PU
                                } else if (CondiClienExe[c] == 0) {
                                    //TODO: Que hacemos con los exentos?
                                    sLineText += 0.0 + ",";
                                }
                                sLineText += CMesCli[c][m]; //Energia
                                writerCSV.write(sLineText);
                                writerCSV.newLine();
                            }
                        }
                    }
                    System.out.println("Finalizado escritura de resultados PjeClienTx.csv");
                } catch (IOException e) {
                    System.out.println("WARNING: No se pudo escribir PjeClienTx.csv. Error: " + e.getMessage());
                    e.printStackTrace(System.out);
                } finally {
                    if (writerCSV != null) {
                        try {
                            writerCSV.close();
                        } catch (IOException e) {
                            System.out.println("No se pudo cerrar conexion con PjeClienTx.csv. Error: " + e.getMessage());
                            e.printStackTrace(System.out);
                        }
                    }
                }
            }
        }
        
        if (horizon == PeajesConstant.HorizonteCalculo.Mensual) {
            LiquiMesRet(MESES[Mes], Ano);
        }
        long tFinalEscritura = System.currentTimeMillis();
        System.out.println("Pagos de Retiro Calculados");
        System.out.println("Tiempo Adquisicion de datos     : " + DosDecimales.format((tFinalLectura - tInicioLectura) / 1000.0) + " seg");
        System.out.println("Tiempo Cálculo                  : " + DosDecimales.format((tFinalCalculo - tInicioCalculo) / 1000.0) + " seg");
        System.out.println("Tiempo Escritura de Resultados  : " + DosDecimales.format((tFinalEscritura - tInicioEscritura) / 1000.0) + " seg");
        System.out.println();
    }

    public static void LiquiMesRet(String mes, int Ano) {
        int m = 0;
        for (int i = 0; i < NUMERO_MESES; i++) {
            if (mes.equals(MESES[i])) {
                m = i;
            }
        }
        String libroSalidaCXLSMes = DirBaseSal + SLASH + "PagoRet" + MESES[m] + ".xlsx";
        if (!USE_MEMORY_WRITER) {
            Escribe.crearLibro(libroSalidaCXLSMes);
            Escribe.creaLiquidacionMes(m,
                    "Pago de Peajes por Retiro",
                    RetEmpSinAjuTxO,
                    TotRetEmpTxRE2288OO,
                    TotRetEmpTxO, /*det mensual*/
                    TotRetItEmpTxO,
                    RetEmpSinAjuO,
                    TotRetEmpRE2288O,
                    TotRetEmpO, /*tot mensual*/
                    TotItRetEmpO,
                    "Empresa",
                    nomEmpO,
                    "Transmisor",
                    nombreTx,
                    "Tabla 2-1: Pagos de Peajes de Retiro por Suministrador",
                    "Tabla 2-2: Pago de Retiro por RE2288",
                    "Tabla 2-3: Pagos de Peajes de Retiro Incluyendo Pago de Retiro por RE2288",
                    "Tabla 2-4: IT de Retiro",
                    libroSalidaCXLSMes, MESES[m], Ano,
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaProrrataMes(m,
                    "Participación de Retiros [%]", prorrMesC, "Participación " + MESES[m],
                    "Cliente", nomCli,
                    "Línea", nomLinTx,
                    "AIC", zonaLinTx,
                    libroSalidaCXLSMes, "PartRet" + MESES[m],
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaTabla1C_long(m,
                    "Pago de Peaje por Clientes " + MESES[m] + " [$]", peajeClienTxNOExenO,
                    "Cliente", nombreCliNOExenO,
                    "Transmisor", nombreTx,
                    libroSalidaCXLSMes, "Pagos",
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaTabla2CDx_double(m,
                    "Pago de Peaje por Contratos con Distribuidoras" + MESES[m] + " [$]", pjeEmpDxTx,
                    "Suministrador", nomSumi,
                    "Transmisor", nombreTx,
                    "Distribuidora", nomDx,
                    facDx,
                    libroSalidaCXLSMes, "PagosDx",
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            if (numClienExentos != 0) {
                Escribe.creaTabla1C_long(m,
                        "Pago Peaje Exento " + MESES[m] + " [$]", peajeClienTxExenO,
                        "Cliente", nombreClientesExenO,
                        "Transmisor", nombreTx,
                        libroSalidaCXLSMes, "PagosExentos",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaTabla2C_double(m,
                        "Ajustes de Pagos correspondientes a " + MESES[m] + " por Central [$]", AjusClienExenCenTxO,
                        "Central", nomCenO,
                        "Transmisor", nombreTx,
                        "Inyeccion Mes", GenPromMesCenO,
                        libroSalidaCXLSMes, "Ajuste" + MESES[m], "#,###,##0;[Red]-#,###,##0;\"-\"");
            }
            Escribe.creaTabla1C_float(m,
                    "Peajes Unitarios " + MESES[m] + " [$/MWh]", PUO,
                    "Barra", nomBarO,
                    "Transmisor", nombreTx,
                    libroSalidaCXLSMes, "PeajesUnitarios",
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
            Escribe.creaTabla1C_float(m,
                    "Peajes RE2288 " + MESES[m] + " [$/MWh]", TotRetEmpTxRE2288O,
                    "Suministrador", nomSumiRM88O,
                    "Transmisor", nombreTx,
                    libroSalidaCXLSMes, "PeajesRE2288",
                    "#,###,##0;[Red]-#,###,##0;\"-\"");
        } else {

            try {
                XSSFWorkbook wb_salida = Escribe.crearLibroVacio(libroSalidaCXLSMes);
                Escribe.creaLiquidacionMes(m,
                        "Pago de Peajes por Retiro",
                        RetEmpSinAjuTxO,
                        TotRetEmpTxRE2288OO,
                        TotRetEmpTxO, /*det mensual*/
                        TotRetItEmpTxO,
                        RetEmpSinAjuO,
                        TotRetEmpRE2288O,
                        TotRetEmpO, /*tot mensual*/
                        TotItRetEmpO,
                        "Empresa",
                        nomEmpO,
                        "Transmisor",
                        nombreTx,
                        "Tabla 2-1: Pagos de Peajes de Retiro por Suministrador",
                        "Tabla 2-2: Pago de Retiro por RE2288",
                        "Tabla 2-3: Pagos de Peajes de Retiro Incluyendo Pago de Retiro por RE2288",
                        "Tabla 2-4: IT de Retiro",
                        wb_salida, MESES[m], Ano,
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaProrrataMes(m,
                        "Participación de Retiros [%]", prorrMesC, "Participación " + MESES[m],
                        "Cliente", nomCli,
                        "Línea", nomLinTx,
                        "AIC", zonaLinTx,
                        wb_salida, "PartRet" + MESES[m],
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaTabla1C_long(m,
                        "Pago de Peaje por Clientes " + MESES[m] + " [$]", peajeClienTxNOExenO,
                        "Cliente", nombreCliNOExenO,
                        "Transmisor", nombreTx,
                        wb_salida, "Pagos",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaTabla2CDx_double(m,
                        "Pago de Peaje por Contratos con Distribuidoras" + MESES[m] + " [$]", pjeEmpDxTx,
                        "Suministrador", nomSumi,
                        "Transmisor", nombreTx,
                        "Distribuidora", nomDx,
                        facDx,
                        wb_salida, "PagosDx",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                if (numClienExentos != 0) {
                    Escribe.creaTabla1C_long(m,
                            "Pago Peaje Exento " + MESES[m] + " [$]", peajeClienTxExenO,
                            "Cliente", nombreClientesExenO,
                            "Transmisor", nombreTx,
                            wb_salida, "PagosExentos",
                            "#,###,##0;[Red]-#,###,##0;\"-\"");
                    Escribe.creaTabla2C_double(m,
                            "Ajustes de Pagos correspondientes a " + MESES[m] + " por Central [$]", AjusClienExenCenTxO,
                            "Central", nomCenO,
                            "Transmisor", nombreTx,
                            "Inyeccion Mes", GenPromMesCenO,
                            wb_salida, "Ajuste" + MESES[m], "#,###,##0;[Red]-#,###,##0;\"-\"");
                }
                Escribe.creaTabla1C_float(m,
                        "Peajes Unitarios " + MESES[m] + " [$/MWh]", PUO,
                        "Barra", nomBarO,
                        "Transmisor", nombreTx,
                        wb_salida, "PeajesUnitarios",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");
                Escribe.creaTabla1C_float(m,
                        "Peajes RE2288 " + MESES[m] + " [$/MWh]", TotRetEmpTxRE2288O,
                        "Suministrador", nomSumiRM88O,
                        "Transmisor", nombreTx,
                        wb_salida, "PeajesRE2288",
                        "#,###,##0;[Red]-#,###,##0;\"-\"");

                Escribe.guardaLibroDisco(wb_salida, libroSalidaCXLSMes);
                wb_salida.close();

            } catch (IOException e) {
                System.out.println("Error al escribir resultados de pagos retiros a archivo " + libroSalidaCXLSMes);
                System.out.println(e.getMessage());
                e.printStackTrace(System.err);
            }
        }

        System.out.println("Archivo Pago de Retiro Mensual creado");
        System.out.println();
        
    }


}
