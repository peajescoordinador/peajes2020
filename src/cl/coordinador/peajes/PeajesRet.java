package cl.coordinador.peajes;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

import java.io.*;
import java.text.DecimalFormat;
/**
 *
 * @author vtoro
 */


public class PeajesRet {

   private static String slash = File.separator;
    private static final int numMeses = 12;
    static String[] nomMes = {"Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul",
            "Ago", "Sep", "Oct", "Nov", "Dic"};
    static String DirBaseSal;
    static double[][][]  RetEmpSinAjuTxO;
    static double[][] RetEmpSinAjuO;
    static double[][][] AjusEmpCTxO;
    static double[][] AjusAnualEmpCTxO;
    static double[][] TotRetEmpO;
    static double[][] TotItRetEmpO;
    static double[] TotAnualRetEmpO;
    static double[][][] TotRetEmpTxO;
    static double[][][] TotRetItEmpTxO;
    static double[][][] TotRetEmpTxRE2288O;
    static double[][][] TotRetItEmpTxRE2288O;
    static double[][] TotAnualRetEmpTxO;
    static double[][] TotAnualRetEmpTxRE2288O;
    static double[][] AjusEmpCO;
    static double[] AjusAnualEmpCO;
    static String[] nomEmpO;
    static int numEmp;
    static String[] nombreTx;
    static String[] nomCli;
    static String[] nomLinIT;
    static double[][][] prorrMesC;
    static double[][] prorrMesCTot;
    static String[] nomLinTx;
    static int[] zonaLin;
    static int[] zonaLinTx;
    static double[][][] PUO;
    static String[] nomBarO;
    static String[] nomCliO;
    static double[][][] peajeClienTxNOExenO;
    static String[] nombreCliNOExenO;
    static String[] nombreClientesExenO;
    static double[][][] peajeClienTxExenO;
    static double[][][] AjusClienExenCenTxO ;
    static String[] nomCenO ;
    static double[] GenAnoxCenO;
    static double[][] GenPromMesCenO;
    static String[] nomEmpC;
    static int numEmpC;
    static String[] nomEmp;
    static int numTx;
    static int[] ne;
    static int[] ne2288;
    static String[] nomEmpCO;
    static String[] nomSumiRM88O;
    static double[][] TotAnualPjeRetEmpTxO;
    static double[][] TotConRe2288AnualPjeRetEmpTxO;
    static double[][] TotAnualPjeRetEmpTxRE2288O;
    static double[] TotAnualPjeRetEmpO;
    static double[] TotConRe2288AnualPjeRetEmpO;
    static double[] TotAnualPjeRetEmpRE2288O;
    static  double[][] pjeAnualClienTxExenO;
    static double[] pjeAnualClienExenO;
    static int numClienExentos;
    static double[][][][] pjeEmpDxTx ;
    static  String[] nomDx;
    static String[] nomSumi;
    static String[] nomSumiRM88;
    static double[][][] facDx;
    static double[][] proEfirme;
    static double[][][] TotRetEmpTxRE2288OO;
    static     double[][] TotRetEmpRE2288O;

    public static void calculaPeajesRet(File DirEntrada, File DirSalida, int Ano,
            boolean LiquidacionReliquidacion) {

        String DirBaseEnt = DirEntrada.toString();
        DirBaseSal = DirSalida.toString();
        DecimalFormat DosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();

        // Libro peajes
        String libroEntrada = DirBaseSal + slash + "Peaje" + Ano + ".xlsx";

        /************
         * lee Peajes e IT
         ************/
        double[][] longAux = new double[1000][numMeses];
        double[][] longAuxVATT = new double[1000][numMeses];
        double[][] longAuxIT = new double[1000][numMeses];
        double[][] longAuxITG = new double[1000][numMeses];
        double[][] longAuxITP = new double[1000][numMeses];
        String[] TxtTemp = new String[1000];
        String[] TxtTempIT = new String[1000];
        int numLinea = Lee.leePeajes(libroEntrada, TxtTemp, longAux);
        int numLineaVATT = Lee.leeIT(libroEntrada, TxtTempIT, longAuxVATT,"VATT");
        int numLineaIT = Lee.leeIT(libroEntrada, TxtTempIT, longAuxIT,"ITER");
        int numLineaITG = Lee.leeIT(libroEntrada, TxtTempIT, longAuxITG,"ITEG");
        int numLineaITP = Lee.leeIT(libroEntrada, TxtTempIT, longAuxITP,"ITP");
        String[] nomLineasN = new String[numLinea];
        double[][] PeajeN = new double[numLinea][12];
        double[][] VATTN = new double[numLineaVATT][12];
        double[][] IT_N = new double[numLineaIT][12];
        double[][] IT_NG = new double[numLineaITG][12];     
        double[][] ITP_N = new double[numLineaITP][12];
        double[] PeajeNMes = new double[numMeses];
        double[] VATTNMes = new double[numMeses];
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
        libroEntrada = DirBaseEnt + slash + "Ent" + Ano + ".xlsx";

        /**********
         * lee VATT
         **********/
        double[][] Aux = new double[2500][numMeses];
        String[] TxtTemp1 = new String[2500];
        String[] TxtTemp2 = new String[2500];
        int numLineasVATT = Lee.leeVATT(libroEntrada, TxtTemp1, TxtTemp2,
                Aux);
        String[] nomLinVATT = new String[numLineasVATT];
        String[] nomProp = new String[numLineasVATT];
        String[] TxtTemp3 = new String[numLineasVATT];
        for (int i = 0; i < numLineasVATT; i++) {
            TxtTemp3[i] = "";
        }
        numTx = 0;
        for (int i = 0; i < numLineasVATT; i++) {
            nomLinVATT[i] = TxtTemp1[i];
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
        for (int i = 0; i < numTx; i++) {
            nombreTx[i] = TxtTemp3[i];
        }
       /*nomLinTx = new String[numLinTx];
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
        String[] TxtTemp0 = new String[600];
        String[] Exen = new String[2500];

        int numCli = Lee.leeClientes(libroEntrada, TxtTemp1, Exen);
        nomCli = new String[numCli];
        System.arraycopy(TxtTemp1, 0, nomCli, 0, numCli);
        TxtTemp0 = new String[numCli];
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

        String[] tmpo = new String[numCli];
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
        for (int j = 0; j < numBarC; j++) {
            nomBar[j] = TxtTemp2[j];
        }


        //Busca clientes exentos
        int[] indiceClienExen = new int[numCli];
        double[] CondiClienExe = new double[numCli];
        int[] indiceClienNOExen = new int[numCli];

        numClienExentos=0;
        int numClienNOExentos=0;
          for (int j=0; j<numCli; j++) {
            double exento= Double.valueOf(Exen[j]).doubleValue();
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
        for (int j = 0; j < numBarCNoEx; j++) {
            nomBarNoEx[j] = TxtTemp2[j];
        }

        /***************
         * Lee Centrales
         ***************/
        TxtTemp1 = new String[600];
        float[] Temp1 = new float[600];
        float[] Temp2= new float[600];
        int numCen = Lee.leeCentrales(libroEntrada, TxtTemp1,Temp1,Temp2);
        String[] nomCen = new String[numCen];
        for(int i=0; i < numCen; i++){
            nomCen[i] = TxtTemp1[i];
        }
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
        for (int j=0; j<numEmp; j++) {
            nomEmp[j] = TxtTemp1[j];
        }

        /************
         * lee L’neas
         *************/
        TxtTemp1 = new String[2000];
        int numLineas = Lee.leeDeflin(libroEntrada, TxtTemp1, Aux);
        double[][] paramLineas = new double[numLineas][10];
        String[] nomLin = new String[numLineas];
        for (int i = 0; i < numLineas; i++) {
            for (int j = 0; j <= 8; j++) {
                paramLineas[i][j] = Aux[i][j];
            }
            nomLin[i] = TxtTemp1[i];
        }

        /**********************
         * lee L’neas Troncales
         **********************/
        TxtTemp1 = new String[2000];
        TxtTemp2 = new String[2000];
        int[] intAux1 = new int[600];
        int[][] intAux2 = new int[600][numMeses];
        int numLinIT = Lee.leeLintron(libroEntrada, TxtTemp1,
                nomLin, TxtTemp2,intAux1, intAux2);
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
         TxtTemp3=new String[numLinIT];
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
        nomLinTx = new String[numLinTx];//solo registros Ïnico L’nea#Transmisor de hoja lintron
        String[] nomPropTx = new String[numLinTx];
        zonaLinTx= new int[numLinTx];

        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
            nomPropTx[i]=TxtTemp2[i];
            zonaLinTx[i]=TxtTemp5[i];
        }
         /**************
         * lee Porrratas Distribuidoras
         **************/

        String[] TxtTe = new String[300];
        String[] TxtTe1 = new String[300];
        String[] TxtTe2 = new String[300];
        double[][][] facDxaux= new double[100][100][numMeses];
        double[][] prorrataEfirmeAux = new double[numMeses][300];

        int num[]= Lee.leeDistribuidoras(libroEntrada,TxtTe,TxtTe1,facDxaux);
        int sumRM88 = Lee.leeProrrataEfirme(libroEntrada, TxtTe2, prorrataEfirmeAux);
        
        int numSumi=num[1];
        int numDx=num[0];

        nomDx= new String[numDx];
        nomSumi= new String[numSumi];
        nomSumiRM88 =  new String[sumRM88];

        facDx= new double[numDx][numSumi][numMeses];
        proEfirme= new double[sumRM88][numMeses];

        for(int i=0;i<numDx;i++){
            nomDx[i]=TxtTe[i];
        }
        for(int i=0;i<numSumi;i++){
            nomSumi[i]=TxtTe1[i];
        }
        System.arraycopy(TxtTe2, 0, nomSumiRM88, 0, sumRM88);
        
        for(int i=0;i<numDx;i++)
          for(int j=0;j<numSumi;j++)
              for(int k=0;k<numMeses;k++)
                  facDx[i][j][k]=facDxaux[i][j][k];

        
        for(int j=0;j<sumRM88;j++)
            for(int k=0;k<numMeses;k++)
                proEfirme[j][k]=prorrataEfirmeAux[k][j];

        // Libro Prorrata
        String libroEntradaP = DirBaseSal + slash + "Prorrata" + Ano + ".xlsx";

          /***************
         * Lee Inyeccion Centrales
         ***************/
        double[][] GenerMensual = new double[numCen][numMeses];
        double[] GeneTotMesProm= new double[numMeses];
        double[][] GenPromMesCen= new double[numCen][numMeses];
        double[] GenAnoxCen= new double[numCen];
        int [][] MesesAct=new int[numCen][numMeses];
        int [] numMesesAct=new int[numCen];
        Lee.leeGeneracionMes(libroEntradaP,GenerMensual);

        for (int i=0;i<numCen;i++){
            for(int m=0; m<numMeses;m++){
            MesesAct[i][m]=0;
            if(GenerMensual[i][m]!=0){
               GenAnoxCen[i]+=GenerMensual[i][m];
               MesesAct[i][m]=1;
               numMesesAct[i]+=1;
            }
            }

            for(int m=0; m<numMeses;m++){
            GenPromMesCen[i][m]=0;
            if(MesesAct[i][m]==1){
            GenPromMesCen[i][m]=GenAnoxCen[i]/numMesesAct[i];
            GeneTotMesProm[m]+= GenPromMesCen[i][m];
            }
            }
        }
         /***************
         * Lee Consumo Mensuales por Cliente
         ***************/
        double[][] CMesCli = new double[numCli][numMeses];
        double[][][] CUE = new double[numCli][3][numMeses];
        double[] ECUAnual = new double[2];
        
        
        //System.out.println( numCli);
        Lee.leeConsumoMes(libroEntradaP,CMesCli,CUE);

       /* for(int i=0;i<numCli;i++){
            ECUAnual[0]+=CUE[i][0];
            ECUAnual[1]+=CUE[i][1];
        }
        * 
        */


        /**************************
         * lee Prorratas de Consumo
         **************************/
        prorrMesC = new double[numLinTx][numCli][numMeses];//las prorratas se encuentran en el orden nomLinTx (rgistros unicos hoja lintron)
        prorrMesCTot = new double[numLinTx][numMeses];
        Lee.leeProrratasC(libroEntradaP, prorrMesC);
        
        for (int l = 0 ; l < numLinTx; l++){
            for (int m = 0 ; m < numMeses; m ++ ){
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
        
        
        double[][][] peajeLinCon = new double[numLinTx][numCli][numMeses];
        double[][][] ItLinCli = new double[numLinTx][numCli][numMeses];
        double[][][] peajeCliTx = new double[numCli][numTx][numMeses];
        double[][][] ItCliTx = new double[numCli][numTx][numMeses];
        double[][] peajeCli = new double[numCli][numMeses];
        double[][] ItCli = new double[numCli][numMeses];

        for (int l = 0; l < numLinTx; l++) {
            //System.out.println(nomLinTxO[l]);
            String[] tmp = nomLineasN[l].split("#");
            int l2 = Calc.Buscar( nomLineasN[l], nomLinTx);//antes buscaba nomLinTxO con nomLinTx lineas de peajes ordenadas que era la salida de prorratas pero no estaba ordenada
            if(l2==-1){//ahora la variable nomLinTx son los registros unicos en orden de lintron (asi estaba salida de prorratas)
             System.out.println();
             System.out.println("L’nea Trocal - "+nomLineasN[l]+" - en archivo Peaje"+Ano+".xls no se encuentra en la hoja 'lintron' del archivo Ent"+Ano+".xlsx");
             System.out.println("Debe asegurarse que la L’neas en el archivo AVI_COMA.xls sean las mismas de la hoja 'lintron' y ejecutar el bot—n Peajes");
            }
            else{
                //System.out.println(l+" "+nomLineasN[l]+" "+l2+" "+PeajeN[l][1]+" "+nomLinTx[l2]+prorrMesC[l2][15][1]);
                for (int j = 0; j < numCli; j++) {
                    for (int m = 0; m < numMeses; m++) {
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
        double[][][] PagoBarra= new double[numBarC][numTx][numMeses];
        double[][][] ItBarra= new double[numBarC][numTx][numMeses];
        double[][] PagoAnoBarra= new double[numBarC][numTx];
        double[][] ConsBarra= new double[numBarC][numMeses];
        double[][][] ECUbarra= new double[numBarC][3][12];
        double[][][] PU= new double[numBarC][numTx][numMeses];
        double[][][] ITU= new double[numBarC][numTx][numMeses];
        

        for(int i=0;i<numBarC;i++){
            for (int m = 0; m < numMeses; m++){
            ECUbarra[i][0][m]=0;
            ECUbarra[i][1][m]=0;
            ECUbarra[i][2][m]=0;
            }
        }
        for(int i=0;i<numCli;i++){
            if(CondiClienExe[i]==-1){
             int l = Calc.Buscar(BPeajeC[i], nomBar);// falta poner solo las barras de clientes No exentos
               for(int m=0;m<numMeses;m++){
                    ECUbarra[l][0][m]+=CUE[i][0][m];
                    ECUbarra[l][1][m]+=CUE[i][1][m];
                    ECUbarra[l][2][m]+=CUE[i][2][m];
                    ECUAnual[0]+=CUE[i][0][m];
                    ECUAnual[1]+=CUE[i][1][m];
               }
              for(int m=0;m<numMeses;m++){
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
                for (int m=0; m<numMeses; m++){
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
      double[][] SumaCUE= new double[numBarC][numMeses];
      double[][][] ProrrCU= new double[numBarC][2][numMeses];
      double[][] ProrrCUano= new double[numBarC][2];
      double[][] ECUbarraAno= new double[numBarC][3];
      double[][][] PagoCU= new double[numBarC][numTx][2];
      double[] PagoCUAnual= new double[2];

      for(int b=0; b<numBarC;b++){
          for(int m=0;m<numMeses;m++){
          SumaCUE[b][m]=ECUbarra[b][0][m]+ ECUbarra[b][1][m]+ ECUbarra[b][2][m];
          ECUbarraAno[b][0] += ECUbarra[b][0][m];
          ECUbarraAno[b][1] += ECUbarra[b][1][m];
          ECUbarraAno[b][2] += ECUbarra[b][0][m]+ ECUbarra[b][1][m]+ ECUbarra[b][2][m];
          }
      }
       for(int b=0; b<numBarC;b++){
           for(int m=0;m<numMeses;m++){
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
        double[][][] pjeCliTxNOExen = new double[numClienNOExentos][numTx][numMeses];
        double[][] pjeCliNOExen = new double[numClienNOExentos][numMeses];
        double[][][] ItCliTxNOExen = new double[numClienNOExentos][numTx][numMeses];
        double[][] ItCliNOExen = new double[numClienNOExentos][numMeses];
        double[][] TotMesTxCliNOExen = new double[numTx][numMeses];
        String Barra[];

                for (int i=0; i<numClienNOExentos;i++){
                    Barra=nombreCliNOExen[i].split("#");
                    int l2 = Calc.Buscar(Barra[2], nomBar);
                    for (int t=0; t<numTx; t++) {
                        for (int m=0; m<numMeses; m++) {
                            
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
         * Pagos por Empresa Clientes No Excentos
         *****************************/
        double[][][] pjeEmpSinAjuTx = new double[numEmp][numTx][numMeses];// el nombre que posee en peajes ret es: pjeEmpCNoExTx
        double[][] pjeEmpSinAju = new double[numEmp][numMeses];
        double[][][] ItEmpSinAjuTx = new double[numEmp][numTx][numMeses];
        double[][] ItEmpSinAju = new double[numEmp][numMeses];
        double[][][] pjeEmpSinAjuTxRE2288 = new double[numEmp][numTx][numMeses];// el nombre que posee en peajes ret es: pjeEmpCNoExTx
        double[][] pjeEmpSinAjuRE2288 = new double[numEmp][numMeses];
        double[][][] ItEmpSinAjuTxRE2288 = new double[numEmp][numTx][numMeses];
        double[][] ItEmpSinAjuRE2288 = new double[numEmp][numMeses];

          /*****************************
         * Calcula pago por empresa y agrega pago de Distribuidoras a los Suministradores
         *****************************/

        pjeEmpDxTx = new double[numDx][numSumi][numTx][numMeses];// el nombre que posee en peajes ret es: pjeEmpCNoExTx

        for (int j = 0; j < numClienNOExentos; j++) {
            String[] tmp = nombreCliNOExen[j].split("#");//Busca la Empresa del Cliente
            int l = Calc.Buscar(tmp[1], nomEmp);
            if(l!=-1){
                 for (int t=0; t<numTx; t++) {
                        for (int m=0; m<numMeses; m++) {
                            pjeEmpSinAjuTx[l][t][m]+=pjeCliTxNOExen[j][t][m];
                            pjeEmpSinAju[l][m]+=pjeCliTxNOExen[j][t][m];
                            ItEmpSinAjuTx[l][t][m]+=ItCliTxNOExen[j][t][m];
                            ItEmpSinAju[l][m]+=ItCliTxNOExen[j][t][m];
                            
                        }
                 }
            }
            else{//si no est‡ como empresa de generaci—n Busca si es una Distribuidora
                int l1 = Calc.Buscar(tmp[1], nomDx);
                 if(l1!=-1){
                    for(int i=0;i<numSumi;i++){
                        int l2= Calc.Buscar(nomSumi[i], nomEmp);//asigna la distribuidara a los suministradores
                        if(l2!=-1){
                            for (int t=0; t<numTx; t++) {
                                if (nomSumi[i].equals("RE2288")){
                                    for (int m=0; m<numMeses; m++) {
                                        for(int s=0;s<sumRM88;s++){
                                            pjeEmpSinAjuTxRE2288[s][t][m]+= pjeCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            pjeEmpSinAjuRE2288[s][m]+=pjeCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            ItEmpSinAjuTxRE2288[s][t][m]+= ItCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                            ItEmpSinAjuRE2288[s][m]+=ItCliTxNOExen[j][t][m]*facDx[l1][i][m]*proEfirme[s][m];
                                        }
                                    }    
                                }
                                else {
                                    for (int m=0; m<numMeses; m++) {
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
                    System.out.println("El Suministrador "+tmp[1]+" en 'clientes'"+" "+j+" no est‡ asignado como Distribuidora o como empresa de Generaci—n");
                }
            }
        }
        
        
        //Calcula prorrata para cuadro IT 
        double[][][] ProrrataRetCenLin1 = new double[numLinTx][numEmp][numMeses];
        double[][][] ProrrataRetCenLin2 = new double[numLinTx][numEmp][numMeses];
        double[][][] ProrrataRetCenLin3 = new double[numLinTx][numEmp][numMeses];

        for (int j = 0; j < numClienNOExentos; j++) {
            int c;
            String[] tmp = nombreCliNOExen[j].split("#");//Busca la Empresa del Cliente
            int l = Calc.Buscar(tmp[1], nomEmp);
            if(l!=-1){
                 for (int t=0; t<numLinTx; t++) {
                        for (int m=0; m<numMeses; m++) {
                           c = l;
                           //ProrrataRetCenLin[numLinTx][numEmp][numMeses];
                           //prorrMesC[numLinTx][numCli][numMeses]
                           ProrrataRetCenLin1[t][c][m] += prorrMesC[t][j][m];
                        }
                 }
            }
            else{//si no est‡ como empresa de generaci—n Busca si es una Distribuidora
                int l1 = Calc.Buscar(tmp[1], nomDx);
                 if(l1!=-1){
                    for(int i=0;i<numSumi;i++){
                        int l2= Calc.Buscar(nomSumi[i], nomEmp);//asigna la distribuidara a los suministradores
                        if(l2!=-1){
                            for (int t=0; t<numLinTx; t++) {
                                if (nomSumi[i].equals("RE2288")){
                                    for (int m=0; m<numMeses; m++) {
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
                                    for (int m=0; m<numMeses; m++) {
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
                    System.out.println("El Suministrador "+tmp[1]+" en 'clientes'"+" "+j+" no est‡ asignado como Distribuidora o como empresa de Generaci—n");
                }
            }
        }
 
        //imprime prorratas
        
        try
	{
            FileWriter writer = new FileWriter(DirBaseSal + slash +"prorratas_pago_ret.csv");
           
            
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
            
            for (int m=0; m<numMeses; m++) {
                for (int i = 0 ; i < numEmp; i++){
                    for (int t=0; t<numLinTx; t++) {
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
        }
        
        catch(IOException e)
	{
	     e.printStackTrace();
             //continue;
	} 
        
          /*****************************
         * Pagos Clientes exentos
         *****************************/
        double[][][] peajeClienTxExen = new double[numCli][numTx][numMeses];
        double[][] peajeClienExen = new double[numCli][numMeses];
        double[][] pjeAnualClienTxExen = new double[numCli][numTx];
        double[] pjeAnualClienExen = new double[numCli];
        double[][] TotMesTxClienExen = new double[numTx][numMeses];

         for (int i=0; i<numClienExentos;i++){
                    for (int m=0; m<numMeses; m++) {
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
       double[][] FactAjusClienExenCen = new double[numCen][numMeses];
       double[][][] AjusClienExenCenTx = new double[numCen][numTx][numMeses];
       for(int j=0;j<numCen;j++){
       for(int m=0;m<numMeses;m++){
         FactAjusClienExenCen[j][m]=GenPromMesCen[j][m]/GeneTotMesProm[m];
          for(int t=0;t<numTx;t++){
             AjusClienExenCenTx[j][t][m]=TotMesTxClienExen[t][m]*FactAjusClienExenCen[j][m];
          }
       }
       }
         /*****************************
         * Asigna Ajuste de Excentos a las empresas
         *****************************/
       double[][][] pagoEmpAjusTx= new double[numEmp][numTx][numMeses];
       double[][] pagoEmpAjus= new double[numEmp][numMeses];
            for (int j = 0; j < numCen; j++) {
            String[] tmp = nomCen[j].split("#");
            int l = Calc.Buscar(tmp[0], nomEmp);
                for (int m = 0; m < numMeses; m++) {
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
        

        //Pago Total (No exento + Ajustex Excento)
        double[][][] TotRetEmpTx = new double[numEmp][numTx][numMeses];
        double[][] TotRetEmp = new double[numEmp][numMeses];
        double[][][] TotItRetEmpTx = new double[numEmp][numTx][numMeses];
        double[][] TotItRetEmp = new double[numEmp][numMeses];
        double[][][] TotRetEmpTxRE2288 = new double[numEmp][numTx][numMeses];
        double[][] TotRetEmpRE2288 = new double[numEmp][numMeses];
        
        double[][][] TotItRetEmpTxRE2288= new double[numEmp][numTx][numMeses];
        double[][] TotItRetEmpRE2288= new double[numEmp][numMeses];
        double[][] TotAnualItRetEmpGTxRE2288= new double[numEmp][numTx];
        double[] TotAnualItRetEmpGRE2288= new double[numEmp];
        
         for (int j = 0; j < numEmp; j++) {
            for (int m = 0; m < numMeses; m++) {
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
            for (int m = 0; m < numMeses; m++) {
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
         // Ordena los archivos de salida para PU
        int[] nbne = Calc.OrdenarBurbujaStr(nomBarNoEx);
        String[] nomBarNoExO = new String[numBarCNoEx];
        for (int i=0; i < numBarCNoEx; i++){
            nomBarNoExO[i] = nomBar[nbne[i]];
        }


        // -------------------------------------------------------------------
        double[][][] prorrMesCO = new double[numLinTx][numCli][numMeses];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCli; j++) {
                for (int k = 0; k < numMeses; k++) {
                    prorrMesCO[i][j][k] = prorrMesC[i][nc[j]][k];
                }
            }
        }
        // -------------------------------------------------------------------
        double[][][] peajeLinCO = new double[numLinTx][numCli][numMeses];
        double[] SumMensualPjeLin = new double[numMeses];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCli; j++) {
                for (int k = 0; k < numMeses; k++) {
                    peajeLinCO[i][j][k] = peajeLinCon[i][nc[j]][k];
                    SumMensualPjeLin[k]+=peajeLinCO[i][j][k];
                }
            }
        }
        // -------------------------------------------------------------------
        peajeClienTxNOExenO = new double[numClienNOExentos][numTx][numMeses];
        for (int i = 0; i < numClienNOExentos; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    peajeClienTxNOExenO[i][j][k] = pjeCliTxNOExen[ncNOExen[i]][j][k];
                }
            }
        }
        // -------------------------------------------------------------------
        double[][] peajeClienNOExenO = new double[numClienNOExentos][numMeses];
        for (int i = 0; i < numClienNOExentos; i++) {
            for (int j = 0; j < numMeses; j++) {
                peajeClienNOExenO[i][j] = pjeCliNOExen[ncNOExen[i]][j];
            }
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
        double[][] peajeEmpCO = new double[numEmpC][numMeses];
        for (int i = 0; i < numEmpC; i++) {
            for (int j = 0; j < numMeses; j++) {
                peajeEmpCO[i][j] = pjeEmpSinAju[nc[i]][j];
            }
        }
          // ---------------------------------------------------------------------
        peajeClienTxExenO = new double[numClienExentos][numTx][numMeses];
        pjeAnualClienTxExenO = new double[numClienExentos][numTx];
        pjeAnualClienExenO= new double[numClienExentos];
        for (int i=0; i < numClienExentos; i++){
        for (int j=0; j < numTx; j++){
        for (int k=0; k < numMeses; k++){
            peajeClienTxExenO[i][j][k] = peajeClienTxExen[ncExen[i]][j][k];
            pjeAnualClienTxExenO[i][j]=pjeAnualClienTxExen[ncExen[i]][j];
            pjeAnualClienExenO[i]=pjeAnualClienExen[ncExen[i]];
        }
        }
        }
          // -------------------------------------------------------------------
        double[][] peajeClienExenO = new double[numClienExentos][numMeses];
        for (int i = 0; i < numClienExentos; i++) {
            for (int j = 0; j < numMeses; j++) {
                peajeClienExenO[i][j] = peajeClienExen[ncExen[i]][j];
            }
        }
          // ---------------------------------------------------------------------
       AjusClienExenCenTxO = new double[numCen][numTx][numMeses];
       GenAnoxCenO=new double[numCen];//ajuste
       GenPromMesCenO=new double[numCen][numMeses];
        for (int j=0; j < numCen; j++)
        for (int t=0; t < numTx; t++)
        for (int m=0; m < numMeses; m++){
           AjusClienExenCenTxO[j][t][m]= AjusClienExenCenTx[n[j]][t][m];
           GenAnoxCenO[j]=GenAnoxCen[n[j]];//ajuste
           GenPromMesCenO[j][m]=GenPromMesCen[n[j]][m];
        }
        // ---------------------------------------------------------------------
        AjusEmpCTxO = new double[numEmp][numTx][numMeses];
        AjusAnualEmpCTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    AjusEmpCTxO[i][j][k] = pagoEmpAjusTx[ne[i]][j][k];
                     AjusAnualEmpCTxO[i][j]+=AjusEmpCTxO[i][j][k];
                }
            }
        }
        // ---------------------------------------------------------------------
        AjusEmpCO = new double[numEmp][numMeses];
        AjusAnualEmpCO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numMeses; j++) {
                AjusEmpCO[i][j] = pagoEmpAjus[ne[i]][j];
                AjusAnualEmpCO[i]+= AjusEmpCO[i][j];
            }
        }
          // ---------------------------------------------------------------------
        TotRetEmpTxO = new double[numEmp][numTx][numMeses];
        TotAnualRetEmpTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotRetEmpTxO[i][j][k] = TotRetEmpTx[ne[i]][j][k];//Esto es lo que estar’a malo
                     TotAnualRetEmpTxO[i][j]+=TotRetEmpTxO[i][j][k];
                }
            }
        }
        
        
        ne2288 = Calc.OrdenarBurbujaStr(nomSumiRM88);
        nomSumiRM88O = new String[sumRM88];
        for (int i = 0; i < sumRM88; i++) {
            nomSumiRM88O[i] = nomSumiRM88[ne2288[i]];
        }
        
        
        TotRetEmpTxRE2288O = new double[sumRM88][numTx][numMeses];
        TotAnualRetEmpTxRE2288O = new double[sumRM88][numTx];
        double[][][] TotItRetEmpTxRE2288O = new double[sumRM88][numTx][numMeses];
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotRetEmpTxRE2288O[i][j][k] = TotRetEmpTxRE2288[ne2288[i]][j][k];
                    TotItRetEmpTxRE2288O[i][j][k] = TotItRetEmpTxRE2288[ne2288[i]][j][k];
                     TotAnualRetEmpTxRE2288O[i][j]+=TotRetEmpTxRE2288O[i][j][k];
                }
            }
        }
        
        
        
        
        TotRetItEmpTxO = new double[numEmp][numTx][numMeses];
        //TotAnualRetEmpTxO = new double[numEmp][numTx];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotRetItEmpTxO[i][j][k] = TotItRetEmpTx[ne[i]][j][k];
                     //TotAnualRetEmpTxO[i][j]+=TotRetEmpTxO[i][j][k];
                }
            }
        }
        
        
        
                  // ---------------------------------------------------------------------
       RetEmpSinAjuTxO = new double[numEmp][numTx][numMeses];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                     RetEmpSinAjuTxO[i][j][k] =  pjeEmpSinAjuTx[ne[i]][j][k];
                }
            }
        }
        // ---------------------------------------------------------------------
        TotRetEmpO = new double[numEmp][numMeses];
        TotItRetEmpO = new double[numEmp][numMeses];
        double[] TotMensualRetEmp=new double[numMeses];
        
        TotAnualRetEmpO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numMeses; j++) {
                TotRetEmpO[i][j] = TotRetEmp[ne[i]][j];
                TotItRetEmpO[i][j] = TotItRetEmp[ne[i]][j];
                TotAnualRetEmpO[i]+=TotRetEmp[ne[i]][j];
                TotMensualRetEmp[j]+=TotRetEmp[ne[i]][j];
            }
        }
        // ---------------------------------------------------------------------
        RetEmpSinAjuO = new double[numEmp][numMeses];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numMeses; j++) {
                RetEmpSinAjuO[i][j] = pjeEmpSinAju[ne[i]][j];
            }
        }
        // ---------------------------------------------------------------------
        double[][][] TotPjeRetEmpTxO = new double[numEmp][numTx][numMeses];
        TotAnualPjeRetEmpTxO = new double[numEmp][numTx];
        TotAnualPjeRetEmpO = new double[numEmp];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotAnualPjeRetEmpTxO[i][j]=TotAnualPjeRetEmpGTx[ne[i]][j]; //sumar
                    TotAnualPjeRetEmpO[i]=TotAnualPjeRetEmpG[ne[i]]; //sumar
                    TotPjeRetEmpTxO[i][j][k] = pjeEmpSinAjuTx[ne[i]][j][k];
                }
            }
        }
        
        TotAnualPjeRetEmpRE2288O = new double[numEmp];
        TotAnualPjeRetEmpTxRE2288O = new double[sumRM88][numTx];
        double[] TotMensualPjeRetEmpRE2288O = new double[numMeses];
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
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
                
        for (int j = 0; j < numMeses; j++) {
            TotMensualRetEmp[j]+=TotMensualPjeRetEmpRE2288O[j];
            
        }
        double[][][] TotRetEmpTxRE2288O = new double[sumRM88][numTx][numMeses];
        double[][] TotAnualPjeRetEmpGTxRE2288O = new double[numEmp][numTx];
        
        for (int i = 0; i < sumRM88; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotRetEmpTxRE2288O[i][j][k]=TotRetEmpTxRE2288[ne2288[i]][j][k];
                }
                TotAnualPjeRetEmpGTxRE2288O[i][j]=TotAnualPjeRetEmpGTxRE2288[ne2288[i]][j];
            }
        }
        
        
        TotRetEmpTxRE2288OO = new double[numEmp][numTx][numMeses];
        TotRetEmpRE2288O=new double[numEmp][numMeses];
        for (int i = 0; i < numEmp; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < numMeses; k++) {
                    TotRetEmpTxO[i][j][k]= re2288toNomsumi[i] == -1? TotRetEmpTxO[i][j][k]:TotRetEmpTxO[i][j][k]+TotRetEmpTxRE2288O[re2288toNomsumi[i]][j][k];//Esto es lo que estar’a malo
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
      PUO = new double[numBarC][numTx][numMeses];
        for (int j=0; j < numBarC; j++)
        for (int t=0; t < numTx; t++)
        for (int m=0; m < numMeses; m++){
           PUO[j][t][m]= PU[nb[j]][t][m];
          }
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
       long tInicioEscritura = System.currentTimeMillis();

        /*
         * Escritura de Resultados
         * =======================
         */
        String libroSalidaCXLS = DirBaseSal + slash +
                "PagoRet" + Ano + ".xlsx";
        Escribe.crearLibro(libroSalidaCXLS);
        Escribe.creaH2F_3d2_long(
                "Pago de Peaje por Línea y Cliente [$]", peajeLinCO,
                "Línea", nomLineasN,
                "Cliente", nomCliO,
                "Mes", nomMes,
                libroSalidaCXLS, "PjeClienLin",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH2F_3d2_long(
                "Pago Peaje por Cliente y Transmisor (Clientes No Exentos) [$]", peajeClienTxNOExenO,
                "Cliente", nombreCliNOExenO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "PjeClienTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        /*Escribe.creaH1F_2d_long(
                "Pago Peaje por Cliente (Clientes No Exentos) [$]", peajeClienNOExenO,
                "Central", nombreCliNOExenO,
                "Mes", nomMes,
                libroSalidaCXLS, "PjexCliente",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
          Escribe.creaH2F_3d_long(
                "Pago Peaje por Empresa y Transmisor (Clientes No Exentos)[$]", peajeEmpCTxO,
                "Empresa", nomEmpCO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "PjeEmpTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Pago Peaje por Empresa (Clientes No Exentos) [$]", peajeEmpCO,
                "Empresa", nomEmpCO,
                "Mes", nomMes,
                libroSalidaCXLS, "PjeEmp",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
       *
       */
        if(numClienExentos!=0){
         Escribe.creaH2F_3d2_long(
                "Pago Peaje de Cliente Exento y Transmisor[$]", peajeClienTxExenO,
                "Cliente", nombreClientesExenO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS,"PjeClienTxExen","#,###,##0;[Red]-#,###,##0;\"-\"");
        }
          /*Escribe.creaH1F_2d_long(
                "Pago Peaje por Cliente Exento [$]", peajeClienExenO,
                "Cliente",nombreClientesExenO,
                "Mes", nomMes,
                libroSalidaCXLS, "PjexClienteExen",
                "#,###,##0;[Red]-#,###,##0;\"-\"");*/
          Escribe.creaH3F_3d_double(
                "Pago por Ajuste de Retiros Exentos por Central y Transmisor [$]", AjusClienExenCenTxO,
                "Central", nomCenO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                "Inyeccion Anual",GenAnoxCenO,
                libroSalidaCXLS,"AjusExenTx","#,###,##0;[Red]-#,###,##0;\"-\"");
          /*Escribe.creaH2F_3d_double(
                "Pago por Ajuste por Empresa y Transmisor [$]", AjusEmpCTxO,
                "Empresa", nomEmpO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "AjusEmpTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");*/
         Escribe.creaH1F_2d_double(
                "Ajuste por Empresa [$]", AjusEmpCO,
                "Empresa", nomEmpO,
                "Mes", nomMes,
                libroSalidaCXLS, "AjusEmp",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
          /* Escribe.creaH2F_3d_double(
                "Pago por Empresa y Transmisor por RM88 [$]", AjusRM88TxO,
                "Empresa", nomEmpO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "AjusRM88Tx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
          Escribe.creaH1F_2d_double(
                "Pago por Empresa por RM88 [$]", AjusRM88O,
                "Empresa", nomEmpO,
                "Mes", nomMes,
                libroSalidaCXLS, "AjusRM88",
                "#,###,##0;[Red]-#,###,##0;\"-\"");*/
          Escribe.creaH3F_4d_double(
                "Pago por Contratos con Distribuidoras [$]", pjeEmpDxTx,
                "Sumnistrador", nomSumi,
                "Distrubuidora", nomDx,
                "Mes", nomMes,
                "Transmisor",nombreTx,
                libroSalidaCXLS,"PagosDx","#,###,##0;[Red]-#,###,##0;\"-\"");
          
          Escribe.creaH2F_3d_double(
                "Pagos de Peaje de Retiro RE2288 por Empresa y Transmisor [$] (Incluye ajuste por Excentos)", TotRetEmpTxRE2288O,
                "Empresa", nomSumiRM88O,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "PagosRE2288",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
          
         // Escribe.CopiaHoja(DirBaseEnt + slash + "Ent" + Ano + ".xlsx",libroSalidaCXLS, "Distribuidoras");
          Escribe.creaH2F_3d_double(
                "Pagos de Peaje de Retiro por Empresa y Transmisor [$] (Incluye ajuste por Excentos)", TotRetEmpTxO,
                "Empresa", nomEmpO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "PagoEmpTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
         /*Escribe.creaH1F_2d_double(
                "Pago Total de Peajes de Retiro por Empresa (Peaje Clientes No Exentos + Ajuste + RM88)[$]", TotRetEmpO,
                "Empresa", nomEmpO,
                "Mes", nomMes,
                libroSalidaCXLS, "TotRetEmp",
                "#,###,##0;[Red]-#,###,##0;\"-\"");*/
          Escribe.crea_SalidaCU(
                "Cargo Unitario [$/MWh]",
                "Barra", nomBarO,
                "Transmisor", nombreTx,
                "Consumo","Consumo CU2","Consumo CU30", ECUbarraO,
                "Prorrata","Prorrata CU2","Prorrata CU30", ProrrCUO,
                "Pagos", "Pago CU2","Pago CU30",PagoCUO,
                PagoAnoBarraO,
                libroSalidaCXLS, "CargoUnitario",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
          Escribe.creaH2F_3d_double(
                "Pago Unitario [$/MWh]", PUO,
                "Barra", nomBarO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaCXLS, "PUnit",
                "#,###,##0.###;[Red]-#,###,##0.###;\"-\"");
          Escribe.crea_verificaRet(
                  "Verifica Pagos de Retiro",libroEntrada,
                  "CUE","CUE2","CUE30","Pago","Consumo",
                  PagoCUAnual,ECUAnual,
                  "Mes",nomMes,
                  "Calculo", TotMensualRetEmp,
                  "Prorrata L’nea",SumMensualPjeLin,
                  "Diferencia",
                  "verifica","#,###,##0;[Red]-#,###,##0;\"-\"");
          Escribe.crea_verificaCalcPeajes(
                  "Verifica Cálculo de Peajes",libroEntrada,
                  "Mes",nomMes,
                  "Peajes", PeajeNMes,
                  "Pago Ret","Pago Iny","Diferencia",
                  "verifica","#,###,##0;[Red]-#,###,##0;\"-\"");
          long tFinalEscritura = System.currentTimeMillis();
          System.out.println("Pagos de Retiro Anual Calculados");
         System.out.println("Tiempo Adquisicion de datos     : "+DosDecimales.format((tFinalLectura-tInicioLectura)/1000.0)+" s");
         System.out.println("Tiempo Cálculo                  : "+DosDecimales.format((tFinalCalculo-tInicioCalculo)/1000.0)+" s");
         System.out.println("Tiempo Escritura de Resultados  : "+DosDecimales.format((tFinalEscritura-tInicioEscritura)/1000.0)+" s");
         System.out.println();
    }

     public static void LiquiMesRet(String mes, int Ano) {
          int m=0;
          for(int i=0;i<numMeses;i++){
              if(mes.equals(nomMes[i]))
                  m=i;
          }
     String libroSalidaGXLSMes= DirBaseSal + slash +"PagoRet" + nomMes[m] + ".xlsx";
     Escribe.crearLibro(libroSalidaGXLSMes);

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
                libroSalidaGXLSMes, nomMes[m],Ano,
                "#,###,##0;[Red]-#,###,##0;\"-\"");

     Escribe.creaProrrataMes(m,
                "Participación de Retiros [%]",prorrMesC,"Participación "+nomMes[m],
                "Cliente",nomCli,
                "L’nea",  nomLinTx,
                "AIC", zonaLinTx,
                libroSalidaGXLSMes, "PartRet"+nomMes[m],
                "#,###,##0;[Red]-#,###,##0;\"-\"");
      Escribe.creaTabla1C_long(m,
                "Pago de Peaje por Clientes "+nomMes[m]+" [$]",peajeClienTxNOExenO,
                "Cliente",nombreCliNOExenO,
                "Transmisor",  nombreTx,
                libroSalidaGXLSMes, "Pagos",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
      Escribe.creaTabla2CDx_double(m,
                "Pago de Peaje por Contratos con Distribuidoras"+nomMes[m]+" [$]",pjeEmpDxTx,
                "Suministrador",nomSumi,
                "Transmisor",  nombreTx,
                "Distribuidora",nomDx,
                facDx,
                libroSalidaGXLSMes, "PagosDx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
      if(numClienExentos!=0){
      Escribe.creaTabla1C_long(m,
                "Pago Peaje Exento "+nomMes[m]+" [$]", peajeClienTxExenO,
                "Cliente", nombreClientesExenO,
                "Transmisor", nombreTx,
                libroSalidaGXLSMes, "PagosExentos",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaTabla2C_double(m,
                "Ajustes de Pagos correspondientes a "+nomMes[m] +" por Central [$]", AjusClienExenCenTxO,
                "Central", nomCenO,
                "Transmisor", nombreTx,
                "Inyeccion Mes",GenPromMesCenO,
                libroSalidaGXLSMes,"Ajuste"+nomMes[m],"#,###,##0;[Red]-#,###,##0;\"-\"");
      }
        Escribe.creaTabla1C_float(m,
                "Peajes Unitarios "+nomMes[m]+" [$/MWh]",PUO,
                "Barra",nomBarO,
                "Transmisor",  nombreTx,
                libroSalidaGXLSMes, "PeajesUnitarios",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaTabla1C_float(m,
                "Peajes RE2288 "+nomMes[m]+" [$/MWh]",TotRetEmpTxRE2288O,
                "Suministrador",nomSumiRM88O,
                "Transmisor",  nombreTx,
                libroSalidaGXLSMes, "PeajesRE2288",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        
        
        
         System.out.println("Peajes de Retiros Calculados");
         System.out.println();
      }





}
