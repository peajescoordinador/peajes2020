package cl.coordinador.peajes;
/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
import static cl.coordinador.peajes.PeajesConstant.MAX_COMPRESSION_RATIO;
import static cl.coordinador.peajes.PeajesConstant.MESES;
import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.SLASH;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DecimalFormat;

/**
 *
 * @author vtoro
 */
public class PeajesIny {

    static double peajeEmpTxO[][][]; //TODO: Encapsulate
    private static double prorrMesGO[][][];
    private static double peajeLinGO[][][];
    private static double ItLinGO[][][];
    static double peajeCenTxO[][][]; //TODO: Encapsulate
    private static double peajeCenO[][];
    static double peajeEmpGO[][]; //TODO: Encapsulate
    static float MGNCO[]; //TODO: Encapsulate
    static double[] PotNetaO; //TODO: Encapsulate
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
    private static int numCen;
    static String[] nombreTx; //TODO: Encapsulate
    private static int numTx;
    static String[] nomEmpGO; //TODO: Encapsulate
    private static int numEmpG;
    static String[] nomLinTx; //TODO: Encapsulate
    private static int numLinTx;
    static String[] nomGenO; //TODO: Encapsulate
    static String[] nomMGNCO; //TODO: Encapsulate
    private static String libroSalidaGXLSMes;
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

    public static void calculaPeajesIny(File DirEntrada, File DirSalida,
            int Ano, boolean LiquidacionReliquidacion) {

        String DirBaseEnt = DirEntrada.toString();
        DirBaseSal = DirSalida.toString();
        DecimalFormat DosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();
        cargandoInfo=true;
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);
        String libroEntrada = DirBaseSal + SLASH + "Peaje" + Ano + ".xlsx";

        /************
         * lee Peajes e IT
         ************/
        double[][] longAux = new double[1000][NUMERO_MESES];
        double[][] longAuxIT = new double[1000][NUMERO_MESES];
        double[][] longAuxITR = new double[1000][NUMERO_MESES];
        double[][] longAuxVATT = new double[1000][NUMERO_MESES];
        double[][] longAuxITP = new double[1000][NUMERO_MESES];
        String[] TxtTemp = new String[1000];
        String[] TxtTempIT = new String[1000];
        int numLinea = Lee.leePeajes(libroEntrada, TxtTemp, longAux);
        int numLineaIT = Lee.leeIT(libroEntrada, TxtTempIT, longAuxIT,"ITEG");
        int numLineaITR = Lee.leeIT(libroEntrada, TxtTempIT, longAuxITR,"ITER");
        int numLineaVATT = Lee.leeIT(libroEntrada, TxtTempIT, longAuxVATT,"VATT");
        int numLineaITP = Lee.leeIT(libroEntrada, TxtTempIT, longAuxITP,"ITP");
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
        libroEntrada = DirBaseEnt + SLASH + "Ent" + Ano + ".xlsx";

        /**********
         * lee VATT
         **********/
        double[][] Aux = new double[1500][NUMERO_MESES];
        String[] TxtTemp1 = new String[1500];
        String[] TxtTemp2 = new String[1500];
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
        for (int i = 0; i < numTx; i++) {
            nombreTx[i] = TxtTemp3[i];
        }
        /*
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
        TxtTemp1 = new String[600];
        double PotNetaTot=0;
        float[] Temp1 = new float[600];
        float[] Temp2= new float[600];
        numCen = Lee.leeCentrales(libroEntrada, TxtTemp1,Temp1,Temp2);
        nomGen = new String[numCen];
        double[] PotNeta = new double[numCen];
        float[] MGNC = new float[numCen];
        int numMGNC=0;
        int[] indMGNC=new int[numCen];

        for (int i = 0; i < numCen; i++) {
            nomGen[i] = TxtTemp1[i];
            PotNeta[i] = Temp1[i];
            PotNetaTot+=PotNeta[i];
            MGNC[i] = Temp2[i];
            //if(MGNC[i]==1){
                indMGNC[numMGNC]=i;
                numMGNC++;
            //}
        }


        TxtTemp1 = new String[numCen];
        for (int i = 0; i < numCen; i++) {
            TxtTemp1[i] = "";
        }
       numEmpG = 0;
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
        int[][] intAux2 = new int[600][NUMERO_MESES];
        int numLinIT = Lee.leeLintron(libroEntrada, TxtTemp1,
                nomLin,TxtTemp2, intAux1, intAux2);
        String TxtTemp4[]=new String[numLinIT];
        nomLinIT = new String[numLinIT];
        zonaLinIT= new int[numLinIT];        
        int[] indZonaLinIT=new int[numLinIT];
        String[] nomZonaLinIT=new String[numLinIT];
        String[] propietario = new String[numLinIT];
        for (int i = 0; i < numLinIT; i++) {
            TxtTemp4[i]="";
            nomLinIT[i] = TxtTemp1[i];
            zonaLinIT[i]=intAux2[i][0];
            propietario[i]=TxtTemp2[i];
            if(zonaLinIT[i]==1){
            indZonaLinIT[i]=0;
            nomZonaLinIT[i]="N";
            }
            else if(zonaLinIT[i]==0){
            indZonaLinIT[i]=1;
            nomZonaLinIT[i]="A";
            }
            else if(zonaLinIT[i]==-1){
            indZonaLinIT[i]=2;
            nomZonaLinIT[i]="S";
            }
        }
        int[] indZonaLinPe=new int[numLinea];
        zonaLinPe= new int[numLinea];
        for (int i = 0; i < numLinea; i++) {
         String[] tmp = nomLineasN[i].split("#");
            int l = Calc.Buscar(tmp[0], nomLinIT);
            if(l==-1){
             System.out.println("Error!!!");
             System.out.println("L’nea Trocal - "+tmp[0]+" - en archivo Peaje"+Ano+".xls no se encuentra en la hoja 'lintron' del archivo Ent"+Ano+".xlsx");
             System.out.println("Debe asegurarse que la L’neas del archivo AVI_COMA.xls se encuentren en la hoja 'lintron' y ejecutar el bot—n Peajes");
            }
            else{
            zonaLinPe[i]=zonaLinIT[l];
            indZonaLinPe[i]=indZonaLinIT[l];
            }
        }

         TxtTemp3=new String[numLinIT];


        numLinTx = 0;
        for (int i = 0; i <numLinIT; i++) {
            int l = Calc.Buscar(nomLinIT[i] + "#" + propietario[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinIT[i] + "#" + propietario[i];
                TxtTemp2[numLinTx]=propietario[i];
                numLinTx++;
            }
        }
        nomLinTx = new String[numLinTx];//solo registros Ïnico L’nea#Transmisor de hoja lintron
        String[] nomPropTx = new String[numLinTx];

        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
            //System.out.println(nomLinTx[i]);
            nomPropTx[i]=TxtTemp2[i];
        }
        int[] indZonaLinTx=new int[numLinTx];
        zonaLinTx= new int[numLinTx];
         for (int i = 0; i < numLinTx; i++) {
         String[] tmp = nomLinTx[i].split("#");
            int l2 = Calc.Buscar(tmp[0], nomLinIT);
            zonaLinTx[i]=zonaLinIT[l2];
            indZonaLinTx[i]=indZonaLinIT[l2];
         }

        // Libro Prorrata
        String libroEntradaP = DirBaseSal + SLASH + "Prorrata" + Ano + ".xlsx";

        /*****************************
         * lee Prorratas de Generaci—n
         *****************************/
        prorrMesGenTx = new double[numLinTx][numCen][NUMERO_MESES];
        prorrMesGenTxTot = new double[numLinTx][NUMERO_MESES];
        Lee.leeProrratasGx(libroEntradaP, prorrMesGenTx);
        for (int l = 0 ; l < numLinTx; l++){
            for (int m = 0 ; m < NUMERO_MESES; m ++ ){
                for (int c = 0 ; c < numCen; c ++ ){
                prorrMesGenTxTot [l][m]+=prorrMesGenTx[l][c][m];
                }
            }
        }
        
        
        /***************************/
        /**Lee Inyeccion Centrales**/
        /***************************/
        double[][] GenerMensual = new double[numCen][NUMERO_MESES];
        double[] GeneTotMesProm= new double[NUMERO_MESES];
        double[][] GenPromMesCen= new double[numCen][NUMERO_MESES];
        double[] GenAnoxCen= new double[numCen];
        int [][] MesesAct=new int[numCen][NUMERO_MESES];
        int [] numMesesAct=new int[numCen];
        Lee.leeGeneracionMes(libroEntradaP,GenerMensual);

        /*LEE MESES CENTRALES ACTIVO*/
        for (int i=0;i<numCen;i++){
            for(int m=0; m<NUMERO_MESES;m++){
                MesesAct[i][m]+=1;  //AGREGAR RUTINA Q LEE MANT CEN 
            }
        }
        
        
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
        
        
        
        long tFinalLectura = System.currentTimeMillis();
        long tInicioCalculo = System.currentTimeMillis();
        cargandoInfo=false;
        calcPagos=true;

        /******************************************
         * Calcula Pagos por Inyecci—n de Centrales
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
         * Calcula Excenci—n de Centrales
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
        
        
        try
	{
            FileWriter writer = new FileWriter(DirBaseSal + SLASH +"prorratas_pago_iny.csv");
           
            
            writer.append("Central");
            writer.append(',');
            writer.append("Linea");
            writer.append(',');
            writer.append("Mes");
            writer.append(',');
            writer.append("Prorrata");
            writer.append('\n');
            
            for (int m=0; m<NUMERO_MESES; m++) {
                for (int i = 0 ; i < numEmpG; i++){
                    for (int t=0; t<numLinTx; t++) {
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
        }
        
        catch(IOException e)
	{
	     e.printStackTrace();
             //continue;
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



        // Ordena los archivos de salida de Inyecci—n por empresas
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
                for (int k = 0; k < NUMERO_MESES; k++) {
                    prorrMesGO[i][j][k] = prorrMesGenTx[i][ng[j]][k];
                }
            }
        }
        // -------------------------------------------------------------------
        peajeLinGO = new double[numLinTx][numCen][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCen; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    peajeLinGO[i][j][k] = peajeLinCen[i][ng[j]][k];
                }
            }
        }
        // -------------------------------------------------------------------
        ItLinGO = new double[numLinTx][numCen][NUMERO_MESES];
        for (int i = 0; i < numLinTx; i++) {
            for (int j = 0; j < numCen; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    ItLinGO[i][j][k] = ItLinCen[i][ng[j]][k];
                }
            }
        }// -------------------------------------------------------------------
        peajeCenTxO = new double[numCen][numTx][NUMERO_MESES];
        for (int i = 0; i < numCen; i++) {
            for (int j = 0; j < numTx; j++) {
                for (int k = 0; k < NUMERO_MESES; k++) {
                    peajeCenTxO[i][j][k] = peajeGenTx[ng[i]][j][k];
                }
            }
        }
        // -------------------------------------------------------------------
        peajeCenO = new double[numCen][NUMERO_MESES];
        for (int i = 0; i < numCen; i++) {
            for (int j = 0; j < NUMERO_MESES; j++) {
                peajeCenO[i][j] = peajeGen[ng[i]][j];
            }
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
                for (int k = 0; k < NUMERO_MESES; k++) {
                    peajeEmpTxO[i][j][k] = peajeEmpGTx[ne[i]][j][k];
                }
            }
        }
        // -------------------------------------------------------------------
        peajeEmpGO = new double[numEmpG][NUMERO_MESES];
        for (int i = 0; i < numEmpG; i++) {
            for (int j = 0; j < NUMERO_MESES; j++) {
                peajeEmpGO[i][j] = peajeEmpG[ne[i]][j];
            }
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
        long tInicioEscritura = System.currentTimeMillis();


        /*
         * Escritura de Resultados
         * =======================
         */
        String libroSalidaGXLS = DirBaseSal + SLASH + "PagoIny" + Ano + ".xlsx";
        Escribe.crearLibro(libroSalidaGXLS);
        Escribe.creaH2F_3d_long(
                "Pago de Peaje por Línea y Central [$]", peajeLinGO,
                "Línea", nomLineasN,
                "Central", nomGenO,
                
                
                "Factor de Excención",MGNCO,
                
                
                "Mes", MESES,
                libroSalidaGXLS, "PjeCenLin",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        /*Escribe.creaH2F_3d_long(
                "Pago Peaje por Central y Transmisor [$]", peajeCenTxO,
                "Central", nomGenO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaGXLS, "PjeCenTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Pago Peaje por Central [$]", peajeCenO,
                "Central", nomGenO,
                "Mes", nomMes,
                libroSalidaGXLS, "PjexCen",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH2F_3d_long(
                "Pago Peaje por Empresa y Transmisor [$]", peajeEmpTxO,
                "Empresa", nomEmpGO,
                "Transmisor", nombreTx,
                "Mes", nomMes,
                libroSalidaGXLS, "PjeEmpTx",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.creaH1F_2d_long(
                "Pago Peaje por Empresa [$]", peajeEmpGO,
                "Empresa", nomEmpGO,
                "Mes", nomMes,
                libroSalidaGXLS, "PjeEmp",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
         * 
         */

        for(int m=0;m<NUMERO_MESES;m++){
            Escribe.creaPIny(m,
                "Pago Peaje por Empresa y Transmisor [$]",peajeEmpTxO,
                 AjusMGNCEmpTxO,PagoEmpTxO,
                 peajeEmpGO,AjusMGNCEmpO,PagoEmpO,
                "Empresa",nomEmpGO,
                "Transmisor", nombreTx,
                libroSalidaGXLS, MESES[m],
                "#,###,##0;[Red]-#,###,##0;\"-\"");

            Escribe.creaDetallePIny(m,
                "Detalle de Pagos por Central [$]",peajeGenTxZonaO,peajeCenTxO,ExcenCenO,
                 AjusMGNCTxO,PagoTotCenTxO,
                 peajeGenO, ExcTotCenO,AjusMGNCTotO,PagoTotCenO,
                 CapConjExcep,FCorrec,
                "Central",nomGenO,
                "Transmisor", nombreTx,
                "Factor Excención", MGNCO,
                //"PNeta", PotNetaO,
                //"Inyecci—n Mensual", GenPromMesCenO,
                //"Factor",facPagoO ,
                libroSalidaGXLS, MESES[m],
                "#,###,##0;[Red]-#,###,##0;\"-\"");
        }
        Escribe.crea_verificaIny(
                  "Verifica Pagos de Inyecci—n",libroEntrada,
                  "Mes",MESES,
                  "Calculo", PagoInyMes,
                  "Prorrata L’nea",pagoInyMesLin,
                  "Diferencia",
                  "verifica","#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_verificaCalcPeajes(
                  "Verifica c‡lculo de Peajes",libroEntrada,
                  "Mes",MESES,
                  "Peajes", PeajeNMes,
                  "Pago Ret","Pago Iny","Diferencia",
                  "verifica","#,###,##0;[Red]-#,###,##0;\"-\"");
        EscribirPagos=true;
        long tFinalEscritura = System.currentTimeMillis();
        System.out.println("Pagos de Inyecci—n Anual Calculados");
        System.out.println("Tiempo Adquisicion de datos     : "+DosDecimales.format((tFinalLectura-tInicioLectura)/(1000.0*60))+" m");
        System.out.println("Tiempo C‡lculo                  : "+DosDecimales.format((tFinalCalculo-tInicioCalculo)/(1000.0*60))+" m");
        System.out.println("Tiempo Escritura de Resultados  : "+DosDecimales.format((tFinalEscritura-tInicioEscritura)/(1000.0*60))+" m");
        System.out.println();
    }

      public static void LiquiMesIny(String mes, int Ano) {
          int m=0;
          for(int i=0;i<NUMERO_MESES;i++){
              if(mes.equals(MESES[i]))
                  m=i;
          }
     libroSalidaGXLSMes= DirBaseSal + SLASH +"PagoIny" + MESES[m] + ".xlsx";
     Escribe.crearLibro(libroSalidaGXLSMes);
     Escribe.creaLiquidacionMesIny(m,
                "Pago de Peajes de Inyecci—n",peajeCenTxO,
                 AjusMGNCTxO,PagoTotCenTxO,
                 peajeGenO,AjusMGNCTotO,PagoTotCenO,
                "Central",nomGenO,
                "Transmisor", nombreTx,
                
                "MGNC", MGNCO,
                "PNeta", PotNetaO,
                
                "Inyeccion [GWh]",GenPromMesCenO,
                
                "Factor",facPagoO ,
                
                libroSalidaGXLSMes, MESES[m],Ano,"#,###,##0;[Red]-#,###,##0;\"-\"");
     Escribe.creaProrrataMes(m,
                "Participaci—n de Inyecciones [%]",prorrMesGenTx,"Participaci—n "+MESES[m],
                "Cliente",nomGen,
                "L’nea",  nomLinTx,
                "AIC", zonaLinTx,
                libroSalidaGXLSMes, "PartIny"+MESES[m],
                "#,###,##0;[Red]-#,###,##0;\"-\"");
    Escribe.creaProrrataMes_long(m,
                "Pagos por Inyecci—n "+MESES[m]+" [$]",
                peajeLinGO,
                "Pago "+MESES[m],
                "Central",
                nomGenO,
                "L’nea",
                nomLineasN,
                "AIC",
                zonaLinPe,
                "Pago IT",
                ItLinGO,
                libroSalidaGXLSMes,
                "PagoxLinea",
                "#,###,##0;[Red]-#,###,##0;\"-\"");
    System.out.println("Archivo Pago de Inyecci—n Mensual creado");
    System.out.println();

           }

       public static boolean cargando(){
        return cargandoInfo;
       }
       public static boolean calculando(){
        return calcPagos;
       }
       public static boolean escribiendo(){
        return calcPagos;
       }
        public static void Comenzar(final File DirIn, final File DirOut, final int AnoAEvaluar, final boolean LiquidacionReliquidacion){
        final SwingWorker worker = new SwingWorker() {
            @Override
            public Object construct() {
                    calculaPeajesIny(DirIn, DirOut, AnoAEvaluar,LiquidacionReliquidacion);
                return true;
            }
        };
    }



}
