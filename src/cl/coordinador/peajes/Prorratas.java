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
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.FloatBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.util.Scanner;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Modela y asigna perdidas a consumos
 *
 * @author
 */
public class Prorratas {

    private static final boolean USE_BUFFEREDSTREAM = true;
    private static final boolean USE_FILECHANNEL = false;
    private static final boolean USE_SCANNER = false;
    private static final boolean USE_MAPPED_NAMES = true;
    
    private static int etapa;
    private static int numEtapas=0;
    private static String nombreSlack;
    private static boolean cargandoInfo=false;
    private static boolean calculandoFlujos=false;
    private static boolean calculandoProrr=false;
    private static boolean guardandoDatos=false;
    private static boolean completo=false;

    private static final boolean USE_FACTORY = false; //Temp switch for the thread factory
    
    //Arreglos comunes para cada etapa (para ejecucion en paralelo):
    private static float [][][] Gx; //Generacion de PLP
    private static boolean[][] barrasActivas; //Barras activas luego de revisar mantenimientos de lineas
    private static int[][] paramGener; //datos de inyeccion en planilla Ent
    private static float[][] Consumos; //datos de consumo en planilla Ent
    private static float[][] FallaEtaHid; //energia no suministrada de etapa
    private static float[][] perdidasPLPMayor110; //perdidas de etapa en PLP
    private static float[][][] Flujo; //Flujo DC mediante GLDF por linea, etapa, hidrologia
    private static float[][][] prorrGx; //Arreglo de salida donde se escribiran las prorratas de generacion
    private static float[][][] prorrCx; //Arreglo de salida donde se escribiran las prorratas de consumos
    private static int[][] orientBarTroncal; //datos de orientacion lineas troncal en planilla Ent
    private static int[][] paramBarTroncal; //datos de barras troncal en planilla Ent
    private static float[][] ConsumosClaves; //consumos reales en planilla Ent
    private static int[][] datosClaves;  //datos de consumos en planilla Ent
    
    public static void CalculaProrratas(File DirEntrada, File DirSalida, int AnoAEvaluar, int tipoCalc, int AnoBase,
            int NumeroHidrologias ,int NumeroEtapasAno, int NumeroSlack,int ValorOffset,boolean ActClientes) throws IOException, FileNotFoundException {
        
        int numGen; //Numero de generadores en planilla centralesPLP (rango 'plpcnfce') de archivo Ent
        int numLin; //Numero de lineas de transmision en archivo Ent
        int numLinTron; //Numero de lineas de transmision troncal en archivo Ent
        int numBarras; //Numero de barras en archivo Ent
        int numHid=NumeroHidrologias; //AnoIni-1962 - Numero de hidrologias a considerar en calculo (definidas por usuario)
        final int offset=ValorOffset;//(AnoIni==2004?0:12);        
        String DirBaseEntrada=DirEntrada.toString();
        String DirBaseSalida=DirSalida.toString();
        final String ArchivoDespachoGeneradores= DirBaseEntrada + SLASH + "plpcen.csv";
        final String ArchivoPerdidasLineas = DirBaseEntrada + SLASH + "plplin.csv";
	//indices de etapas relevantes para escritura de resultados
        final int etapaPeriodoIni=NumeroEtapasAno*(AnoAEvaluar-AnoBase)+offset;//(tipoCalc==0?offset:144*(Ano-AnoIni)+offset);
        final int etapaPeriodoFin=NumeroEtapasAno*(AnoAEvaluar-AnoBase+1)+offset;//(tipoCalc==0?offset+144:144*(Ano-AnoIni+1)+offset);
        numEtapas=etapaPeriodoFin-etapaPeriodoIni;
        String [] TxtTemp1; //almacenamiento temporal de texto 1
        String[] TxtTemp2; //almacenamiento temporal de texto 2
        String[] TxtTemp3; //almacenamiento temporal de texto 3
        int[] IntTemp; //almacenamiento temporal de enteros
        DecimalFormat dosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();
        cargandoInfo=true;
        String rutaLibroEnt = DirBaseEntrada + SLASH + "Ent" + AnoAEvaluar + ".xlsx";
        String[] EnergiaCU={"CUE2","CUE30","EUnit"};
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);

        /**************
         * lee de Meses
         **************/
        System.out.println("Importando Informacion y Parametros");
        XSSFWorkbook wb_Ent = new XSSFWorkbook(new FileInputStream( rutaLibroEnt ));
        int[] intAux1=new int[600];
        int numSp = Lee.leeMeses(wb_Ent, intAux1, MESES);
        int[] paramEtapa = new int[numEtapas];
        System.arraycopy(intAux1, 0, paramEtapa, 0, numEtapas);

        /*
         * Lectura de parametros de lineas
         * ===============================
         */
        TxtTemp1=new String[2500];
        double[][] Aux=new double[2500][11];
        numLin = Lee.leeDeflin(wb_Ent, TxtTemp1, Aux);
        float[][] paramLineas;
        System.out.println("Lineas: "+numLin);
        paramLineas = new float[numLin][10];
        String [] nombreLineas = new String[numLin];
        for(int i=0; i < numLin; i++){
            for(int j=0; j <= 9; j++){
                paramLineas[i][j] = (float)Aux[i][j];
            }
            nombreLineas[i] = TxtTemp1[i];
        }

        /*
         * lee Lineas Troncales
         * ====================
         */
        intAux1=new int[2500];
        int[][] intAux2 = new int[2500][3];
        TxtTemp1=new String[2500];
        TxtTemp2=new String[2500];
        numLinTron = Lee.leeLintron(wb_Ent, TxtTemp1, nombreLineas,TxtTemp2, intAux1, intAux2);
        String[] nomLinTron = new String[numLinTron];
        int[] indiceLintron = new int[numLinTron];
        int[][] datosLintron = new int[numLinTron][3];
//        String[]zonaIT=new String[numLinTron];
        String[] nomProp=new String[numLinTron];
        String[] LinTronProp=new String[numLinTron];
         String TxtTemp4[]=new String[numLinTron];
        for(int l=0; l < numLinTron; l++){
            nomLinTron[l] = TxtTemp1[l];
            nomProp[l]=TxtTemp2[l];
            LinTronProp[l]=nomLinTron[l]+"#"+nomProp[l];
            indiceLintron[l] = intAux1[l];
            datosLintron[l][0] = intAux2[l][0];
            datosLintron[l][1] = intAux2[l][1];
            datosLintron[l][2] = intAux2[l][2];
            TxtTemp4[l]="";
        }


        TxtTemp3=new String[numLinTron];
        IntTemp=new int[numLinTron];

        int numLinTx = 0;
        for (int i = 0; i <numLinTron; i++) {
//             if(datosLintron[i][0]==1) zonaIT[i]="N";
//             if(datosLintron[i][0]==0) zonaIT[i]="A";
//             if(datosLintron[i][0]==-1) zonaIT[i]="S";
            int l = Calc.Buscar(nomLinTron[i] + "#" + nomProp[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinTron[i] + "#" + nomProp[i];
                TxtTemp2[numLinTx]=nomProp[i];
                IntTemp[numLinTx]=datosLintron[i][0];
                if(datosLintron[i][0]==1){
                    TxtTemp3[numLinTx]="N";
                }
                else if(datosLintron[i][0]==0){
                     TxtTemp3[numLinTx]="A";
                }
                 else if(datosLintron[i][0]==-1){
                     TxtTemp3[numLinTx]="S";
                 }
                numLinTx++;
            }
        }
        String[] nomLinTx = new String[numLinTx];//solo registros inicio Linea-Transmisor
//        String[] nomPropTx = new String[numLinTx];
        String[] zona = new String[numLinTx];
        int[] datosLinIT = new int[numLinTx];
        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
//            nomPropTx[i]=TxtTemp2[i];
            zona[i] = TxtTemp3[i];
            datosLinIT[i]=IntTemp[i];
        }


        /*
         * lee de Barras
         * =============
         */
        TxtTemp1=new String[2500];
        int[][] intAux3=new int[2500][4];
        numBarras = Lee.leeDefbar(wb_Ent, TxtTemp1, intAux3);
        String [] nomBar = new String[numBarras];
        paramBarTroncal = new int[numBarras][3];
        int numBarrasTroncales = 0;
        for(int i=0;i<numBarras;i++){
            nomBar[i] = TxtTemp1[i];
            // 1 si la barra es troncal, 0 en caso contrario
            paramBarTroncal[i][0] = intAux3[i][0];
            // 0 si la barra esta en el AIC, 1 si esta en el norte y -1 si esta en el sur
            paramBarTroncal[i][1] = intAux3[i][1];
            // 1 si la barra está en el SIC, -1 si está en el SING
            paramBarTroncal[i][2] = intAux3[i][2];
            if(paramBarTroncal[i][0] == 1){
                numBarrasTroncales++;
            }
        }
        nombreSlack=nomBar[NumeroSlack-1];

        /*
         * Lectura de consumos
         * ===================
         */
        Consumos = new float[numBarras][numEtapas];
        Lee.leeConsumoxBarra(wb_Ent,Consumos,numBarras,numEtapas);
        float[][] consumoNormalizado=new float[numBarras][numEtapas];
        boolean[][] barrasConsumo = new boolean[numBarras][numEtapas];
        float[] ConsEta = new float[numEtapas];
        System.out.println("Barras: "+numBarras);
        for(int b=0; b < numBarras; b++){
            for(int e=0;e<numEtapas;e++){
                barrasConsumo[b][e] = (Consumos[b][e] != 0);
                ConsEta[e] += Consumos[b][e];
            }
        }
        for(int e=0; e < numEtapas; e++){
            for(int b=0; b < numBarras; b++){
            consumoNormalizado[b][e] = Consumos[b][e]/ConsEta[e];
            }
        }
        int[] duracionEta = new int[numEtapas];
        Lee.leeEtapas(wb_Ent,duracionEta,numEtapas);

        /*
         * Lectura de mantenimientos de lineas
         * ===================================
         */
        // cambios en condicion operativa para cada linea en cada etapa.
        int[][] LinMan = new int[numLin][numEtapas];
        for(int i=0; i < numLin; i++) {
            for(int j=0; j < numEtapas; j++) {
                LinMan[i][j] = -1;
            }
        }
        Lee.leeLinman(wb_Ent, LinMan, nombreLineas, numEtapas);

        /*******************
        Lectura de Centrales
        ********************/
        TxtTemp1 = new String[700];
        String [] TxtTemp1_2 = new String[700];
        float[] Temp1 = new float[700];
        float[] Temp2= new float[700];
        int numCen = Lee.leeCentrales(wb_Ent, TxtTemp1,Temp1,Temp2);
        String[] nombreCentrales = new String[numCen];
        System.arraycopy(TxtTemp1, 0, nombreCentrales, 0, numCen);

        /******************************
        Lectura de datos de generadores
        *******************************/
        TxtTemp1 = new String[1000];
        
        numGen = Lee.leePlpcnfe(wb_Ent,TxtTemp1,
                intAux3,nombreCentrales);
        
        int numGen_Sin_Fallas = Lee.leePlpcnfe(wb_Ent,TxtTemp1_2,nombreCentrales);
        
        
        System.out.println("Generadores: "+numGen);
        paramGener = new int[numGen][3];
        String [] nomGen = new String[numGen];
        String [] nomGen_Sin_Fallas = new String[numGen_Sin_Fallas];
//        boolean[] barrasGeneracion = new boolean[numBarras];
//        for(int i=0; i < numBarras; i++){
//            barrasGeneracion[i] = false;
//        }
        for(int i=0; i < numGen; i++){
            nomGen[i] = TxtTemp1[i];
           // System.out.println("Peajes "+nomGen[i]);
            paramGener[i][1] = intAux3[i][1];
            paramGener[i][0] = intAux3[i][0];
//            barrasGeneracion[paramGener[i][0]] = true;
        }
        System.arraycopy(TxtTemp1_2, 0, nomGen_Sin_Fallas, 0, numGen_Sin_Fallas);
            
            
            
        /*****************************************
        Lectura de orientacion de barras troncales
        ******************************************/
        orientBarTroncal=new int[numBarras][numLin];
        for(int i=0; i < numBarras; i++){
            for(int j=0; j < numLin; j++){
                orientBarTroncal[i][j]=0;
            }
        }
        Lee.leeOrient(wb_Ent, orientBarTroncal, nomBar,
                nombreLineas);

        /**************************
        Lectura de Suministradores
        **************************/
//        TxtTemp1 = new String[100];
        //int numSum = Lee.leeSumin(libroEntrada, TxtTemp1);
        //String [] nomSum = new String[numSum];
        //for(int i=0; i < numSum; i++){
          //  nomSum[i] = TxtTemp1[i];
        //}
        /*
         * Lectura de Consumos de Claves
         * =============================
         */

        float[][] Temporal1 = new float[2500][numEtapas];
        float[][] Temporal2 = new float[2500][NUMERO_MESES];
        float[][][] Temporal3 = new float[2500][3][NUMERO_MESES];
        int numClaves;

        if (ActClientes) {
            numClaves = Lee.leeConsumos2(rutaLibroEnt, Temporal1, Temporal2, numEtapas,
                     paramEtapa, duracionEta, Temporal3);
        } else {
            numClaves = Lee.leeConsumos(wb_Ent, Temporal1, Temporal2, numEtapas,
                     paramEtapa, duracionEta, Temporal3);
        }

        ConsumosClaves = new float[numClaves][numEtapas];
        for (int i = 0; i < numClaves; i++) {
            System.arraycopy(Temporal1[i], 0, ConsumosClaves[i], 0, numEtapas);
        }
        Temporal1 = null;
        
        float[][] ConsClaveMes = new float[numClaves][NUMERO_MESES];
        for (int i = 0; i < numClaves; i++) {
            System.arraycopy(Temporal2[i], 0, ConsClaveMes[i], 0, NUMERO_MESES);
        }
        Temporal2 = null;
        
        float[][][] ECU = new float[numClaves][3][NUMERO_MESES];
        for (int i = 0; i < numClaves; i++) {
            System.arraycopy(Temporal3[i], 0, ECU[i], 0, 3);
        }
        Temporal3 = null;

        /*for(int i=0; i<numClaves;i++){
            System.arraycopy(Temporal1[i], 0, ConsumosClaves[i], 0, numEtapas);
            System.arraycopy(Temporal2[i], 0, ConsClaveMes[i], 0, numMeses);
            System.arraycopy(Temporal3[i], 0, ECU[i], 0, 3);
        }
        */
        
        
        
        /******************
        Lectura de Clientes
        *******************/
        TxtTemp1 = new String[2500];
        String[] Exen = new String[2500];
        int numCli = Lee.leeClientes(wb_Ent, TxtTemp1,Exen);
        String [] nomCli = new String[numCli];
        System.arraycopy(TxtTemp1, 0, nomCli, 0, numCli);

        /*************************
        Lectura de Datos de Claves
        **************************/
        TxtTemp1 = new String[2500];
        TxtTemp2 = new String[2500];
        int clav = Lee.leeBarcli(wb_Ent, TxtTemp1,
                TxtTemp2, intAux3, nomCli, nomBar);
        String[] nombreClaves = new String[numClaves];
//        String[] nombreClClientes = new String[numClaves];
        datosClaves = new int[numClaves][4];
        for(int i=0; i < numClaves; i++) {
            nombreClaves[i]=TxtTemp1[i];
//            nombreClClientes[i]=TxtTemp2[i];
            datosClaves[i][0]=intAux3[i][0];
            datosClaves[i][1]=intAux3[i][1];
            datosClaves[i][2]=intAux3[i][2];
            datosClaves[i][3]=intAux3[i][3];
        }

        TxtTemp2 = null;
        //Escribe la energia consumida por Cliente 
        //* ==============================================

        float[][] CMes = new float[numCli][NUMERO_MESES];
        float[][][] ECUCli = new float[numCli][3][NUMERO_MESES];
        float[] ConsCliAno = new float[numClaves];

        for(int j=0; j<numClaves; j++){
            for(int m=0; m<NUMERO_MESES; m++){
            ECUCli[datosClaves[j][2]][0][m]+=ECU[j][0][m];
            ECUCli[datosClaves[j][2]][1][m]+=ECU[j][1][m];
            ECUCli[datosClaves[j][2]][2][m]+=ECU[j][2][m];
            CMes[datosClaves[j][2]][m]+=ConsClaveMes[j][m];
            ConsCliAno[datosClaves[j][2]]+=ConsClaveMes[j][m];
            }
        }

        /*
         * Lectura de Generacion (Despachos) y energia no suministrada
         * ===========================================================
         */
        System.out.println("Inicio lectura archivo despacho PLP...");
        long time_dispatch = System.currentTimeMillis();
        int cuenta = 0; //Contador de lineas
        File testReadFile = new File(ArchivoDespachoGeneradores);
        BufferedReader input = null;
        Gx = new float[numGen][numEtapas][numHid]; //Despacho PLP
        FallaEtaHid = new float[numEtapas][numHid];   //Falla PLP
        java.util.Map<String, Integer> m_nomGen = new java.util.TreeMap<String , Integer>();
        java.util.Map<String, Integer> m_nomGen_Sin_Fallas = new java.util.TreeMap<String , Integer>();
        
        //Chequeamos consistencia en las contantes:
        assert(!(USE_BUFFEREDSTREAM & (USE_SCANNER | USE_FILECHANNEL))) : "Only one option USE_BUFFEREDSTREAM, USE_SCANNER or USE_FILECHANNEL should be true! Did you mess with these constants?";
        assert(!(USE_FILECHANNEL & (USE_BUFFEREDSTREAM | USE_SCANNER))) : "Only one option USE_BUFFEREDSTREAM, USE_SCANNER or USE_FILECHANNEL should be true! Did you mess with these constants?";
        assert(!(USE_SCANNER & (USE_BUFFEREDSTREAM | USE_FILECHANNEL))) : "Only one option USE_BUFFEREDSTREAM, USE_SCANNER or USE_FILECHANNEL should be true! Did you mess with these constants?";
        
        if (USE_MAPPED_NAMES) {
            int cont=0;
            for (String g: nomGen){
                m_nomGen.put(g, cont);
                cont++;
            }
            cont=0;
            for (String g: nomGen_Sin_Fallas){
                m_nomGen_Sin_Fallas.put(g, cont);
                cont++;
            }
        }
        
        if (USE_BUFFEREDSTREAM) {
        try {
            //input = new BufferedReader( new FileReader(testReadFile) );
            input = new BufferedReader( new InputStreamReader(new FileInputStream(testReadFile), "Latin1"));
            String line = null;
            cuenta=0;
            int indGen=0;
            int indGen2=0;
            int indHid=0;
            int indEta=0;
            float Pgen=0;
            float ENS=0;
            String sNomGenPrev = "";
            Integer indexGen = null;
            Integer indexGenSinFallas = null;
            while ((line = input.readLine()) != null){
                if(cuenta>0){
                    if((line.substring(0,5).trim()).equals("MEDIA")==false){
                        if (USE_MAPPED_NAMES) {
                            String sNomGenActual = (line.substring(32, line.indexOf(",", 32))).trim();
                            if (!sNomGenActual.equals(sNomGenPrev)) {
                                indexGen = m_nomGen.get(sNomGenActual);
                                indexGenSinFallas = m_nomGen_Sin_Fallas.get(sNomGenActual);
                                sNomGenPrev = sNomGenActual;
                            }
                            if (indexGenSinFallas != null) {
                                indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                                indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                                if (indexGen != null) {
                                    if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                        Pgen = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                                        Gx[indexGen][indEta - etapaPeriodoIni][indHid] = Pgen; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                    }
                                }
                            } else {
                                if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                    ENS = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                                    FallaEtaHid[indEta - etapaPeriodoIni][indHid] += ENS; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                }
                            }
                        } else {
                            indGen = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen);
                            indGen2 = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen_Sin_Fallas);
                            
                            if(indGen2>-1){
                          if(indGen>-1){  
                            
                            //System.out.println(indGen +" "+nomGen[indGen]);
                            indHid=Integer.valueOf((line.substring(4,line.indexOf(",",4))).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,line.indexOf(",",8))).trim())-1;
                            //Pgen=Float.valueOf((line.substring(151,158)).trim()); //Peajes
                            Pgen=Float.valueOf((line.substring(103,line.indexOf(",",103))).trim()); 
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                //System.out.println(indGen +" "+nomGen[indGen]);
                                Gx[indGen][indEta-etapaPeriodoIni][indHid]=Pgen; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                //System.out.println(Gx[indGen][indEta-etapaPeriodoIni][indHid]);
                            }
                            
                          }
                            
                        }
                        
                        
                        else{//suma las fallas
                            indHid=Integer.valueOf((line.substring(4,line.indexOf(",",4))).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,line.indexOf(",",8))).trim())-1;
                            //ENS=Float.valueOf((line.substring(151,158)).trim());
                            ENS=Float.valueOf((line.substring(103,line.indexOf(",",103))).trim());
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                FallaEtaHid[indEta-etapaPeriodoIni][indHid]+=ENS; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                            }
                        }
                        }
                        
                    }
                }
                cuenta++;
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace(System.out);
        } catch (IOException ex) {
            ex.printStackTrace(System.out);
        } catch (NumberFormatException e) {
            throw new IOException("No fue posible convertir valor en linea '" + cuenta + "' archivo plp '" + ArchivoDespachoGeneradores +"'", e);
        } finally {
            try {
                if (input != null) {
                    input.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace(System.out);
            }
        }
        }

        if (USE_FILECHANNEL) {
            RandomAccessFile aFile = new RandomAccessFile(ArchivoDespachoGeneradores, "r");
            FileChannel inChannel = aFile.getChannel();
            ByteBuffer buffer = ByteBuffer.allocate(1024 * 1000); //1GB buffer
            String line = "";
            int indGen = 0;
            int indGen2 = 0;
            int indHid = 0;
            int indEta = 0;
            float Pgen = 0;
            float ENS = 0;
            while (inChannel.read(buffer) > 0) {
                buffer.flip();
                for (int i = 0; i < buffer.limit(); i++) {
                    boolean isEOL = false;
                    char c = ((char) buffer.get());
                    if (c == '\r' || c == '\n') {
                        isEOL = true;
                    } else {
                        line += c;
                    }
                    if (isEOL) {
                        //Procesar linea:
                        if (cuenta > 0) {
                            if ((line.substring(0, 5).trim()).equals("MEDIA") == false) {
                                indGen = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen);
                                indGen2 = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen_Sin_Fallas);
                                if (indGen2 > -1) {
                                    if (indGen > -1) {
                                        indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                                        indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                                        Pgen = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                                        if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                            Gx[indGen][indEta - etapaPeriodoIni][indHid] = Pgen; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                        }
                                    }
                                } else {//suma las fallas
                                    indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                                    indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                                    ENS = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                                    if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                        FallaEtaHid[indEta - etapaPeriodoIni][indHid] += ENS; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                    }
                                }
                            }
                        }
                        line = "";
                        cuenta++;
                    }
                }
                buffer.clear(); // do something with the data and clear/compact it.
            }
            inChannel.close();
            aFile.close();
        }
        
        if (USE_SCANNER) {
            RandomAccessFile aFile = new RandomAccessFile(ArchivoDespachoGeneradores, "r");
            FileChannel inChannel = aFile.getChannel();
            Scanner fileScanner = new Scanner(inChannel, StandardCharsets.ISO_8859_1.name());
//            Scanner fileScanner = new Scanner(new BufferedReader(new InputStreamReader(new FileInputStream(ArchivoDespachoGeneradores), StandardCharsets.ISO_8859_1.name())));
            String line = "";
            int indGen = 0;
            int indGen2 = 0;
            int indHid = 0;
            int indEta = 0;
            float Pgen = 0;
            float ENS = 0;
            while (fileScanner.hasNextLine()) {
                line = fileScanner.nextLine();
                if (cuenta > 0) {
                    if ((line.substring(0, 5).trim()).equals("MEDIA") == false) {
                        indGen = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen);
                        indGen2 = Calc.Buscar((line.substring(32, line.indexOf(",", 32))).trim(), nomGen_Sin_Fallas);
                        if (indGen2 > -1) {
                            if (indGen > -1) {
                                indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                                indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                                Pgen = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                                if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                    Gx[indGen][indEta - etapaPeriodoIni][indHid] = Pgen; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                                }
                            }
                        } else {//suma las fallas
                            indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                            indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                            ENS = Float.valueOf((line.substring(103, line.indexOf(",", 103))).trim());
                            if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                FallaEtaHid[indEta - etapaPeriodoIni][indHid] += ENS; //TODO: This can cause an java.lang.ArrayIndexOutOfBoundsException when the selected hydro is lower than the values in plp file!
                            }
                        }
                    }
                }
                cuenta++;
            }
        }
        
        System.out.println("Fin lectura archivo despacho PLP : " + ((System.currentTimeMillis() - time_dispatch)/1000) + "[seg]");
        
        /*
         * Lectura de datos de Lineas del sistema reducido
         * ===============================================
         */
        long time_flow = System.currentTimeMillis();
        TxtTemp1=new String[600];
        float[][] Aux1 = new float[600][1];
        int numLinSistRed = Lee.leeLinPLP(wb_Ent, TxtTemp1, Aux1);
        int[] paramLinSistRed=new int[numLinSistRed];
        String [] nombreLineasSistRed = new String[numLinSistRed];
        for(int i=0;i<numLinSistRed;i++){
            nombreLineasSistRed[i]=TxtTemp1[i];
            //tension lineas sistema reducido
            paramLinSistRed[i]=(int)(Math.round(Aux[i][0]));
        }

        TxtTemp1 = null;
        
        /*
         * Lectura de Perdidas en lineas de tension >110 kV
         * ================================================
         */
        System.out.println("Leyendo archivo flujos lineas PLP...");
        testReadFile = new File(ArchivoPerdidasLineas);
        input = null;
        perdidasPLPMayor110 = new float[numEtapas][numHid];
        java.util.Map<String, Integer> m_nombreLineasSistRed = new java.util.TreeMap<String, Integer>();
        if (USE_MAPPED_NAMES) {
            int cont = 0;
            for (String l: nombreLineasSistRed) {
                m_nombreLineasSistRed.put(l, cont);
                cont++;
            }
        }
        try {
            input = new BufferedReader( new InputStreamReader(new
                    FileInputStream(testReadFile), "Latin1"));
            String line = null;
            cuenta=0;
            int indLin=0;
            int indHid=0;
            int indEta=0;
            float Perd=0;
            String sNomLineaPrev = "";
            Integer indexLin = null;
            while (( line = input.readLine()) != null){
                if(cuenta>0){
                    if (USE_MAPPED_NAMES) {
                        if ((line.substring(0, 5).trim()).equals("MEDIA") == false && (line.substring(0, 3).trim()).equals("Sim") == true) {
                            String sNomLineaActual = (line.substring(32, 79)).trim();
                            if (!sNomLineaActual.equals(sNomLineaPrev)) {
                                indexLin = m_nombreLineasSistRed.get(sNomLineaActual);
                                sNomLineaPrev = sNomLineaActual;
                            }
                            if (indexLin != null) {
                                if (paramLinSistRed[indexLin] > 110) {
                                    indHid = Integer.valueOf((line.substring(4, line.indexOf(",", 4))).trim()) - 1;
                                    indEta = Integer.valueOf((line.substring(8, line.indexOf(",", 8))).trim()) - 1;
                                    if (indEta < etapaPeriodoFin && indEta >= etapaPeriodoIni) {
                                        Perd = Float.valueOf(line.substring(111, line.indexOf(",", 111)).trim());
                                        perdidasPLPMayor110[indEta - etapaPeriodoIni][indHid] += Perd; //TODO
                                    }
                                }
                            }
                        }
                    } else {
                    if((line.substring(0,5).trim()).equals("MEDIA")==false && (line.substring(0,3).trim()).equals("Sim")==true){
                        indLin=Calc.Buscar((line.substring(32,79)).trim(),nombreLineasSistRed);
                        if(indLin>-1){
                            //System.out.println(indLin+" "+nombreLineasSistRed[indLin]);
                            indHid=Integer.valueOf((line.substring(4,line.indexOf(",",4))).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,line.indexOf(",",8))).trim())-1;
                            Perd=Float.valueOf(line.substring(111,line.indexOf(",",111)).trim());
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                if(paramLinSistRed[indLin]>110){
                                    perdidasPLPMayor110[indEta-etapaPeriodoIni][indHid]+=Perd; //TODO
                                }
                            }
                        }
                    }
                    }
                }
                cuenta++;
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace(System.out);
        } catch (IOException ex) {
            ex.printStackTrace(System.out);
        } catch (NumberFormatException e) {
            throw new IOException("No fue posible convertir valor en linea '" + cuenta + "' archivo plp '" + ArchivoPerdidasLineas +"'", e);
        } finally {
            try {
                if (input != null) {
                    input.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace(System.out);
            }
        }
        System.out.println("Fin lectura archivo flujos lineas PLP : " + ((System.currentTimeMillis() - time_flow)/1000) + "[seg]");
        
        /*
         * Escalamiento de la demanda
         * ==========================
         */
//        float[][]GxEtaHid = new float[numEtapas][numHid];
//        for(int h=0; h<numHid; h++){
//            for(int e=0; e<numEtapas; e++){
//                for(int g=0; g<numGen; g++){
//                    GxEtaHid[e][h]+=Gx[g][e][h];
//                }
//            }
//        }
        //DEPRECATED: Eliminado consumo escalado por etapas por no tener uso posterior:
//        float[][][] consumoEscalado=new float[numBarras][numEtapas][numHid];
//        for(int b=0;b<numBarras;b++){
//            for(int e=0;e<numEtapas;e++){
//                for(int h=0;h<numHid;h++){
//                    consumoEscalado[b][e][h]=(float)(consumoNormalizado[b][e]*GxEtaHid[e][h]); //Inutil?
//                }
//            }
//        }
        
        long tFinalLectura = System.currentTimeMillis();
        cargandoInfo=false;
        calculandoFlujos=true;
        long tInicioCalculo = System.currentTimeMillis();


        //Chequeo de consistencia:
        barrasActivas = Calc.ChequeoConsistencia(paramLineas, LinMan, numBarras, numEtapas);

        int nBarraSlack = Calc.Buscar(nombreSlack,nomBar);
        //
        Flujo = new float[numLin][numEtapas][numHid];

        //
        prorrGx = new float[numLin][numGen][numEtapas];
//        float[][] prorrEtaGx = new float[numLin][numGen];
        //
        prorrCx = new float[numLin][numCli][numEtapas];
//        float[][] prorrEtaCons = new float[numLin][numCli];
        //
//        float[][] GGDFref = new float[numLin][numHid];
//        float[][][] GGDFEtapa = new float[numBarras][numLin][numHid];
//        float[][] GLDFref = new float[numLin][numHid];
//        float[][][] GLDFEtapa = new float[numBarras][numLin][numHid];
        //
        System.out.println("Calculando...");
        //
        
        int[] centralesFlujo;
        int[] lineasFlujo;
//        centralesFlujo = Lee.leeCentralesFlujo(rutaLibroEnt, nomGen,"centrales_flujo");
        centralesFlujo = Lee.leeCentralesFlujo(wb_Ent, nomGen,"centrales_flujo", false); //TODO: move to config file
//        lineasFlujo = Lee.leeCentralesFlujo(rutaLibroEnt, nombreLineas,"lineas_flujo");
        lineasFlujo = Lee.leeCentralesFlujo(wb_Ent, nombreLineas,"lineas_flujo", false); //TODO: move to config file
        
        //Escritura del header archivo prorratas.csv:
        FileWriter writerProrratas = new FileWriter(DirBaseSalida + SLASH + "prorratas.csv");

        writerProrratas.append("Etapa");
        writerProrratas.append(',');
        writerProrratas.append("Hidrologia");
        writerProrratas.append(',');
        writerProrratas.append("Línea");
        writerProrratas.append(',');
        writerProrratas.append("Central");
        writerProrratas.append(',');
        writerProrratas.append("Prorrata");
        writerProrratas.append(',');
        writerProrratas.append("Gx");
        writerProrratas.append(',');
        writerProrratas.append("GGDF");
        writerProrratas.append('\n');

        writerProrratas.flush();
        writerProrratas.close();
        
        //Escritura del header archivo consumos.csv:
        FileWriter writerConsumos = new FileWriter(DirBaseSalida + SLASH + "consumos.csv");
        writerConsumos.append("Hidrologia,Etapa");
        for (int b = 0; b < numBarras; b++) {
            writerConsumos.append(",");
            writerConsumos.append(Float.toString(b));
        }
        writerConsumos.append("\n");
        
        
        
        /***************************************************/
        /*INICIO ITERACIONES (PARALELIZAR EL SIGUIENTE FOR)*/
        /**********LAS ITERACIONES SON POR ETAPA************/
        /***************************************************/
        
        //Lee opciones especificas del executor:
        String sMaxThreads = PeajesCDEC.getOptionValue("Max Threads", PeajesConstant.DataType.INTEGER);
        assert(sMaxThreads != null): "Como puede ser nulo esta importante llave? Cambiaste archivo config?";
        int nMaxThreads = Integer.parseInt(sMaxThreads);
        
        String sTimeOut = PeajesCDEC.getOptionValue("Thread timeout (en minutos)", PeajesConstant.DataType.INTEGER);
        assert(sMaxThreads != null): "Como puede ser nulo 'Thread timeout'? Cambiaste archivo config?";
        int nTimeOut = Integer.parseInt(sTimeOut);
        
        //Lee opciones de uso de memoria:
        String sMemoryUsage = PeajesCDEC.getOptionValue("Uso Memoria::STRING::Auto::Max::Min", PeajesConstant.DataType.STRING);
        PeajesConstant.UsoMemoria memory = PeajesConstant.UsoMemoria.valueOf(sMemoryUsage);
        boolean useDisk = false;
        switch (memory) {
            case Auto:
                useDisk = (nMaxThreads > 4);
                break;
            case Min:
                useDisk = true;
                break;
            case Max:
                useDisk = false;
                break;
        }
        System.out.println("Usando executor: Total etapas=" + numEtapas + ". Threads Paralelo=" + nMaxThreads + ". Uso Memoria=" + !useDisk);
        ExecutorService exeService;
        if (USE_FACTORY) {
            if (nMaxThreads > 1) {
                exeService = Executors.newFixedThreadPool(nMaxThreads);
            } else {
                exeService = Executors.newSingleThreadExecutor();
            }
        } else {
            exeService = new ExtendedExecutor(nMaxThreads, nTimeOut);
        }
        
        long initExecutorTime = System.currentTimeMillis();
        for(etapa=0;etapa<numEtapas;etapa++) {
//            System.out.println("etapa : "+etapa);
            float[][] paramLinEta = new float[numLin][10];
            for (int l = 0; l < numLin; l++) {
                for (int i = 0; i <= 5; i++) {
                    paramLinEta[l][i] = paramLineas[l][i];
                }
                if (LinMan[l][etapa] != (-1)) {
                    paramLinEta[l][5] = LinMan[l][etapa];	//Operativa
                }
            }
            for (int l = 0; l < numLinTron; l++) {
                paramLinEta[indiceLintron[l]][6] = 1;
                paramLinEta[indiceLintron[l]][7] = datosLintron[l][0];
                paramLinEta[indiceLintron[l]][8] = datosLintron[l][1];
                paramLinEta[indiceLintron[l]][9] = datosLintron[l][2];
            }
            
            exeService.submit(new ProrratasExe(etapa, nBarraSlack, numGen, numLin, numBarras, numHid, DirBaseSalida, numClaves, paramLinEta, useDisk));
        }
        long elapsed = System.currentTimeMillis() - initExecutorTime;
        System.out.println("===========Submitted all tasks: Time: " + elapsed / 1000 + "[sec]===========");
        exeService.shutdown();
        try {
            exeService.awaitTermination(1, TimeUnit.DAYS);
        } catch (InterruptedException e) {
            e.printStackTrace(System.out);
        } catch (Exception e) {
            e.printStackTrace(System.out);
        }
        elapsed = System.currentTimeMillis() - initExecutorTime;
        System.out.println("===========Finished Parallel Execution! Total time: " + elapsed / 1000 + "[sec]===========");

        writerConsumos.flush();
        writerConsumos.close();

        long tfinIteraciones = System.currentTimeMillis();

        calculandoFlujos = false;
        calculandoProrr = true;
        
        /**********************************/
        /******** FIN ITERACIONES *********/
        /**********************************/
        
        //Escribe flujos ajustados por hidrologia:
        FileWriter writerFlujosHidrologia = new FileWriter(DirBaseSalida + SLASH + "flujos_hidrologia.csv");
        //header:
        writerFlujosHidrologia.append("Hidrologia,Etapa");
        for (int l = 0; l < lineasFlujo.length; l++) {
            writerFlujosHidrologia.append(',');
            writerFlujosHidrologia.append(nombreLineas[lineasFlujo[l]]);
        }
        writerFlujosHidrologia.append('\n');
        //valores:
        for (int hh = 0; hh < numHid; hh++) {
            for (int e = 0; e < numEtapas; e++) {
                writerFlujosHidrologia.append(String.valueOf(hh));
                writerFlujosHidrologia.append(',');
                writerFlujosHidrologia.append(String.valueOf(e));
                for (int l = 0; l < lineasFlujo.length; l++) {
                    writerFlujosHidrologia.append(',');
                    writerFlujosHidrologia.append(Float.toString(Flujo[lineasFlujo[l]][e][hh]));
                }
                writerFlujosHidrologia.append('\n');
            }
        }
        writerFlujosHidrologia.flush();
        writerFlujosHidrologia.close();
        
        //Escribe flujos medios (promedios de hidrologias):
        FileWriter writerFlujosMedios = new FileWriter(DirBaseSalida + SLASH + "flujos_medios.csv");
        float[][] FlujoMedio = new float[numLin][numEtapas];
        for (int e = 0; e < numEtapas; e++) {
            for (int l = 0; l < numLin; l++) {
                for (int h = 0; h < numHid; h++) {
                    FlujoMedio[l][e] += Flujo[l][e][h] / numHid;
                }
            }
        }
        //header:
        writerFlujosMedios.append("Etapa");
        for (int l = 0; l < lineasFlujo.length; l++) {
            writerFlujosMedios.append(',');
            writerFlujosMedios.append(nombreLineas[lineasFlujo[l]]);
        }
        writerFlujosMedios.append('\n');

        //valores:
        for (int e = 0; e < numEtapas; e++) {
            writerFlujosMedios.append(String.valueOf(e));
            for (int l = 0; l < lineasFlujo.length; l++) {
                writerFlujosMedios.append(',');
                writerFlujosMedios.append(Float.toString(FlujoMedio[lineasFlujo[l]][e]));
            }
            writerFlujosMedios.append('\n');
        }
        writerFlujosMedios.flush();
        writerFlujosMedios.close();
        
        //Calcula consumo mensual y anual de Energia:
        float ConsumoAnualEnergial = 0;
        int etapasPeriodo = etapaPeriodoFin - etapaPeriodoIni;
        float[] ConsumoMensualEnergia = new float[NUMERO_MESES];
        int mes = 0;
        for (int e = 0; e < etapasPeriodo; e++) {
            mes = (int) Math.floor((double) e / (etapasPeriodo / NUMERO_MESES));
            ConsumoAnualEnergial += ConsEta[e] * (float) duracionEta[e];
            ConsumoMensualEnergia[mes] += ConsEta[e] * (float) duracionEta[e];
        }
        
        //Calcula prorratas mensuales y anuales para todas las lineas:
        double[][] prorrAnoG=new double[numLin][numGen];
        double[][][] prorrMesG=new double[numLin][numGen][NUMERO_MESES];
        double[][] prorrAnoC=new double[numLin][numCli];
        double[][][] prorrMesC=new double[numLin][numCli][NUMERO_MESES];
        for(int l=0;l<numLin;l++){
            for(int e=0;e<etapasPeriodo;e++){
                mes=(int)Math.floor((double)e/(NumeroEtapasAno/NUMERO_MESES));
                for(int g=0;g<numGen;g++){
                    //System.out.println(prorrGx[l][g][e]);
                    prorrAnoG[l][g] += prorrGx[l][g][e] * ( ConsEta[e] * duracionEta[e] / ConsumoAnualEnergial );
                    prorrMesG[l][g][mes] += prorrGx[l][g][e]*( ConsEta[e] * duracionEta[e] / ConsumoMensualEnergia[mes] );
                }
                for(int c=0;c<numCli;c++){
                    prorrAnoC[l][c] += prorrCx[l][c][e] * (ConsEta[e] * duracionEta[e] / ConsumoAnualEnergial);
                    prorrMesC[l][c][mes]+= prorrCx[l][c][e]*(ConsEta[e] *duracionEta[e]/ConsumoMensualEnergia[mes]);
                }
            }
        }
        
        // Calcula prorratas mensuales y anuales para para las lineas troncales:
        double[][] prorrAnoTroncG = new double[numLinTx][numGen];
        double[][] prorrAnoTroncC = new double[numLinTx][numCli];
        double[][][] prorrMesTroncG=new double[numLinTx][numGen][NUMERO_MESES];
        double[][][] prorrMesTroncC=new double[numLinTx][numCli][NUMERO_MESES];
        double[][] ProrrVerMesLinG = new double[numLinTron][NUMERO_MESES];
        double[][] ProrrVerMesLinC = new double[numLinTron][NUMERO_MESES];
        for (int l = 0; l < numLinTron; l++) {
            int l2 = Calc.Buscar(LinTronProp[l], nomLinTx);
            //System.out.println(l+" "+LinTronProp[l]+" "+nomLinTx[l2]+" "+l2);
            for (int g = 0; g < numGen; g++) {
                prorrAnoTroncG[l2][g] += prorrAnoG[indiceLintron[l]][g];
                for (int m = 0; m < NUMERO_MESES; m++) {
                    prorrMesTroncG[l2][g][m] += prorrMesG[indiceLintron[l]][g][m];
                    ProrrVerMesLinG[l][m] += prorrMesG[indiceLintron[l]][g][m];
                }
            }
            for (int c = 0; c < numCli; c++) {
                prorrAnoTroncC[l2][c] += prorrAnoC[indiceLintron[l]][c];
                for (int m = 0; m < NUMERO_MESES; m++) {
                    prorrMesTroncC[l2][c][m] += prorrMesC[indiceLintron[l]][c][m];
                    ProrrVerMesLinC[l][m] += prorrMesC[indiceLintron[l]][c][m];
                }
            }
        }
        double[] sumPorrMesG = new double[numGen];
        double[][] sumProrrMesLinG = new double[numLinTx][NUMERO_MESES];
        double[][] sumProrrMesLinC = new double[numLinTx][NUMERO_MESES];
        for (int l=0; l<numLinTx; l++) {
            for(int m=0; m<NUMERO_MESES; m++) {
                for (int g=0; g<numGen; g++) {
                    sumPorrMesG[g] += prorrMesTroncG[l][g][m];
                    sumProrrMesLinG[l][m] += prorrMesTroncG[l][g][m];
                }
                for (int c=0; c<numCli; c++)
                    sumProrrMesLinC[l][m] += prorrMesTroncC[l][c][m];
            }
        }
        // prorratas por Linea
        double[][] prorrataLinea = new double[numLinTx][NUMERO_MESES];
        double[][] prorrataLineaTron = new double[numLinTron][NUMERO_MESES];
        for(int l=0; l<numLinTx; l++)
            for(int m=0; m<NUMERO_MESES; m++)
                prorrataLinea[l][m] =
                        sumProrrMesLinC[l][m] + sumProrrMesLinG[l][m];
        for(int l=0; l<numLinTron; l++)
            for(int m=0; m<NUMERO_MESES; m++)
                prorrataLineaTron[l][m] =
                        ProrrVerMesLinC[l][m] + ProrrVerMesLinG[l][m];
        // Factor de correccion
        double[][] FactorG = new double[numLinTx][NUMERO_MESES];
        double[][] FactorC = new double[numLinTx][NUMERO_MESES];
        for (int l = 0; l < numLinTx; l++) {
            for (int m = 0; m < NUMERO_MESES; m++) {
                if (datosLinIT[l] == 0) {
                    double FdenG = sumProrrMesLinG[l][m];
                    if (Math.round(1000000000 * FdenG) == 0) {
                        FdenG = 1.0;
                    }
                    FactorG[l][m] = 0.8 / FdenG;
                    double FdenC = sumProrrMesLinC[l][m];
                    if (Math.round(1000000000 * FdenC) == 0) {
                        FdenC = 1.0;
                    }
                    FactorC[l][m] = 0.2 / FdenC;
                } else {
                    double Fden = prorrataLinea[l][m];
                    if (Math.round(1000000000 * Fden) == 0) {
                        Fden = 1.0;
                    }
                    FactorG[l][m] = 1 / Fden;
                    FactorC[l][m] = 1 / Fden;
                }
            }
        }
        
        // Procesa salida prorratas de generacion
        double[][][] prorrMesLinG = new double[numLinTx][numCen][NUMERO_MESES];
        double[][] generacionMes = new double[numCen][NUMERO_MESES];
        for(int m=0; m<NUMERO_MESES; m++) {
            for(int c=0; c<numCen; c++) {
                generacionMes[c][m] = 0;
                for(int l=0; l<numLinTx; l++) {
                    prorrMesLinG[l][c][m] = 0;
                }
            }
        }
        for (int l=0; l<numLinTx; l++) {
            for(int m=0; m<NUMERO_MESES; m++) {
                for (int g=0; g<numGen; g++) {
                    if (sumPorrMesG[g] != 0) {
                        if (paramGener[g][1] == -1)
                            System.out.println("Generador " + g +" no asignado en 'centrales'");             
                        prorrMesLinG[l][paramGener[g][1]][m]
                                += prorrMesTroncG[l][g][m]*FactorG[l][m];
                    }
                }
            }
        }
        for (int g=0; g<numGen; g++) {
            if (sumPorrMesG[g] != 0) {
                for(int e=0;e<numEtapas;e++) {
                    for(int h=0; h<numHid; h++){
                        //System.out.println(Gx[g][e][h]);
                        generacionMes[paramGener[g][1]][paramEtapa[e]]
                                += Gx[g][e][h]*duracionEta[e]/numHid/1000;
                        //System.out.println(generacionMes[paramGener[g][1]][paramEtapa[e]]);
                    }
                }
            }
        }
        
        // Procesa salida final de prorratas de consumo
        double[][][] prorrMesLinC = new double[numLinTx][numCli][NUMERO_MESES];
        for (int l = 0; l < numLinTx; l++) {
            for (int c = 0; c < numCli; c++) {
                for (int m = 0; m < NUMERO_MESES; m++) {
                    prorrMesLinC[l][c][m]
                            += prorrMesTroncC[l][c][m] * FactorC[l][m];
                }
            }
        }
        String[] nombreEtapas=new String[numEtapas];
        String[] nombreHid=new String[numHid];
        for(int a=0; a<1; a++) {
            for(int e=etapaPeriodoIni;e<etapaPeriodoFin;e++){
                nombreEtapas[e-etapaPeriodoIni]="";
                nombreEtapas[e-etapaPeriodoIni]+=(e-etapaPeriodoIni+1);
            }
            for(int h=0; h<numHid; h++){
                nombreHid[h] = "";
                nombreHid[h] += (h+1);
            }
        }

        calculandoProrr=false;
        long tFinalCalculo = System.currentTimeMillis();
        long tInicioEscritura = System.currentTimeMillis();
        guardandoDatos=true;

        /*
         * Escritura de Resultados
         * =======================
         */
        String libroSalidaXLS = DirBaseSalida + SLASH + "Prorrata" + AnoAEvaluar + ".xlsx";
//        Escribe.crearLibro(libroSalidaXLS);
        XSSFWorkbook wb_salida = Escribe.crearLibroVacio (libroSalidaXLS);
        
        Escribe.creaH3F_3d_double(
                "Prorratas de Generación", prorrMesLinG,
                "Línea", nomLinTx,
                "Central", nombreCentrales,
                "Zona", zona,
                "Mes", MESES,
                wb_salida, "ProrrGMes", "0.000%;[Red]-0.000%;\"-\"");
        System.out.println("Acaba de crear la hoja xls ProrrGMes");
        Escribe.creaH3F_3d_double(
                "Prorratas de Consumo", prorrMesLinC,
                "Línea", nomLinTx,
                "Cliente", nomCli,
                "Zona", zona,
                "Mes", MESES,
                wb_salida, "ProrrCMes", "0.000%;[Red]-0.000%;\"-\"");
        System.out.println("Acaba de crear la hoja xls ProrrCMes");
        Escribe.creaH1F_2d_double(
                "Prorratas por Línea", prorrataLinea,
                "Línea", nomLinTx,
                "Mes", MESES,
                wb_salida, "ProrrLin", "0.000%;[Red]-0.000%;\"-\"");
        System.out.println("Acaba de crear la hoja xls ProrrLin");
        Escribe.creaH1F_2d_double(
                "Generación [GWh]", generacionMes,
                "Central", nombreCentrales,
                "Mes", MESES,
                wb_salida, "GMes", "0.0;[Red]-0.0;\"-\"");
        System.out.println("Acaba de crear la hoja xls GMes");
        Escribe.creaH1F_2d_float(
                "Consumo [MWh]", ConsClaveMes,
                "Cliente", nombreClaves,
                "Mes", MESES,
                wb_salida, "CMes", "0.0;[Red]-0.0;\"-\"");
        System.out.println("Acaba de crear la hoja xls CMes");
        Escribe.creaH1F_2d_double(
                "Detalle de prorratas de Generación", Calc.transponer(prorrAnoTroncG),
                "Central", nomGen,
                "Línea", nomLinTx,
                wb_salida, "ProrrG", "0.000%");
        System.out.println("Acaba de crear la hoja xls ProrrG");
        Escribe.creaH1F_2d_double(
                "Detalle de prorratas de Consumo", Calc.transponer(prorrAnoTroncC),
                "Clave", nomCli,
                "Linea", nomLinTx,
                wb_salida, "ProrrC", "0.000%");
        System.out.println("Acaba de crear la hoja xls ProrrC");
        Escribe.creaH1FT_2d_float(
                "Consumo [MWh]", CMes, ECUCli,
                "Cliente", nomCli,
                "Mes", MESES, EnergiaCU, "CU",
                wb_salida, "CMesCli", "0.0;[Red]-0.0;\"-\"");
        System.out.println("Acaba de crear la hoja xls CMesCli");
        Escribe.crea_verifProrrPeaj(prorrataLineaTron,
                nomLinTron,
                wb_Ent, "verProrr", "0.000%;[Red]-0.000%;\"-\"", 12);
        System.out.println("Acaba de actualiza planilla Ent hoja xls verProrr");
        
        // Graba y cierra conexion:
        Escribe.guardaLibroDisco(wb_Ent, rutaLibroEnt);
        Escribe.guardaLibroDisco(wb_salida, libroSalidaXLS);
        wb_Ent.close();
        wb_salida.close();
         
        guardandoDatos=false;
        long tFinalEscritura = System.currentTimeMillis();
        System.out.println("Tiempo Adquisición de datos     : "+dosDecimales.format((tFinalLectura-tInicioLectura)/1000.0)+" s");
        System.out.println("Tiempo Calculos                 : "+dosDecimales.format((tFinalCalculo-tInicioCalculo)/1000.0)+" s");
        System.out.println("Tiempo Iteraciones              : "+dosDecimales.format((tfinIteraciones-tInicioCalculo)/1000.0)+" s");
        System.out.println("Tiempo Escritura de Resultados  : "+dosDecimales.format((tFinalEscritura-tInicioEscritura)/1000.0)+" s");
        System.out.println("Tiempo total                    : "+dosDecimales.format((tFinalEscritura-tInicioLectura)/1000.0)+" s");

        completo=true;
    }

    public static void calculaEtapa(int etapa, int nBarraSlack, int numGen, int numLin, int numBarras, int numHid, String DirBaseSalida, int numClaves, float[][] paramLinEta, boolean useDisk) throws IOException {
        Matriz Ybarra;
        Matriz Xbarra;
        float[][] flujoDCEtapa = new float[numLin][numHid];
        float[] flujoDCHid = new float[numLin];
        float[][] GLDFref;
//        float[][][] GLDFEtapa; //Movido a objeto
        GGDF GLDFEtapa = null;
        float[][] GGDFref;
//        float[][][] GGDFEtapa; //Movido a objeto
        GGDF GGDFEtapa = null;
        
        float[][] prorrEtaGx;
        float[][] prorrEtaCons;
        
        float[]GxEtaHid = new float[numHid];
        for (int h = 0; h < numHid; h++) {
            for (int g = 0; g < numGen; g++) {
                GxEtaHid[h] += Gx[g][etapa][h];
            }
        }
        
        /*
        * Calcula matriz de Admitancias y matriz de Impedancias
        * =====================================================
        */
        try {
            int barrasEliminadas=0;
            // Calcula Ybarra considerando todas las barras, activas e inactivas
            Ybarra=new Matriz(Calc.CalculaYBarra(paramLinEta,numBarras,numLin));
            // Elimina de Ybarra las filas y columnas correspondientes a barras inactivas y la slack,
            // de manera de obtener una matriz invertible
            for(int b=0;b<numBarras;b++){
                if(barrasActivas[b][etapa]==false || b==nBarraSlack){
                    Ybarra=(Ybarra.EliminarFila(b-barrasEliminadas)).EliminarColumna(b-barrasEliminadas);
                    barrasEliminadas++;
                }
            }
            Xbarra=(Ybarra.InversionRapida()).uminus();
            /* Se agregan las filas y columnas de las barras inactivas y la slack rellenas con ceros,
            * de manera de mantener coeherencia en los indices de barras
            */
            for(int b=0;b<numBarras;b++){
                if(barrasActivas[b][etapa]==false || b==nBarraSlack){
                    Xbarra=(Xbarra.InsertarCerosFila(b)).InsertarCerosColumna(b);
                }
            }
            /* Calcula Factores de Desplazamiento A y
            GLDF barra referencia y GLDF resto del sistema. */
            float[][] GSDF = Calc.CalculaGSDF(Xbarra,paramLinEta,barrasActivas, etapa);
            GLDFref=Calc.CalculaGLDFRef(GSDF,paramLinEta,paramGener,etapa,Gx);
    //        GLDFEtapa=Calc.CalculaGLDF(GSDF,GLDFref,paramLinEta,etapa);
            GLDFEtapa=Calc.calculaGLDF(GSDF,GLDFref,paramLinEta, useDisk);
            // Calcula Flujo DC y asignacion de perdidas
            // -----------------------------------------
            float[] R=new float[numLin];                   // resistencias en p.u
            float[] perdI2R=new float[numLin];             // perdidas de cada linea segun I*I*R
            float[] perdidas=new float[numLin];            // perdidas de cada linea segun diferencia entre Gx y Demanda
            float[] perdMayor110=new float[numLin];        // perdidas de cada linea segun diferencia entre Gx y Demanda
            float[] perdMenor110=new float[numLin];        // perdidas de cada linea segun diferencia entre Gx y Demanda
            float[] perdRealesSistema=new float[numHid];   // perdidas del sistema
            float[] perdI2RSistMayor110=new float[numHid]; // perdidas de todas las lineas de tension > 110kV
            float[] perdI2RSistMenor110=new float[numHid]; // perdidas de todas las lineas de tension <= 110kV
            float conSist;
            float[][] conAjustado=new float[numBarras][numHid];
            float[] genSist=new float[numHid];
            float[] conMasPerd= new float[numBarras];      // consumos con asignacion de perdidas por iteracion [MW]
            // Consumos con asignacion de perdidas por iteracion [MW]
            float[][] conMasPerdEta= new float[numBarras][numHid];
            for (int h = 0; h < numHid; h++) {
                genSist[h] = GxEtaHid[h];
                for (int b = 0; b > numBarras; b++) {
                    conAjustado[b][h] = 0;
                }
            }
            for (int l = 0; l < numLin; l++) {
                R[l] = paramLinEta[l][3];                    // resistencia en p.u.
            }

    //        FileWriter writerConsumos = new FileWriter(DirBaseSalida + SLASH + "consumos.csv");
            //<--Inicio ciclo hidro
            for (int h = 0; h < numHid; h++) {

    //            writerConsumos.append(Float.toString(h));
    //            writerConsumos.append(",");
    //            writerConsumos.append(Float.toString(etapa));

                for (int l = 0; l < numLin; l++) {
                    perdI2R[l] = 0;
                    perdidas[l] = 0;
                    flujoDCHid[l] = 0;
                }
                perdRealesSistema[h] = 0;
                perdI2RSistMayor110[h] = 0;
                perdI2RSistMenor110[h] = 0;
                conSist = 0;
                for (int b = 0; b < numBarras; b++) {
                    conSist += Consumos[b][etapa];
                }
                for (int b = 0; b < numBarras; b++) {
                    conAjustado[b][h] += Consumos[b][etapa] * (conSist - FallaEtaHid[etapa][h]) / conSist;
    //                writerConsumos.append(",");
    //                writerConsumos.append(Float.toString(conAjustado[b][h]));
                }
    //            writerConsumos.append("\n");

                perdRealesSistema[h] = genSist[h] - (conSist - FallaEtaHid[etapa][h]);
                // Calculo de Flujo DC
    //            flujoDCHid = Calc.FlujoDC_GLDF(GLDFEtapa, conAjustado, h, etapa);//flujos en MW
                flujoDCHid = Calc.flujoDC_GLDF(GLDFEtapa, conAjustado, h);//flujos en MW
                //System.out.println("Flujo DC "+flujoDCHid[586]);
                for (int l = 0; l < numLin; l++) {
                    flujoDCEtapa[l][h] = flujoDCHid[l];
                }
                // Calcula perdidas
                for (int l = 0; l < numLin; l++) {
                    if (flujoDCHid[l] != 0) {
                        float sBase = 100;
                        perdI2R[l] = sBase * (R[l] * (flujoDCHid[l] / sBase) * (flujoDCHid[l] / sBase));	//perdidas en MW
                        //System.out.println("Perdidas cuadraticas "+ perdI2R[l]);
                    }
                    if (paramLinEta[l][2] > 110) {
                        perdI2RSistMayor110[h] += perdI2R[l];
                    } else {
                        perdI2RSistMenor110[h] += perdI2R[l];
                    }
                }
                // Perdidas Reales prorrateadas en las lineas de acuerdo al I2R de cada una
                perdMayor110 = Calc.ProrrPerdidas(perdidasPLPMayor110[etapa][h], perdI2R, paramLinEta, "Mayor_110", h);
                perdMenor110 = Calc.ProrrPerdidas((perdRealesSistema[h] - perdidasPLPMayor110[etapa][h]), perdI2R, paramLinEta, "Menor_Igual_110", h);
                for (int l = 0; l < numLin; l++) {
                    perdidas[l] = perdMayor110[l] + perdMenor110[l];
                }
                // Asigna perdidas a consumos
    //            conMasPerd = Calc.AsignaPerdidas(flujoDCHid, GLDFEtapa, perdidas, paramLinEta, conAjustado, etapa, h);
                conMasPerd = Calc.asignaPerdidas(flujoDCHid, GLDFEtapa, perdidas, paramLinEta, conAjustado, h);
            }
            //<--Fin ciclo hidro

            for (int h = 0; h < numHid; h++) {
                for (int l = 0; l < numLin; l++) {
                    Flujo[l][etapa][h] = flujoDCEtapa[l][h];
                }
                for (int b = 0; b < numBarras; b++) {
                    conMasPerdEta[b][h] = conMasPerd[b];
                }
            }
            /*
            * Calcula GGDF barra referencia y GGDF resto del sistema.
            */
            GGDFref=Calc.CalculaGGDFRef(GSDF,conMasPerdEta, paramLinEta);
    //        GGDFEtapa=Calc.CalculaGGDF(GSDF,GGDFref,paramLinEta,etapa);
            GGDFEtapa=Calc.calculaGGDF(GSDF,GGDFref,paramLinEta, useDisk);
            /*
            * Calcula prorratas promedio por etapa
            */
    //        prorrEtaGx=Calc.CalculaProrrGx(flujoDCEtapa, GGDFEtapa, Gx, paramGener, paramLinEta, paramBarTroncal,
    //                orientBarTroncal, etapa, centralesFlujo, lineasFlujo,GSDF,GGDFref );
            prorrEtaGx=Calc.calculaProrrGx(flujoDCEtapa, GGDFEtapa, Gx, paramGener, paramLinEta, paramBarTroncal, orientBarTroncal, etapa);
            prorrEtaCons=Calc.calculaProrrCons(flujoDCEtapa, GLDFEtapa, ConsumosClaves, datosClaves, paramLinEta, paramBarTroncal, orientBarTroncal, etapa);
            for (int l = 0; l < numLin; l++) {
                for (int g = 0; g < numGen; g++) {
                    prorrGx[l][g][etapa] = prorrEtaGx[l][g];
                }
                for (int c = 0; c < numClaves; c++) {
                    prorrCx[l][datosClaves[c][2]][etapa] += prorrEtaCons[l][c];
                }
            }
        } finally {
            if (GGDFEtapa != null) {
                GGDFEtapa.close();
            }
            if (GLDFEtapa != null) {
                GLDFEtapa.close();
            }
        }
        System.out.println("Finalizado calculo etapa : "+ etapa);
    }
    
    public Prorratas() {
    }
	
    public static float obtenerProgreso(){
        float progreso;
        if (numEtapas==0)
            progreso=0;
        else
            progreso=(float)(etapa+1)/(numEtapas);
        return progreso;
    }

    public static boolean cargando(){
        return cargandoInfo;
    }

    public static boolean calculaFlujos(){
        return calculandoFlujos;
    }

    public static boolean calculaProrratas(){
        return calculandoProrr;
    }

    public static boolean guardando(){
        return guardandoDatos;
    }

    public static boolean terminado(){
            return completo;
    }

    public static void Comenzar(final File DirIn, final File DirOut, final int AnoAEvaluar, final int tipoCalculo, final int AnoBase,
            final int NumHidro, final int NumEtapasAno, final int NumSlack, final int Offset, final boolean Cli) {
        javax.swing.SwingWorker worker = new javax.swing.SwingWorker() {

            @Override
            protected Object doInBackground() throws Exception {
                try {
                    CalculaProrratas(DirIn, DirOut, AnoAEvaluar, tipoCalculo, AnoBase, NumHidro, NumEtapasAno, NumSlack, Offset, Cli);
                } catch (IOException e) {
                    System.out.println(e);
                } catch (Exception e) {
                    e.printStackTrace(System.out);
                }
                return true;
            }
        };
        worker.execute();

    }
}

class ProrratasExe implements Runnable {
    private int etapa;
    private int nBarraSlack;
    private int numGen;
    private int numLin;
    private int numBarras;
    private int numHid;
    private String DirBaseSalida;
    private int numClaves;
    private float[][] paramLinEta;
    private boolean useDisk;

    public ProrratasExe(int etapa, int nBarraSlack, int numGen, int numLin, int numBarras, int numHid, String DirBaseSalida, int numClaves, float[][] paramLinEta, boolean useDisk) {
        this.etapa = etapa;
        this.nBarraSlack = nBarraSlack;
        this.numGen = numGen;
        this.numLin = numLin;
        this.numBarras = numBarras;
        this.numHid = numHid;
        this.DirBaseSalida = DirBaseSalida;
        this.numClaves = numClaves;
        this.paramLinEta = paramLinEta;
        this.useDisk = useDisk;
    }
    
    @Override
    public void run() {
        try {
            Prorratas.calculaEtapa(etapa, nBarraSlack, numGen, numLin, numBarras, numHid, DirBaseSalida, numClaves, paramLinEta, useDisk);
        } catch (IOException e) {
            e.printStackTrace(System.out);
        } catch (Exception e) {
            e.printStackTrace(System.out);
        }
    }
    
}

class ExtendedExecutor extends ThreadPoolExecutor {

    /**
     * Creates a new fixed-sized thread pool executor with the defined number of
     * threads and schedule time out
     *
     * @param maxThreads maximum number of threads
     * @param maxTimeOut maximum time-out before cancelling pending threads (in
     * minutes)
     */
    public ExtendedExecutor(int maxThreads, int maxTimeOut) {
        super(maxThreads, // core threads
                maxThreads, // max threads
                maxTimeOut, // timeout
                TimeUnit.MINUTES, // timeout units
                new LinkedBlockingQueue<Runnable>() // work queue
        );
    }

    @Override
    protected void afterExecute(Runnable r, Throwable t) {
        super.afterExecute(r, t);
        if (t == null && r instanceof Future<?>) {
            try {
                Future<?> future = (Future<?>) r;
                if (future.isDone()) {
                    future.get();
                }
            } catch (CancellationException ce) {
                t = ce;
            } catch (ExecutionException ee) {
                t = ee.getCause();
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
            }
        }
        if (t != null) {
            System.out.println(t);
        }
    }
}

/**
 * Clase auxiliar para almacener los objetos GGDF en archivo temporal
 * @author  Frank Leanez at www.flconsulting.cl
 */
class GGDF {
    
    private RandomAccessFile fp;
    private File f_bin;
    private final int numBarras;
    private final int numLineas;
    private final int numHidro;
    private float[][][] D;

    /**
     * Crea un nuevo objeto GGDF (o GLDF) a partir del archivo binario de
     * referencia
     *
     * @param f_bin archivo binario con datos de GGDF (o GLDF)
     * @param numBarras numero total de barras
     * @param numLineas numero total de lineas de transmision
     * @param numHidro numero total de hidrologias
     * @throws FileNotFoundException si hay problemas en acceder al archivo de
     * datos
     */
    public GGDF(File f_bin, int numBarras, int numLineas, int numHidro) throws FileNotFoundException {
        this.fp = new RandomAccessFile(f_bin, "r");
        this.f_bin = f_bin;
        this.numBarras = numBarras;
        this.numLineas = numLineas;
        this.numHidro = numHidro;
    }
    
    /**
     * Crea un nuevo objeto GGDF (o GLDF) a partir de la matriz del argumento
     *
     * @param E matriz 3D con los valores GGDF (barras, lineas, hidrologias)
     */
    public GGDF(float[][][] D) {
        this.D = D;
        this.numBarras = D.length;
        this.numLineas = D[0].length;
        this.numHidro = D[0][0].length;
    }
    
    /**
     * Chequea si este GDDF fue contruido desde una matriz en memoria o un
     * archivo binario
     *
     * @return true si almacena sus valores en memoria. false si use disco
     */
    boolean isMemoryGGDF() {
        return (D != null);
    }
    
    /**
     * WARNING: Evite en lo posible usar! Solo cuando se desee leer algun numero
     * puntual
     * <br>Obtiene los valores GGDF para la barra, linea e hidrologia
     * definida
     *
     * @param barra numero (0-base) de hidrologia segun orden en que fueron
     * ingresadas
     * @param linea numero (0-base) de linea segun orden en que fueron
     * ingresadas
     * @param hidro numero (0-base) de hidro segun orden en que fueron
     * ingresadas
     * @return valor del GGDF
     */
    float get(int barra, int linea, int hidro) throws IOException {
        if (isMemoryGGDF()) {
            return D[barra][linea][hidro];
        } else {
            long pos = position3D(hidro, linea, barra, numHidro, numLineas, numBarras) * 4; //4 bytes (float)
            fp.seek(pos);
            return fp.readFloat();
        }
    }
    
    /**
     * Obtiene los valores GGDF para todas las barras
     *
     * @param hidro numero (0-base) de hidrologia segun orden en que fueron
     * ingresadas
     * @param linea numero (0-base) de linea segun orden en que fueron
     * ingresadas
     * @return un arreglo de tamano getNumBarras() con todos (incluyento ceros)
     * los valores guardados
     */
    float[] get(int hidro, int linea) throws IOException {
        if (isMemoryGGDF()) {
            float[] f = new float[numBarras];
            for (int b = 0; b < numBarras; b++) {
                f[b] = D[b][linea][hidro];
            }
            return f;
        } else {
            long pos = position3D(hidro, linea, 0, numHidro, numLineas, numBarras) * 4; //4 bytes (float)
            fp.seek(pos);
            byte[] b = new byte[numBarras * 4];
            int nRead = fp.read(b, 0, numBarras * 4);
            return decode(b);
        }
    }
    
    /**
     * Obtiene la matriz clasica valores GGDF para la hidrologia ingresada
     * <br>La dimension del arreglo esta dado por getNumLineas() x
     * getNumBarras()
     * <br>El orden en que GGDF esta definido es la conversion a arreglo
     * "desplazando hacia la derecha" la matriz rectangular D(n, l). Es decir:
     * <br> GGDF[0, 1, 3, 4, 5 ... l x b] -> [D(0,0), D(0,1), D(0,2), ...
     * D(0,n-1), D(1,0), D(1,1), ... D(1,n-1), ... D(2,0), D(2,1)...
     *
     * @param hidro numero (0-base) de hidrologia segun orden en que fueron
     * ingresadas
     * @return un arreglo de tamano getNumLineas() x getNumBarras() con todos
     * (incluyento ceros) los valores guardados
     */
    float[] get(int hidro) throws IOException {
        if (isMemoryGGDF()) {
            float[] f = new float[numLineas * numBarras];
            int cont = 0;
            for (int l = 0; l < numLineas; l++) {
                for (int b = 0; b < numBarras; b++) {
                    f[cont] = D[b][l][hidro];
                    cont++;
                }
            }
            return f;
        } else {
            long pos = position3D(hidro, 0, 0, numHidro, numLineas, numBarras) * 4; //4 bytes (float)
            fp.seek(pos);
            byte[] b = new byte[numLineas * numBarras * 4]; //sin buffer?
            int nRead = fp.read(b, 0, numLineas * numBarras * 4);
            return decode(b);
        }
    }
    
//    float[] get(int barra, int linea) {
//        float[] lRet = new float[numHidro];
//        try {
//            long pos = position3D(barra, linea, 0, numBarras, numLineas, numHidro) * 4; //4 bytes (float)
//            long endPos = position3D(barra, linea, numHidro - 1, numBarras, numLineas, numHidro) * 4; //4 bytes (float)
//            fp.seek(pos);
//            int h = 0;
//            while (pos < endPos) {
//                lRet[h] = fp.readFloat();
//                h++;
//            }
//        } catch (IOException e) {
//            lRet = new float[numHidro];
//        }
//        return lRet;
//    }
    
    /**
     * Cierra los stream y channels abiertos al archivo binario
     */
    void close() {
        if (!isMemoryGGDF()) {
            try {
                if (fp != null) {
                    fp.close();
                }
            } catch (IOException e) {
                System.out.println("WARNING: temp file couldn't be cleaned. Must be done manually" + e.getMessage());
            }
        }
    }
    
    private static long position3D(int x, int y, int z, int maxX, int maxY, int maxZ) {
        long pos = ((long) maxY * (long) maxZ) * (long) x + ((long) maxZ * (long) y) + (long) z;
        long max = (long) maxX * (long) maxY * (long) maxZ;
        return Math.min(pos, max);
    }
    
    /**
     * Point al archivo binario que almacena los datos
     *
     * @return archivo binario que almacena los datos. null si este es un GGDF
     * almacenado en memoria
     */
    public File getF_bin() {
        return f_bin;
    }

    /**
     * Numero total de barras empleado para construir este GGDF
     * @return Numero total de barras empleado para construir este GGDF
     */
    public int getNumBarras() {
        return numBarras;
    }

    /**
     * Numero total de lineas empleado para construir este GGDF
     * @return Numero total de lineas empleado para construir este GGDF
     */
    public int getNumLineas() {
        return numLineas;
    }

    /**
     * Numero total de hidrologias empleado para construir este GGDF
     * @return Numero total de hidrologias empleado para construir este GGDF
     */
    public int getNumHidro() {
        return numHidro;
    }
    
    /**
     * Codificador Float array to Byte array 
     * <br>Autor original: Alexander Markov
     *
     * @param floatArray input arrray
     * @return output array
     */
    static byte[] encode(float floatArray[]) {
        byte byteArray[] = new byte[floatArray.length * 4];

        // wrap the byte array to the byte buffer 
        ByteBuffer byteBuf = ByteBuffer.wrap(byteArray);

        // create a view of the byte buffer as a float buffer 
        FloatBuffer floatBuf = byteBuf.asFloatBuffer();

        // now put the float array to the float buffer, 
        // it is actually stored to the byte array 
        floatBuf.put(floatArray);

        return byteArray;
    }

    /**
     * Decodificador Float array to Byte array 
     * <br>Autor original: Alexander Markov
     *
     * @param byteArray input array
     * @return output array
     */
    static float[] decode(byte byteArray[]) {
        float floatArray[] = new float[byteArray.length / 4];

        // wrap the source byte array to the byte buffer 
        ByteBuffer byteBuf = ByteBuffer.wrap(byteArray);

        // create a view of the byte buffer as a float buffer 
        FloatBuffer floatBuf = byteBuf.asFloatBuffer();

        // now get the data from the float buffer to the float array, 
        // it is actually retrieved from the byte array 
        floatBuf.get(floatArray);

        return floatArray;
    }
    
}