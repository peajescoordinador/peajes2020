package cl.coordinador.peajes;

/**
 * CalcPEF
 * @(#)Peajes.java
 *
 *
 * CDEC-SIC. Direccion de Peajes
 * @version 3.10 2008/Marzo
 * Modela y asigna p?rdidas a consumos
 */
import static cl.coordinador.peajes.PeajesConstant.MAX_COMPRESSION_RATIO;
import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.SLASH;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.text.DecimalFormat;

public class Prorratas {
	
    private static int etapa;
    private static int numEtapas=0;
    private static String nombreSlack;
    private static boolean cargandoInfo=false;
    private static boolean calculandoFlujos=false;
    private static boolean calculandoProrr=false;
    private static boolean guardandoDatos=false;
    private static boolean completo=false;


    public static void CalculaProrratas(File DirEntrada, File DirSalida, int Ano, int tipoCalc, int AnoIni,
            int NumeroHidrologias ,int NumeroEtapasAno, int NumeroSlack,int ValorOffset,boolean ActClientes) throws IOException, FileNotFoundException {
        
        final int numHid=NumeroHidrologias;//AnoIni-1962;
        final int offset=ValorOffset;//(AnoIni==2004?0:12);        
        String DirBaseEntrada=DirEntrada.toString();
        String DirBaseSalida=DirSalida.toString();
        final String ArchivoDespachoGeneradores= DirBaseEntrada + SLASH + "plpcen.csv";
        final String ArchivoPerdidasLineas = DirBaseEntrada + SLASH + "plplin.csv";
	//indices de etapas relevantes para escritura de resultados
        final int etapaPeriodoIni=NumeroEtapasAno*(Ano-AnoIni)+offset;//(tipoCalc==0?offset:144*(Ano-AnoIni)+offset);
        final int etapaPeriodoFin=NumeroEtapasAno*(Ano-AnoIni+1)+offset;//(tipoCalc==0?offset+144:144*(Ano-AnoIni+1)+offset);
        numEtapas=etapaPeriodoFin-etapaPeriodoIni;
        int cuenta;
        String [] TxtTemp1=new String[1000]; //almacenamiento temporal de texto
        Matriz Ybarra;	// Ybarra (n x n)
        Matriz Xbarra;	// Fila y columna de Slack se insertan con ceros (n x n)
        DecimalFormat DosDecimales=new DecimalFormat("0.00");
        long tInicioLectura = System.currentTimeMillis();
        cargandoInfo=true;
        String unicodeMessage = "Importando Informacion y Parametros";
        PrintStream out = new PrintStream(System.out, true, "UTF-8");
        out.println(unicodeMessage);
        String libroEntrada = DirBaseEntrada + SLASH + "Ent" + Ano + ".xlsx";
        String[] nombreMeses = {"Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul",
        "Ago", "Sep", "Oct", "Nov", "Dic"};
        String[] EnergiaCU={"CUE2","CUE30","EUnit"};
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(MAX_COMPRESSION_RATIO);

        /**************
         * lee de Meses
         **************/
        int numSp;
        int[] intAux1=new int[600];
        numSp = Lee.leeMeses(libroEntrada, intAux1, nombreMeses);
        int[] paramEtapa = new int[numEtapas];
        for (int i=0; i<numEtapas; i++)
            paramEtapa[i] = intAux1[i];

        /*
         * Lectura de parametros de lineas
         * ===============================
         */
        TxtTemp1=new String[2500];
        double[][] Aux=new double[2500][11];
        int numLin = Lee.leeDeflin(libroEntrada, TxtTemp1, Aux);
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
         * lee LÍneas Troncales
         * ====================
         */
        intAux1=new int[2500];
        int[][] intAux2 = new int[2500][3];
        TxtTemp1=new String[2500];
        String[] TxtTemp2=new String[2500];
        int numLinTron = Lee.leeLintron(libroEntrada, TxtTemp1,
                nombreLineas,TxtTemp2, intAux1, intAux2);
        String[] nomLinTron = new String[numLinTron];
        int[] indiceLintron = new int[numLinTron];
        int[][] datosLintron = new int[numLinTron][3];
        String[]zonaIT=new String[numLinTron];
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


        String[] TxtTemp3=new String[numLinTron];
         int[] TxtTemp5=new int[numLinTron];

        int numLinTx = 0;
        for (int i = 0; i <numLinTron; i++) {
             if(datosLintron[i][0]==1) zonaIT[i]="N";
             if(datosLintron[i][0]==0) zonaIT[i]="A";
             if(datosLintron[i][0]==-1) zonaIT[i]="S";
            int l = Calc.Buscar(nomLinTron[i] + "#" + nomProp[i], TxtTemp4);
            if (l == -1) {
                TxtTemp4[numLinTx] = nomLinTron[i] + "#" + nomProp[i];
                TxtTemp2[numLinTx]=nomProp[i];
                TxtTemp5[numLinTx]=datosLintron[i][0];
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
        String[] nomLinTx = new String[numLinTx];//solo registros ìnico LÍnea#Transmisor
        String[] nomPropTx = new String[numLinTx];
        String[] zona = new String[numLinTx];
        int[] datosLinIT = new int[numLinTx];
        for (int i = 0; i < numLinTx; i++) {
            nomLinTx[i] = TxtTemp4[i];
            nomPropTx[i]=TxtTemp2[i];
            zona[i] = TxtTemp3[i];
            datosLinIT[i]=TxtTemp5[i];
        }


        /*
         * lee de Barras
         * =============
         */
        TxtTemp1=new String[2500];
        int[][] intAux3=new int[2500][4];
        int numBarras = Lee.leeDefbar(libroEntrada, TxtTemp1, intAux3);
        String [] nomBar = new String[numBarras];
        int[][] paramBarTroncal = new int[numBarras][3];
        int numBarrasTroncales = 0;
        for(int i=0;i<numBarras;i++){
            nomBar[i] = TxtTemp1[i];
            // 1 si la barra es troncal, 0 en caso contrario
            paramBarTroncal[i][0] = intAux3[i][0];
            // 0 si la barra estˆ en el AIC, 1 si estˆ en el norte y -1 si estˆ en el sur
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
        float[][] Consumos = new float[numBarras][numEtapas];
        Lee.leeConsumoxBarra(libroEntrada,Consumos,numBarras,numEtapas);
        float[][] conNormalizado=new float[numBarras][numEtapas];
        boolean[][] barrasConsumo = new boolean[numBarras][numEtapas];
        float[] ConsEta = new float[numEtapas];
        System.out.println("Barras: "+numBarras);
        for(int b=0; b < numBarras; b++){
            for(int e=0;e<numEtapas;e++){
                barrasConsumo[b][e] = (Consumos[b][e]==0? false:true);
                ConsEta[e] += Consumos[b][e];
            }
        }
        for(int e=0; e < numEtapas; e++){
            for(int b=0; b < numBarras; b++){
            conNormalizado[b][e] = Consumos[b][e]/ConsEta[e];
            }
        }
        int[] duracionEta = new int[numEtapas];
        Lee.leeEtapas(libroEntrada,duracionEta,numEtapas);

        /*
         * Lectura de mantenimientos de lÍneas
         * ===================================
         */
        // cambios en condicion operativa para cada linea en cada etapa.
        int[][] LinMan = new int[numLin][numEtapas];
        for(int i=0; i < numLin; i++) {
            for(int j=0; j < numEtapas; j++) {
                LinMan[i][j] = -1;
            }
        }
        Lee.leeLinman(libroEntrada, LinMan, nombreLineas, numEtapas);

        /*******************
        Lectura de Centrales
        ********************/
        TxtTemp1 = new String[700];
        String [] TxtTemp1_2 = new String[700];
        float[] Temp1 = new float[700];
        float[] Temp2= new float[700];
        int numCen = Lee.leeCentrales(libroEntrada, TxtTemp1,Temp1,Temp2);
        String[] nombreCentrales = new String[numCen];
        System.arraycopy(TxtTemp1, 0, nombreCentrales, 0, numCen);

        /******************************
        Lectura de datos de generadores
        *******************************/
        TxtTemp1 = new String[1000];
        
        int numGen = Lee.leePlpcnfe(libroEntrada,TxtTemp1,
                intAux3,nombreCentrales);
        
        int numGen_Sin_Fallas = Lee.leePlpcnfe(libroEntrada,TxtTemp1_2,nombreCentrales);
        
        
        System.out.println("Generadores: "+numGen);
        int[][] paramGener = new int[numGen][3];
        String [] nomGen = new String[numGen];
        String [] nomGen_Sin_Fallas = new String[numGen_Sin_Fallas];
        boolean[] barrasGeneracion = new boolean[numBarras];
        for(int i=0; i < numBarras; i++){
            barrasGeneracion[i] = false;
        }
        for(int i=0; i < numGen; i++){
            nomGen[i] = TxtTemp1[i];
           // System.out.println("Peajes "+nomGen[i]);
            paramGener[i][1] = intAux3[i][1];
            paramGener[i][0] = intAux3[i][0];
            barrasGeneracion[paramGener[i][0]] = true;
        }
        for(int i=0; i < numGen_Sin_Fallas; i++){
            nomGen_Sin_Fallas[i] = TxtTemp1_2[i];
            
            
        }
            
            
            
        /*****************************************
        Lectura de orientaci„n de barras troncales
        ******************************************/
        int[][] orientBarTroncal=new int[numBarras][numLin];
        for(int i=0; i < numBarras; i++){
            for(int j=0; j < numLin; j++){
                orientBarTroncal[i][j]=0;
            }
        }
        Lee.leeOrient(libroEntrada, orientBarTroncal, nomBar,
                nombreLineas);

        /**************************
        Lectura de Suministradores
        **************************/
        TxtTemp1 = new String[100];
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

        //System.out.println(ActClientes);
        if(ActClientes=false){
        numClaves=Lee.leeConsumos2(libroEntrada, Temporal1,Temporal2,numEtapas
                ,paramEtapa,duracionEta,Temporal3);
        }
        else{
            numClaves=Lee.leeConsumos(libroEntrada, Temporal1,Temporal2,numEtapas
                ,paramEtapa,duracionEta,Temporal3);//debe ademˆs escribir la hoja con los clientes
        }

        float[][] ConsumosClaves = new float[numClaves][numEtapas];
        for(int i=0; i<numClaves;i++)            System.arraycopy(Temporal1[i], 0, ConsumosClaves[i], 0, numEtapas);      
        Temporal1 = null;
        
        float[][] ConsClaveMes = new float[numClaves][NUMERO_MESES];
        for(int i=0; i<numClaves;i++) System.arraycopy(Temporal2[i], 0, ConsClaveMes[i], 0, NUMERO_MESES);
        Temporal2 = null;
        
        float[][][] ECU= new float[numClaves][3][NUMERO_MESES];
        for(int i=0; i<numClaves;i++) System.arraycopy(Temporal3[i], 0, ECU[i], 0, 3);
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
        int numCli = Lee.leeClientes(libroEntrada, TxtTemp1,Exen);
        String [] nomCli = new String[numCli];
        System.arraycopy(TxtTemp1, 0, nomCli, 0, numCli);

        /*************************
        Lectura de Datos de Claves
        **************************/
        TxtTemp1 = new String[2500];
        TxtTemp2 = new String[2500];
        int clav = Lee.leeBarcli(libroEntrada, TxtTemp1,
                TxtTemp2, intAux3, nomCli, nomBar);
        String[] nombreClaves = new String[numClaves];
        String[] nombreClClientes = new String[numClaves];
        int[][] datosClaves = new int[numClaves][4];
        for(int i=0; i < numClaves; i++) {
            nombreClaves[i]=TxtTemp1[i];
            nombreClClientes[i]=TxtTemp2[i];
            datosClaves[i][0]=intAux3[i][0];
            datosClaves[i][1]=intAux3[i][1];
            datosClaves[i][2]=intAux3[i][2];
            datosClaves[i][3]=intAux3[i][3];
        }

        TxtTemp2 = null;
        //Escribe la energÍa consumida por Cliente 
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
         * Lectura de Generacion (Despachos) y energ’a no suministrada
         * ===========================================================
         */
        BufferedReader input = null;
        File testReadFile = new File(ArchivoDespachoGeneradores);
        input = null;
        float[][][] Gx = new float[numGen][numEtapas][numHid];
        float[][]FallaEtaHid = new float[numEtapas][numHid];
        try {
            //input = new BufferedReader( new FileReader(testReadFile) );
            input = new BufferedReader( new InputStreamReader(new
                    FileInputStream(testReadFile), "Latin1"));
            String line = null;
            cuenta=0;
            int indGen=0;
            int indGen2=0;
            int indHid=0;
            int indEta=0;
            float Pgen=0;
            float ENS=0;
            /*while ((line = input.readLine()) != null){
                if(cuenta>0){
                    if((line.substring(0,5).trim()).equals("MEDIA")==false){
                        indGen = Calc.Buscar((line.substring(32,line.indexOf(",",32))).trim(),nomGen);
                        //System.out.println("PLP "+(line.substring(32,line.indexOf(",",32))).trim());
                        if(indGen>-1){
                            //System.out.println(indGen +" "+nomGen[indGen]);
                            indHid=Integer.valueOf((line.substring(4,6)).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,11)).trim())-1;
                            Pgen=Float.valueOf((line.substring(103,110)).trim());
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                Gx[indGen][indEta-etapaPeriodoIni][indHid]=Pgen;
                                //System.out.println(Gx[indGen][indEta-etapaPeriodoIni][indHid]);
                            }
                        }
                        else{//suma las fallas
                            indHid=Integer.valueOf((line.substring(4,6)).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,11)).trim())-1;
                            ENS=Float.valueOf((line.substring(103,110)).trim());
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                FallaEtaHid[indEta-etapaPeriodoIni][indHid]+=ENS;
                            }
                        }
                    }
                }
                cuenta++;
            }*/
                    
            while ((line = input.readLine()) != null){
                if(cuenta>0){
                    if((line.substring(0,5).trim()).equals("MEDIA")==false){
                        
                        indGen = Calc.Buscar((line.substring(32,line.indexOf(",",32))).trim(),nomGen);
                        //System.out.println("PLP "+(line.substring(32,line.indexOf(",",32))).trim() + " " + indGen);
                        indGen2 = Calc.Buscar((line.substring(32,line.indexOf(",",32))).trim(),nomGen_Sin_Fallas);
                        
                        if(indGen2>-1){
                          if(indGen>-1){  
                            
                            //System.out.println(indGen +" "+nomGen[indGen]);
                            indHid=Integer.valueOf((line.substring(4,line.indexOf(",",4))).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,line.indexOf(",",8))).trim())-1;
                            //Pgen=Float.valueOf((line.substring(151,158)).trim()); //Peajes
                            Pgen=Float.valueOf((line.substring(103,line.indexOf(",",103))).trim()); 
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                //System.out.println(indGen +" "+nomGen[indGen]);
                                Gx[indGen][indEta-etapaPeriodoIni][indHid]=Pgen;
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
                                FallaEtaHid[indEta-etapaPeriodoIni][indHid]+=ENS;
                            }
                        }
                    }
                }
                cuenta++;
            }
        }
        catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        catch (IOException ex){
            ex.printStackTrace();
        }
        finally {
            try {
                if (input!= null) {
                    input.close();
                }
            }
            catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        
        /*
         * Lectura de datos de Lineas del sistema reducido
         * ===============================================
         */
        TxtTemp1=new String[600];
        float[][] Aux1 = new float[600][1];
        int numLinSistRed = Lee.leeLinPLP(libroEntrada, TxtTemp1, Aux1);
        int[] paramLinSistRed=new int[numLinSistRed];
        String [] nombreLineasSistRed = new String[numLinSistRed];
        for(int i=0;i<numLinSistRed;i++){
            nombreLineasSistRed[i]=TxtTemp1[i];
            //tensi„n lineas sistema reducido
            paramLinSistRed[i]=(int)(Math.round(Aux[i][0]));
        }

        TxtTemp1 = null;
        
        /*
         * Lectura de Perdidas en lineas de tension >110 kV
         * ================================================
         */
        testReadFile = new File(ArchivoPerdidasLineas);
        input = null;
        float[][] perdidasPLPMayor110 = new float[numEtapas][numHid];
        try {
            input = new BufferedReader( new InputStreamReader(new
                    FileInputStream(testReadFile), "Latin1"));
            String line = null;
            cuenta=0;
            int indLin=0;
            int indHid=0;
            int indEta=0;
            float Perd=0;
            while (( line = input.readLine()) != null){
                if(cuenta>0){
                    if((line.substring(0,5).trim()).equals("MEDIA")==false && (line.substring(0,3).trim()).equals("Sim")==true){
                        indLin=Calc.Buscar((line.substring(32,79)).trim(),nombreLineasSistRed);
                        if(indLin>-1){
                            //System.out.println(indLin+" "+nombreLineasSistRed[indLin]);
                            indHid=Integer.valueOf((line.substring(4,line.indexOf(",",4))).trim())-1;
                            indEta=Integer.valueOf((line.substring(8,line.indexOf(",",8))).trim())-1;
                            Perd=Float.valueOf(line.substring(111,line.indexOf(",",111)).trim());
                            if(indEta<etapaPeriodoFin && indEta>=etapaPeriodoIni){
                                if(paramLinSistRed[indLin]>110){
                                    perdidasPLPMayor110[indEta-etapaPeriodoIni][indHid]+=Perd;
                                }
                            }
                        }
                    }
                }
                cuenta++;
            }
        }
        catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        catch (IOException ex){
            ex.printStackTrace();
        }
        finally {
            try {
                if (input!= null) {
                    input.close();
                }
            }
            catch (IOException ex) {
                ex.printStackTrace();
            }
        }

        /*
         * Escalamiento de la demanda
         * ==========================
         */
        float[][][] conEscalado=new float[numBarras][numEtapas][numHid];
        float[][]GxEtaHid = new float[numEtapas][numHid];
        for(int h=0; h<numHid; h++){
            for(int e=0; e<numEtapas; e++){
                for(int g=0; g<numGen; g++){
                    GxEtaHid[e][h]+=Gx[g][e][h];
                    //System.out.println(GxEtaHid[e][h]);
                }
            }
        }
        for(int b=0;b<numBarras;b++){
            for(int e=0;e<numEtapas;e++){
                for(int h=0;h<numHid;h++){
                    conEscalado[b][e][h]=(float)(conNormalizado[b][e]*GxEtaHid[e][h]);
                }
            }
        }	    
        long tFinalLectura = System.currentTimeMillis();
        cargandoInfo=false;
        calculandoFlujos=true;
        long tInicioCalculo = System.currentTimeMillis();

        /*
         * Chequeo de consistencia
         * =======================
         */
        boolean[][] barrasActivas = new boolean[numBarras][numEtapas];
        barrasActivas = Calc.ChequeoConsistencia(paramLineas, LinMan,
                numBarras, numEtapas);

        /*
         * Iteraciones
         * ===========
         */
        int BarraSlack = Calc.Buscar(nombreSlack,nomBar);
        float[][] paramLinEta = new float[numLin][10];
        //
        float[][][] Flujo = new float[numLin][numEtapas][numHid];
        float[][] flujoDCEtapa = new float[numLin][numHid];
        float[] flujoDCHid = new float[numLin];
        //
        float[][][] prorrGx = new float[numLin][numGen][numEtapas];
        float[][] prorrEtaGx = new float[numLin][numGen];
        //
        float[][][] prorrCx = new float[numLin][numCli][numEtapas];
        float[][] prorrEtaCons = new float[numLin][numCli];
        //
        float[][] GGDFref = new float[numLin][numHid];
        float[][][] GGDFEtapa = new float[numBarras][numLin][numHid];
        float[][] GLDFref = new float[numLin][numHid];
        float[][][] GLDFEtapa = new float[numBarras][numLin][numHid];
        //
        System.out.println("Calculando...");
        //
        
        
        int[] centralesFlujo = Lee.leeCentralesFlujo(libroEntrada, nomGen,"centrales_flujo");
        int[] lineasFlujo = Lee.leeCentralesFlujo(libroEntrada, nombreLineas,"lineas_flujo");
        
        
        try
	{
            FileWriter writer = new FileWriter(DirBaseSalida + SLASH +"prorratas.csv");
           
            
            writer.append("Etapa");
            writer.append(',');
            writer.append("Hidrologia");
            writer.append(',');
            writer.append("Línea");
            writer.append(',');
            writer.append("Central");
            writer.append(',');
            writer.append("Prorrata");
            writer.append(',');
            writer.append("Gx");
            writer.append(',');
            writer.append("GGDF");
            writer.append('\n');
            
            writer.flush();
            writer.close();
        }
        
        catch(IOException e)
	{
	     e.printStackTrace();
	} 
        
        
        
        /***************************************************/
        /*INICIO ITERACIONES (PARALELIZAR EL SIGUIENTE FOR)*/
        /**********LAS ITERACIONES SON POR ETAPA************/
        /***************************************************/
        
       // try
        //{
            FileWriter writer2 = new FileWriter(DirBaseSalida + SLASH +"consumos.csv");
            writer2.append("Hidrologia,Etapa");
            for(int b=0;b<numBarras;b++){
                writer2.append(",");
                writer2.append(Float.toString(b));
            }
            writer2.append("\n");
       // }
        //catch(IOException e)
        
        //{
       //     e.printStackTrace();
       // } 
        
        for(etapa=0;etapa<numEtapas;etapa++) {
            System.out.println("etapa : "+etapa);
            for(int l=0; l < numLin; l++) {
                for(int i=0; i <= 5; i++)
                    paramLinEta[l][i] = paramLineas[l][i];
                if(LinMan[l][etapa] != (-1)){
                    paramLinEta[l][5] = LinMan[l][etapa];	//Operativa
                }
            }
            for(int l=0; l < numLinTron; l++) {
                paramLinEta[indiceLintron[l]][6] = 1;
                paramLinEta[indiceLintron[l]][7] = datosLintron[l][0];
                paramLinEta[indiceLintron[l]][8] = datosLintron[l][1];
                paramLinEta[indiceLintron[l]][9] = datosLintron[l][2];           
                
                
            }
            /*
             * Calcula matriz de Admitancias y matriz de Impedancias
             * =====================================================
             */
            int barrasEliminadas=0;
            // Calcula Ybarra considerando todas las barras, activas e inactivas
            Ybarra=new Matriz(Calc.CalculaYBarra(paramLinEta,numBarras,numLin));
            // Elimina de Ybarra las filas y columnas correspondientes a barras inactivas y la slack,
            // de manera de obtener una matriz invertible
            for(int b=0;b<numBarras;b++){
                if(barrasActivas[b][etapa]==false || b==BarraSlack){
                    Ybarra=(Ybarra.EliminarFila(b-barrasEliminadas)).EliminarColumna(b-barrasEliminadas);
                    barrasEliminadas++;
                }
            }
            Xbarra=(Ybarra.InversionRapida()).uminus();
            /* Se agregan las filas y columnas de las barras inactivas y la slack rellenas con ceros,
             * de manera de mantener coeherencia en los indices de barras
             */
            for(int b=0;b<numBarras;b++){
                if(barrasActivas[b][etapa]==false || b==BarraSlack){
                    Xbarra=(Xbarra.InsertarCerosFila(b)).InsertarCerosColumna(b);
                }
            }

            /* Calcula Factores de Desplazamiento A y
               GLDF barra referencia y GLDF resto del sistema. */
            float[][] GSDF = Calc.CalculaGSDF(Xbarra,paramLinEta,barrasActivas,etapa);
            GLDFref=Calc.CalculaGLDFRef(GSDF,paramLinEta,paramGener,etapa,Gx);
            GLDFEtapa=Calc.CalculaGLDF(GSDF,GLDFref,paramLinEta,etapa);

            // Calcula Flujo DC y asignacion de perdidas
            // -----------------------------------------
            float[] R=new float[numLin];                   // resistencias en p.u
            float[] perdI2R=new float[numLin];             // perdidas de cada l’nea segÏn I*I*R
            float[] perdidas=new float[numLin];            // perdidas de cada l’nea segun diferencia entre Gx y Demanda
            float[] perdMayor110=new float[numLin];        // perdidas de cada l’nea segun diferencia entre Gx y Demanda
            float[] perdMenor110=new float[numLin];        // perdidas de cada l’nea segun diferencia entre Gx y Demanda
            float[] perdRealesSistema=new float[numHid];   // perdidas del sistema
            float[] perdI2RSistMayor110=new float[numHid]; // perdidas de todas las l’neas de tensi—n > 110kV
            float[] perdI2RSistMenor110=new float[numHid]; // perdidas de todas las l’neas de tensi—n <= 110kV
            float conSist;
            float[][] conAjustado=new float[numBarras][numHid];
            float[] genSist=new float[numHid];
            float[] conMasPerd= new float[numBarras];      // consumos con asignacion de p?rdidas por iteraci—n [MW]
            // Consumos con asignacion de p?rdidas por iteracion [MW]
            float[][] conMasPerdEta= new float[numBarras][numHid];
            for(int h=0;h<numHid;h++){
                genSist[h] = GxEtaHid[etapa][h];
                for(int b=0;b>numBarras;b++)
                    conAjustado[b][h] = 0;
            }
            for(int l=0;l<numLin;l++)
                R[l]=paramLinEta[l][3];                    // resistencia en p.u.
            for(int h=0;h<numHid;h++){
                
                try
                        {
                        writer2.append(Float.toString(h));
                        writer2.append(",");
                        writer2.append(Float.toString(etapa));
                    
                    }
                    catch(IOException e)
        
                    {
                        e.printStackTrace();
                    } 
                
                
                for(int l=0;l<numLin;l++){
                    perdI2R[l]=0;
                    perdidas[l]=0;
                    flujoDCHid[l]=0;
                }
                perdRealesSistema[h]=0;
                perdI2RSistMayor110[h]=0;
                perdI2RSistMenor110[h]=0;
                conSist=0;
                for(int b=0;b<numBarras;b++)
                    conSist += Consumos[b][etapa];
                for(int b=0;b<numBarras;b++){
                    conAjustado[b][h] += Consumos[b][etapa]*(conSist-FallaEtaHid[etapa][h])/conSist;
                    try
                        {
                    
                        writer2.append(",");
                        writer2.append(Float.toString(conAjustado[b][h]));
                    }
                    catch(IOException e)
        
                    {
                        e.printStackTrace();
                    } 
                    
                }
                
                writer2.append("\n");
                
                
                
                
            
           
       
        
        
        
                
                
                
                perdRealesSistema[h] =
                        genSist[h]-(conSist-FallaEtaHid[etapa][h]);
                // Calculo de Flujo DC
                flujoDCHid=Calc.FlujoDC_GLDF(GLDFEtapa,conAjustado,h,etapa);//flujos en MW
                //System.out.println("Flujo DC "+flujoDCHid[586]);
                for(int l=0;l<numLin;l++)
                    flujoDCEtapa[l][h]=flujoDCHid[l];
                // Calcula perdidas
                for(int l=0;l<numLin;l++) {
                    if (flujoDCHid[l]!=0) {
                        float sBase = 100;
                        perdI2R[l]=sBase*(R[l]*(flujoDCHid[l]/sBase)*(flujoDCHid[l]/sBase));	//perdidas en MW
                        //System.out.println("Perdidas cuadraticas "+ perdI2R[l]);
                    }
                    if(paramLinEta[l][2]>110)
                        perdI2RSistMayor110[h]+=perdI2R[l];
                    else
                        perdI2RSistMenor110[h]+=perdI2R[l];
                }
                // Perdidas Reales prorrateadas en las lineas de acuerdo al I2R de cada una
                perdMayor110=Calc.ProrrPerdidas(perdidasPLPMayor110[etapa][h],perdI2R,paramLinEta,"Mayor_110",h);
                perdMenor110=Calc.ProrrPerdidas((perdRealesSistema[h]-perdidasPLPMayor110[etapa][h]),
                                                        perdI2R,paramLinEta,"Menor_Igual_110",h);
                for(int l=0;l<numLin;l++)
                    perdidas[l]=perdMayor110[l]+perdMenor110[l];
                // Asigna perdidas a consumos
                conMasPerd=Calc.AsignaPerdidas(flujoDCHid, GLDFEtapa,
                        perdidas, paramLinEta, conAjustado, etapa, h);
            }
            for(int h=0; h<numHid; h++){
                for(int l=0; l<numLin; l++)
                    Flujo[l][etapa][h] = flujoDCEtapa[l][h];
                for(int b=0; b<numBarras; b++)
                    conMasPerdEta[b][h] = conMasPerd[b];
            }
            /*
             * Calcula GGDF barra referencia y GGDF resto del sistema.
             */
            GGDFref=Calc.CalculaGGDFRef(GSDF,conMasPerdEta, paramLinEta);
            GGDFEtapa=Calc.CalculaGGDF(GSDF,GGDFref,paramLinEta,etapa);
            //System.out.println("GGDF "+ GGDFEtapa[31][603][1]);
            /*
             * Calcula prorratas promedio por etapa
             */
            prorrEtaGx=Calc.CalculaProrrGx(flujoDCEtapa, GGDFEtapa, Gx, paramGener, paramLinEta, paramBarTroncal,
                    orientBarTroncal, etapa, centralesFlujo, lineasFlujo,DirBaseSalida,GSDF,GGDFref );
            
            prorrEtaCons=Calc.CalculaProrrCons(flujoDCEtapa, GLDFEtapa,
                    ConsumosClaves, datosClaves, paramLinEta,
                    paramBarTroncal, orientBarTroncal, etapa);
            for(int l=0;l<numLin;l++) {
                for(int g=0;g<numGen;g++)
                    prorrGx[l][g][etapa] = prorrEtaGx[l][g];
                for(int c=0;c<numClaves;c++) {
                    prorrCx[l][datosClaves[c][2]][etapa] += prorrEtaCons[l][c];
                }
            }
        }   
        
           
            writer2.flush();
            writer2.close();
        
        /**********************************/
        /******** FIN ITERACIONES *********/
        /**********************************/
        
        
        
        long tfinIteraciones = System.currentTimeMillis();
        perdidasPLPMayor110 = null;
        GGDFref = null;
        GGDFEtapa = null;
        GLDFref = null;
        GLDFEtapa = null;

        calculandoFlujos=false;
        calculandoProrr=true;

        /*
         * Calcula prorratas promedio anuales y mensuales
         * =================================================
         */
        int etapasPeriodo = etapaPeriodoFin-etapaPeriodoIni;
        float[][] FlujoMedio = new float[numLin][numEtapas];
        float ConsAnoEner = 0;
        float[] ConsMesEner = new float[NUMERO_MESES];
        int mes=0;
        for(int e=0;e<etapasPeriodo;e++){
            mes=(int)Math.floor((double)e/(etapasPeriodo/NUMERO_MESES));
            ConsAnoEner += ConsEta[e]*(float)duracionEta[e];
            ConsMesEner[mes] += ConsEta[e]*(float)duracionEta[e];
        }
        for(int e=0;e<numEtapas;e++){
            for(int l=0;l<numLin;l++){
                for(int h=0;h<numHid;h++){
                    FlujoMedio[l][e]+=Flujo[l][e][h]/numHid;
                    
                }
            }
        }
        
        
        
        
        try
	{
            FileWriter writer = new FileWriter(DirBaseSalida + SLASH +"flujos_hidrologia.csv");
            
            writer.append("Hidrologia,Etapa");
	    
            
            for(int l=0;l<lineasFlujo.length;l++){
                    writer.append(',');
                    writer.append(nombreLineas[lineasFlujo[l]]);
                                                        
                }
            writer.append('\n');   
            
            for (int hh = 0 ; hh < numHid ; hh++) {  
                for(int e=0;e<numEtapas;e++){
                    writer.append(String.valueOf(hh));
                    writer.append(',');
                    writer.append(String.valueOf(e));
                    for(int l=0;l<lineasFlujo.length;l++){
                        writer.append(',');
                        writer.append(Float.toString(Flujo[lineasFlujo[l]][e][hh]));
                                                          
                    }
                    writer.append('\n');
                }
        }
            writer.flush();
            writer.close();
        }
        
        catch(IOException e)
	{
	     e.printStackTrace();
	} 
        
         try
	{
            FileWriter writer = new FileWriter(DirBaseSalida + SLASH +"flujos_medios.csv");
            
            writer.append("Etapa");
	    
            
            for(int l=0;l<lineasFlujo.length;l++){
                    writer.append(',');
                    writer.append(nombreLineas[lineasFlujo[l]]);
                                                        
                }
            writer.append('\n');   
            
            
                for(int e=0;e<numEtapas;e++){

                    writer.append(String.valueOf(e));
                    for(int l=0;l<lineasFlujo.length;l++){
                        writer.append(',');
                        writer.append(Float.toString(FlujoMedio[lineasFlujo[l]][e]));
                                                          
                    }
                    writer.append('\n');
                }
        
            writer.flush();
            writer.close();
        }
        
        catch(IOException e)
	{
	     e.printStackTrace();
	} 
        

        
        double[][] prorrAnoG=new double[numLin][numGen];
        double[][][] prorrMesG=new double[numLin][numGen][NUMERO_MESES];
        double[][] prorrAnoC=new double[numLin][numCli];
        double[][][] prorrMesC=new double[numLin][numCli][NUMERO_MESES];
        // Calcula para todas las lineas
        for(int l=0;l<numLin;l++){
            for(int e=0;e<etapasPeriodo;e++){
                mes=(int)Math.floor((double)e/(NumeroEtapasAno/NUMERO_MESES));
                for(int g=0;g<numGen;g++){
                    //System.out.println(prorrGx[l][g][e]);
                    prorrAnoG[l][g] += prorrGx[l][g][e] * ( ConsEta[e] * duracionEta[e] / ConsAnoEner );
                    prorrMesG[l][g][mes] += prorrGx[l][g][e]*( ConsEta[e] * duracionEta[e] / ConsMesEner[mes] );
                }
                for(int c=0;c<numCli;c++){
                    prorrAnoC[l][c]+=
                            prorrCx[l][c][e]*(ConsEta[e]
                            *duracionEta[e]/ConsAnoEner);
                    prorrMesC[l][c][mes]+=
                            prorrCx[l][c][e]*(ConsEta[e]
                            *duracionEta[e]/ConsMesEner[mes]);
                }
            }
        }
        // Filtra lineas troncales
        double[][] prorrAnoTroncG = new double[numLinTx][numGen];
        double[][] prorrAnoTroncC = new double[numLinTx][numCli];
        double[][][] prorrMesTroncG=new double[numLinTx][numGen][NUMERO_MESES];
        double[][][] prorrMesTroncC=new double[numLinTx][numCli][NUMERO_MESES];
        double[][] ProrrVerMesLinG = new double[numLinTron][NUMERO_MESES];
        double[][] ProrrVerMesLinC = new double[numLinTron][NUMERO_MESES];
        for(int l=0; l<numLinTron; l++){
            int l2= Calc.Buscar(LinTronProp[l],nomLinTx);
            //System.out.println(l+" "+LinTronProp[l]+" "+nomLinTx[l2]+" "+l2);
            for(int g=0; g<numGen; g++){
                prorrAnoTroncG[l2][g] += prorrAnoG[indiceLintron[l]][g];
                for(int m=0; m<NUMERO_MESES; m++){
                    prorrMesTroncG[l2][g][m] += prorrMesG[indiceLintron[l]][g][m];
                ProrrVerMesLinG[l][m]+=prorrMesG[indiceLintron[l]][g][m];
                }
            }
            for(int c=0; c<numCli; c++){
                prorrAnoTroncC[l2][c] += prorrAnoC[indiceLintron[l]][c];
                for(int m=0; m<NUMERO_MESES; m++){
                    prorrMesTroncC[l2][c][m] += prorrMesC[indiceLintron[l]][c][m];
                    ProrrVerMesLinC[l][m]+=prorrMesC[indiceLintron[l]][c][m];
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
        for (int l=0; l<numLinTx; l++) {
            for(int m=0; m<NUMERO_MESES; m++) {
                if (datosLinIT[l] == 0) {
                    double FdenG = sumProrrMesLinG[l][m];
                    if (Math.round(1000000000*FdenG) == 0)
                        FdenG = 1.0;
                    FactorG[l][m] = 0.8/FdenG;
                    double FdenC = sumProrrMesLinC[l][m];
                    if (Math.round(1000000000*FdenC) == 0)
                        FdenC = 1.0;
                    FactorC[l][m] = 0.2/FdenC;
                }
                else {
                    double Fden = prorrataLinea[l][m];
                    if (Math.round(1000000000*Fden) == 0)
                        Fden = 1.0;
                    FactorG[l][m] = 1/Fden;
                    FactorC[l][m] = 1/Fden;
                }
            }
        }
        // Procesa salida prorratas de generaci„n
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
        int l1=0;
        for (int l=0; l<numLinTx; l++) {
            //l1 = Calc.Buscar(nomLinTx[l].split("#")[0],nomLinTron);
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
        for(int l=0; l<numLinTx; l++){
            //l1 = Calc.Buscar(nomLinTx[l].split("#")[0],nomLinTron);
            for(int c=0; c<numCli; c++){
                for(int m=0; m<NUMERO_MESES; m++){
                prorrMesLinC[l][c][m]
                        += prorrMesTroncC[l][c][m]*FactorC[l][m];
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
        String libroSalidaXLS = DirBaseSalida + SLASH + "Prorrata" + Ano + ".xlsx";
        Escribe.crearLibro(libroSalidaXLS);
        Escribe.creaH3F_3d_double(
                "Prorratas de Generación", prorrMesLinG,
                "Línea", nomLinTx,
                "Central", nombreCentrales,
                "Zona",zona,
                "Mes", nombreMeses,
                libroSalidaXLS,"ProrrGMes","0.000%;[Red]-0.000%;\"-\"");
        Escribe.creaH3F_3d_double(
                "Prorratas de Consumo", prorrMesLinC,
                "Línea",nomLinTx,
                "Cliente",nomCli,
                "Zona",zona,
                "Mes", nombreMeses,
                libroSalidaXLS,"ProrrCMes","0.000%;[Red]-0.000%;\"-\"");
        Escribe.creaH1F_2d_double(
                "Prorratas por L’nea", prorrataLinea,
                "Línea", nomLinTx,
                "Mes", nombreMeses,
                libroSalidaXLS, "ProrrLin","0.000%;[Red]-0.000%;\"-\"");
        Escribe.creaH1F_2d_double(
                "Generación [GWh]", generacionMes,
                "Central", nombreCentrales,
                "Mes", nombreMeses,
                libroSalidaXLS, "GMes","0.0;[Red]-0.0;\"-\"");
        Escribe.creaH1F_2d_float(
                "Consumo [MWh]",ConsClaveMes,
                "Cliente", nombreClaves,
                "Mes", nombreMeses,
                libroSalidaXLS, "CMes","0.0;[Red]-0.0;\"-\"");
        Escribe.creaH1F_2d_double(
                "Detalle de prorratas de Generación", Calc.transponer(prorrAnoTroncG),
                "Central", nomGen,
                "L’nea", nomLinTx,
                libroSalidaXLS, "ProrrG","0.000%");
        Escribe.creaH1F_2d_double(
                "Detalle de prorratas de Consumo", Calc.transponer(prorrAnoTroncC),
                "Clave", nomCli,
                "Linea", nomLinTx,
                libroSalidaXLS, "ProrrC","0.000%");
        Escribe.creaH1FT_2d_float(
                "Consumo [MWh]", CMes, ECUCli,
                "Cliente", nomCli,
                "Mes", nombreMeses,EnergiaCU,"CU",
                libroSalidaXLS, "CMesCli","0.0;[Red]-0.0;\"-\"");
         Escribe.crea_verifProrrPeaj(prorrataLineaTron,
                 nomLinTron,
                libroEntrada, "verProrr","0.000%;[Red]-0.000%;\"-\"",12);



        guardandoDatos=false;
        long tFinalEscritura = System.currentTimeMillis();
        System.out.println("Tiempo Adquisición de datos     : "+DosDecimales.format((tFinalLectura-tInicioLectura)/1000.0)+" s");
        System.out.println("Tiempo Calculos                 : "+DosDecimales.format((tFinalCalculo-tInicioCalculo)/1000.0)+" s");
        System.out.println("Tiempo Iteraciones              : "+DosDecimales.format((tfinIteraciones-tInicioCalculo)/1000.0)+" s");
        System.out.println("Tiempo Escritura de Resultados  : "+DosDecimales.format((tFinalEscritura-tInicioEscritura)/1000.0)+" s");
        System.out.println("Tiempo total                    : "+DosDecimales.format((tFinalEscritura-tInicioLectura)/1000.0)+" s");

        completo=true;
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
                }
                return true;
            }
        };
        worker.execute();

    }
}



    
