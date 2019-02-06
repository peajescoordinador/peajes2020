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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import java.util.Properties;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author aramos
 */
public class Lee {

    public static int leeClientes(String libroEntrada, String[] TextoTemporal1, String[] Exento) {
        int numClientes = 0;
        Cell c2 = null;
        Cell c3 = null;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("clientes");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=2) {
                Row r = s.getRow(crefs[i].getRow());
                c2 = r.getCell(crefs[i].getCol());
                c3 = r.getCell(crefs[i+1].getCol());//Ajuste
                Exento[numClientes]=c3.toString().trim();//Ajuste
                TextoTemporal1[numClientes] = c2.toString().trim();// Nombre
                numClientes++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numClientes;
    }

    public static int leeCentrales(String libroEntrada, String[] TextoTemporal,float[] Potencia, float[] MedioGene) {
        int numCentrales = 0;
        Cell c1 = null;
        Cell c2 = null;
        Cell c3 = null;
        Cell c4 = null;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("centrales");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=4) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                c2 = r.getCell(crefs[i+1].getCol());
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                Potencia[numCentrales]= (float) c3.getNumericCellValue();
                MedioGene[numCentrales]= (float) c4.getNumericCellValue();
                TextoTemporal[numCentrales] = c2.toString().trim()+"#"+c1.toString().trim(); // Nombre
                
                //System.out.println( TextoTemporal[numCentrales]);
                
                numCentrales++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numCentrales;
    }

 public static int leeVATT(String libroEntrada, String[] TextoTemporal1, String[] TextoTemporal2, double[][] VATT) {
        int numLineasPeajes = 0;
        try {
            Cell c1 = null;
            int nomRangoInd;
            Name nomRango;
            AreaReference aref;
            CellReference[] crefs;
            Sheet s;
            ////POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));

            // Lectura de datos
            nomRangoInd = wb.getNameIndex("lineasVATT");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i++) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                TextoTemporal1[i] = c1.toString().trim();
            }
            // Lectura de datos
            nomRangoInd = wb.getNameIndex("transmisores");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i++) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                TextoTemporal2[i] = c1.toString().trim();
            }
            // Lectura de datos
            nomRangoInd = wb.getNameIndex("VATT");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=NUMERO_MESES) {
                Row r = s.getRow(crefs[i].getRow());
                for (int j=0; j<NUMERO_MESES; j++) {
                    c1 = r.getCell(crefs[i + j].getCol());
                    VATT[numLineasPeajes][j] = c1.getNumericCellValue();
                }
                numLineasPeajes++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineasPeajes;
    }
  public static void leeEscribeArchivoVATT(String libroEntrada,String libroAVICOMA,int Ano) {

        try {
            Sheet s;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroAVICOMA ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroAVICOMA ));
            s=wb.getSheet("VATT");
            // Lectura de datos
            int NT=s.getLastRowNum()-3;
            String[] Tramo=new String[NT];
            String[] Prop=new String[NT];
            //String[] Comen=new String[NT];
            double[][] VATT=new double[NT][NUMERO_MESES];

            int aux=17+12*(Ano-2005);
            //System.out.println(NT);

            // Lectura de datos
            for (int i=0; i<NT; i++) {
                Tramo[i]= s.getRow(i+1).getCell(1).toString();
                Prop[i]=s.getRow(i+1).getCell(2).toString();
                for(int k=0;k<NUMERO_MESES;k++){
                    VATT[i][k]=s.getRow(i+1).getCell(k+aux).getNumericCellValue();
                }
            }
            s=wb.getSheet("DatosIndex");
            for (int i=0; i<NT; i++) {
                //Comen[i]= s.getRow(i+1).getCell(13).toString();
            }
             //Escribe lo clientes en hoja clientes
            //POIFSFileSystem fsEnt = new //POIFSFileSystem(new FileInputStream(libroEntrada));
            XSSFWorkbook wbEnt = new XSSFWorkbook(new FileInputStream(libroEntrada));
            s=wbEnt.getSheet("VATT");
            String nomhoja=s.getSheetName();

         Cell cellTx = null;
         Cell cell= null;
         Cell cellTram = null;
         Cell cellDat = null;
         Row row=null;
         short fila = 5;

         Font font = wbEnt.createFont();
            font.setFontHeightInPoints((short)10);
            font.setFontName("Century Gothic");
            CellStyle estilo = wbEnt.createCellStyle();
            estilo.setFont(font);

            CellStyle estilo1 = wbEnt.createCellStyle();
            estilo1.setFont(font);
            estilo1.setBorderRight(BorderStyle.THIN);
            estilo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

            CellStyle estilo2 = wbEnt.createCellStyle();
            estilo2.setFont(font);
            estilo2.setBorderLeft(BorderStyle.THIN);
            estilo2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

         // Datos
            for(int i=0;i<NT;i++){
             row = s.createRow(fila); fila++;
             cellTram = row.createCell(1);
             cellTram.setCellValue(Tramo[i]);
             cellTram.setCellStyle(estilo);
             cellTx = row.createCell(2);
             cellTx.setCellValue(Prop[i]);
             cellTx.setCellStyle(estilo1);
             for(int k=0;k<NUMERO_MESES;k++){
                 cellDat = row.createCell(k+3);
                 cellDat.setCellValue(VATT[i][k]/12);
                 cellDat.setCellStyle(estilo);
             }
            /* cell = row.createCell(15);
             cell.setCellValue(Comen[i]);
             cell.setCellStyle(estilo2);
             * 
             */
            }
            // Crea nombre de rango de salida         
            wbEnt.removeName("lineasVATT");
            Name nombreTramo = wbEnt.createName();
            nombreTramo.setNameName("lineasVATT"); // Nombre del rango
            CellReference cellRef2 = new CellReference(cellTram.getRowIndex(), cellTram.getColumnIndex());
            String reference2 = "VATT"+"!B6:"+cellRef2.formatAsString(); // area reference
            nombreTramo.setRefersToFormula(reference2);
            
            wbEnt.removeName("transmisores");
            Name nombreTx = wbEnt.createName();
            nombreTx.setNameName("transmisores"); // Nombre del rango
            CellReference cellRef1 = new CellReference(cellTx.getRowIndex(), cellTx.getColumnIndex());
            String reference1 = "VATT"+"!C6:"+cellRef1.formatAsString(); // area reference
            nombreTx.setRefersToFormula(reference1);

            wbEnt.removeName("VATT");
            Name nombreDatos = wbEnt.createName();
            nombreDatos.setNameName("VATT"); // Nombre del rango
            CellReference cellRef = new CellReference(cellDat.getRowIndex(), cellDat.getColumnIndex());
            String reference = "VATT"+"!D6:"+cellRef.formatAsString(); // area reference
            nombreDatos.setRefersToFormula(reference);

            FileOutputStream archivoSalida = new FileOutputStream( libroEntrada );
            wbEnt.write(archivoSalida);
            archivoSalida.close();
            System.out.println( "Acaba de extraer los VATT y asignarlos en la planilla Ent"+Ano+".xlsx" );
            
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }

    }
   public static void leeEscribeIndices(String libroEntrada,String libroAVICOMA,int Ano) {
       Cell cell=null;
        try {
            Sheet s;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroAVICOMA ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroAVICOMA ));
            s=wb.getSheet("Indices");
            // Lectura de datos
            int UltimoDato=0;
            int FilaAno=5+(Ano-2004)*12;
            double[] dolar=new double[NUMERO_MESES];
            boolean fin=false;

            for (int i=0; i<NUMERO_MESES; i++) {

                if(fin==false){
                    cell=s.getRow(FilaAno+i).getCell(2);
                    //System.out.println(cell);
                    if(cell==null){
                        dolar[i]=0;
                        UltimoDato=i;
                        fin=true;
                    }
                    else
                    dolar[i]=cell.getNumericCellValue();
                }
                else
                dolar[i]=0;
            }
             //Escribe los datos en libro de entrada
            //POIFSFileSystem fsEnt = new //POIFSFileSystem(new FileInputStream(libroEntrada));
            XSSFWorkbook wbEnt = new XSSFWorkbook(new FileInputStream(libroEntrada));
            Sheet hoja=wbEnt.getSheet("indices");
            String nhoja=hoja.getSheetName();

         Cell cellDat = null;
         Row row=null;
         short fila = 4;

         Font font = wbEnt.createFont();
            font.setFontHeightInPoints((short)10);
            font.setFontName("Century Gothic");
            CellStyle estilo = wbEnt.createCellStyle();
            estilo.setFont(font);

            CellStyle estilo1 = wbEnt.createCellStyle();
            estilo1.setFont(font);
            estilo1.setBorderRight(BorderStyle.THIN);
            estilo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

            CellStyle estilo2 = wbEnt.createCellStyle();
            estilo2.setFont(font);
            estilo2.setBorderLeft(BorderStyle.THIN);
            estilo2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

         // Datos
            for(int i=0;i<NUMERO_MESES;i++){
             row = hoja.getRow(fila); fila++;
                 cellDat = row.createCell(2);
                 if(dolar[i]==0){
                 cellDat.setCellValue(dolar[UltimoDato-1]);
                 cellDat.setCellStyle(estilo);
                 }
                 else{
                 cellDat.setCellValue( dolar[i]);
                 cellDat.setCellStyle(estilo);
                 }
            }
            // Crea nombre de rango de salida
            wbEnt.removeName("dolar");
            Name nombreDatos = wbEnt.createName();
            nombreDatos.setNameName("dolar"); // Nombre del rango
            CellReference cellRef = new CellReference(cellDat.getRowIndex(), cellDat.getColumnIndex());
            String reference = nhoja+"!C5:"+cellRef.formatAsString(); // area reference
            nombreDatos.setRefersToFormula(reference);

            FileOutputStream archivoSalida = new FileOutputStream( libroEntrada );
            wbEnt.write(archivoSalida);
            archivoSalida.close();
            System.out.println( "Acaba de extraer el Dolar y asignarlo en la planilla Ent"+Ano+".xlsx" );

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }

    }

    public static int leeDeflin(String libroEntrada, String[] TextoTemporal1, double[][] Aux) {
        int numLineas = 0;
        double zBase;
        double sBase = 100;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("deflin");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=12) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c2 = null;
                Cell c3 = null;
                Cell c4 = null;
                Cell c5 = null;
                Cell c6 = null;
                Cell c7 = null;
                Cell c8 = null;
                Cell c9 = null;
                Cell c10 = null;
                Cell c11 = null;
                Cell c12= null;
                c2 = r.getCell(crefs[i+1].getCol());
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                c5 = r.getCell(crefs[i+4].getCol());
                c6 = r.getCell(crefs[i+5].getCol());
                c7 = r.getCell(crefs[i+6].getCol());
                c8 = r.getCell(crefs[i+7].getCol());
                c9 = r.getCell(crefs[i+8].getCol());
                c10 = r.getCell(crefs[i+9].getCol());
                c11 = r.getCell(crefs[i+10].getCol());
                c12 = r.getCell(crefs[i+11].getCol());
                TextoTemporal1[numLineas]=c2.toString().trim(); //Nombre
                Aux[numLineas][0] = (int) c3.getNumericCellValue() - 1; // Barra_A
                Aux[numLineas][1] = (int) c4.getNumericCellValue() - 1; // Barra_B
                Aux[numLineas][2] = c5.getNumericCellValue(); // V_[kV]
                zBase=Aux[numLineas][2]*Aux[numLineas][2]/sBase;
                Aux[numLineas][3] = c6.getNumericCellValue()/zBase; // R_[ohm]/zBase
                Aux[numLineas][4] = c7.getNumericCellValue()/zBase; // X_[ohm]/zBase
                Aux[numLineas][5] = (int) c8.getNumericCellValue(); //Operativa
                Aux[numLineas][6] = (int) c9.getNumericCellValue(); //Troncal
                Aux[numLineas][7] = (int) c10.getNumericCellValue(); //Zona
                Aux[numLineas][8] = (int) c11.getNumericCellValue(); //dir
                Aux[numLineas][9] = (int) c12.getNumericCellValue(); //Area
                numLineas++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineas;
    }

    public static int leeLintron(String libroEntrada, String[] TextoTemporal, String[] nombreLineas,String[] nomTx, int[] intAux1, int[][] intAux2) {
        int numLineasT = 0;
        int numLineasT2 = 0;
        int aux;
                Cell c1 = null;
                Cell c2 = null;
                Cell c3 = null;
                Cell c4 = null;
                Cell c5 = null;
                Cell c6 = null;
                
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            AreaReference arefDatos;
            CellReference[] crefsDatos;
            AreaReference arefTx;
            CellReference[] crefsTx;
            int nomRangoInd = wb.getNameIndex("lintron");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            int nomRangoIndDatos = wb.getNameIndex("datosLintron");
            Name nomRangoDatos = wb.getNameAt(nomRangoIndDatos);
            arefDatos = new AreaReference(nomRangoDatos.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsDatos = arefDatos.getAllReferencedCells();
            int nomRangoindDatosTx = wb.getNameIndex("transmisorIT");
            Name nomRangoDatosTx = wb.getNameAt(nomRangoindDatosTx);
            arefTx = new AreaReference(nomRangoDatosTx.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsTx = arefTx.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=2) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                c2 = r.getCell(crefs[i+1].getCol());
                
                TextoTemporal[numLineasT] = c1.toString().trim();
                aux=Calc.Buscar(c2.toString().trim(), nombreLineas);
                if(aux==-1) {
                    System.out.println("La línea "+c2.toString().trim()+" en 'lintron' no se encuentra en 'lineas'");
                }
                intAux1[numLineasT] = aux;

                numLineasT++;
            }
            
            
            Sheet m = wb.getSheet(crefsDatos[0].getSheetName());
            for (int i=0; i<crefsDatos.length; i+=3) {
                Row r = m.getRow(crefsDatos[i].getRow());

                c3 = r.getCell(crefsDatos[i].getCol());
                c4 = r.getCell(crefsDatos[i+1].getCol());
                c5 = r.getCell(crefsDatos[i+2].getCol()); 
        
                intAux2[numLineasT2][0] = (int) c3.getNumericCellValue();
                intAux2[numLineasT2][1] = (int) c4.getNumericCellValue();
                intAux2[numLineasT2][2] = (int) c5.getNumericCellValue();
                numLineasT2++;
            }
            
            
            
            
            
            
            
            
            for (int i=0; i<crefsTx.length; i++) {
                Row r = s.getRow(crefsTx[i].getRow());
                c1=r.getCell(crefsTx[i].getCol());
                nomTx[i]=c1.toString().trim();
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineasT;
    }

    public static int leeLintronIT(String libroEntrada, String[] TextoTemporal,
            String[] TextoTemporal2, String[] nombreLineas, int[] intAux1,
            double[][] ITE, double[][] ITP) {
        int numLineasT = 0;
        int iTemp1;
        String TextoTmp;
        double ITEAux;
        double ITPAux;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            AreaReference arefTx;
            CellReference[] crefsTx;
            AreaReference arefITE;
            CellReference[] crefsITE;
            AreaReference arefITP;
            CellReference[] crefsITP;
            int nomRangoInd = wb.getNameIndex("lintron");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            int nomRangoIndTx = wb.getNameIndex("transmisorIT");
            Name nomRangoTx = wb.getNameAt(nomRangoIndTx);
            arefTx = new AreaReference(nomRangoTx.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsTx = arefTx.getAllReferencedCells();
            int nomRangoIndITE = wb.getNameIndex("ITE");
            Name nomRangoITE = wb.getNameAt(nomRangoIndITE);
            arefITE = new AreaReference(nomRangoITE.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITE = arefITE.getAllReferencedCells();
            int nomRangoIndITP = wb.getNameIndex("ITP");
            Name nomRangoITP = wb.getNameAt(nomRangoIndITP);
            arefITP = new AreaReference(nomRangoITP.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITP = arefITP.getAllReferencedCells();
            for (int i=0; i<crefs.length; i+=2) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c1 = null;
                Cell c2 = null;
                c1 = r.getCell(crefs[i].getCol());
                c2 = r.getCell(crefs[i+1].getCol());
                TextoTemporal[numLineasT] = c1.toString().trim();
                intAux1[numLineasT] = Calc.Buscar(c2.toString().trim(), nombreLineas);
                Cell cTx = r.getCell(crefsTx[0].getCol());
                TextoTemporal2[numLineasT] = c1.toString().trim()+"#"+cTx.toString().trim();
                for (int m=0; m<NUMERO_MESES; m++) {
                    Cell cITE = r.getCell(crefsITE[m].getCol());
                    ITE[numLineasT][m] = (float) cITE.getNumericCellValue();
                    Cell cITP = r.getCell(crefsITP[m].getCol());
                    ITP[numLineasT][m] = (float) cITP.getNumericCellValue();
                }
                if (numLineasT > 0) {
                    for (int k=numLineasT; k>0; k--) {
                        if (intAux1[k] < intAux1[k - 1]) {
                            TextoTmp = TextoTemporal[k];
                            TextoTemporal[k] = TextoTemporal[k - 1];
                            TextoTemporal[k - 1] = TextoTmp;
                            TextoTmp = TextoTemporal2[k];
                            TextoTemporal2[k] = TextoTemporal2[k - 1];
                            TextoTemporal2[k - 1] = TextoTmp;
                            for (int m=0; m<NUMERO_MESES; m++) {
                                ITEAux = ITE[k][m];
                                ITPAux = ITP[k][m];
                                ITE[k][m] = ITE[k - 1][m];
                                ITP[k][m] = ITP[k - 1][m];
                                ITE[k - 1][m] = ITEAux;
                                ITP[k - 1][m] = ITPAux;
                            }
                            iTemp1 = intAux1[k];
                            intAux1[k] = intAux1[k-1];
                            intAux1[k-1] = iTemp1;
                        }
                    }
                }
                numLineasT++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineasT;
    }

    public static int leePeajes(String libroEntrada, String[] nombreLineas, double[][] longAux) {
        int numLineasT = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream(libroEntrada));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(libroEntrada));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("Peajes");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i = 0; i < crefs.length; i += 12) {
                Row r = s.getRow(crefs[i].getRow());
                nombreLineas[numLineasT] = r.getCell(crefs[i].getCol() - 1).getStringCellValue();
                for (int j = 0; j < 12; j += 1) {
                    longAux[numLineasT][j] = r.getCell(crefs[j].getCol()).getNumericCellValue();
                }
                numLineasT++;
            }
        } catch (java.io.FileNotFoundException e) {
            System.out.println("No se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return numLineasT;
    }
    
    

    public static int leeIT(String libroEntrada, String[] nombreLineas, double[][] longAux,String NombreRango) {
        int numLineasT = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream(libroEntrada));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(libroEntrada));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex(NombreRango);
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i = 0; i < crefs.length; i += 12) {
                Row r = s.getRow(crefs[i].getRow());
                nombreLineas[numLineasT] = r.getCell(crefs[i].getCol() - 1).getStringCellValue();
                for (int j = 0; j < 12; j += 1) {
                    longAux[numLineasT][j] = r.getCell(crefs[j].getCol()).getNumericCellValue();
                }
                numLineasT++;
            }
        } catch (java.io.FileNotFoundException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return numLineasT;
    }

    
    
    public static int leeLintronIT2(String libroEntrada, String[] TextoTemporal,
            String[] LineasT, String[] nombreLineas, int[] intAux1,
            double[][] ITE,double[][] ITEG ,double[][] ITER,double[][] ITP,int[] numLineasIT) {
        String[] txtTmp=new String[2500];
        String[] txtTmp1=new String[2500];
        int numLineasT = 0;
        numLineasIT[0] = 0;
        for (int i = 0; i < 2500; i++) {
            txtTmp[i] = "";
        }
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            AreaReference arefTx;
            CellReference[] crefsTx;
            AreaReference arefITE;
            CellReference[] crefsITE;
            AreaReference arefITEG;
            CellReference[] crefsITEG;
            AreaReference arefITER;
            CellReference[] crefsITER;
            AreaReference arefITP;
            CellReference[] crefsITP;
            int nomRangoInd = wb.getNameIndex("lintron");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            
            int nomRangoIndTx = wb.getNameIndex("transmisorIT");
            Name nomRangoTx = wb.getNameAt(nomRangoIndTx);
            arefTx = new AreaReference(nomRangoTx.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsTx = arefTx.getAllReferencedCells();
            
            int nomRangoIndITE = wb.getNameIndex("ITE");
            Name nomRangoITE = wb.getNameAt(nomRangoIndITE);
            arefITE = new AreaReference(nomRangoITE.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITE = arefITE.getAllReferencedCells();
            
            int nomRangoIndITEG = wb.getNameIndex("ITEG");
            Name nomRangoITEG = wb.getNameAt(nomRangoIndITEG);
            arefITEG = new AreaReference(nomRangoITEG.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITEG = arefITEG.getAllReferencedCells();
            
            int nomRangoIndITER = wb.getNameIndex("ITER");
            Name nomRangoITER = wb.getNameAt(nomRangoIndITER);
            arefITER = new AreaReference(nomRangoITER.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITER = arefITER.getAllReferencedCells();
            
            int nomRangoIndITP = wb.getNameIndex("ITP");
            Name nomRangoITP = wb.getNameAt(nomRangoIndITP);
            arefITP = new AreaReference(nomRangoITP.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsITP = arefITP.getAllReferencedCells();

            for (int i=0; i<crefs.length; i+=2) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c1 = null;
                Cell c2 = null;
                c1 = r.getCell(crefs[i].getCol());
                c2 = r.getCell(crefs[i+1].getCol());
                TextoTemporal[numLineasIT[0]] = c1.getStringCellValue();
                intAux1[numLineasIT[0]] = Calc.Buscar(c2.toString().trim(), nombreLineas);
                Cell cTx = r.getCell(crefsTx[0].getCol());
                txtTmp1[numLineasIT[0]] = c1.getStringCellValue()+"#"+cTx.getStringCellValue();//Lênea#Transmisor

            int t = Calc.Buscar(txtTmp1[numLineasIT[0]], txtTmp);
            if (t == -1) {
                txtTmp[numLineasT] =txtTmp1[numLineasIT[0]];
                for (int m=0; m<NUMERO_MESES; m++) {
                    Cell cITE = r.getCell(crefsITE[m].getCol());
                    ITE[numLineasT][m] =  cITE.getNumericCellValue();
                    Cell cITEG = r.getCell(crefsITEG[m].getCol());
                    ITEG[numLineasT][m] =  cITEG.getNumericCellValue();
                    Cell cITER = r.getCell(crefsITER[m].getCol());
                    ITER[numLineasT][m] =  cITER.getNumericCellValue();
                    
                    Cell cITP = r.getCell(crefsITP[m].getCol());
                    ITP[numLineasT][m] =  cITP.getNumericCellValue();
                }
                numLineasT++;
            }
            numLineasIT[0]++;
            }
            for(int i=0;i<numLineasT;i++){
                LineasT[i]=txtTmp[i];// registros “nicos Lênea#Transmisor
            //System.out.println(LineasT[i]+" "+ITE[i][0]);
            }
        }

        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineasT;
    }

    static void leeProrratasGx(String libroEntrada, double[][][] prorrataMesGx) {
        int numLineas = prorrataMesGx.length;
        int numCentrales = prorrataMesGx[0].length;
        int k = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("ProrrGMes");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            if(crefs.length!=numLineas*numCentrales*12)
                System.out.println("Se eliminaron o agregaron Centrales, pero no se han calculado las Prorratas considerando esta modificación");
            for (int i=0; i<numLineas; i++) {
                for (int j=0; j<numCentrales; j++) {
                    Row r = s.getRow(crefs[k*NUMERO_MESES].getRow());
                    for (int m=0; m<NUMERO_MESES; m++) {
                        Cell c1 = null;
                        c1 = r.getCell(crefs[m].getCol());
                        prorrataMesGx[i][j][m] = c1.getNumericCellValue();
                    }
                    k++;
                }
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

      static void leeProrratasC(String libroEntrada, double[][][] prorrataMesC) {
        int numLineas = prorrataMesC.length;
        int numClientes = prorrataMesC[0].length;
        int k = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("ProrrCMes");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            System.out.println(numLineas+" "+numClientes);
            for (int i=0; i<numLineas; i++) {
                for (int j=0; j<numClientes; j++) {
                    Row r = s.getRow(crefs[k*NUMERO_MESES].getRow());
                    for (int m=0; m<NUMERO_MESES; m++) {
                        Cell c1 = null;
                        c1 = r.getCell(crefs[m].getCol());
                        //System.out.println("líneas "+i+" cliente "+j+ " mes "+m);
                        prorrataMesC[i][j][m] = c1.getNumericCellValue();
                    }
                    k++;
                }
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

       static void leeGeneracionMes(String libroEntrada, double[][] GenMes) {//agregado para ajuste
        int numCentrales = GenMes.length;
           try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("GMes");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
              for (int i=0; i<crefs.length; i=i+(NUMERO_MESES))  {
                Row r = s.getRow(crefs[i].getRow());
                Cell cdes = null;
                for (int m=0; m<NUMERO_MESES; m++) {
                cdes = r.getCell(crefs[m].getCol());
                GenMes[i/NUMERO_MESES][m] = (double) cdes.getNumericCellValue();
                 }
              }

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }
   static void leeConsumoMes(String libroEntrada, double[][] CMes, double[][][] CU) {//agregado para ajuste
        int numConsumos =0;
        int cuenta=0;
           try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("CMesCli");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());

            AreaReference aref1;
            CellReference[] crefs1;
            int nomRangoInd1 = wb.getNameIndex("CU");
            Name nomRango1 = wb.getNameAt(nomRangoInd1);
            aref1 = new AreaReference(nomRango1.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();

              for (int i=0; i<crefs.length; i=i+(NUMERO_MESES))  {
                Row r = s.getRow(crefs[i].getRow());
                for (int m=0; m<NUMERO_MESES; m++) {
                Cell cdes = r.getCell(crefs[m].getCol());
                CMes[numConsumos][m] =cdes.getNumericCellValue();
                 }
                numConsumos++;
              }

            for (int i=0; i<crefs1.length; i+=3*NUMERO_MESES)  {
                Row r1 = s.getRow(crefs1[i].getRow());
                
                for (int m=0; m<NUMERO_MESES; m++) {
                    Cell c1 = r1.getCell(crefs1[i+m].getCol());
                    Cell c2 = r1.getCell(crefs1[i+NUMERO_MESES+m].getCol());
                    Cell c3 = r1.getCell(crefs1[i+2*NUMERO_MESES+m].getCol());

                    CU[cuenta][0][m] =  c1.getNumericCellValue();
                    CU[cuenta][1][m] =  c2.getNumericCellValue();
                    CU[cuenta][2][m] =  c3.getNumericCellValue();
                }
                cuenta++;
                 }

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }


    }
    static int[] leeLineasFlujo(String libroEntrada, String nombreLineas[]) {
        int[] lineas;
        String Linea;
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));   
            AreaReference arefFlujo = new AreaReference(wb.getNameAt(wb.getNameIndex("lineas_flujo")).getRefersToFormula(), wb.getSpreadsheetVersion());
            CellReference[] crefsFlujo = arefFlujo.getAllReferencedCells();
            lineas = new int[crefsFlujo.length];
            Sheet s = wb.getSheet(crefsFlujo[0].getSheetName());
            for (int i = 0 ; i< crefsFlujo.length;i++){
                Row r = s.getRow(crefsFlujo[i].getRow());
                Linea = r.getCell(crefsFlujo[i].getCol()).toString().trim();
                System.out.println(Linea);
                lineas[i] = Calc.Buscar(Linea,nombreLineas);
                System.out.println(lineas[i]);
                
            }
            return lineas;
            
            
        }
        
   
                
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return lineas = new int[1];
    }
    
    
     static int[] leeCentralesFlujo(String libroEntrada, String nombreLineas[], String areaNombre) {
        int[] lineas;
        String Linea;
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));   
            AreaReference arefFlujo = new AreaReference(wb.getNameAt(wb.getNameIndex(areaNombre)).getRefersToFormula(), wb.getSpreadsheetVersion());
            CellReference[] crefsFlujo = arefFlujo.getAllReferencedCells();
            lineas = new int[crefsFlujo.length];
            Sheet s = wb.getSheet(crefsFlujo[0].getSheetName());
            for (int i = 0 ; i< crefsFlujo.length;i++){
                Row r = s.getRow(crefsFlujo[i].getRow());
                Linea = r.getCell(crefsFlujo[i].getCol()).toString().trim();
                System.out.println(Linea);
                lineas[i] = Calc.Buscar(Linea,nombreLineas);
                System.out.println(lineas[i]);
                
            }
            return lineas;
        }
        
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return lineas = new int[1];
    }
    
    static void leeIndices(String libroEntrada, double[] dolar, double[] interes) {
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            // dolar
            AreaReference arefDolar = new AreaReference(wb.getNameAt(wb.getNameIndex("dolar")).getRefersToFormula(), wb.getSpreadsheetVersion());
            CellReference[] crefsDolar = arefDolar.getAllReferencedCells();
            Sheet sDolar = wb.getSheet(crefsDolar[0].getSheetName());
            // inter?s
            AreaReference arefInteres = new AreaReference(wb.getNameAt(wb.getNameIndex("interes")).getRefersToFormula(), wb.getSpreadsheetVersion());
            CellReference[] crefsInteres = arefInteres.getAllReferencedCells();
            Sheet sInteres = wb.getSheet(crefsInteres[0].getSheetName());
            for (int m=0; m<NUMERO_MESES; m++) {
                // dolar
                Row rDolar = sDolar.getRow(crefsDolar[m].getRow());
                Cell cDolar = null;
                cDolar = rDolar.getCell(crefsDolar[m].getCol());
                dolar[m] = cDolar.getNumericCellValue();
                // inter?s
                Row rInteres = sInteres.getRow(crefsInteres[m].getRow());
                Cell cInteres = null;
                cInteres = rInteres.getCell(crefsInteres[m].getCol());
                interes[m] = cInteres.getNumericCellValue();
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

    public static int leeDefbar(String libroEntrada, String[] TextoTemporal1, int[][] intAux3) {
        int numBarras = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("defbar");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=5) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c2 = r.getCell(crefs[i+1].getCol());
                Cell c3 = r.getCell(crefs[i+2].getCol());
                Cell c4 = r.getCell(crefs[i+3].getCol());
                Cell c5 = r.getCell(crefs[i+4].getCol());
                TextoTemporal1[numBarras] = c2.toString().trim(); // Nombre
                intAux3[numBarras][0] = (int) c3.getNumericCellValue(); // 1 si la barra es troncal, 0 en caso contrario
                intAux3[numBarras][1] = (int) c4.getNumericCellValue(); // 0 si la barra estö en el AIC, 1 si estö en el norte y -1 si estö en el sur
                intAux3[numBarras][2] = (int) c5.getNumericCellValue(); // 1 si la barra está en el SIC, -1 si la barra está en el SING
                numBarras++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numBarras;
    }

public static void leeConsumoxBarra(String libroEntrada, float[][] Consumos, int numBarras, int numEtapas) {
        int cuenta = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("consxbarra");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            if (crefs.length/numEtapas != numBarras) {
                System.out.println("Largo registro "+crefs.length+" "+" numero etapas "+numEtapas);
                System.out.println("Numero barras "+numBarras);
                System.out.println( "Registro de Consumos por Barra mal asignado");
                System.exit(0);
            }
            for (int i=0; i<crefs.length; i+=numEtapas) {
                Row r = s.getRow(crefs[i].getRow());
                for (int j=0; j<numEtapas; j++) {
                    Cell c = r.getCell(crefs[i+j].getCol());
                    Consumos[cuenta][j] = (float) c.getNumericCellValue();
                }
                cuenta++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

    public static void leeEtapas(String libroEntrada, int[] duracionEtapas, int numEtapas) {
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("etapas");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            if (crefs.length != numEtapas) {
                System.out.println( "Registro de Etapas mal asignado");
                System.exit(0);
            }
            for (int i=0; i<crefs.length; i++) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c = r.getCell(crefs[i].getCol());
                duracionEtapas[i] = (int) c.getNumericCellValue();
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

public static void leeLinman(String libroEntrada, int[][] LinMan, String[] nombreLineas, int numEtapas) {
        String Linea;
        int indiceLinea;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("linman");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=(numEtapas+1)) {
                Row r = s.getRow(crefs[i].getRow());
                Linea = r.getCell(crefs[i].getCol()).toString().trim();
                indiceLinea = Calc.Buscar(Linea,nombreLineas);
                //System.out.println(Linea);
                if(indiceLinea==-1)System.out.println("La línea -"+Linea+"- en 'linman' no se encuentra definida en 'lineas'");
                for (int j=0; j<numEtapas; j++) {
                    
                    Cell c = r.getCell(crefs[i+j+1].getCol());
                    //System.out.println(i+j+1);
                    //System.out.println(c.getCellType());
                    //System.out.println(c.getStringCellValue());
                    LinMan[indiceLinea][j] = (int) c.getNumericCellValue();
                }
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }


    public static int leePlpcnfe(String libroEntrada, String[] TextoTemporal1,
            int[][] intAux3, String[] nombreCentrales) {
        int numGeneradores = 0;
        int aux;
        try {
            Cell c2 = null;
            Cell c3 = null;
            Cell c4 = null;
            Cell c5 = null;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("plpcnfce");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=5) {
                Row r = s.getRow(crefs[i].getRow());
                c2 = r.getCell(crefs[i+1].getCol());
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                c5 = r.getCell(crefs[i+4].getCol());
                TextoTemporal1[numGeneradores]=c2.toString().trim(); // Nombre
                if (c3.getStringCellValue().compareTo("")!=0) {
                    aux = Calc.Buscar(c4.toString().trim()+"#"+c3.toString().trim(),nombreCentrales);
                if(aux==-1) {
                    System.out.println("El generador PLP "+c2.toString().trim()+" de "+c4.toString().trim()+" en 'centralesPLP' " +
                            "no posee una central de peajes asociada en 'centrales'");
                }
                intAux3[numGeneradores][1] = aux;
                intAux3[numGeneradores][0] = (int) c5.getNumericCellValue()-1; // barra de conexiãn
                if(intAux3[numGeneradores][0]==-1){
                    System.out.println("La barra del Generador: "+c4.toString().trim()+"#"+c3.toString().trim()+" se encuentra mal asignada");
                }
                numGeneradores++;
            }
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numGeneradores;
    }

    public static int leePlpcnfe(String libroEntrada, String[] TextoTemporal1,
             String[] nombreCentrales) {
        int numGeneradores = 0;
        int aux;
        try {
            Cell c2 = null;
            Cell c3 = null;
            Cell c4 = null;
            Cell c5 = null;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("plpcnfce");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=5) {
                Row r = s.getRow(crefs[i].getRow());
                c2 = r.getCell(crefs[i+1].getCol());
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                c5 = r.getCell(crefs[i+4].getCol());
                TextoTemporal1[numGeneradores]=c2.toString().trim(); // Nombre
                if (c3.getStringCellValue().compareTo("")!=0) {
                    aux = Calc.Buscar(c4.toString().trim()+"#"+c3.toString().trim(),nombreCentrales);
                if(aux==-1) {
                    System.out.println("El generador PLP "+c2.toString().trim()+" de "+c4.toString().trim()+" en 'centralesPLP' " +
                            "no posee una central de peajes asociada en 'centrales'");
                }
                
                
            }
                numGeneradores++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numGeneradores;
    }
    
    
    
    
    public static int leeSumin(String libroEntrada, String[] TextoTemporal1) {
        int numSumin = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("sumin");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i++) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c = null;
                c = r.getCell(crefs[i].getCol());
                TextoTemporal1[numSumin] = c.toString().trim(); // Nombre
                numSumin++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numSumin;
    }

public static void leeOrient(String libroEntrada, int[][] orientBarTroncal, String[] nombreBarras, String[] nombreLineas) {
        String Linea;
        int indiceLinea;
        String Barra;
        int indiceBarra;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref1, aref2;
            CellReference[] crefs1, crefs2;
            int nomRangoInd2 = wb.getNameIndex("orientCol");
            int nomRangoInd3 = wb.getNameIndex("orientFil");
            Name nomRango1 = wb.getNameAt(nomRangoInd2);
            Name nomRango2 = wb.getNameAt(nomRangoInd3);
            aref1 = new AreaReference(nomRango1.getRefersToFormula(), wb.getSpreadsheetVersion());
            aref2 = new AreaReference(nomRango2.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();
            crefs2 = aref2.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs1[0].getSheetName());
            int nCol = crefs1.length;
            int nFil = crefs2.length;
            for (int i=0; i<nCol; i++) {
                Row r1 = s.getRow(crefs1[i].getRow());
                Linea = r1.getCell(crefs1[i].getCol()).toString().trim();
                indiceLinea = Calc.Buscar(Linea,nombreLineas);
                for (int j=0; j<nFil; j++) {
                    Row r2 = s.getRow(crefs2[j].getRow());
                    Barra = r2.getCell(crefs2[j].getCol()).toString().trim();
                    indiceBarra = Calc.Buscar(Barra,nombreBarras);
                    if(indiceBarra==-1){
                        System.out.println("La barra "+Barra+" en 'orient' no se encuentra en la hoja 'barras'");
                    System.exit(0);
                    }
                    if(indiceLinea==-1){
                        System.out.println("La línea "+Linea+" en 'orient' no se encuentra en la hoja 'lineas'");
                    System.exit(0);
                    }
                    
                    Cell c = r2.getCell(crefs1[i].getCol());
                    orientBarTroncal[indiceBarra][indiceLinea] = (int) c.getNumericCellValue();
                }
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
    }

    public static int leeBarcli(String libroEntrada, String[] TextoTemporal1, String[] TextoTemporal2,
            int[][] intAux3, String[] nombreClientes, String[] nombreBarras) {
        int numClaves = 0;
        Row fila = null;
        Cell cell = null;
        Cell c2e = null;
        Cell c3e = null;
        Cell c4e = null;
        Cell c5e = null;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));

            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("barcli");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());

         
            for (int i=0; i<crefs.length; i+=6) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c2 = null;
                Cell c3 = null;
                Cell c4 = null;
                Cell c5 = null;
                Cell c6= null;
                c2 = r.getCell(crefs[i+1].getCol());
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                c5 = r.getCell(crefs[i+4].getCol());
                c6 = r.getCell(crefs[i+5].getCol());
                TextoTemporal1[numClaves]=c2.toString().trim();  //Clave
                //intAux3[numClaves][1]=Calc.Buscar(c4.toString().trim(),nombreSumin); // Suministrador
                intAux3[numClaves][0]=Calc.Buscar(c5.toString().trim(),nombreBarras); // Barras
                if(intAux3[numClaves][0]==-1){
                System.out.println("La barra de consumo "+c5.toString().trim()+ " en 'consumos' no se encuentra en hoja 'barras'");
                }
                intAux3[numClaves][2]=Calc.Buscar(c3.toString().trim()+"#"+c4.toString().trim()+"#"+c5.toString().trim(),nombreClientes); // Cliente
                if(intAux3[numClaves][2]==-1){
                System.out.println("El consumo "+c3.toString().trim()+"#"+c4.toString().trim()+"#"+c5.toString().trim()+
                        " no tiene un Cliente asociado en hoja 'Clientes'");
                }
                intAux3[numClaves][3]= (int) c6.getNumericCellValue();
                    numClaves++;
            }

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numClaves;
    }

    public static int leeConsumos(String libroEntrada, float[][] ConsumosClaves,float[][] ConsClaveMes,
            int numEtapas,int[] paramEtapa,int[] duracionEta, float[][][] ECU) {
        int numClaves = 0;
        int cuenta1 = 0;
        int cuenta2 = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("consumos");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());

            AreaReference aref1;
            CellReference[] crefs1;
            int nomRangoInd1 = wb.getNameIndex("CU");
            Name nomRango1 = wb.getNameAt(nomRangoInd1);
            aref1 = new AreaReference(nomRango1.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();

            for (int i=0; i<crefs.length; i+=numEtapas) {
                Row r = s.getRow(crefs[i].getRow());
                for (int j=0; j<numEtapas; j++) {
                    Cell c = null;
                    c = r.getCell(crefs[i+j].getCol());
                    ConsumosClaves[numClaves][j] = (float)c.getNumericCellValue();
                }
                numClaves++;
            }
            
           
            for(int i=0;i<crefs1.length;i+=3*(NUMERO_MESES)){
            Row r1 = s.getRow(crefs1[i].getRow());
            
            for (int m=0; m<NUMERO_MESES; m++) {
                Cell c1 = null;
                Cell c2 = null;
                Cell c3 = null;
                    c1 = r1.getCell(crefs1[i+m].getCol());
                    c2 = r1.getCell(crefs1[i+NUMERO_MESES+m].getCol());
                    c3 = r1.getCell(crefs1[i+2*NUMERO_MESES+m].getCol());
                    ECU[cuenta1][0][m] = (float) c1.getNumericCellValue();
                    ECU[cuenta1][1][m] = (float) c2.getNumericCellValue();
                    ECU[cuenta1][2][m] = (float) c3.getNumericCellValue();
            }
                    cuenta1++;
            }
            //Calcula la energêa mensual por Clave
           for (int j=0; j<numClaves; j++) {
                for(int e=0;e<numEtapas;e++) {
                        ConsClaveMes[j][paramEtapa[e]]
                                += ConsumosClaves[j][e]*duracionEta[e];
                }
           }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numClaves;
    }
 public static int leeCU(String libroEntrada, double[] ECU2,  double[] ECU30,int[] intAux3,String[] nombreBarras) {
        int numClaves = 0;
        int cuenta = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));

            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("CU");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());

            AreaReference aref1;
            CellReference[] crefs1;
            int nomRangoInd1 = wb.getNameIndex("barcli");
            Name nomRango1 = wb.getNameAt(nomRangoInd1);
            aref1 = new AreaReference(nomRango1.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();

               for (int i=0; i<crefs1.length; i+=5) {
                Row r1 = s.getRow(crefs1[i].getRow());
                Cell c5 = null;
                c5 = r1.getCell(crefs1[i+4].getCol());
                intAux3[numClaves]=Calc.Buscar(c5.toString().trim(),nombreBarras); // Barras

                Row r = s.getRow(crefs[cuenta].getRow());
                Cell c1 = null;
                Cell c2 = null;
                    c1 = r.getCell(crefs[cuenta].getCol());
                    c2 = r.getCell(crefs[cuenta+1].getCol());
                    ECU2[numClaves] = c1.getNumericCellValue();
                    ECU30[numClaves] = c2.getNumericCellValue();
                cuenta=cuenta+2;

                numClaves++;
               }

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numClaves;
    }

        public static int leeConsumos2(String libroEntrada, float[][] ConsumosClaves,float[][] ConsClaveMes,
                          int numEtapas,int[] paramEtapa,int[] duracionEta, float[][][] ECU) {
        int numClaves = 0;
        int cuenta1 = 0;
        int cuenta2 = 0;
        String[] ClaveCli=new String[2500];
        XSSFSheet hoja = null;

        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("consumos");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());

            AreaReference aref1;
            CellReference[] crefs1;
            int nomRangoInd1 = wb.getNameIndex("CU");
            Name nomRango1 = wb.getNameAt(nomRangoInd1);
            aref1 = new AreaReference(nomRango1.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();

            for (int i=0; i<crefs.length; i+=numEtapas) {
                Row r = s.getRow(crefs[i].getRow());
                for (int j=0; j<numEtapas; j++) {
                    Cell c = null;
                    c = r.getCell(crefs[i+j].getCol());
                    ConsumosClaves[numClaves][j] = (float)c.getNumericCellValue();
                }
                numClaves++;
            }
            
                       
            
            for(int i=0;i<crefs1.length;i+=3*(NUMERO_MESES)){
            Row r1 = s.getRow(crefs1[i].getRow());
            
            for (int m=0; m<NUMERO_MESES; m++) {
                Cell c1 = null;
                Cell c2 = null;
                Cell c3 = null;
                    c1 = r1.getCell(crefs1[i+m].getCol());
                    c2 = r1.getCell(crefs1[i+NUMERO_MESES+m].getCol());
                    c3 = r1.getCell(crefs1[i+2*NUMERO_MESES+m].getCol());
                    ECU[cuenta1][0][m] = (float) c1.getNumericCellValue();
                    ECU[cuenta1][1][m] = (float) c2.getNumericCellValue();
                    ECU[cuenta1][2][m] = (float) c3.getNumericCellValue();
            }
                    cuenta1++;
            }
            //Calcula la energêa mensual por Clave
           for (int j=0; j<numClaves; j++) {
                for(int e=0;e<numEtapas;e++) {
                        ConsClaveMes[j][paramEtapa[e]]
                                += ConsumosClaves[j][e]*duracionEta[e];
                }
           }
            //Extrae los clientes asociados a los conusmos
            int nomRangoInd2 = wb.getNameIndex("barcli");
            Name nomRango2= wb.getNameAt(nomRangoInd2);
            aref = new AreaReference(nomRango2.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();

            for (int i=0; i<crefs.length; i+=5) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c3 = null;
                Cell c4 = null;
                Cell c5 = null;
                c3 = r.getCell(crefs[i+2].getCol());
                c4 = r.getCell(crefs[i+3].getCol());
                c5 = r.getCell(crefs[i+4].getCol());
                ClaveCli[cuenta2]=c3.toString().trim()+"#"+c4.toString().trim()+"#"+c5.toString().trim(); // Cliente
                cuenta2++;
                }
        String[] TxtTemp0 = new String[numClaves];
        for (int i = 0; i < numClaves; i++) {
            TxtTemp0[i] = "";
        }
        int numCli = 0;
        for (int j = 0; j < numClaves; j++) {
            int l = Calc.Buscar(ClaveCli[j], TxtTemp0);
            if (l == -1) {
                TxtTemp0[numCli] = ClaveCli[j];
                numCli++;
            }
        }
        String[] nomCli=new String[numCli];
        for (int j = 0; j < numCli; j++) {
            nomCli[j] = TxtTemp0[j];
        }

        String[] nomCliO=new String[numCli];
        int[] nc = Calc.OrdenarBurbujaStr(nomCli);
        nomCliO = new String[numCli];
        for (int i = 0; i < numCli; i++) {
            nomCliO[i] = nomCli[nc[i]];
        }


        //Escribe los clientes en hoja clientes del archivo de entrada
         hoja = wb.getSheet("clientes");
         Cell cell = null;
         Row row=null;
         short fila = 5;

         Font font = wb.createFont();
            font.setFontHeightInPoints((short)10);
            font.setFontName("Century Gothic");
            CellStyle estilo = wb.createCellStyle();
            estilo.setFont(font);
            CellStyle estilo1 = wb.createCellStyle();
            estilo1.setFont(font);
            estilo1.setAlignment(HorizontalAlignment.CENTER);


         // Titulos Secundarios

            for(int i=0;i<numCli;i++){
             row = hoja.createRow(fila); fila++;
                cell = row.createCell(1);
            cell.setCellValue(i+1);
            cell.setCellStyle(estilo1);
            cell = row.createCell(2);

            cell.setCellValue(nomCliO[i]);
            cell.setCellStyle(estilo);
            cell = row.createCell(3);
            cell.setCellValue(-1);
            cell.setCellStyle(estilo);

            }
            // Crea nombre de rango de salida
            Name nombreCel = wb.createName();
            wb.removeName("clientes");
            nombreCel.setNameName("clientes"); // Nombre del rango
            CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
            String reference = "clientes"+"!C6:"+cellRef.formatAsString(); // area reference
            nombreCel.setRefersToFormula(reference);

            FileOutputStream archivoSalida = new FileOutputStream( libroEntrada );
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println( "Acaba de extraer y escribir en la hoja 'clientes' los Clientes asociados a los consumos" );
            System.out.println( "Recuerde indicar los clientes excentos antes de calcular los pagos" );
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numClaves;
    }

    public static int leeLinPLP(String libroEntrada, String[] TextoTemporal1, float[][] Aux) {
        int numLineasSistRed = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("linPLP");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=3) {
                Row r = s.getRow(crefs[i].getRow());
                Cell c2 = r.getCell(crefs[i+1].getCol());
                Cell c3 = r.getCell(crefs[i+2].getCol());
                TextoTemporal1[numLineasSistRed] = c2.toString().trim(); // Nombre
                Aux[numLineasSistRed][0] = (float) c3.getNumericCellValue(); // Tensiãn
                numLineasSistRed++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numLineasSistRed;
    }

        public static int leeMeses(String libroEntrada, int[] intAux, String[] nombreMeses) {
        int numSubperiodos = 0;
        Cell c1 = null;
        Cell c2 = null;
        int numBloques;
        int et = 0;
        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));


            AreaReference aref;
            CellReference[] crefs;
            int nomRangoInd = wb.getNameIndex("meses");
            Name nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            Sheet s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=2) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                c2 = r.getCell(crefs[i+1].getCol());
                numBloques = (int) c2.getNumericCellValue();
                for (int j=0; j<numBloques; j++) {
                    intAux[et] = Calc.Buscar(c1.toString().trim(),nombreMeses);
                    et++;
                }
                numSubperiodos++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numSubperiodos;
    }

        public static int leeEfirme(String libroEntrada, String[] TextoTemporal1, double[][] Efirme) {
        int numEmpre = 0;
        try {
            Cell c1 = null;
            int nomRangoInd;
            Name nomRango;
            AreaReference aref;
            CellReference[] crefs;
            Sheet s;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            // Lectura de datos
            nomRangoInd = wb.getNameIndex("EmpEfir");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i++) {
                Row r = s.getRow(crefs[i].getRow());
                c1 = r.getCell(crefs[i].getCol());
                TextoTemporal1[i] = c1.toString().trim();
            }

            // Lectura de datos
            nomRangoInd = wb.getNameIndex("Efirme");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (int i=0; i<crefs.length; i+=NUMERO_MESES) {
                Row r = s.getRow(crefs[i].getRow());
                for (int j=0; j<NUMERO_MESES; j++) {
                    c1 = r.getCell(crefs[i + j].getCol());
                    Efirme[numEmpre][j] = c1.getNumericCellValue();
                }
                numEmpre++;
            }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return numEmpre;
    }
 public static int[] leeDistribuidoras(String libroEntrada, String[] TextoTemporal1,String[] TextoTemporal2, double[][][] Prorrata) {
        int[] cuenta=new int[2];
        int numDx;
        int aux=0;
        int numSum=0;
        try {
            Cell c1 = null;
            Cell c2 = null;
            int nomRangoInd;
            Name nomRango;
            AreaReference aref;
            CellReference[] crefs;
            AreaReference aref1;
            CellReference[] crefs1;
            Sheet s;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            // Lectura de datos

            nomRangoInd = wb.getNameIndex("Distr");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (numDx=0; numDx<crefs.length; numDx++) {
                Row r = s.getRow(crefs[numDx].getRow());
                c1 = r.getCell(crefs[numDx].getCol());
                TextoTemporal1[numDx] = c1.toString().trim();
            }
            int sum=0;
            int tmp=0;
            for(int i=0;i<NUMERO_MESES;i++){
                tmp=i+1;
            nomRangoInd = wb.getNameIndex("ProrrDx"+tmp);
            nomRango = wb.getNameAt(nomRangoInd);
            aref1 = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();
            sum=crefs1.length/numDx;
            for (numSum=0; numSum<sum; numSum++) {
                Row r = s.getRow(crefs1[numSum].getRow());
                c2 = r.getCell(crefs1[numSum].getCol());
                TextoTemporal2[numSum] = c2.toString().trim();
                //System.out.println(TextoTemporal2[numSum]);
            }
            // Lectura de datos
            aux=0;
            for (int j=sum; j<crefs1.length; j+=numSum) {
                Row r = s.getRow(crefs1[j].getRow());
                for (int k=0; k<numSum; k++) {
                    c1 = r.getCell(crefs1[j + k].getCol());
                    //System.out.println(c1.getNumericCellValue());
                    Prorrata[aux][k][i] = c1.getNumericCellValue();
                }
                aux++;
                //System.out.println(aux);
            }
            cuenta[0]=numDx;
            cuenta[1]=numSum;
        }
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return cuenta;
    }
 
    public static int leeProrrataEfirme(String libroEntrada, String[] TextoTemporal, double[][] Prorrata) {
        int cuenta = 0;
        int numSum;
        try {
            Cell c1 = null;
            Cell c2 = null;
            int nomRangoInd;
            Name nomRango;
            AreaReference aref;
            CellReference[] crefs;
            AreaReference aref1;
            CellReference[] crefs1;
            Sheet s;
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream( libroEntrada ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream( libroEntrada ));
            // Lectura de datos

            nomRangoInd = wb.getNameIndex("sumRM88");
            nomRango = wb.getNameAt(nomRangoInd);
            aref = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs = aref.getAllReferencedCells();
            s = wb.getSheet(crefs[0].getSheetName());
            for (numSum=0; numSum<crefs.length; numSum++) {
                Row r = s.getRow(crefs[numSum].getRow());
                c1 = r.getCell(crefs[numSum].getCol());
                TextoTemporal[numSum] = c1.toString().trim();
                //System.out.println(TextoTemporal[numSum]);
            }
            
            nomRangoInd = wb.getNameIndex("ProrrRM88");
            nomRango = wb.getNameAt(nomRangoInd);
            aref1 = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefs1 = aref1.getAllReferencedCells();
            // Lectura de datos
            //System.out.println(crefs1.length);
            for (int suministrador=0; suministrador<numSum; suministrador++) {
                Row r = s.getRow(crefs1[suministrador*NUMERO_MESES].getRow());
                for (int mes=0; mes<NUMERO_MESES; mes++) {
                    c1 = r.getCell(crefs1[mes + suministrador*NUMERO_MESES].getCol());
                    //System.out.println(c1.getNumericCellValue());
                    Prorrata[mes][suministrador] = c1.getNumericCellValue();
                }
                //System.out.println(aux);
            }
            cuenta =numSum;
        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }
        return cuenta;
 }

    public static void leePropiedades(Properties propiedades, String ruta){
        InputStream is = null;
        try {
            File f = new File(ruta);
            is = new FileInputStream( f );
            propiedades.load(is);
        }
        catch ( Exception e ) {
            e.printStackTrace();
            is = null;
        }
        if (is != null) {
            try {
                is.close();
            }
            catch ( Exception e ) {
            e.printStackTrace();
            }
        }


    }


}

