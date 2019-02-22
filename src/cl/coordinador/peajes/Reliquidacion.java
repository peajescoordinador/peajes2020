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

import static cl.coordinador.peajes.PeajesConstant.MESES;
import static cl.coordinador.peajes.PeajesConstant.NUMERO_MESES;
import static cl.coordinador.peajes.PeajesConstant.SLASH;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.util.StringTokenizer;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase de reliquidacion:
 * <li>craer un archivo que se llame reliquidacion con el resumen de los IT de
 * facturaccion luego a este archivo se le iran agregadon hojas</li>
 * <li>leer el archivo liquidacion mes del directorio de la liquidaci‹n y copiar
 * tablas y pegarlas en un nuevo archivo</li>
 * <li>leer el archivo liquidacion mes del directorio de la RELIQUIDACION y
 * copiar tablas y pegarlas en EL MISMO ARCHIVO</li>
 * <li>pegar funciones suma y resta para creear cuadro de liquidacion</li>
 * <li>copiar el ultimo cuadro en una hoja de pagos</li>
 *
 * @author vtoro
 */
public class Reliquidacion {

public static void Reliquidacion(File DirectorioLiquidacion, File DirectorioReliquidacion, String mes, int Ano,String Fpago){

    String libroLiquidacion= DirectorioLiquidacion + SLASH +"Liquidación" + mes + ".xlsx";
    String libroReliquidacion = DirectorioReliquidacion + SLASH +"Liquidación" + mes + ".xlsx";
    File ArchivoReli= new File(DirectorioReliquidacion + SLASH +"Liquidación" + mes + ".xlsx");

    int nummes=0;

    for (int i = 0; i < NUMERO_MESES; i++) {
        if (mes.equals(MESES[i])) {
            nummes = i;
            i = i + 1;
            if (i < 10) {
                mes = "0" + i;
            }
        }
    }

    String libroRit = DirectorioReliquidacion + SLASH +"rit" + mes + ".xlsx";
    mes=MESES[nummes];

        Cell cLiq = null;
        Cell cReliq = null;
        Cell cRit = null;
        Cell cRitRel = null;
        Cell cRitIT = null;
        CellReference c4 = null;
        CellReference RefI;
        CellReference RefF;
        XSSFSheet  hojarit=null;
        XSSFSheet  hojacuadro=null;
        Row rowliq=null;
        Row rowReliq=null;
        Row rowRit=null;
        Row rowcuadro=null;
        Row rowRitRel=null;
        Row rowRitIT=null;

        try {
            //POIFSFileSystem fs = new //POIFSFileSystem(new FileInputStream(libroLiquidacion));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(libroLiquidacion));
            //POIFSFileSystem fsRel = new //POIFSFileSystem(new FileInputStream(libroReliquidacion));
            XSSFWorkbook wbRel = new XSSFWorkbook(new FileInputStream(libroReliquidacion));

            //POIFSFileSystem fsrit = new //POIFSFileSystem(new FileInputStream(libroRit));
            XSSFWorkbook wbrit = new XSSFWorkbook(new FileInputStream(libroRit));
            hojarit = wbrit.getSheet("Detalle");
            hojacuadro=wbrit.getSheet("Cuadro de Pago");
            if(hojacuadro==null)
                hojacuadro = wbrit.createSheet("Cuadro de Pago");
            else{
                wbrit.removeSheetAt(wbrit.getSheetIndex(hojacuadro));
                hojacuadro = wbrit.createSheet("Cuadro de Pago");
            }
            if(hojarit==null)
                hojarit = wbrit.createSheet("Detalle");
            else{
                wbrit.removeSheetAt(wbrit.getSheetIndex(hojarit));
                hojarit = wbrit.createSheet("Detalle");
            }

            hojarit.setPrintGridlines(false);
            hojarit.setDisplayGridlines(false);
            hojacuadro.setPrintGridlines(false);
            hojacuadro.setDisplayGridlines(false);
            
            Font font = wbrit.createFont();
            font.setFontHeightInPoints((short)8);
            font.setFontName("Century Gothic");
            DataFormat formato1 = wbrit.createDataFormat();
            CellStyle estiloDatos1 = wbrit.createCellStyle();
            StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"", ";");
            String formatoPos1 = formatoCompleto1.nextToken();
            estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
            estiloDatos1.setFont(font);
            estiloDatos1.setBorderRight(BorderStyle.THIN);
            estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

            CellStyle estiloDatosA = wbrit.createCellStyle();
            estiloDatosA.setDataFormat(formato1.getFormat(formatoPos1));
            estiloDatosA.setFont(font);
            estiloDatosA.setBorderRight(BorderStyle.THIN);
            estiloDatosA.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
            
            
            CellStyle estiloTexto = wbrit.createCellStyle();
            estiloTexto.setFont(font);


            Font fontTituloFila = wbrit.createFont();
            fontTituloFila.setFontHeightInPoints((short)8);
            fontTituloFila.setFontName("Century Gothic");
            fontTituloFila.setBold(true);
            CellStyle estiloTituloFila = wbrit.createCellStyle();
            estiloTituloFila.setFont(fontTituloFila);
            estiloTituloFila.setBorderRight(BorderStyle.THIN);
            estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFila.setBorderLeft(BorderStyle.THIN);
            estiloTituloFila.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFila.setBorderBottom(BorderStyle.THIN);
            estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFila.setBorderTop(BorderStyle.THIN);
            estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

            CellStyle estiloTituloFilaA = wbrit.createCellStyle();
            estiloTituloFilaA.setFont(fontTituloFila);
            estiloTituloFilaA.setBorderRight(BorderStyle.THIN);
            estiloTituloFilaA.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFilaA.setBorderLeft(BorderStyle.THIN);
            estiloTituloFilaA.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFilaA.setBorderBottom(BorderStyle.THIN);
            estiloTituloFilaA.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFilaA.setBorderTop(BorderStyle.THIN);
            estiloTituloFilaA.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloTituloFilaA.setAlignment(HorizontalAlignment.CENTER);
            estiloTituloFilaA.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
            estiloTituloFilaA.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font fontTitulo = wbrit.createFont();
            fontTitulo.setFontHeightInPoints((short)9);
            fontTitulo.setFontName("Century Gothic");
            fontTitulo.setBold(true);
            CellStyle estiloTitulo = wbrit.createCellStyle();
            estiloTitulo.setFont(fontTitulo);

            DataFormat formato4 = wbrit.createDataFormat();
            CellStyle estiloDatos4 = wbrit.createCellStyle();
            StringTokenizer formatoCompleto4 = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"", ";");
            String formatoPos4 = formatoCompleto4.nextToken();
            estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
            estiloDatos4.setFont(font);
            estiloDatos4.setBorderRight(BorderStyle.THIN);
            estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloDatos4.setBorderBottom(BorderStyle.THIN);
            estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloDatos4.setBorderTop(BorderStyle.THIN);
            estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

            CellStyle estiloDatos4A = wbrit.createCellStyle();
            estiloDatos4A.setDataFormat(formato4.getFormat(formatoPos4));
            estiloDatos4A.setFont(font);
            estiloDatos4A.setBorderRight(BorderStyle.THIN);
            estiloDatos4A.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloDatos4A.setBorderBottom(BorderStyle.THIN);
            estiloDatos4A.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloDatos4A.setBorderTop(BorderStyle.THIN);
            estiloDatos4A.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
            estiloDatos4A.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
            estiloDatos4A.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            AreaReference arefliq;
            CellReference[] crefsliq;
            AreaReference arefReliq;
            CellReference[] crefsReliq;

            Name nomRango = wb.getName("Total");
            arefliq = new AreaReference(nomRango.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsliq = arefliq.getAllReferencedCells();
            Sheet s = wb.getSheet(crefsliq[0].getSheetName());
            c4=arefliq.getLastCell();
           int filas=c4.getRow()-4;
           short columnas=c4.getCol();
           
            Name nomRangoRe = wbRel.getName("Total");
            arefReliq = new AreaReference(nomRangoRe.getRefersToFormula(), wb.getSpreadsheetVersion());
            crefsReliq = arefReliq.getAllReferencedCells();
            Sheet sRel = wbRel.getSheet(crefsReliq[0].getSheetName());
            
           int k=0;
           CellType tipo = CellType.BLANK;
           int aux=0;
           int a=0;
           int indcolTot=300;
           double valor=0;
           XSSFFormulaEvaluator formula=new XSSFFormulaEvaluator(wbrit);


            for (int i=0; i<filas; i++) {
                     rowRit = hojarit.createRow(i+5);//hoja detalle, celdas para pagos con IT estimado
                     rowcuadro = hojacuadro.createRow(i+5);
                     rowRitRel = hojarit.createRow(i+10+filas);//hoja detalle, celdas para pagos con IT Real
                     rowRitIT = hojarit.createRow(i+14+filas*2);//hoja detalle, celdas para diferencia entre IT estiamdo - IT Real
                     //Lee datos en archivo de liquidaci‹n con IT estimado
                     for (int j=0; j<columnas+1; j++) {
                         rowliq = s.getRow(crefsliq[i*columnas].getRow());
                         cLiq = rowliq.getCell(j);
                         cRit = rowRit.createCell(j);
                         if(cLiq==null){
                             cRit.setCellValue("");
                             if(i>1){
                                 aux=aux+1;
                                 if(aux==2){
                                      indcolTot=j;
                                      System.out.println(indcolTot);
                                 }
                             }
                         }
                         //copia datos en archivo de reliquidaci‹n rit.xls
                         else{
                             tipo=cLiq.getCellType();
                             if(null!=tipo)switch (tipo) {
                                 case NUMERIC:
                                     cRit.setCellValue(cLiq.getNumericCellValue());
                                     cRit.setCellStyle(estiloDatos1);
                                     break;
                                 case STRING:
                                     cRit.setCellValue(cLiq.toString());
                                     cRit.setCellStyle(estiloTituloFila);
                                     break;
                                 case FORMULA:
                                     cRit.setCellFormula(cLiq.getCellFormula());
                                     cRit.setCellStyle(estiloDatos4);
                                     break;
                                 default:
                                     break;
                             }
                             hojarit.autoSizeColumn(j);
                         }


                         rowReliq = sRel.getRow(crefsReliq[i*columnas].getRow());
                         cReliq = rowReliq.getCell(j);//Lee datos de archivo liquidacion con IT Real
                         cRitRel = rowRitRel.createCell(j);
                         cRitIT = rowRitIT.createCell(j);
                         if(cReliq==null){
                         cRitRel.setCellValue("");
                         }
                         else{
                             tipo=cReliq.getCellType();
                             if(null!=tipo)switch (tipo) {
                                 case NUMERIC:
                                     cRitRel.setCellValue(cReliq.getNumericCellValue());
                                     cRitRel.setCellStyle(estiloDatos1);
                                     RefI = new CellReference(cRit.getRowIndex(), cRit.getColumnIndex());
                                     RefF = new CellReference(cRitRel.getRowIndex(), cRitRel.getColumnIndex());
                                     cRitIT.setCellFormula(RefF.formatAsString()+"-"+RefI.formatAsString());
                                     if(j<indcolTot){
                                         a=j;
                                         estiloDatosA.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
                                         estiloDatosA.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                         cRitIT.setCellStyle(estiloDatosA);
                                         XSSFFormulaEvaluator.evaluateAllFormulaCells(wbrit);
                                         rowcuadro.createCell(a).setCellValue(cRitIT.getNumericCellValue());
                                         rowcuadro.getCell(a).setCellStyle(estiloDatosA);
                                     }
                                     else{
                                         cRitIT.setCellStyle(estiloDatos1);
                                     }       break;
                                 case STRING:
                                     cRitRel.setCellValue(cReliq.toString());
                                     cRitRel.setCellStyle(estiloTituloFila);
                                     cRitIT.setCellValue(cReliq.toString());
                                     if(j<indcolTot){
                                         a=j;
                                         cRitIT.setCellStyle(estiloTituloFilaA);
                                         rowcuadro.createCell(a).setCellValue(cRitIT.toString());
                                         rowcuadro.getCell(a).setCellStyle(estiloTituloFilaA);
                                     }
                                     else{
                                         cRitIT.setCellStyle(estiloTituloFila);
                                     }       break;
                                 case FORMULA:
                                     RefI = new CellReference(14+filas, cRitRel.getColumnIndex());
                                     RefF = new CellReference(cRitRel.getRowIndex()-1, cRitRel.getColumnIndex());
                                     cRitRel.setCellFormula("sum("+RefI.formatAsString()+":"+RefF.formatAsString()+")");
                                     cRitRel.setCellStyle(estiloDatos4);
                                     RefI = new CellReference(cRit.getRowIndex(), cRit.getColumnIndex());
                                     RefF = new CellReference(cRitRel.getRowIndex(), cRitRel.getColumnIndex());
                                     cRitIT.setCellFormula(RefF.formatAsString()+"-"+RefI.formatAsString());
                                     if(j<indcolTot){
                                         a=j;
                                         cRitIT.setCellStyle(estiloDatos4A);
                                         XSSFFormulaEvaluator.evaluateAllFormulaCells(wbrit);
                                         rowcuadro.createCell(a).setCellValue(cRitIT.getNumericCellValue());
                                         rowcuadro.getCell(a).setCellStyle(estiloDatos4A);
                                     }
                                     else{
                                         cRitIT.setCellStyle(estiloDatos4);
                                     }       break;
                                 default:
                                     break;
                             }
                             hojarit.autoSizeColumn(j);
                             hojacuadro.autoSizeColumn(a);
                         }
                     }
            }

            hojarit.createRow(3).createCell(1).setCellValue("1.- Liquidación de Peajes correspondiente a "+ mes+" de "+Ano+
                    " (Valores en $ indexados a "+ mes+" de "+Ano+ " ) con IT Estimados");
            hojarit.getRow(3).getCell(1).setCellStyle(estiloTitulo);
            hojarit.createRow(filas+8).createCell(1).setCellValue("2.- Cálculo de Peajes correspondiente a "+ mes+" de "+Ano+
                    "  (Valores en $ indexados a "+ mes+" de "+Ano+ ") con IT Reales");
            hojarit.getRow(filas+8).getCell(1).setCellStyle(estiloTitulo);
            hojarit.createRow(12+filas*2).createCell(1).setCellValue("3.-Reliquidación de Ingresos Tarifarios correspondiente a "+ mes+" de "+Ano+
                    " (Valores en $ indexados a "+ mes+" de "+Ano+ ")");
            hojarit.getRow(12+filas*2).getCell(1).setCellStyle(estiloTitulo);
            hojarit.createRow(16+filas*3).createCell(1).setCellValue("(1) Valores Positivos: Usuarios Pagan, Valores Negativos: Usuarios Reciben");
            hojarit.getRow(12+filas*2).getCell(1).setCellStyle(estiloTexto);
            hojarit.createRow(17+filas*3).createCell(1).setCellValue("(2) Suma de Cuadros N° 2 y N° 3");
            hojarit.getRow(12+filas*2).getCell(1).setCellStyle(estiloTexto);
            
            hojacuadro.createRow(2).createCell(1).setCellValue("Reliquidación de Ingresos Tarifarios");
            hojacuadro.getRow(2).getCell(1).setCellStyle(estiloTitulo);
            hojacuadro.createRow(3).createCell(1).setCellValue( mes+" de "+Ano);
            hojacuadro.getRow(3).getCell(1).setCellStyle(estiloTitulo);
            hojacuadro.createRow(4).createCell(1).setCellValue("(Valores en $ - fecha de pago hasta: "+Fpago+")");
            hojacuadro.getRow(4).getCell(1).setCellStyle(estiloTitulo);

            FileOutputStream archivoSalida = new FileOutputStream( libroRit );
            wbrit.write(archivoSalida);
            archivoSalida.close();
            ArchivoReli.delete();

        }
        catch (java.io.FileNotFoundException e) {
                System.out.println( "No se se puede acceder al archivo " + e.getMessage());
        }
        catch (Exception e) {
                e.printStackTrace();
        }

    }
}
