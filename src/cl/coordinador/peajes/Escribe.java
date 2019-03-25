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
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.StringTokenizer;
import java.util.Properties;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author aramos
 */
public class Escribe {

    static public void crearLibro( String nomLibro ) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            FileOutputStream archivoSalida = new FileOutputStream( nomLibro );
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println( "Acaba de crear el archivo xls vacio " + nomLibro );
        }
        catch (java.io.FileNotFoundException e) {
            System.out.println( "No se se puede acceder la archivo " + e.getMessage());
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public XSSFWorkbook crearLibroVacio(String nomLibro) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
        wb.write(archivoSalida);
        archivoSalida.close();
        System.out.println("Acaba de crear el archivo xls vacio " + nomLibro);
        return wb;
    }
    
    static public void guardaLibroDisco(XSSFWorkbook wb, String nomLibro) throws IOException {
        FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
        wb.write(archivoSalida);
        archivoSalida.close();
    }
    
    static public void creaH1F_2d_double(String titulo, double Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH1F_2d_double(titulo, Datos, tituloFilas, nombreFilas, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static public void creaH1F_2d_double(String titulo, double Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja;
        Cell cellTC ;
        Cell cellTF;
        Cell cell = null;
        Row row;
        short fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 2);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 2);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 2; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 2);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

    }

    static public void creaH1F_2d_long(String titulo, double Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH1F_2d_long(titulo, Datos, tituloFilas, nombreFilas, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH1F_2d_long(String titulo, double Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        CellStyle estiloDatos1 = wb.createCellStyle();
        estiloDatos1.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderBottom(BorderStyle.THIN);
        estiloDatos1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setBorderTop(BorderStyle.THIN);
        estiloDatos1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 2);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        if (numFilas != nombreFilas.length) {
            System.out.println("Error!!!");
            System.out.println("Las líneas registradas en " + nomHoja + " no coinciden con las líneas en 'VATT'");
            numFilas = nombreFilas.length;
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas[i]);
            //System.out.println(nombreFilas[i]+" "+Datos[i][0]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 2);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 2; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 2);
                cell.setCellStyle(estiloDatos);
            }
        }
        //Escribe la suma mensual
        Row rowFin = hoja.createRow(fila);
        fila++;
        for (int j = 0; j < 12; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2);
            Cell cellSumF = row.getCell(j + 2);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());

            cell = rowFin.createCell(j + 2);
            cell.setCellStyle(estiloDatos1);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos1);
        }
        cellTC = rowFin.createCell(1);
        cellTC.setCellValue("Total");
        cellTC.setCellStyle(estiloTituloFila);

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

    }

    static public void creaH1F_2d_float(String titulo, float Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));

            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH1F_2d_float(String titulo, float Datos[][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 2);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 2);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 2; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 2);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

    }

    static public void creaH1FT_2d_float(String titulo, float Datos[][], float Datos1[][][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[], String nombreColumnas1[], String nomRango,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH1FT_2d_float(String titulo, float Datos[][], float Datos1[][][],
            String tituloFilas, String nombreFilas[],
            String tituloColumnas, String nombreColumnas[], String nombreColumnas1[], String nomRango,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Cell cell1 = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;
        //int numCol1=Datos1[0].length*Datos1[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 2);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        for (int j = 0; j < Datos1[0].length; j++) {
            cellTC = row.createCell(j + numCol + 2);
            cellTC.setCellValue(nombreColumnas1[j]);
            cellTC.setCellStyle(estiloTituloTer);

        }

        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 2);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);

            }
            for (int j = 0; j < Datos1[0].length; j++) {
                for (int m = 0; m < Datos1[0][0].length; m++) {
                    cell1 = row.createCell(j * Datos1[0][0].length + m + numCol + 2);
                    cell1.setCellValue(Datos1[i][j][m]);
                    cell1.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);

        Name nombreCel1 = wb.createName();
        nombreCel1.setNameName(nomRango); // Nombre del rango 
        CellReference cellRef1 = new CellReference(cell1.getRowIndex(), cell1.getColumnIndex());
        String reference1 = nomHoja + "!$O$6:" + cellRef1.formatAsString(); // area reference
        nombreCel1.setRefersToFormula(reference1);

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 2; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 2);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

    }

    static public void creaH2F_2d_float(String titulo, float Datos[][],
            String tituloFilas1, String nombreFilas1[],
            String tituloFilas2, String nombreFilas2[],
            String tituloColumnas, String nombreColumnas[],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH2F_2d_float(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH2F_2d_float(String titulo, float Datos[][],
            String tituloFilas1, String nombreFilas1[],
            String tituloFilas2, String nombreFilas2[],
            String tituloColumnas, String nombreColumnas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 3);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            cellTF = row.createCell(2);
            cellTF.setCellValue(nombreFilas2[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 3);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 3; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 3);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaH2F_2d_double(String titulo, double Datos[][],
            String tituloFilas1, String nombreFilas1[],
            String tituloFilas2, String nombreFilas2[],
            String tituloColumnas, String nombreColumnas[],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH2F_2d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH2F_2d_double(String titulo, double Datos[][],
            String tituloFilas1, String nombreFilas1[],
            String tituloFilas2, String nombreFilas2[],
            String tituloColumnas, String nombreColumnas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int j = 0; j < numCol; j++) {
            cellTC = row.createCell(j + 3);
            cellTC.setCellValue(nombreColumnas[j]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            cellTF = row.createCell(2);
            cellTF.setCellValue(nombreFilas2[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int j = 0; j < numCol; j++) {
                cell = row.createCell(j + 3);
                cell.setCellValue(Datos[i][j]);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$C$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < numCol + 3; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < numCol; j++) {
                cell = row.getCell(j + 3);
                cell.setCellStyle(estiloDatos);
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(numCol + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaH2F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH2F_3d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH2F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        int fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 3);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 3);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 3; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    cell = row.getCell(k + 3);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaH3F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, String[] nombreFilas3,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH3F_3d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, nombreFilas3, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH3F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, String[] nombreFilas3,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        int fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(4);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 4);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 3);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell(2);
                cellTF.setCellValue(nombreFilas3[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 4);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$E$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    cell = row.getCell(k + 4);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
 
    static public void creaH2F_3d_long(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, float[] nombreFilas3,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH2F_3d_long(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, nombreFilas3, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH2F_3d_long(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, float[] nombreFilas3,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        int fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        for (int k = 0; k < dim3; k++) {
            //cellTC = row.createCell((int)k+3);
            cellTC = row.createCell((int) k + 4);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                cellTF = row.createCell((int) 3);
                cellTF.setCellValue(nombreFilas3[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                for (int k = 0; k < dim3; k++) {
                    //cell = row.createCell(k+3);
                    cell = row.createCell(k + 4);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        //for (int i = 1; i < dim3 + 3; i++)
        for (int i = 1; i < dim3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    //cell = row.getCell(k+3);
                    cell = row.getCell(k + 4);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaH3F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, double[] InyAnual,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH3F_3d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, tituloFilas3, InyAnual, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH3F_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, double[] InyAnual,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,##0.00;\"-\"", ";");
        String formatoPos = formatoCompleto.nextToken();
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos.setFont(font);
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(4);//cambio
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;

        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 4);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                cell = row.createCell(3);
                cell.setCellValue(InyAnual[i]);
                cell.setCellStyle(estiloDatos2);

                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 4);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 3);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
 
    static public void creaH2F_3d2_long(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH2F_3d2_long(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH2F_3d2_long(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        int fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 3);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 3);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 3; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    cell = row.getCell(k + 3);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaH3F_3d2_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, double[] InyAnual,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH3F_3d2_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, tituloFilas3, InyAnual, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH3F_3d2_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, double[] InyAnual,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,##0.00;\"-\"", ";");
        String formatoPos = formatoCompleto.nextToken();
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos.setFont(font);
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(4);//cambio
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;

        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 4);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);

                cell = row.createCell(3);
                cell.setCellValue(InyAnual[i]);
                cell.setCellStyle(estiloDatos2);

                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 4);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 3);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaHxF_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaHxF_3d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
 
    static public void creaHxF_3d_double(String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("##.#0%", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k + 3);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);//barras
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);//transmisor
                cellTF.setCellStyle(estiloTituloFilaSec);
                for (int k = 0; k < dim3; k++) {
                    if (k == 2) {
                        cell.setCellStyle(estiloDatos1);
                        cell = row.createCell(k + 3);
                        cell.setCellValue(Datos[i][j][k]);
                        cell.setCellStyle(estiloDatos1);
                    } else {
                        cell = row.createCell(k + 3);
                        cell.setCellValue(Datos[i][j][k]);
                        cell.setCellStyle(estiloDatos);
                    }
                }
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 3; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    cell = row.getCell(k + 3);
                    cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void crea_SalidaCU(String titulo,
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloDatos1, String Sub1Datos1, String Sub2Datos1, double Datos1[][],
            String tituloDatos2, String Sub1Datos2, String Sub2Datos2, double Datos2[][],
            String tituloDatos3, String Sub1Datos3, String Sub2Datos3, double Datos3[][][],
            double DatosTot[][],
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_SalidaCU(titulo, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloDatos1, Sub1Datos1, Sub2Datos1, Datos1, tituloDatos2, Sub1Datos2, Sub2Datos2, Datos2, tituloDatos3, Sub1Datos3, Sub2Datos3, Datos3, DatosTot, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_SalidaCU(String titulo,
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloDatos1, String Sub1Datos1, String Sub2Datos1, double Datos1[][],
            String tituloDatos2, String Sub1Datos2, String Sub2Datos2, double Datos2[][],
            String tituloDatos3, String Sub1Datos3, String Sub2Datos3, double Datos3[][][],
            double DatosTot[][],
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        CellStyle estiloTitulo1 = wb.createCellStyle();
        estiloTitulo1.setFont(fontTitulo);
        estiloTitulo1.setBorderRight(BorderStyle.THIN);
        estiloTitulo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderLeft(BorderStyle.THIN);
        estiloTitulo1.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderBottom(BorderStyle.THIN);
        estiloTitulo1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderTop(BorderStyle.THIN);
        estiloTitulo1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,##0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("0.00%", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer(formatoDatos, ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);
        estiloDatos3.setBorderRight(BorderStyle.THIN);
        estiloDatos3.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        // Dimensiones del arreglo
        int dim1 = Datos3.length;//barras
        int dim2 = Datos3[0].length;//Tx
        int dim3 = Datos3[0][0].length;//N de cargos
        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        for (int i = 3; i < 3 + dim3 * 3; i++) {
            cellTC = row.createCell(i);
            cellTC.setCellStyle(estiloTitulo1);
        }
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloDatos1);
        cellTC.setCellStyle(estiloTitulo1);
        CellReference cellRefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        CellReference cellRefF = new CellReference(cellTC.getRowIndex(), 2 + dim3);
        String reference = cellRefI.formatAsString() + ":" + cellRefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = row.createCell(3 + dim3);
        cellTC.setCellValue(tituloDatos2);
        cellTC.setCellStyle(estiloTitulo1);
        cellRefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        cellRefF = new CellReference(cellTC.getRowIndex(), 2 + dim3 * 2);
        reference = cellRefI.formatAsString() + ":" + cellRefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = row.createCell(3 + dim3 * 2);
        cellTC.setCellValue(tituloDatos3);
        cellTC.setCellStyle(estiloTitulo1);
        cellRefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        cellRefF = new CellReference(cellTC.getRowIndex(), 3 + dim3 * 3);
        reference = cellRefI.formatAsString() + ":" + cellRefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);

        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(Sub1Datos1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(Sub2Datos1);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3 + dim3);
        cellTC.setCellValue(Sub1Datos2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4 + dim3);
        cellTC.setCellValue(Sub2Datos2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3 + dim3 * 2);
        cellTC.setCellValue(Sub1Datos3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4 + dim3 * 2);
        cellTC.setCellValue(Sub2Datos3);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3 + dim3 * 3);
        cellTC.setCellValue("Pago Total");
        cellTC.setCellStyle(estiloTituloFila);
        // Titulos Filas y Datos
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);//barras
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell((int) 2);
                cellTF.setCellValue(nombreFilas2[j]);//transmisor
                cellTF.setCellStyle(estiloTituloFilaSec);

                for (int k = 0; k < dim3; k++) {
                    cell = row.createCell(k + 3);
                    cell.setCellValue(Datos1[i][k]);
                    cell.setCellStyle(estiloDatos1);

                    cell = row.createCell(k + 3 + dim3);
                    cell.setCellValue(Datos2[i][k]);
                    cell.setCellStyle(estiloDatos2);

                    cell = row.createCell(k + 3 + dim3 * 2);
                    cell.setCellValue(Datos3[i][j][k]);
                    cell.setCellStyle(estiloDatos3);
                }
                cell = row.createCell(3 + dim3 * 3);
                cell.setCellValue(DatosTot[i][j]);
                cell.setCellStyle(estiloDatos3);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 * 3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1 * dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            cell = row.getCell(4);
            cell.setCellStyle(estiloDatos1);
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaDetallePIny(int mes,
            String titulo, double Datos1[][][][], double Datos[][][], double Datos2[][][],
            double Datos3[][][], double Datos4[][][],
            double DatosTot1[][], double DatosTot2[][],
            double DatosTot3[][], double DatosTot4[][],
            double CapExce[], double FC[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, float[] DatosFilas3,
            //String tituloFilas4, double[] DatosFilas4,
            //String tituloFilas5, double[][] DatosFilas5,
            //String tituloFilas6, double[][] DatosFilas6,
            String nomLibro, String nombreMes, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            String nomHoja="Detalle"+nombreMes;
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaDetallePIny(mes, titulo, Datos1, Datos, Datos2, Datos3, Datos4, DatosTot1, DatosTot2, DatosTot3, DatosTot4, CapExce, FC, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, DatosFilas3, wb, nombreMes, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaDetallePIny(int mes,
            String titulo, double Datos1[][][][], double Datos[][][], double Datos2[][][],
            double Datos3[][][], double Datos4[][][],
            double DatosTot1[][], double DatosTot2[][],
            double DatosTot3[][], double DatosTot4[][],
            double CapExce[], double FC[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, float[] DatosFilas3,
            XSSFWorkbook wb, String nombreMes, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Cell cell1 = null;
        Cell cellT = null;
        Row row = null;
        Row row1 = null;
        String nomHoja = "Detalle" + nombreMes;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        CellStyle estiloTitulo1 = wb.createCellStyle();
        estiloTitulo1.setFont(fontTitulo);
        estiloTitulo1.setBorderRight(BorderStyle.THIN);
        estiloTitulo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderLeft(BorderStyle.THIN);
        estiloTitulo1.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderBottom(BorderStyle.THIN);
        estiloTitulo1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderTop(BorderStyle.THIN);
        estiloTitulo1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("0.000%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;
        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        Row rowTmp = row;

        row = hoja.createRow(fila);
        fila++;

        //for(int i=6;i<6+7*dim2+4;i++){
        for (int i = 3; i < 3 + 7 * dim2 + 4; i++) {
            cellTC = row.createCell(i);
            cellTC.setCellStyle(estiloTituloFila);
        }

        for (int i = 0; i < dim2; i++) {
            //cellTC = row.createCell(6+i*4);
            cellTC = row.createCell(3 + i * 4);
            cellTC.setCellValue(nombreFilas2[i]);
            cellTC.setCellStyle(estiloTitulo1);
            CellReference RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
            //CellReference RefF = new CellReference(cellTC.getRowIndex(), 6+i*4+3);
            CellReference RefF = new CellReference(cellTC.getRowIndex(), 3 + i * 4 + 3);
            String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
        }

        int a = 1;
        for (int aux = 0; aux < 3; aux++) {
            for (int j = 0; j < dim2; j++) {
                //cellTF = row.createCell(6+(4+aux)*dim2+(aux+1)+j);
                cellTF = row.createCell(3 + (4 + aux) * dim2 + (aux + 1) + j);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFila);
                CellReference RefI = new CellReference(cellTF.getRowIndex(), cellTF.getColumnIndex());
                CellReference RefF = new CellReference(cellTF.getRowIndex() + 1, cellTF.getColumnIndex());
                String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
                hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
            }
            //cellTF = row.createCell(6+(4+aux)*dim2+(aux+1)+dim2);
            cellTF = row.createCell(3 + (4 + aux) * dim2 + (aux + 1) + dim2);
            cellTF.setCellValue("Total");
            cellTF.setCellStyle(estiloTituloFila);
            a++;
            CellReference RefI = new CellReference(cellTF.getRowIndex(), cellTF.getColumnIndex());
            CellReference RefF = new CellReference(cellTF.getRowIndex() + 1, cellTF.getColumnIndex());
            String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
        }
        //cellTF = row.createCell(6+dim2*4);
        cellTF = row.createCell(3 + dim2 * 4);
        cellTF.setCellValue("Total General");
        cellTF.setCellStyle(estiloTituloFila);
        CellReference RefI = new CellReference(cellTF.getRowIndex(), cellTF.getColumnIndex());
        CellReference RefF = new CellReference(cellTF.getRowIndex() + 1, cellTF.getColumnIndex());
        String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        //cellTC = row.createCell(3);
        //cellTC.setCellValue(tituloFilas4);
        //cellTC.setCellStyle(estiloTituloFila);
        //cellTC = row.createCell(4);
        //cellTC.setCellValue(tituloFilas5);
        //cellTC.setCellStyle(estiloTituloFila);
        //cellTC = row.createCell(5);
        //cellTC.setCellValue(tituloFilas6);
        //cellTC.setCellStyle(estiloTituloFila);
        // Titulos Filas y Datos
        for (int i = 0; i < dim2; i++) {
            //cellTF = row.createCell(6+i*4);
            cellTF = row.createCell(3 + i * 4);
            cellTF.setCellValue("N");
            cellTF.setCellStyle(estiloTituloFila);

            //cellTF = row.createCell(6+i*4+1);
            cellTF = row.createCell(3 + i * 4 + 1);
            cellTF.setCellValue("A");
            cellTF.setCellStyle(estiloTituloFila);

            //cellTF = row.createCell(6+i*4+2);
            cellTF = row.createCell(3 + i * 4 + 2);
            cellTF.setCellValue("S");
            cellTF.setCellStyle(estiloTituloFila);

            //cellTF = row.createCell(6+i*4+3);
            cellTF = row.createCell(3 + i * 4 + 3);
            cellTF.setCellValue("Total");
            cellTF.setCellStyle(estiloTituloFila);
        }

        //for(int i=6+4*dim2;i<6+7*dim2+4;i++){
        for (int i = 3 + 4 * dim2; i < 3 + 7 * dim2 + 4; i++) {
            cellTC = row.createCell(i);
            cellTC.setCellStyle(estiloTituloFila);
        }

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            cellTF = row.createCell(2);
            cellTF.setCellValue(DatosFilas3[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            //cellTF = row.createCell(3);
            //cellTF.setCellValue(DatosFilas4[i]);
            //cellTF.setCellStyle(estiloDatos1);
            //cellTF = row.createCell(4);
            //cellTF.setCellValue(DatosFilas5[i][mes]);
            //cellTF.setCellStyle(estiloDatos1);
            //cellTF = row.createCell(5);
            //cellTF.setCellValue(DatosFilas6[i][mes]);
            //cellTF.setCellStyle(estiloDatos1);
            //Datos
            int aux = dim2 + 1;
            int aux1 = 0;
            for (int j = 0; j < dim2; j++) {
                for (int z = 0; z < 3; z++) {
                    //cell = row.createCell(j+aux1*3+z+6);
                    cell = row.createCell(j + aux1 * 3 + z + 3);//6
                    cell.setCellStyle(estiloDatos);
                    cell.setCellValue(Datos1[i][j][z][mes]);
                }
                //cellT = row.createCell(j+aux1*3+9);
                cellT = row.createCell(j + aux1 * 3 + 6);//6
                cellT.setCellStyle(estiloDatos2);
                cellT.setCellValue(Datos[i][j][mes]);
                aux1++;

                cell.setCellStyle(estiloDatos);
                //cell = row.createCell(j+6+4*aux-3);//10
                cell = row.createCell(j + 3 + 4 * aux - 3);
                cell.setCellValue(Datos2[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                //cell = row.createCell(j+6+5*aux-3);//14
                cell = row.createCell(j + 3 + 5 * aux - 3);
                cell.setCellValue(Datos3[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                //cell = row.createCell(j+6+6*aux-3);//18
                cell = row.createCell(j + 3 + 6 * aux - 3);
                cell.setCellValue(Datos4[i][j][mes]);
                cell.setCellStyle(estiloDatos);
            }
            //cellT = row.createCell(5+4*aux-3);//9
            cellT = row.createCell(2 + 4 * aux - 3);
            cellT.setCellStyle(estiloDatos2);
            cellT.setCellValue(DatosTot1[i][mes]);

            cellT.setCellStyle(estiloDatos2);
            //cellT = row.createCell(5+5*aux-3);//13
            cellT = row.createCell(2 + 5 * aux - 3);
            cellT.setCellValue(DatosTot2[i][mes]);

            cellT.setCellStyle(estiloDatos2);
            //cellT = row.createCell(5+aux*6-3);//17
            cellT = row.createCell(2 + aux * 6 - 3);
            cellT.setCellValue(DatosTot3[i][mes]);

            cellT.setCellStyle(estiloDatos2);
            //cellT = row.createCell(5+aux*7-3);//21
            cellT = row.createCell(2 + aux * 7 - 3);
            cellT.setCellValue(DatosTot4[i][mes]);
            cellT.setCellStyle(estiloDatos2);
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        //Escribe la suma mensual
        Row rowFin = hoja.createRow(fila);
        fila++;
        //for(int j=6;j<dim2*7+4;j++){
        for (int j = 3; j < dim2 * 7 + 4; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j);
            Cell cellSumF = row.getCell(j);
            RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = rowFin.createCell(j);
            //cell.setCellStyle(estiloDatos);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }

        cellTC = rowFin.createCell(2);
        cellTC.setCellValue("Total");
        cellTC.setCellStyle(estiloTituloFila);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        //for (int i = 1; i < dim2*4+6; i++)
        for (int i = 1; i < dim2 * 4 + 3; i++) {
            hoja.autoSizeColumn(i);
        }
        //Capacidad Exceptuada
        row1 = hoja.getRow(3);
        Cell cellTC3 = row1.createCell(1);
        Cell cellTC4 = row1.createCell(2);
        cellTC3.setCellValue("Capacidad Conjunta Exceptuada");
        cellTC3.setCellStyle(estiloDatos);
        cellTC4.setCellValue(CapExce[mes]);
        cellTC4.setCellStyle(estiloDatos3);

        row1 = hoja.getRow(4);
        Cell cellTC5 = row1.createCell(1);
        Cell cellTC6 = row1.createCell(2);
        cellTC5.setCellValue("Factor de Corrección");
        cellTC5.setCellStyle(estiloDatos);
        cellTC6.setCellValue(FC[mes]);
        cellTC6.setCellStyle(estiloDatos3);
        // Titulo Principal

        //for(int i=6;i<6+7*dim2+4;i++){
        for (int i = 3; i < 3 + 7 * dim2 + 4; i++) {
            cellTC = rowTmp.createCell(i);
            cellTC.setCellStyle(estiloTitulo1);
        }
        //cellTC = rowTmp.createCell(6);
        cellTC = rowTmp.createCell(3);
        cellTC.setCellValue("Pagos por Inyección");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        //RefF = new CellReference(cellTC.getRowIndex(),5+(4*dim2+1) );
        RefF = new CellReference(cellTC.getRowIndex(), 2 + (4 * dim2 + 1));
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        //cellTC = rowTmp.createCell(6+(4*dim2+1));
        cellTC = rowTmp.createCell(3 + (4 * dim2 + 1));
        cellTC.setCellValue("Exención MGNC");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        //RefF = new CellReference(cellTC.getRowIndex(),5+(5*dim2+2));
        RefF = new CellReference(cellTC.getRowIndex(), 2 + (5 * dim2 + 2));
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        //cellTC = rowTmp.createCell(6+(5*dim2+2));
        cellTC = rowTmp.createCell(3 + (5 * dim2 + 2));
        cellTC.setCellValue("Ajuste por MGNC");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        //RefF = new CellReference(cellTC.getRowIndex(),5+(6*dim2+3));
        RefF = new CellReference(cellTC.getRowIndex(), 2 + (6 * dim2 + 3));
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        //cellTC = rowTmp.createCell(6+(6*dim2+3));
        cellTC = rowTmp.createCell(3 + (6 * dim2 + 3));
        cellTC.setCellValue("Pago Incluyendo Ajuste por MGNC");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        //RefF = new CellReference(cellTC.getRowIndex(),5+(7*dim2+4));
        RefF = new CellReference(cellTC.getRowIndex(), 2 + (7 * dim2 + 4));
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaPIny(int mes,
            String titulo, double Datos1[][][], double Datos2[][][],
            double Datos3[][][],
            double DatosTot1[][], double DatosTot2[][],
            double DatosTot3[][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nombreMes, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaPIny(mes, titulo, Datos1, Datos2, Datos3, DatosTot1, DatosTot2, DatosTot3, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, wb, nombreMes, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nombreMes);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaPIny(int mes,
            String titulo, double Datos1[][][], double Datos2[][][],
            double Datos3[][][],
            double DatosTot1[][], double DatosTot2[][],
            double DatosTot3[][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nombreMes, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        String nomHoja = nombreMes;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        CellStyle estiloTitulo1 = wb.createCellStyle();
        estiloTitulo1.setFont(fontTitulo);
        estiloTitulo1.setBorderRight(BorderStyle.THIN);
        estiloTitulo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderLeft(BorderStyle.THIN);
        estiloTitulo1.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderBottom(BorderStyle.THIN);
        estiloTitulo1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderTop(BorderStyle.THIN);
        estiloTitulo1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer(formatoDatos, ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderTop(BorderStyle.THIN);
        estiloDatos1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setBorderBottom(BorderStyle.THIN);
        estiloDatos1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        Row rowTmp = row;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);

        // Titulos Filas y Datos
        int a = 1;
        for (int aux = 0; aux < 3; aux++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = row.createCell((int) 1 + j + aux * dim2 + a);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFila);
            }
            cellTF = row.createCell(1 + a * (dim2 + 1));
            cellTF.setCellValue("Total");
            cellTF.setCellStyle(estiloTituloFila);

            a++;
        }
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            cellTF = row.createCell(2);
            //Datos
            int aux = dim2 + 1;
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 2 + aux);//6
                cell.setCellValue(Datos2[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 2 + aux * 2);//10
                cell.setCellValue(Datos3[i][j][mes]);

                cell.setCellStyle(estiloDatos);

            }

            cell = row.createCell(1 + aux);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(1 + aux * 2);//9
            cell.setCellValue(DatosTot2[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(1 + aux * 3);//13
            cell.setCellValue(DatosTot3[i][mes]);
            cell.setCellStyle(estiloDatos2);

        }
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim2 * 3 + 6; i++) {
            hoja.autoSizeColumn(i);
        }

        // Titulo Principal
        for (int i = 2; i < 2 + 3 * (dim2 + 1); i++) {
            cellTC = rowTmp.createCell(i);
            cellTC.setCellStyle(estiloTitulo1);
        }

        cellTC = rowTmp.getCell(2);
        cellTC.setCellValue("Pagos por Inyección");
        cellTC.setCellStyle(estiloTitulo1);
        CellReference RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        CellReference RefF = new CellReference(cellTC.getRowIndex(), 1 + (dim2 + 1));
        String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = rowTmp.getCell(2 + (dim2 + 1));
        //cellTC.setCellValue("Ajuste por MGNC");
        cellTC.setCellValue("Exenciones");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        RefF = new CellReference(cellTC.getRowIndex(), 1 + (dim2 + 1) * 2);
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = rowTmp.getCell(2 + (dim2 + 1) * 2);
        //cellTC.setCellValue("Pago Incluyendo Ajuste por MGNC");
        cellTC.setCellValue("Pago Incluyendo Exenciones");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        RefF = new CellReference(cellTC.getRowIndex(), 1 + (dim2 + 1) * 3);
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        //Escribe la suma mensual
        Row rowFin = hoja.createRow(fila);
        fila++;
        for (int j = 0; j < (dim2 + 1) * 3; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2);
            Cell cellSumF = row.getCell(j + 2);
            RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = rowFin.createCell(j + 2);
            cell.setCellStyle(estiloDatos1);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos1);
        }
        cellTC = rowFin.createCell(1);
        cellTC.setCellValue("Total");
        cellTC.setCellStyle(estiloTituloFila);

        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
    }
    
    static public void creaPagosRet_3d_long(String titulo, long Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaPagosRet_3d_long(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaPagosRet_3d_long(String titulo, long Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        for (int k = 0; k < dim3; k++) {
            cellTC = row.createCell((int) k * dim2 + 2);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        row = hoja.createRow(fila);
        fila++;
        for (int k = 0; k < dim3; k++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = row.createCell(2 + j + k * dim2);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFila);
            }
        }
        // Titulos Filas y Datos
        short filaTmp = fila;

        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            for (int k = 0; k < dim3; k++) {
                for (int j = 0; j < dim2; j++) {
                    cell = row.createCell(2 + j + k * dim2);
                    cell.setCellValue(Datos[i][j][k]);
                    cell.setCellStyle(estiloDatos);
                }
            }

        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 3; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim1; i++) {
            for (int j = 0; j < dim2; j++) {
                row = hoja.getRow(fila);
                fila++;
                for (int k = 0; k < dim3; k++) {
                    //cell = row.getCell(k+3);
                    //cell.setCellStyle(estiloDatos);
                }
            }
        }
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaLiquidacionMesIny(int mes,
            String titulo, double Datos1[][][],
            double Datos2[][][], double Datos3[][][],
            double DatosTot1[][], double DatosTot2[][], double DatosTot3[][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas4, float[] DatosFilas4,
            String tituloFilas5, double[] DatosFilas5,
            String tituloFilas3, double[][] DatosFilas3,
            String tituloFilas6, double[][] DatosFilas6,
            String nomLibro, String nomHoja, int Ano, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaLiquidacionMesIny(mes, titulo, Datos1, Datos2, Datos3, DatosTot1, DatosTot2, DatosTot3, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas4, DatosFilas4, tituloFilas5, DatosFilas5, tituloFilas3, DatosFilas3, tituloFilas6, DatosFilas6, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static public void creaLiquidacionMesIny(int mes,
            String titulo, double Datos1[][][],
            double Datos2[][][], double Datos3[][][],
            double DatosTot1[][], double DatosTot2[][], double DatosTot3[][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas4, float[] DatosFilas4,
            String tituloFilas5, double[] DatosFilas5,
            String tituloFilas3, double[][] DatosFilas3,
            String tituloFilas6, double[][] DatosFilas6,
            XSSFWorkbook wb, String nomHoja, int Ano, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        CellStyle estiloTitulo1 = wb.createCellStyle();
        estiloTitulo1.setFont(fontTitulo);
        estiloTitulo1.setBorderRight(BorderStyle.THIN);
        estiloTitulo1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderLeft(BorderStyle.THIN);
        estiloTitulo1.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderBottom(BorderStyle.THIN);
        estiloTitulo1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setBorderTop(BorderStyle.THIN);
        estiloTitulo1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTitulo1.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("0.00;[Red]-0.00", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(MESES[mes] + " " + Ano);
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue("(Valores en $ indexados a " + MESES[mes] + " " + Ano + ")");
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        for (int i = 6; i < 6 + 3 * (dim2 + 1); i++) {
            cellTC = row.createCell(i);
            cellTC.setCellStyle(estiloTitulo1);
        }

        cellTC = row.createCell(6);
        cellTC.setCellValue("Pagos por Inyección");
        cellTC.setCellStyle(estiloTitulo1);
        CellReference RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        CellReference RefF = new CellReference(cellTC.getRowIndex(), 5 + (dim2 + 1));
        String reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = row.createCell(6 + (dim2 + 1));
        cellTC.setCellValue("Ajuste por MGNC");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        RefF = new CellReference(cellTC.getRowIndex(), 5 + (dim2 + 1) * 2);
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        cellTC = row.createCell(6 + (dim2 + 1) * 2);
        cellTC.setCellValue("Pago Incluyendo Ajuste por MGNC");
        cellTC.setCellStyle(estiloTitulo1);
        RefI = new CellReference(cellTC.getRowIndex(), cellTC.getColumnIndex());
        RefF = new CellReference(cellTC.getRowIndex(), 5 + (dim2 + 1) * 3);
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas4);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas5);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(4);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(5);
        cellTC.setCellValue(tituloFilas6);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        int a = 1;
        for (int aux = 0; aux < 3; aux++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = row.createCell((int) 5 + j + aux * dim2 + a);
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFila);
            }
            cellTF = row.createCell(5 + a * (dim2 + 1));
            cellTF.setCellValue("Total");
            cellTF.setCellStyle(estiloTituloFila);
            a++;
        }
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            cellTF = row.createCell(2);
            cellTF.setCellValue(DatosFilas4[i]);
            cellTF.setCellStyle(estiloDatos1);
            cellTF = row.createCell(3);
            cellTF.setCellValue(DatosFilas5[i]);
            cellTF.setCellStyle(estiloDatos1);

            cellTF = row.createCell(4);
            cellTF.setCellValue(DatosFilas3[i][mes]);
            cellTF.setCellStyle(estiloDatos1);

            cellTF = row.createCell(5);
            cellTF.setCellValue(DatosFilas6[i][mes]);
            cellTF.setCellStyle(estiloDatos1);

            //Datos
            int aux = dim2 + 1;
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 6);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 6 + aux);//6
                cell.setCellValue(Datos2[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 6 + aux * 2);//10
                cell.setCellValue(Datos3[i][j][mes]);

                cell.setCellStyle(estiloDatos);
            }
            cell = row.createCell(5 + aux);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(5 + aux * 2);//9
            cell.setCellValue(DatosTot2[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(5 + aux * 3);//13
            cell.setCellValue(DatosTot3[i][mes]);
            cell.setCellStyle(estiloDatos2);
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        reference = nomHoja + "!$G$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        //Escribe la suma mensual
        Row rowFin = hoja.createRow(fila);
        fila++;
        for (int j = 0; j < (dim2 + 1) * 3; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 6);
            Cell cellSumF = row.getCell(j + 6);
            RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = rowFin.createCell(j + 6);
            cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }
        cellTC = rowFin.createCell(5);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = rowFin.createCell(4);
        cellTC.setCellValue("Total");
        cellTC.setCellStyle(estiloTituloFila);
        RefI = new CellReference(cellTC.getRowIndex(), 4);
        RefF = new CellReference(cellTC.getRowIndex(), 5);
        reference = nomHoja + "!" + RefI.formatAsString() + ":" + RefF.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        hoja.autoSizeColumn(1);
        hoja.autoSizeColumn(2);
        hoja.autoSizeColumn(3);
        hoja.autoSizeColumn(4);
        hoja.autoSizeColumn(5);
        for (int i = 6; i < 6 + (dim2 + 1) * 3; i++) {
            hoja.setColumnWidth(i, 6 * 500);
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaLiquidacionMes(int mes,
            String titulo,
            double Datos1[][][],
            double Datos2[][][],
            double Datos3[][][],
            double Datos4[][][],
            double DatosTot1[][],
            double DatosTot2[][],
            double DatosTot3[][],
            double DatosTot4[][],
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String nomgrupo1,
            String nomgrupo2,
            String nomgrupo3,
            String nomgrupo4,
            String nomLibro,
            String nomHoja,
            int Ano,
            String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaLiquidacionMes(mes, titulo, Datos1, Datos2, Datos3, Datos4, DatosTot1, DatosTot2, DatosTot3, DatosTot4, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, nomgrupo1, nomgrupo2, nomgrupo3, nomgrupo4, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaLiquidacionMes(int mes,
            String titulo,
            double Datos1[][][],
            double Datos2[][][],
            double Datos3[][][],
            double Datos4[][][],
            double DatosTot1[][],
            double DatosTot2[][],
            double DatosTot3[][],
            double DatosTot4[][],
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String nomgrupo1,
            String nomgrupo2,
            String nomgrupo3,
            String nomgrupo4,
            XSSFWorkbook wb,
            String nomHoja,
            int Ano,
            String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderLeft(BorderStyle.THIN);
        estiloTituloFila.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderLeft(BorderStyle.THIN);
        estiloTituloFilaSec.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderRight(BorderStyle.THIN);
        estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(MESES[mes] + " " + Ano);
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue("(Valores en $ indexados a " + MESES[mes] + " " + Ano + ")");
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC.setCellStyle(estiloTitulo);

        cellTC = row.createCell(1);
        cellTC.setCellValue(nomgrupo1);
        cellTC.setCellStyle(estiloTitulo2);
        cellTC = row.createCell(2 + (dim2 + 2));
        cellTC.setCellValue(nomgrupo2);
        cellTC.setCellStyle(estiloTitulo2);
        cellTC = row.createCell(3 + (dim2 + 2) * 2);
        cellTC.setCellValue(nomgrupo3);
        cellTC.setCellStyle(estiloTitulo2);
        cellTC = row.createCell(4 + (dim2 + 2) * 3);
        cellTC.setCellValue(nomgrupo4);
        cellTC.setCellStyle(estiloTitulo2);

        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.getRow(7);

        for (int aux = 0; aux < 4; aux++) {//4 tablas
            for (int i = 2 + aux * (dim2 + 3); i < -1 + (aux + 1) * (dim2 + 3); i++) {
                cellTF = row.createCell(i);
                cellTF.setCellStyle(estiloTituloFila);
            }
            Cell cellTCom1 = row.getCell(2 + aux * (dim2 + 3));
            Cell cellTCom2 = row.getCell(1 + aux * (dim2 + 3) + dim2);
            CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
            CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
            String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
            cellTCom1.setCellValue(tituloFilas2);
        }

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Datos
        for (int b = 0; b < 4; b++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = row.createCell((int) 2 + j + b * (dim2 + 3));
                cellTF.setCellValue(nombreFilas2[j]);
                cellTF.setCellStyle(estiloTituloFila);
            }
            cellTF = row.createCell((2 + dim2) * (b + 1) + b);
            cellTF.setCellValue("Total");
            cellTF.setCellStyle(estiloTituloFila);
            cellTC = row.createCell(1 + (dim2 + 3) * b);
            cellTC.setCellValue(tituloFilas1);
            cellTC.setCellStyle(estiloTituloFila);
        }
        short filaTmp = fila;
        int aux = dim2 + 3;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            //Datos
            for (int j = 0; j < dim2; j++) {

                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 2 + aux);//8
                cell.setCellValue(Datos2[i][j][mes]);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 2 + aux * 2);//14
                cell.setCellValue(Datos3[i][j][mes]);
                cell.setCellStyle(estiloDatos);

                cell.setCellStyle(estiloDatos);
                cell = row.createCell(j + 2 + aux * 3);//14
                cell.setCellValue(Datos4[i][j][mes]);
                cell.setCellStyle(estiloDatos);

            }
            cell = row.createCell(aux - 1);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(aux * 2 - 1);//10
            cell.setCellValue(DatosTot2[i][mes]);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(aux * 3 - 1);//15
            cell.setCellValue(DatosTot3[i][mes]);
            cell.setCellStyle(estiloDatos2);

            cell.setCellStyle(estiloDatos2);
            cell = row.createCell(aux * 4 - 1);//
            cell.setCellValue(DatosTot4[i][mes]);
            cell.setCellStyle(estiloDatos2);

            cellTF = row.createCell(1 + aux);//6
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            cellTF = row.createCell(1 + aux * 2);//11
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            cellTF = row.createCell(1 + aux * 3);//11
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

        }

        //Escribe totales
        rowFin = hoja.createRow(fila);
        fila++;
        for (int b = 0; b < 4; b++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = rowFin.createCell(1 + b * aux);
                cellTF.setCellValue("Total General");
                cellTF.setCellStyle(estiloTituloFila);
            }
        }
        for (int b = 0; b < 4; b++) {
            for (int j = 0; j < dim2 + 1; j++) {
                Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2 + b * aux);
                Cell cellSumF = row.getCell(j + 2 + b * aux);
                CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
                CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());

                cell = rowFin.createCell(j + 2 + b * aux);
                cell.setCellStyle(estiloDatos4);
                cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
                cell.setCellStyle(estiloDatos4);
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 4; i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        //hoja.autoSizeColumn(i);
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        //for (int i = 0; i < dim2; i++) {
        //  row = hoja.getRow(fila); fila++;
        //for (int j = 0; j < dim2; j++) {
        //  cell = row.getCell(j+3);
        //cell.setCellStyle(estiloDatos);
        //}
        //}

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        //cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        //reference = nomHoja+"!$B$2:"+cellRef.formatAsString(); // area reference
        //hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    static public void creaProrrataMes(int mes,
            String titulo, double Datos[][][], String tituloDatos,
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, int[] DatosFilas3,
            String nomLibro, String nombreMes, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaProrrataMes(mes, titulo, Datos, tituloDatos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, DatosFilas3, wb, nombreMes, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nombreMes);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaProrrataMes(int mes,
            String titulo, double Datos[][][], String tituloDatos,
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, int[] DatosFilas3,
            XSSFWorkbook wb, String nombreMes, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        String nomHoja = nombreMes;
        int fila = 0;
        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###0.000%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setBorderRight(BorderStyle.THIN);
        estiloDatos3.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos3.setFont(font);
        // Dimensiones del arreglo
        int dim1 = Datos.length;//numero lineas
        int dim2 = Datos[0].length;//numero clientes
        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(tituloDatos);
        cellTC.setCellStyle(estiloTituloFila);
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim2; i++) {//numero de Clientes
            for (int j = 0; j < dim1; j++) {//numero de lineas
                row = hoja.createRow(fila);
                fila++;
                cell = row.createCell(1);
                cell.setCellValue(nombreFilas1[i]);
                cell.setCellStyle(estiloTituloFilaSec);

                cell = row.createCell(2);
                cell.setCellValue(nombreFilas2[j]);
                cell.setCellStyle(estiloDatos1);

                cell = row.createCell(3);
                cell.setCellValue(DatosFilas3[j]);
                cell.setCellStyle(estiloTituloFilaSec);

                cell = row.createCell(4);
                cell.setCellValue(Datos[j][i][mes]);
                cell.setCellStyle(estiloDatos3);
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim2 + 3; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
    }
    
    static public void creaTabla1C_float(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaTabla1C_float(mes, titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaTabla1C_float(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;

        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.00;[Red]-#,###,##0.00;\"-\"");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 2);//2
                cell.setCellValue(Datos[i][j][mes]);
                cell.setCellStyle(estiloDatos1);
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        //nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        //hoja.autoSizeColumn(i);
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaTabla2C_double(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreTx,
            String tituloFilas3, double[][] DatosFilas3,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaTabla2C_double(mes, titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreTx, tituloFilas3, DatosFilas3, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaTabla2C_double(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreTx,
            String tituloFilas3, double[][] DatosFilas3,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;

        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("0.00;[Red]-0.00", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 3 + j);
            cellTF.setCellValue(nombreTx[j]);
            cellTF.setCellStyle(estiloTituloFila);

        }

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            cellTF = row.createCell(2);
            cellTF.setCellValue(DatosFilas3[i][mes]);
            cellTF.setCellStyle(estiloDatos1);

            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 3);//2
                cell.setCellStyle(estiloDatos2);
                cell.setCellValue(Datos[i][j][mes]);
                cell.setCellStyle(estiloDatos2);
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) //hoja.setColumnWidth(i, 5*700);
        {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
                //cell.setCellStyle(estiloDatos);
            }
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaTabla1C_long(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaTabla1C_long(mes, titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaTabla1C_long(int mes,
            String titulo, double Datos[][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;

        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }
        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos2);
                cell.setCellValue(Datos[i][j][mes]);
                cell.setCellStyle(estiloDatos2);
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        //hoja.setColumnWidth(1, 8*1000);
        for (int i = 1; i < (dim2 + 3) * 3; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
                //cell.setCellStyle(estiloDatos);
            }
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaProrrataMes_long(int mes,
            String titulo,
            double Datos[][][],
            String tituloDatos,
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String tituloFilas3,
            int[] DatosFilas3,
            String tituloFilas4,
            double DatosFilas4[][][],
            String nomLibro,
            String nombreMes,
            String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaProrrataMes_long(mes, titulo, Datos, tituloDatos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, DatosFilas3, tituloFilas4, DatosFilas4, wb, nombreMes, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nombreMes);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaProrrataMes_long(int mes,
            String titulo,
            double Datos[][][],
            String tituloDatos,
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String tituloFilas3,
            int[] DatosFilas3,
            String tituloFilas4,
            double DatosFilas4[][][],
            XSSFWorkbook wb,
            String nombreMes,
            String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        String nomHoja = nombreMes;
        int fila = 0;
        
        hoja = wb.createSheet(nomHoja);
        
        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);
        
        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);
        
        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);
        
        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);
        
        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);
        
        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer(formatoDatos, ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);
        
        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###0.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setBorderRight(BorderStyle.THIN);
        estiloDatos3.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;//numero lineas
        int dim2 = Datos[0].length;//numero clientes

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(tituloDatos);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(5);
        cellTC.setCellValue(tituloFilas4);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim2; i++) {//numero de Clientes
            for (int j = 0; j < dim1; j++) {//numero de lineas
                row = hoja.createRow(fila);
                fila++;
                cell = row.createCell(1);
                cell.setCellValue(nombreFilas1[i]);
                cell.setCellStyle(estiloTituloFilaSec);
                
                cell = row.createCell(2);
                cell.setCellValue(nombreFilas2[j]);
                cell.setCellStyle(estiloDatos1);
                
                cell = row.createCell(3);
                cell.setCellValue(DatosFilas3[j]);
                cell.setCellStyle(estiloTituloFilaSec);
                
                cell = row.createCell(4);
                cell.setCellValue(Datos[j][i][mes]);
                cell.setCellStyle(estiloDatos1);
                
                cell = row.createCell(5);
                cell.setCellValue(DatosFilas4[j][i][mes]);
                cell.setCellStyle(estiloDatos1);
                
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim2 + 3; i++) {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
                //cell.setCellStyle(estiloDatos);
            }
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
    }

    static public void creaLiquidacion(int mes,
            String titulo,
            double Datos1[][][],
            double Datos2[][][],
            double Datos3[][][],
            double Datos4[][][],
            double DatosTot1[][],
            double DatosTot2[][],
            double DatosTot3[][],
            double DatosTot4[][],
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String nomgrupo1,
            String nomgrupo2,
            String nomgrupo3,
            String nomgrupo4,
            String nomLibro,
            String nomHoja,
            int Ano,
            String FechaPago,
            String formatoDatos,
            String nota1,
            String nota2) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaLiquidacion(mes, titulo, Datos1, Datos2, Datos3, Datos4, DatosTot1, DatosTot2, DatosTot3, DatosTot4, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, nomgrupo1, nomgrupo2, nomgrupo3, nomgrupo4, wb, nomHoja, Ano, FechaPago, formatoDatos, nota1, nota2);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaLiquidacion(int mes,
            String titulo,
            double Datos1[][][],
            double Datos2[][][],
            double Datos3[][][],
            double Datos4[][][],
            double DatosTot1[][],
            double DatosTot2[][],
            double DatosTot3[][],
            double DatosTot4[][],
            String tituloFilas1,
            String[] nombreFilas1,
            String tituloFilas2,
            String[] nombreFilas2,
            String nomgrupo1,
            String nomgrupo2,
            String nomgrupo3,
            String nomgrupo4,
            XSSFWorkbook wb,
            String nomHoja,
            int Ano,
            String FechaPago,
            String formatoDatos,
            String nota1,
            String nota2) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);

        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTitulo2);

        CellStyle estiloTexto = wb.createCellStyle();
        estiloTexto.setFont(font);

        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTitulo2);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.cloneStyleFrom(estiloTituloTer);
        estiloTituloFila.setBorderLeft(BorderStyle.THIN);
        estiloTituloFila.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        CellStyle estiloTituloFilaA = wb.createCellStyle();
        estiloTituloFilaA.cloneStyleFrom(estiloTituloFila);
        estiloTituloFilaA.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
        estiloTituloFilaA.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle estiloTituloFila2 = wb.createCellStyle();
        estiloTituloFila2.cloneStyleFrom(estiloTituloFila);
        estiloTituloFila2.setVerticalAlignment(VerticalAlignment.JUSTIFY);

        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(font);
        estiloTituloFilaSec.setBorderLeft(BorderStyle.THIN);
        estiloTituloFilaSec.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        CellStyle estiloTituloFilaSecA = wb.createCellStyle();
        estiloTituloFilaSecA.cloneStyleFrom(estiloTituloFilaSec);
        estiloTituloFilaSecA.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
        estiloTituloFilaSecA.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        CellStyle estiloDatosA = wb.createCellStyle();
        estiloDatosA.cloneStyleFrom(estiloDatos);
        estiloDatosA.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
        estiloDatosA.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle estiloDatos5 = wb.createCellStyle();
        StringTokenizer formatoCompleto5 = new StringTokenizer(formatoDatos, ";");
        String formatoPos5 = formatoCompleto5.nextToken();
        estiloDatos5.setDataFormat(formato.getFormat(formatoPos5));
        estiloDatos5.setFont(font);
        estiloDatos5.setFillForegroundColor(FillPatternType.THICK_VERT_BANDS.getCode());

        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderRight(BorderStyle.THIN);
        estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(MESES[mes] + " " + Ano);
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue("(Valores en $ - fecha de pago hasta " + FechaPago + " )");
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC.setCellStyle(estiloTitulo);

        cellTC = row.createCell(1);
        cellTC.setCellValue(nomgrupo1);
        cellTC.setCellStyle(estiloTituloSec);
        cellTC = row.createCell(2 + (dim2 + 2));
        cellTC.setCellValue(nomgrupo2);
        cellTC.setCellStyle(estiloTituloSec);
        cellTC = row.createCell(3 + (dim2 + 2) * 2);
        cellTC.setCellValue(nomgrupo3);
        cellTC.setCellStyle(estiloTituloSec);
        cellTC = row.createCell(4 + (dim2 + 2) * 3);
        cellTC.setCellValue(nomgrupo4);
        cellTC.setCellStyle(estiloTituloSec);

        row = hoja.createRow(fila);
        fila++;
        row.createCell(100).setCellValue("");
        row = hoja.createRow(fila);
        fila++;
        row = hoja.getRow(7);

        for (int aux = 0; aux < 4; aux++) {//Tres tablas
            for (int i = 2 + aux * (dim2 + 3); i < -1 + (aux + 1) * (dim2 + 3); i++) {
                cellTF = row.createCell(i);
                cellTF.setCellStyle(estiloTituloFila);
            }
            Cell cellTCom1 = row.getCell(2 + aux * (dim2 + 3));
            Cell cellTCom2 = row.getCell(1 + aux * (dim2 + 3) + dim2);
            CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
            CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
            String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
            cellTCom1.setCellValue(tituloFilas2);
        }

        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        for (int aux = 0; aux < 4; aux++) {//4 tablas
            row = hoja.getRow(fila - 2);
            Cell cellTC_a1 = row.createCell(1 + aux * (dim2 + 3));
            cellTC_a1.setCellValue(tituloFilas1);
            cellTC_a1.setCellStyle(estiloTituloFila2);
            Cell cellTC_b1 = row.createCell(-1 + (aux + 1) * (dim2 + 3));
            cellTC_b1.setCellValue("Total");
            cellTC_b1.setCellStyle(estiloTituloFila2);
            row = hoja.getRow(fila - 1);
            Cell cellTC_a2 = row.createCell(1 + aux * (dim2 + 3));
            cellTC_a2.setCellStyle(estiloTituloFila2);
            CellReference cellRef1 = new CellReference(cellTC_a1.getRowIndex(), cellTC_a1.getColumnIndex());
            CellReference cellRef2 = new CellReference(cellTC_a2.getRowIndex(), cellTC_a2.getColumnIndex());
            String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
            Cell cellTC_b2 = row.createCell(-1 + (aux + 1) * (dim2 + 3));
            cellTC_b2.setCellStyle(estiloTituloFila2);
            cellRef1 = new CellReference(cellTC_b1.getRowIndex(), cellTC_b1.getColumnIndex());
            cellRef2 = new CellReference(cellTC_b2.getRowIndex(), cellTC_b2.getColumnIndex());
            reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
            hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
        }
        //
        // Titulos Filas y Datos
        for (int b = 0; b < 4; b++) {//4 tablas
            for (int j = 0; j < dim2; j++) {
                cellTF = row.createCell((int) 2 + j + b * (dim2 + 3));
                cellTF.setCellValue(nombreFilas2[j]);
                if (b == 0) {
                    cellTF.setCellStyle(estiloTituloFilaA);
                } else {
                    cellTF.setCellStyle(estiloTituloFila);
                }
            }
        }
        //
        short filaTmp = fila;
        int aux = dim2 + 3;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            // Tabla 1
            cellTF = row.createCell(1);//2
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSecA);
            // Tabla2
            cellTF = row.createCell(1 + aux);//6
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            // Tabla 3
            cellTF = row.createCell(1 + aux * 2);//11
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            //
            cellTF = row.createCell(1 + aux * 3);//
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            // Datos
            for (int j = 0; j < dim2; j++) {
                // Tabla 1
                cell = row.createCell(j + 2);//2
                cell.setCellValue(Datos3[i][j][mes]);
                cell.setCellStyle(estiloDatosA);
                // Tabla2
                cell = row.createCell(j + 2 + aux);//8
                cell.setCellValue(Datos1[i][j][mes]);
                cell.setCellStyle(estiloDatos);
                // Tabla 3
                cell = row.createCell(j + 2 + aux * 2);//14
                cell.setCellValue(Datos2[i][j][mes]);
                cell.setCellStyle(estiloDatos);
                //Tabla 4
                cell = row.createCell(j + 2 + aux * 3);//
                cell.setCellValue(Datos4[i][j][mes]);
                cell.setCellStyle(estiloDatos);
            }
            // Total Tabla 1
            cell = row.createCell(aux - 1);//5
            cell.setCellValue(DatosTot3[i][mes]);
            cell.setCellStyle(estiloDatos2);
            // Total Tabla 2
            cell = row.createCell(aux * 2 - 1);//10
            cell.setCellValue(DatosTot1[i][mes]);
            cell.setCellStyle(estiloDatos2);
            // Total Tabla 3
            cell = row.createCell(aux * 3 - 1);//15
            cell.setCellValue(DatosTot2[i][mes]);
            cell.setCellStyle(estiloDatos2);
            // Total Tabla 4
            cell = row.createCell(aux * 4 - 1);//
            cell.setCellValue(DatosTot4[i][mes]);
            cell.setCellStyle(estiloDatos2);
        }
        //
        // Escribe totales
        estiloTituloFila.setFillForegroundColor(HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex());
        estiloTituloFila.setFillPattern(FillPatternType.NO_FILL);
        rowFin = hoja.createRow(fila);
        fila++;
        for (int b = 0; b < 4; b++) {
            for (int j = 0; j < dim2; j++) {
                cellTF = rowFin.createCell(1 + b * aux);
                cellTF.setCellValue("Total General");
                cellTF.setCellStyle(estiloTituloFila);
            }
        }

        for (int b = 0; b < 4; b++) {
            for (int j = 0; j < dim2 + 1; j++) {
                Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2 + b * aux);
                Cell cellSumF = row.getCell(j + 2 + b * aux);
                CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
                CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
                cell = rowFin.createCell(j + 2 + b * aux);
                cell.setCellStyle(estiloDatos4);
                cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            }
        }
        //
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$B$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 4; i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        //fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(filaTmp);
            filaTmp++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
            }
        }
        //
        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        //
        // Notas
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        Cell cellNota = row.createCell(1);
        cellNota.setCellValue(nota1);
        cellNota.setCellStyle(estiloTexto);
        row = hoja.createRow(fila);
        fila++;
        cellNota = row.createCell(1);
        cellNota.setCellValue(nota2);
        cellNota.setCellStyle(estiloTexto);
    }
   
    static public void crea_AjusteCentrales(int m,
            String titulo, double Datos1[][][], double Datos1a[][][],
            double DatosTot1[][], double DatosTot1a[][],
            String tituloFilas1, String[] nombreFilas1, String[] nombreFilas1a,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nomHoja, int Ano, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_AjusteCentrales(m, titulo, Datos1, Datos1a, DatosTot1, DatosTot1a, tituloFilas1, nombreFilas1, nombreFilas1a, tituloFilas2, nombreFilas2, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
   
    static public void crea_AjusteCentrales(int m,
            String titulo, double Datos1[][][], double Datos1a[][][],
            double DatosTot1[][], double DatosTot1a[][],
            String tituloFilas1, String[] nombreFilas1, String[] nombreFilas1a,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nomHoja, int Ano, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        Row rowFin1 = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);
        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderLeft(BorderStyle.THIN);
        estiloTituloFila.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderLeft(BorderStyle.THIN);
        estiloTituloFilaSec.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderRight(BorderStyle.THIN);
        estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;
        int dim1a = Datos1a.length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        row = hoja.getRow(3);
        for (int i = 2; i < 2 + dim2; i++) {
            cellTF = row.createCell(i);
            cellTF.setCellStyle(estiloTituloFila);
        }

        Cell cellTCom1 = row.getCell(2);
        Cell cellTCom2 = row.getCell(1 + dim2);
        cellTCom1.setCellValue(tituloFilas2);
        CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
        CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
        String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));

        row = hoja.getRow(3);
        for (int i = 5 + dim2; i < 5 + dim2 * 2; i++) {
            cellTF = row.createCell(i);
            cellTF.setCellStyle(estiloTituloFila);
        }

        cellTCom1 = row.getCell(5 + dim2);
        cellTCom1.setCellValue(tituloFilas2);
        cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
        cellRef2 = new CellReference(cellTCom1.getRowIndex(), 4 + dim2 * 2);
        reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString();
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(5 + dim2);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
            cellTF = row.createCell((int) 5 + dim2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }

        cellTF = row.createCell(dim2 + 2);
        cellTF.setCellValue("Total");
        cellTF.setCellStyle(estiloTituloFila);
        cellTF = row.createCell(dim2 * 2 + 5);
        cellTF.setCellValue("Total");
        cellTF.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4 + dim2);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        for (int i = 5; i < dim1a + 7; i++) {
            row = hoja.createRow(i);
        }

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.getRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j][m]);
            }
            cell = row.createCell(dim2 + 2);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i][m]);
        }
        rowFin = row;

        fila = filaTmp;
        for (int i = 0; i < dim1a; i++) {
            row = hoja.getRow(fila);
            fila++;
            cellTF = row.createCell(4 + dim2);
            cellTF.setCellValue(nombreFilas1a[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);
            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 5 + dim2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1a[i][j][m]);
            }
            cell = row.createCell(5 + dim2 * 2);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1a[i][m]);
        }
        rowFin1 = row;
        //Escribe Suma
        cellTF = hoja.getRow(rowFin.getRowNum() + 1).createCell(1);
        cellTF.setCellValue("Total General");
        cellTF.setCellStyle(estiloTituloFila);
        cellTF = hoja.getRow(rowFin1.getRowNum() + 1).createCell(4 + dim2);
        cellTF.setCellValue("Total General");
        cellTF.setCellStyle(estiloTituloFila);

        for (int j = 0; j < dim2 + 1; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2);
            Cell cellSumF = rowFin.getCell(j + 2);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = hoja.getRow(rowFin.getRowNum() + 1).createCell(j + 2);
            cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }
        for (int j = 0; j < dim2 + 1; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(5 + dim2 + j);
            Cell cellSumF = rowFin1.getCell(5 + dim2 + j);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = hoja.getRow(rowFin1.getRowNum() + 1).createCell(5 + dim2 + j);
            cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) //hoja.setColumnWidth(i, 5*700);
        {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
                //cell.setCellStyle(estiloDatos);
            }
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo + " " + Ano);
        cellTC1.setCellStyle(estiloTitulo);
    }
    
    static public void crea_1TablaTx_1C(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nomHoja, int Ano, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_1TablaTx_1C(titulo, Datos1, DatosTot1, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_1TablaTx_1C(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nomHoja, int Ano, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderRight(BorderStyle.THIN);
        estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        row = hoja.getRow(3);
        Cell cellTCom1 = row.createCell(2);
        Cell cellTCom2 = row.createCell(1 + dim2);
        cellTCom1.setCellValue(tituloFilas2);
        cellTCom1.setCellStyle(estiloTituloFila);
        CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
        CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
        String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
        cellTCom2.setCellStyle(estiloTituloFila);

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }
        cellTF = row.createCell(dim2 + 2);
        cellTF.setCellValue("Total");
        cellTF.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            //Datos
            for (int j = 0; j < dim2; j++) {

                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j]);
            }
            cell = row.createCell(dim2 + 2);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i]);
        }

        //Escribe totales
        rowFin = hoja.createRow(fila);
        fila++;
        for (int j = 0; j < dim2; j++) {
            cellTF = rowFin.createCell(1);
            cellTF.setCellValue("Total General");
            cellTF.setCellStyle(estiloTituloFila);
        }

        for (int j = 0; j < dim2 + 1; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2);
            Cell cellSumF = row.getCell(j + 2);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());

            cell = rowFin.createCell(j + 2);
            cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        // CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        //String reference = nomHoja+"!$D$6:"+cellRef.formatAsString(); // area reference
        //nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        //hoja.autoSizeColumn(i);
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
//        fila = filaTmp;
//        for (int i = 0; i < dim2; i++) {
//            row = hoja.getRow(fila);
//            fila++;
////            for (int j = 0; j < dim2; j++) {
////                cell = row.getCell(j + 3);
////                //cell.setCellStyle(estiloDatos);
////            }
//        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
//        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo + " " + Ano);
        cellTC1.setCellStyle(estiloTitulo);
        
    }
    
    static public void crea_1TablaTx_1C_double(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String nomLibro, String nomHoja, int Ano, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_1TablaTx_1C_double(titulo, Datos1, DatosTot1, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_1TablaTx_1C_double(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            XSSFWorkbook wb, String nomHoja, int Ano, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("#,###,##0.#0", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0", ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
        StringTokenizer formatoCompleto4 = new StringTokenizer(formatoDatos, ";");
        String formatoPos4 = formatoCompleto4.nextToken();
        estiloDatos4.setDataFormat(formato4.getFormat(formatoPos4));
        estiloDatos4.setFont(font);
        estiloDatos4.setBorderRight(BorderStyle.THIN);
        estiloDatos4.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderBottom(BorderStyle.THIN);
        estiloDatos4.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos4.setBorderTop(BorderStyle.THIN);
        estiloDatos4.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        row = hoja.getRow(3);
        Cell cellTCom1 = row.createCell(2);
        Cell cellTCom2 = row.createCell(1 + dim2);
        cellTCom1.setCellValue(tituloFilas2);
        cellTCom1.setCellStyle(estiloTituloFila);
        CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
        CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
        String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
        cellTCom2.setCellStyle(estiloTituloFila);

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 2 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }
        cellTF = row.createCell(dim2 + 2);
        cellTF.setCellValue("Total");
        cellTF.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            //Datos
            for (int j = 0; j < dim2; j++) {

                cell = row.createCell(j + 2);//2
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j]);
            }
            cell = row.createCell(dim2 + 2);//5
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i]);
        }

        //Escribe totales
        rowFin = hoja.createRow(fila);
        fila++;
        for (int j = 0; j < dim2; j++) {
            cellTF = rowFin.createCell(1);
            cellTF.setCellValue("Total General");
            cellTF.setCellStyle(estiloTituloFila);
        }

        for (int j = 0; j < dim2 + 1; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 2);
            Cell cellSumF = row.getCell(j + 2);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());

            cell = rowFin.createCell(j + 2);
            cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        // CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        //String reference = nomHoja+"!$D$6:"+cellRef.formatAsString(); // area reference
        //nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        //hoja.autoSizeColumn(i);
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
////        fila = filaTmp;
////        for (int i = 0; i < dim2; i++) {
////            row = hoja.getRow(fila);
////            fila++;
////            for (int j = 0; j < dim2; j++) {
////                cell = row.getCell(j + 3);
////                //cell.setCellStyle(estiloDatos);
////            }
////        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo + " " + Ano);
        cellTC1.setCellStyle(estiloTitulo);
    }
    
    static public void crea_1TablaTx_2C(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, double[] DatosFilas3,
            String nomLibro, String nomHoja, int Ano, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_1TablaTx_2C(titulo, Datos1, DatosTot1, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloFilas3, DatosFilas3, wb, nomHoja, Ano, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_1TablaTx_2C(
            String titulo, double Datos1[][],
            double DatosTot1[],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloFilas3, double[] DatosFilas3,
            XSSFWorkbook wb, String nomHoja, int Ano, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        Row rowFin = null;
        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer(formatoDatos, ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setBorderLeft(BorderStyle.THIN);
        estiloDatos1.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos2.setBorderLeft(BorderStyle.THIN);
        estiloDatos2.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("###.##0%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        DataFormat formato4 = wb.createDataFormat();
        CellStyle estiloDatos4 = wb.createCellStyle();
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

        // Dimensiones del arreglo
        int dim1 = Datos1.length;
        int dim2 = Datos1[0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        row = hoja.getRow(3);
        Cell cellTCom1 = row.createCell(3);
        Cell cellTCom2 = row.createCell(2 + dim2);
        cellTCom1.setCellValue(tituloFilas2);
        cellTCom1.setCellStyle(estiloTituloFila);
        CellReference cellRef1 = new CellReference(cellTCom1.getRowIndex(), cellTCom1.getColumnIndex());
        CellReference cellRef2 = new CellReference(cellTCom2.getRowIndex(), cellTCom2.getColumnIndex());
        String reference1 = nomHoja + "!" + cellRef1.formatAsString() + ":" + cellRef2.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference1));
        cellTCom2.setCellStyle(estiloTituloFila);

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        //cellTC = row.createCell(4);
        //cellTC.setCellValue(tituloFilas4);
        //cellTC.setCellStyle(estiloTituloFila);

        // Nombre Transmisores y Datos
        for (int j = 0; j < dim2; j++) {
            cellTF = row.createCell((int) 3 + j);
            cellTF.setCellValue(nombreFilas2[j]);
            cellTF.setCellStyle(estiloTituloFila);
        }
        cellTF = row.createCell(dim2 + 3);
        cellTF.setCellValue("Total");
        cellTF.setCellStyle(estiloTituloFila);

        short filaTmp = fila;
        for (int i = 0; i < dim1; i++) {
            row = hoja.createRow(fila);
            fila++;
            cellTF = row.createCell(1);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloFilaSec);

            cell = row.createCell(2);
            cell.setCellValue(DatosFilas3[i]);
            cell.setCellStyle(estiloDatos1);

            //Datos
            for (int j = 0; j < dim2; j++) {
                cell = row.createCell(j + 3);
                cell.setCellStyle(estiloDatos);
                cell.setCellValue(Datos1[i][j]);
                cellTF.setCellStyle(estiloDatos);
            }
            cell = row.createCell(dim2 + 3);
            cell.setCellStyle(estiloDatos2);
            cell.setCellValue(DatosTot1[i]);
        }

        //Escribe totales
        rowFin = hoja.createRow(fila);
        fila++;

        cellTF = rowFin.createCell(1);
        cellTF.setCellValue("Total General");
        cellTF.setCellStyle(estiloTituloFila);
        cellTF = rowFin.createCell(2);
        cellTF.setCellStyle(estiloTituloFila);

        for (int j = 0; j < dim2 + 1; j++) {
            Cell cellSumI = hoja.getRow(filaTmp).getCell(j + 3);
            Cell cellSumF = row.getCell(j + 3);
            CellReference RefI = new CellReference(cellSumI.getRowIndex(), cellSumI.getColumnIndex());
            CellReference RefF = new CellReference(cellSumF.getRowIndex(), cellSumF.getColumnIndex());
            cell = rowFin.createCell(j + 3);
            //cell.setCellStyle(estiloDatos4);
            cell.setCellFormula("sum(" + RefI.formatAsString() + ":" + RefF.formatAsString() + ")");
            cell.setCellStyle(estiloDatos4);
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        hoja.setColumnWidth(1, 4 * 2000);
        for (int i = 2; i < (dim2 + 4); i++) {
            hoja.setColumnWidth(i, 5 * 700);
        }
        //hoja.autoSizeColumn(i);
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        //estiloDatos.setDataFormat(formato.getFormat(formatoDatos));

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
    }
    
    static public void crea_verificaRet(
            String titulo, String nomLibro,
            String titulo1, String tit2, String tit3, String tit4, String tit5,
            double Datos1[], double Datos2[],
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3, double DatosFilas3[],
            String TituloFilas4,
            String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_verificaRet(titulo, wb, titulo1, tit2, tit3, tit4, tit5, Datos1, Datos2, TituloFilas1, DatosFilas1, TituloFilas2, DatosFilas2, TituloFilas3, DatosFilas3, TituloFilas4, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_verificaRet(
            String titulo, XSSFWorkbook wb,
            String titulo1, String tit2, String tit3, String tit4, String tit5,
            double Datos1[], double Datos2[],
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3, double DatosFilas3[],
            String TituloFilas4,
            String nomHoja, String formatoDatos) {
        
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cell = null;
        Row row = null;
        short fila = 0;
        
        hoja = wb.getSheet(nomHoja);
        
        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);
        
        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);
        
        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);
        
        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);
        
        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("0.0000");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);
        
        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("###,##0.00");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setBorderBottom(BorderStyle.THIN);
        estiloDatos1.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setBorderTop(BorderStyle.THIN);
        estiloDatos1.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloDatos1.setAlignment(HorizontalAlignment.CENTER);
        
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(titulo);
        cellTC.setCellStyle(estiloTitulo);
        
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        
        cellTC = row.createCell(1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC.setCellValue(tit2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(tit3);
        cellTC.setCellStyle(estiloTituloFila);

        //Pago CUE
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tit4);
        cellTC.setCellStyle(estiloTituloFilaSec);
        cellTC = row.createCell(2);
        cellTC.setCellValue(Datos1[0]);
        cellTC.setCellStyle(estiloDatos2);
        cellTC = row.createCell(3);
        cellTC.setCellValue(Datos1[1]);
        cellTC.setCellStyle(estiloDatos2);

        //Consumo CUE
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tit5);
        cellTC.setCellStyle(estiloTituloFilaSec);
        cellTC = row.createCell(2);
        cellTC.setCellValue(Datos2[0]);
        cellTC.setCellStyle(estiloDatos2);
        cellTC = row.createCell(3);
        cellTC.setCellValue(Datos2[1]);
        cellTC.setCellStyle(estiloDatos2);

        //Calcula CUE
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(titulo1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(Datos1[0] / Datos2[0]);
        cellTC.setCellStyle(estiloDatos1);
        cellTC = row.createCell(3);
        cellTC.setCellValue(Datos1[1] / Datos2[1]);
        cellTC.setCellStyle(estiloDatos1);
        
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        cellTC = row.createCell(1);
        cellTC.setCellValue(TituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(TituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(TituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(TituloFilas4);
        cellTC.setCellStyle(estiloTituloFila);

        // Datos
        for (int i = 0; i < 12; i++) {
            row = hoja.createRow(fila);
            fila++;
            cell = row.createCell(1);
            cell.setCellValue(DatosFilas1[i]);
            cell.setCellStyle(estiloTituloFilaSec);
            
            cell = row.createCell(2);
            cell.setCellValue(DatosFilas2[i]);
            cell.setCellStyle(estiloDatos);
            cell = row.createCell(3);
            cell.setCellValue(DatosFilas3[i]);
            cell.setCellStyle(estiloDatos);
            
            cell = row.createCell(4);
            cell.setCellValue(DatosFilas2[i] - DatosFilas3[i]);
            cell.setCellStyle(estiloDatos3);
        }

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        hoja.setColumnWidth(1, 5 * 256);
        for (int i = 2; i < 5; i++) {
            hoja.autoSizeColumn(i);
        }
    }
    
    static public void crea_verificaIny(
            String titulo, String nomLibro,
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3, double DatosFilas3[],
            String TituloFilas4,
            String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_verificaIny(titulo, wb, TituloFilas1, DatosFilas1, TituloFilas2, DatosFilas2, TituloFilas3, DatosFilas3, TituloFilas4, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_verificaIny(
            String titulo, XSSFWorkbook wb,
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3, double DatosFilas3[],
            String TituloFilas4,
            String nomHoja, String formatoDatos) {
        
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cell = null;
        Row row = null;
        short fila = 22;
        
        hoja = wb.getSheet(nomHoja);

        //hoja.setPrintGridlines(false);
        //hoja.setDisplayGridlines(false);
        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);
        
        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);
        
        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);
        
        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);
        
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer("0.0000");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(titulo);
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        cellTC = row.createCell(1);
        cellTC.setCellValue(TituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(TituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(TituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(TituloFilas4);
        cellTC.setCellStyle(estiloTituloFila);

        // Datos
        for (int i = 0; i < 12; i++) {
            row = hoja.createRow(fila);
            fila++;
            cell = row.createCell(1);
            cell.setCellValue(DatosFilas1[i]);
            cell.setCellStyle(estiloTituloFilaSec);
            
            cell = row.createCell(2);
            cell.setCellValue(DatosFilas2[i]);
            cell.setCellStyle(estiloDatos);
            cell = row.createCell(3);
            cell.setCellValue(DatosFilas3[i]);
            cell.setCellStyle(estiloDatos);
            
            cell = row.createCell(4);
            cell.setCellValue(DatosFilas2[i] - DatosFilas3[i]);
            cell.setCellStyle(estiloDatos2);
        }

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        hoja.setColumnWidth(1, 5 * 256);
        for (int i = 2; i < 5; i++) {
            hoja.autoSizeColumn(i);
        }
    }
    
    static public void crea_verificaCalcPeajes(
            String titulo, String nomLibro,
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3,
            String TituloFilas4, String TituloFilas5,
            String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_verificaCalcPeajes(titulo, wb, TituloFilas1, DatosFilas1, TituloFilas2, DatosFilas2, TituloFilas3, TituloFilas4, TituloFilas5, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_verificaCalcPeajes(
            String titulo, XSSFWorkbook wb,
            String TituloFilas1, String DatosFilas1[],
            String TituloFilas2, double DatosFilas2[],
            String TituloFilas3,
            String TituloFilas4, String TituloFilas5,
            String nomHoja, String formatoDatos) {
        
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cell = null;
        Cell cellIny = null;
        Cell cellRet = null;
        Row row = null;
        short fila = 39;
        
        hoja = wb.getSheet(nomHoja);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);
        
        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);
        
        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);
        
        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        
        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer("#,###,##0;[Red]-#,###,##0;\"-\"");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);
        
        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("0.0000");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(titulo);
        cellTC.setCellStyle(estiloTitulo);
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        cellTC = row.createCell(1);
        cellTC.setCellValue(TituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(TituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(3);
        cellTC.setCellValue(TituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(4);
        cellTC.setCellValue(TituloFilas4);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(5);
        cellTC.setCellValue(TituloFilas5);
        cellTC.setCellStyle(estiloTituloFila);

        // Datos
        for (int i = 0; i < 12; i++) {
            row = hoja.getRow(10 + i);
            cellRet = row.getCell(2);
            
            row = hoja.getRow(27 + i);
            cellIny = row.getCell(2);
            
            row = hoja.createRow(fila);
            fila++;
            cell = row.createCell(1);
            cell.setCellValue(DatosFilas1[i]);
            cell.setCellStyle(estiloTituloFilaSec);
            
            cell = row.createCell(2);
            cell.setCellValue(DatosFilas2[i]);
            cell.setCellStyle(estiloDatos);
            CellReference cellRefPje = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
            
            cell = row.createCell(3);
            cell.setCellValue(cellRet.getNumericCellValue());
            cell.setCellStyle(estiloDatos);
            CellReference cellRefRet = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
            
            cell = row.createCell(4);
            cell.setCellValue(cellIny.getNumericCellValue());
            cell.setCellStyle(estiloDatos);
            CellReference cellRefIny = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
            
            cell = row.createCell(5);
            cell.setCellStyle(estiloDatos1);
            cell.setCellFormula(cellRefPje.formatAsString() + "-" + cellRefIny.formatAsString() + "-" + cellRefRet.formatAsString());
        }

        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        hoja.setColumnWidth(1, 5 * 400);
        for (int i = 2; i < 5; i++) {
            hoja.setColumnWidth(i, 8 * 400);
        }
    }
    
     
    static public void crea_verifProrr(double Datos[][],
            int numFilas, String nombreFilas[],
            String nomLibro, String nomHoja, String formatoDatos, int inicio) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_verifProrr(Datos, numFilas, nombreFilas, wb, nomHoja, formatoDatos, inicio);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de actualizar la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_verifProrr(double Datos[][],
            int numFilas, String nombreFilas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos, int inicio) {
        XSSFSheet hoja;
        Cell cell;
        Row row;
        short fila = 4;

        hoja = wb.getSheet(nomHoja);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        //int numFilas=Datos.length;
        int numCol = Datos[0].length;

        String[] linea = new String[numFilas];

        // Titulos Filas y Datos
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            cell = row.getCell(1);
            linea[i] = cell.getRichStringCellValue().toString().trim();
        }
        for (int i = 0; i < numFilas; i++) {
            int l = Calc.Buscar(linea[i], nombreFilas);
            if (l != -1) {
                for (int j = 0; j < numCol; j++) {
                    row = hoja.getRow(4 + i);
                    cell = row.getCell(j + inicio + 3);
                    if (Datos[l][j] == 0) {
                        cell.setCellValue("x");
                    } else {
                        cell.setCellValue("o");
                    }
                }
            } else {
                System.out.println("WARNING: La línea '" + linea[i] + "' en '" + nomHoja + "' no se encuentra en la hoja 'lintron'");
            }
        }
    }
     
    static public void crea_verifProrrPeaj(double Datos[][],
            String nombrelineas[],
            String nomLibro, String nomHoja, String formatoDatos, int inicio) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            crea_verifProrrPeaj(Datos, nombrelineas, wb, nomHoja, formatoDatos, inicio);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de actualizar la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void crea_verifProrrPeaj(double Datos[][],
            String nombrelineas[],
            XSSFWorkbook wb, String nomHoja, String formatoDatos, int inicio) {
        XSSFSheet hoja = null;
        Cell cellLin = null;
        Cell cell = null;
        Row row = null;
        short fila = 4;

        hoja = wb.getSheet(nomHoja);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        // Dimensiones del arreglo
        int numFilas = Datos.length;
        int numCol = Datos[0].length;

        String[] linea = new String[numFilas];
        String[] lineaT = new String[numFilas];
        String[] tmp = new String[2];

        // Titulos Filas y Datos
        for (int i = 0; i < numFilas; i++) {
            row = hoja.getRow(fila);
            fila++;
            cellLin = row.getCell(1);
            lineaT[i] = cellLin.getRichStringCellValue().toString().trim();
            linea[i] = lineaT[i].split("#")[0];
            //linea[i]=tmp[0];
        }

        for (int i = 0; i < numFilas; i++) {
            int l = Calc.Buscar(linea[i], nombrelineas);
            //if(l!=-1){
            for (int j = 0; j < numCol; j++) {
                row = hoja.getRow(4 + i);//4+l
                cell = row.getCell(j + inicio + 3);
                if (Datos[i][j] == 0) {//daots[l] pero deberia haber dicho datos[i]
                    cell.setCellValue("x");
                } else {
                    cell.setCellValue("o");
                }
            }
            //}
            if (l == -1) {
                System.out.println("La línea " + linea[i] + " en la hoja " + nomHoja + " no se encuentra en la hoja 'lintron'");
            }
        }
    }

    static public void creaH3F_4d_double(String titulo, double Datos[][][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, String[] nombreFilas3,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaH3F_4d_double(titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreFilas2, tituloColumnas, nombreColumnas, tituloFilas3, nombreFilas3, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static public void creaH3F_4d_double(String titulo, double Datos[][][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreFilas2,
            String tituloColumnas, String[] nombreColumnas,
            String tituloFilas3, String[] nombreFilas3,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;
        int fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        StringTokenizer formatoCompleto2 = new StringTokenizer("#,##0.00;\"-\"", ";");
        String formatoPos = formatoCompleto.nextToken();
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos.setFont(font);
        estiloDatos2.setFont(font);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;
        int dim4 = Datos[0][0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        fila++;
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(4);//cambio
        cellTC.setCellValue(tituloColumnas);
        cellTC.setCellStyle(estiloTituloSec);
        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;

        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas2);
        cellTC.setCellStyle(estiloTituloFila);

        cellTC = row.createCell(3);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        for (int k = 0; k < dim4; k++) {
            cellTC = row.createCell((int) k + 4);
            cellTC.setCellValue(nombreColumnas[k]);
            cellTC.setCellStyle(estiloTituloTer);
        }
        // Titulos Filas y Datos
        int filaTmp = fila;
        for (int i = 0; i < dim2; i++) {
            for (int j = 0; j < dim1; j++) {
                for (int p = 0; p < dim3; p++) {
                    row = hoja.createRow(fila);
                    fila++;
                    cellTF = row.createCell(1);
                    cellTF.setCellValue(nombreFilas1[i]);
                    cellTF.setCellStyle(estiloTituloFilaSec);

                    cellTF = row.createCell((int) 2);
                    cellTF.setCellValue(nombreFilas2[j]);
                    cellTF.setCellStyle(estiloTituloFilaSec);

                    cell = row.createCell(3);
                    cell.setCellValue(nombreFilas3[p]);
                    cell.setCellStyle(estiloDatos2);
                    for (int k = 0; k < dim4; k++) {
                        cell = row.createCell(k + 4);
                        cell.setCellValue(Datos[j][i][p][k]);
                        cell.setCellStyle(estiloDatos);
                    }
                }
            }
        }
        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < dim3 + 4; i++) {
            hoja.autoSizeColumn((i));
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim3 + 3);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
    }
    
    static public void creaTabla2CDx_double(int mes,
            String titulo, double Datos[][][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreTx,
            String tituloFilas3, String[] DatosFilas3,
            double[][][] factores,
            String nomLibro, String nomHoja, String formatoDatos) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibro ));
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(nomLibro));
            creaTabla2CDx_double(mes, titulo, Datos, tituloFilas1, nombreFilas1, tituloFilas2, nombreTx, tituloFilas3, DatosFilas3, factores, wb, nomHoja, formatoDatos);
            // Graba y Cierra
            FileOutputStream archivoSalida = new FileOutputStream(nomLibro);
            wb.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);
        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
       
    static public void creaTabla2CDx_double(int mes,
            String titulo, double Datos[][][][],
            String tituloFilas1, String[] nombreFilas1,
            String tituloFilas2, String[] nombreTx,
            String tituloFilas3, String[] DatosFilas3,
            double[][][] factores,
            XSSFWorkbook wb, String nomHoja, String formatoDatos) {
        XSSFSheet hoja = null;
        Cell cellTC = null;
        Cell cellT = null;
        Cell cellTF = null;
        Cell cell = null;
        Row row = null;

        short fila = 0;

        hoja = wb.createSheet(nomHoja);

        hoja.setPrintGridlines(false);
        hoja.setDisplayGridlines(false);

        // Estilos
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Century Gothic");
        CellStyle estilo = wb.createCellStyle();
        estilo.setFont(font);

        Font fontTitulo = wb.createFont();
        fontTitulo.setFontHeightInPoints((short) 10);
        fontTitulo.setFontName("Century Gothic");
        fontTitulo.setBold(true);
        CellStyle estiloTitulo = wb.createCellStyle();
        estiloTitulo.setFont(fontTitulo);

        Font fontTitulo2 = wb.createFont();
        fontTitulo2.setFontHeightInPoints((short) 8);
        fontTitulo2.setFontName("Century Gothic");
        fontTitulo2.setBold(true);
        CellStyle estiloTitulo2 = wb.createCellStyle();
        estiloTitulo2.setFont(fontTitulo2);

        Font fontTituloSec = wb.createFont();
        fontTituloSec.setFontHeightInPoints((short) 8);
        fontTituloSec.setFontName("Century Gothic");
        fontTituloSec.setBold(true);
        CellStyle estiloTituloSec = wb.createCellStyle();
        estiloTituloSec.setFont(fontTituloSec);

        Font fontTituloTer = wb.createFont();
        fontTituloTer.setFontHeightInPoints((short) 8);
        fontTituloTer.setFontName("Century Gothic");
        fontTituloTer.setBold(true);
        CellStyle estiloTituloTer = wb.createCellStyle();
        estiloTituloTer.setFont(fontTituloTer);
        estiloTituloTer.setBorderBottom(BorderStyle.THIN);
        estiloTituloTer.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setBorderTop(BorderStyle.THIN);
        estiloTituloTer.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloTer.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFila = wb.createFont();
        fontTituloFila.setFontHeightInPoints((short) 8);
        fontTituloFila.setFontName("Century Gothic");
        fontTituloFila.setBold(true);
        CellStyle estiloTituloFila = wb.createCellStyle();
        estiloTituloFila.setFont(fontTituloFila);
        estiloTituloFila.setBorderRight(BorderStyle.THIN);
        estiloTituloFila.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderBottom(BorderStyle.THIN);
        estiloTituloFila.setBottomBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setBorderTop(BorderStyle.THIN);
        estiloTituloFila.setTopBorderColor(IndexedColors.PALE_BLUE.getIndex());
        estiloTituloFila.setAlignment(HorizontalAlignment.CENTER);

        Font fontTituloFilaSec = wb.createFont();
        fontTituloFilaSec.setFontHeightInPoints((short) 8);
        fontTituloFilaSec.setFontName("Century Gothic");
        CellStyle estiloTituloFilaSec = wb.createCellStyle();
        estiloTituloFilaSec.setFont(fontTituloFilaSec);
        estiloTituloFilaSec.setBorderRight(BorderStyle.THIN);
        estiloTituloFilaSec.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato1 = wb.createDataFormat();
        CellStyle estiloDatos1 = wb.createCellStyle();
        StringTokenizer formatoCompleto1 = new StringTokenizer("0.00;[Red]-0.00", ";");
        String formatoPos1 = formatoCompleto1.nextToken();
        estiloDatos1.setDataFormat(formato1.getFormat(formatoPos1));
        estiloDatos1.setFont(font);
        estiloDatos1.setBorderRight(BorderStyle.THIN);
        estiloDatos1.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato2 = wb.createDataFormat();
        CellStyle estiloDatos2 = wb.createCellStyle();
        StringTokenizer formatoCompleto2 = new StringTokenizer(formatoDatos, ";");
        String formatoPos2 = formatoCompleto2.nextToken();
        estiloDatos2.setDataFormat(formato2.getFormat(formatoPos2));
        estiloDatos2.setFont(font);
        estiloDatos2.setAlignment(HorizontalAlignment.CENTER);
        estiloDatos2.setBorderRight(BorderStyle.THIN);
        estiloDatos2.setRightBorderColor(IndexedColors.PALE_BLUE.getIndex());

        DataFormat formato = wb.createDataFormat();
        CellStyle estiloDatos = wb.createCellStyle();
        StringTokenizer formatoCompleto = new StringTokenizer(formatoDatos, ";");
        String formatoPos = formatoCompleto.nextToken();
        estiloDatos.setDataFormat(formato.getFormat(formatoPos));
        estiloDatos.setFont(font);

        DataFormat formato3 = wb.createDataFormat();
        CellStyle estiloDatos3 = wb.createCellStyle();
        StringTokenizer formatoCompleto3 = new StringTokenizer("0.00%", ";");
        String formatoPos3 = formatoCompleto3.nextToken();
        estiloDatos3.setDataFormat(formato3.getFormat(formatoPos3));
        estiloDatos3.setFont(font);

        // Dimensiones del arreglo
        int dim1 = Datos.length;
        int dim2 = Datos[0].length;
        int dim3 = Datos[0][0].length;

        // Titulos Secundarios
        fila++;
        row = hoja.createRow(fila);
        fila++;
        row = hoja.createRow(fila);
        fila++;

        // Titulos Filas y Terciarios
        row = hoja.createRow(fila);
        fila++;
        cellTC = row.createCell(1);
        cellTC.setCellValue(tituloFilas1);
        cellTC.setCellStyle(estiloTituloFila);
        cellTC = row.createCell(2);
        cellTC.setCellValue(tituloFilas3);
        cellTC.setCellStyle(estiloTituloFila);

        // Titulos Filas y Datos
        for (int j = 0; j < dim3; j++) {
            cellTF = row.createCell((int) 3 + j);
            cellTF.setCellValue(nombreTx[j]);
            cellTF.setCellStyle(estiloTituloFila);

        }

        short filaTmp = fila;
        for (int i = 0; i < dim2; i++) {
            for (int j = 0; j < dim1; j++) {
                row = hoja.createRow(fila);
                fila++;
                cellTF = row.createCell(1);
                cellTF.setCellValue(nombreFilas1[i]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cellTF = row.createCell(2);
                cellTF.setCellValue(DatosFilas3[j]);
                cellTF.setCellStyle(estiloDatos1);

                //Datos
                for (int t = 0; t < dim3; t++) {
                    cell = row.createCell(t + 3);//2
                    cell.setCellStyle(estiloDatos2);
                    cell.setCellValue(Datos[j][i][t][mes]);
                    cell.setCellStyle(estiloDatos2);
                }
            }
        }
        // escribe cuadro con prorratas
        row = hoja.getRow(3);
        cellTF = row.createCell(5 + dim3);
        cellTF.setCellValue("Empresa");
        cellTF.setCellStyle(estiloTituloTer);
        for (int i = 0; i < dim2; i++) {
            cellTF = row.createCell(i + 6 + dim3);
            cellTF.setCellValue(nombreFilas1[i]);
            cellTF.setCellStyle(estiloTituloTer);
        }

        for (int j = 0; j < dim1; j++) {
            for (int i = 0; i < dim2; i++) {
                row = hoja.getRow(4 + j);
                cellTF = row.createCell(5 + dim3);
                cellTF.setCellValue(DatosFilas3[j]);
                cellTF.setCellStyle(estiloTituloFilaSec);
                cell = row.createCell(i + 6 + dim3);
                cell.setCellValue(factores[j][i][mes]);
                cell.setCellStyle(estiloDatos3);
            }
        }

        // Crea nombre de rango de salida
        Name nombreCel = wb.createName();
        nombreCel.setNameName(nomHoja); // Nombre del rango igual al nombre de la hoja
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String reference = nomHoja + "!$D$6:" + cellRef.formatAsString(); // area reference
        nombreCel.setRefersToFormula(reference);
        // Ajusta anchos
        hoja.setColumnWidth(0, 2 * 256);
        for (int i = 1; i < (dim2 + 3) * 3; i++) //hoja.setColumnWidth(i, 5*700);
        {
            hoja.autoSizeColumn(i);
        }
        // Aplica estilo definitivo despues de ajuste de ancho de columnas
        estiloDatos.setDataFormat(formato.getFormat(formatoDatos));
        fila = filaTmp;
        for (int i = 0; i < dim2; i++) {
            row = hoja.getRow(fila);
            fila++;
            for (int j = 0; j < dim2; j++) {
                cell = row.getCell(j + 3);
                //cell.setCellStyle(estiloDatos);
            }
        }

        // Titulo Principal
        row = hoja.getRow(1);
        Cell cellTC1 = row.createCell(1);
        Cell cellTC2 = row.createCell(dim2 + 2);
        cellTC1.setCellValue(titulo);
        cellTC1.setCellStyle(estiloTitulo);
        cellRef = new CellReference(cellTC2.getRowIndex(), cellTC2.getColumnIndex());
        reference = nomHoja + "!$B$2:" + cellRef.formatAsString(); // area reference
        //hoja.addMergedRegion(CellRangeAddress.valueOf(reference));
        cellT = row.createCell(5 + dim3);
        cellT.setCellValue("Participación de Suministradores sobre pagos de empresas Distribuidoras");
        cellT.setCellStyle(estiloTitulo);

    }
        
    static public void CopiaHoja(String nomLibroE, String nomLibroS, String nomHoja) {
        try {
            //POIFSFileSystem archivoEntrada = new //POIFSFileSystem(new FileInputStream( nomLibroE ));
            XSSFWorkbook wbe = new XSSFWorkbook(new FileInputStream(nomLibroE));
            int indice = wbe.getSheetIndex(nomHoja);

            //POIFSFileSystem archivoSalidaC = new //POIFSFileSystem(new FileInputStream( nomLibroS ));
            XSSFWorkbook wbs = new XSSFWorkbook(new FileInputStream(nomLibroS));
            wbs.cloneSheet(indice);

            FileOutputStream archivoSalida = new FileOutputStream(nomLibroS);
            wbs.write(archivoSalida);
            archivoSalida.close();
            System.out.println("Acaba de crear la hoja xls " + nomHoja);

        } catch (IOException e) {
            System.out.println("No se se puede acceder al archivo " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace(System.out);
        }
    }
    
    static void appendToDebugProrrata(String DirBaseSalida, int[] lineasFlujo, int[] gflux, int e, float[][][] Gx, GGDF D, int[][] datosGener, float[][] A, float[][] Dref, float[][][] Prorratas) {
        BufferedWriter writerCSV = null;
        try {
            writerCSV = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(DirBaseSalida + PeajesConstant.SLASH + "prorratas.csv", true), PeajesConstant.CSV_ENCODING));
            int numHid = D.getNumHidro();
            for (int h = 0; h < numHid; h++) {
                for (int l = 0; l < lineasFlujo.length; l++) {
                    for (int g = 0; g < gflux.length; g++) {
                        //for(int g=0;g<numGeneradores;g++){
                        writerCSV.write(String.valueOf(e));
                        writerCSV.write(',');
                        writerCSV.write(String.valueOf(h));
                        writerCSV.write(',');
                        writerCSV.write(String.valueOf(lineasFlujo[l]));
                        writerCSV.write(',');
                        writerCSV.write(String.valueOf(gflux[g]));
                        writerCSV.write(',');
//                        writer.append(String.valueOf((genEquiv[gflux[g]][lineasFlujo[l]][h]) / (genEquivTotal[lineasFlujo[l]][h])));
                        writerCSV.write("na");
                        writerCSV.write(',');
                        writerCSV.write("" + (Gx[gflux[g]][e][h]));
                        writerCSV.write(',');
//                        writer.append("" + (D[datosGener[gflux[g]][0]][lineasFlujo[l]][h]));
                        writerCSV.write("" + D.get(datosGener[gflux[g]][0], lineasFlujo[l], h));
                        writerCSV.write(',');
                        writerCSV.write("" + (A[datosGener[gflux[g]][0]][lineasFlujo[l]]));
                        writerCSV.write(',');
                        writerCSV.write("" + (Dref[lineasFlujo[l]][h]));
                        writerCSV.write('\n');
                    }
                }
            }
            for (int l = 0; l < lineasFlujo.length; l++) {
                for (int g = 0; g < gflux.length; g++) {
                    writerCSV.write(String.valueOf(e));
                    writerCSV.write(',');
                    writerCSV.write("med");
                    writerCSV.write(',');
                    writerCSV.write(String.valueOf(lineasFlujo[l]));
                    writerCSV.write(',');
                    writerCSV.write(String.valueOf(gflux[g]));
                    writerCSV.write(',');
                    writerCSV.write("" + (Prorratas[lineasFlujo[l]][gflux[g]][e]));
                    writerCSV.write('\n');
                }
            }
        } catch (IOException f) {
            f.printStackTrace(System.out);
        } finally {
            if (writerCSV != null) {
                try {
                    writerCSV.close();
                } catch (IOException ex) {
                    System.out.println("No se pudo cerrar conexion con prorratas.csv. Error: " + ex.getMessage());
                    ex.printStackTrace(System.out);
                }
            }
        }
    }
    
    static void appendToDebugConsumo(String DirBaseSalida, int etapa, float[][]conAjustado) {
        BufferedWriter writerCSV = null;
        try {
            writerCSV = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(DirBaseSalida + PeajesConstant.SLASH + "consumos.csv", true), PeajesConstant.CSV_ENCODING));
            int numBarras = conAjustado.length;
            int numHid = conAjustado[0].length;
            for (int h = 0; h < numHid; h++) {

                writerCSV.write(Float.toString(h));
                writerCSV.write(",");
                writerCSV.write(Float.toString(etapa));
                for (int b = 0; b < numBarras; b++) {
                    writerCSV.write(",");
                    writerCSV.write(Float.toString(conAjustado[b][h]));
                }
                writerCSV.write('\n');
            }
        } catch (IOException f) {
            f.printStackTrace(System.out);
        } finally {
            if (writerCSV != null) {
                try {
                    writerCSV.close();
                } catch (IOException ex) {
                    System.out.println("No se pudo cerrar conexion con consumos.csv. Error: " + ex.getMessage());
                    ex.printStackTrace(System.out);
                }
            }
        }
    }

    static public void EscribePropiedades(Properties propiedades, String ruta) {
        try {
            File f = new File(ruta);
            OutputStream out = new FileOutputStream(f);
            propiedades.store(out, "Parametros Peajator");
            out.close();
        } catch (IOException e) {
            System.out.println("No se puede escribir archivo de configuracion en ruta " + ruta);
            e.printStackTrace(System.out);
        } catch (Exception e) {
            e.printStackTrace(System.out);
        }
    }
    
    /**
     * Esta rutina verifica que el nombre del rango (u hoja) sea valido
     * <br>Es decir, que no contenga caracteres prohibidos por Excel segun
     * documenta getNameName() de poi v4.0.1 (literal):
     * <li> Valid characters The first character of a name must be a letter, an
     * underscore character (_), or a backslash (\). </li>
     * <li>Remaining characters in the name can be letters, numbers, periods,
     * and underscore characters. Cell references disallowed </li>
     * <li>Names cannot be the same as a cell reference, such as Z$100 or R1C1.
     * </li>
     * <li>Spaces are not valid Spaces are not allowed as part of a name. </li>
     * <li>Use the underscore character (_) and period (.) as word separators,
     * such as, Sales_Tax or First.Quarter. Name length A name can contain up to
     * 255 characters. Case sensitivity Names can contain uppercase and
     * lowercase letters. </li>
     *
     * @param nombreOriginal nombre original del rango. null creara una nullpointerexception
     * @return nombre verificado del rango
     */
    static public String creaNameExcelSeguro(String nombreOriginal) {
        if (nombreOriginal.isEmpty()) {
            return nombreOriginal;
        }
        String nombreSeguro = nombreOriginal.replaceAll("[^\\p{IsAlphabetic}^\\p{IsDigit}]", "_");
        try {
            String sFirstNumber = nombreSeguro.substring(0, 1);
            int nFirstNumber = Integer.parseInt(sFirstNumber);
            nombreSeguro = nombreSeguro.replaceFirst(sFirstNumber, "N" + nFirstNumber);
        } catch (NumberFormatException e) {
            //Continuar
        }
        return nombreSeguro;
    }
    
}

