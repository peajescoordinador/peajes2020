package cl.coordinador.peajes;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author vtoro
 */

import java.io.*;
public class EscribeArchivosFinales {
    private static final int numMeses = 12;
    private static String slash = File.separator;
    static String[] nomMes = {"Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul",
            "Ago", "Sep", "Oct", "Nov", "Dic"};

   public static void EscribeLiqMes(String mes, int Ano, File DirBaseSal,File DirBaseEnt,String FechPago) {
          int mesLiq=0;
          for(int i=0;i<numMeses;i++){
              if(mes.equals(nomMes[i]))
                  mesLiq=i;
          }
    //Guarda en el arreglo de pago total e IT, sólo los pagos por inyecciones de cada empresa
    double[][][] RetMasInyEmpTx=new double[PeajesRet.numEmp][PeajesRet.numTx][numMeses];
    double[][][] RetMasInyItEmpTx=new double[PeajesRet.numEmp][PeajesRet.numTx][numMeses];
    double[][] TotPagoEmp=new double[PeajesRet.numEmp][numMeses];
    double[][] TotItEmp=new double[PeajesRet.numEmp][numMeses];
    for (int j = 0; j < PeajesRet.numEmp; j++) {
        for (int t = 0; t < PeajesRet.numTx; t++) {
            for (int m = 0; m < numMeses; m++) {
            RetMasInyEmpTx[j][t][m]=PeajesIny.PagoEmpTxO[j][t][m];
            RetMasInyItEmpTx[j][t][m]= PeajesIny.ItEmpTxO[j][t][m];
        }
    }
    }
    //Agrega al pago total los pagos por retiros asociado a la empresa suministradora
    for (int j = 0; j < PeajesRet.numEmp; j++) {
                for (int m = 0; m < numMeses; m++) {
                for (int t = 0; t < PeajesRet.numTx; t++) {
                    RetMasInyEmpTx[j][t][m]=RetMasInyEmpTx[j][t][m]+PeajesRet.TotRetEmpTxO[j][t][m];
                    RetMasInyItEmpTx[j][t][m]=RetMasInyItEmpTx[j][t][m]+PeajesRet.TotRetItEmpTxO[j][t][m];
                    TotPagoEmp[j][m]+=RetMasInyEmpTx[j][t][m];
                    TotItEmp[j][m]+=RetMasInyItEmpTx[j][t][m];
                }
                }
               }

     //Escribre Archivos
     String libroSalidaGXLSMes= DirBaseSal + slash +"Liquidación" + nomMes[mesLiq] + ".xlsx";

     Escribe.crearLibro(libroSalidaGXLSMes);
     //esta función es igual a creaLiquidacionMes, excepto por el formato de datos, unificar para tener una sola funci—n
      Escribe.creaLiquidacion(mesLiq,
                "Liquidación de Peajes del Sistema de Transmisión Troncal",
                PeajesIny.PagoEmpTxO,
                PeajesRet.TotRetEmpTxO,
                RetMasInyEmpTx,
                RetMasInyItEmpTx, /* aca va It por emp y Tx*/
                PeajesIny.PagoEmpO,
                PeajesRet.TotRetEmpO,
                TotPagoEmp,
                TotItEmp,   /* aca va TotIt*/
                "Empresa Usuaria",
                PeajesRet.nomEmpO,
                "Transmisor",
                PeajesRet.nombreTx,
                "Cuadro N° 1: Cuadro de Pagos Total (Ver Notas 1 y 2)",
                "Cuadro N° 2: Pagos por Inyección",
                "Cuadro N° 3: Pagos por Retiro",
                "Cuadro N° 4: IT Total",
                libroSalidaGXLSMes,"Total",Ano,FechPago,
                "#,###,##0;[Red]-#,###,##0;\"-\"",
                "(1) Valores Positivos: Usuarios Pagan, Valores Negativos: Usuarios Reciben",
                "(2) Suma de Cuadros NÁ 2 y NÁ 3");
      
     Escribe.creaLiquidacionMes(mesLiq,
                "Pago de Peajes por Retiro",
                PeajesRet.RetEmpSinAjuTxO,
                 PeajesRet.TotRetEmpTxRE2288OO,
                 PeajesRet.TotRetEmpTxO,
                 PeajesRet.TotRetItEmpTxO,
                 PeajesRet.RetEmpSinAjuO,
                 PeajesRet.TotRetEmpRE2288O,
                 PeajesRet.TotRetEmpO,
                 PeajesRet.TotItRetEmpO,
                "Empresa Usuaria",
                PeajesRet.nomEmpO,
                "Transmisor", 
                PeajesRet.nombreTx,
                "Tabla 2-1: Pagos de Peajes de Retiro por Suministrador",
                "Tabla 2-2: Pago de Retiro por RE2288",
                "Tabla 2-3: Pagos de Peajes de Retiro Incluyendo Pago de Retiro por RE2288",
                "Tabla 2-4: IT de Retiro",
                libroSalidaGXLSMes,
                "Peajes_Retiros",
                Ano,
                "#,###,##0;[Red]-#,###,##0;\"-\"");
     
     Escribe.creaProrrataMes(mesLiq,
                "Participación de Retiros [%]",PeajesRet.prorrMesC,"Participación "+nomMes[mesLiq],
                "Cliente",PeajesRet.nomCli,
                "Línea",  PeajesRet.nomLinTx,
                "AIC", PeajesRet.zonaLinTx,
                libroSalidaGXLSMes, "PartRet"+nomMes[mesLiq],
                "#,###,##0;[Red]-#,###,##0;\"-\"");
     Escribe.creaLiquidacionMes(mesLiq,
                "Pago de Peajes por Inyección [$]",
                PeajesIny.peajeEmpTxO,
                PeajesIny.AjusMGNCEmpTxO,
                PeajesIny.PagoEmpTxO,
                PeajesIny.ItEmpTxO,
                PeajesIny.peajeEmpGO,
                PeajesIny.AjusMGNCEmpO,
                PeajesIny.PagoEmpO,
                PeajesIny.ItEmpO,
                "Empresa",
                PeajesIny.nomEmpGO,
                "Transmisor",
                PeajesIny.nombreTx,
                "Tabla 1-1: Pagos de Peajes de Inyección por Empresa",
                "Tabla 1-2: Ajuste por Exención de MGNC",
                "Tabla 1-3: Pagos de Peajes de Inyección Incluyendo Ajuste por MGNC",
                "Tabla 1-4: IT por Inyección Incluyendo Ajuste por MGNC",
                libroSalidaGXLSMes,
                "Pje_Inyección",
                Ano,
                "#,###,##0;[Red]-#,###,##0;\"-\"");
Escribe.creaLiquidacionMesIny(mesLiq,
                "Pago de Peajes de Inyección",
                PeajesIny.peajeCenTxO,
                PeajesIny.AjusMGNCTxO,
                PeajesIny.PagoTotCenTxO,
                PeajesIny.peajeGenO,
                PeajesIny.AjusMGNCTotO,
                PeajesIny.PagoTotCenO,
                "Central",PeajesIny.nomGenO,
                "Transmisor", PeajesIny.nombreTx,
                
                "MGNC", PeajesIny.MGNCO,
                "PNeta", PeajesIny.PotNetaO,
                
                "Inyeccion [GWh]",PeajesIny.GenPromMesCenO,
                
                "Factor",PeajesIny.facPagoO ,
                
                libroSalidaGXLSMes,"PagosXCentral",Ano,
                "#,###,##0;[Red]-#,###,##0;\"-\"");
 Escribe.crea_AjusteCentrales(mesLiq,
            "Pagos Exentos "+nomMes[mesLiq]+" (Valores en $) por Central",
            PeajesIny.ExcenMGNCTxO,
            PeajesIny.AjusMGNCTxO,
            PeajesIny.ExcenMGNCO,
            PeajesIny.AjusMGNCTotO,
            "Central",PeajesIny.nomMGNCO,PeajesIny.nomGenO,
            "Transmisor", PeajesIny.nombreTx,
            libroSalidaGXLSMes, "Ajus"+nomMes[mesLiq], Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
Escribe.creaProrrataMes(mesLiq,
                "Participación de Inyecciones [%]",PeajesIny.prorrMesGenTx,
                "Participación "+nomMes[mesLiq],
                "Cliente",PeajesIny.nomGen,
                "L’nea",  PeajesIny.nomLinTx,
                "AIC", PeajesIny.zonaLinTx,
                libroSalidaGXLSMes, "PartIny"+nomMes[mesLiq],
                "#,###,##0;[Red]-#,###,##0;\"-\"");
System.out.println();
System.out.println("Archivos de Liquidación Mensual terminados");
System.out.println();

      }
    public static void EscribeLiqAno(String mes, int Ano, File DirBaseSal) {
        //Escribre Archivos
        String libroSalidaGXLSAno= DirBaseSal + slash +"Cuadros" +Ano + ".xlsx";
        
        

        
        
        
        
        Escribe.crearLibro(libroSalidaGXLSAno);
        Escribe.crea_1TablaTx_1C(
            "Pagos de Peajes de Inyección por Empresas (Valores en $)",PeajesIny.PeajeAnualEmpGTxO,
            PeajesIny. PeajeAnualEmpGO,
            "Empresa",PeajesIny.nomEmpGO,
            "Transmisor", PeajesIny.nombreTx,
            libroSalidaGXLSAno, "PagIny", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_1TablaTx_2C(
            "Pagos de Peajes de Inyección por MGNC (Valores en $)",PeajesIny.peajeAnualMGNCTxO,
            PeajesIny.peajeAnualMGNCO,
            "Central",PeajesIny.nomMGNCO,
            "Transmisor", PeajesIny.nombreTx,
            "P. Neta [MW]",PeajesIny.PotNetaMGNCO,
            libroSalidaGXLSAno, "PagMGNC", Ano, "#,##0.00");
        Escribe.crea_1TablaTx_2C(
            "Excención de Pagos de Peajes de Inyección por MGNC (Valores en $)",PeajesIny.ExcenAnualMGNCTxO,
            PeajesIny.ExcenAnualMGNCO,
            "Central",PeajesIny.nomMGNCO,
            "Transmisor", PeajesIny.nombreTx,
            "Factor Pago[%]",PeajesIny.facPagoMGNCO,
            libroSalidaGXLSAno, "ExcMGNC", Ano, "##0.00%");
        Escribe.crea_1TablaTx_1C(
                 "Ajustes por Exención de MGNC (Valores en $)",PeajesIny.AjusMGNCAnualEmpTxO,
                 PeajesIny.AjusMGNCAnualEmpO,
                 "Central",PeajesIny.nomEmpGO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "AjusIny", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_1TablaTx_1C(
                 "Pagos de peajes de Inyección incluyendo ajustes por Exención de MGNC (Valores en $)",PeajesIny.PagoAnualEmpGTxO,
                 PeajesIny.PagoAnualEmpGO,
                 "Central",PeajesIny.nomEmpGO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "Pago_AjIny", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_1TablaTx_1C(
                 "Pagos de peajes de Retiro por Suministrador (Valores en $)",PeajesRet.TotAnualPjeRetEmpTxO,
                 PeajesRet.TotAnualPjeRetEmpO,
                 "Central",PeajesRet.nomEmpO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "PagoRet", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_1TablaTx_1C(
                 "Pagos de peajes de Retiro RE2288 por Suministrador (Valores en $)",PeajesRet.TotAnualPjeRetEmpTxRE2288O,
                 PeajesRet.TotAnualPjeRetEmpRE2288O,
                 "Central",PeajesRet.nomSumiRM88O,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "PagoRE2288", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");

        Escribe.crea_1TablaTx_1C(
                 "Pagos de peajes de Retiro por Suministrador Incluyendo RE2288 (Valores en $)",PeajesRet.TotConRe2288AnualPjeRetEmpTxO,
                 PeajesRet.TotConRe2288AnualPjeRetEmpO,
                 "Central",PeajesRet.nomEmpO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "PagoRet_RE2288", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");

        if(PeajesRet.numClienExentos!=0){
        Escribe.crea_1TablaTx_1C(
                 "Clientes con Peajes Exceptuados por Contratos hasta el 6/05/2002 (Valores en $)",PeajesRet.pjeAnualClienTxExenO,
                 PeajesRet.pjeAnualClienExenO,
                 "Central",PeajesRet.nombreClientesExenO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "PagoRetEx", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
        Escribe.crea_1TablaTx_1C_double(
                 "Ajuste de Peajes por clientes no regulados, contratados antes del 6/5/2002 (Valores en $)",PeajesRet.AjusAnualEmpCTxO,//poner los anuales
                 PeajesRet.AjusAnualEmpCO,
                 "Central",PeajesRet.nomEmpO,
                 "Transmisor", PeajesIny.nombreTx,
                 libroSalidaGXLSAno, "AjusteRet", Ano, "#,###,##0;[Red]-#,###,##0;\"-\"");
    }
   System.out.println();
   System.out.println("Cuadro Anual terminado");
   System.out.println();
    }


     }


