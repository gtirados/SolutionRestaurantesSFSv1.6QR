USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_NOTACREDITO_SFS]    Script Date: 09/28/2021 22:15:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
SP_FACTURAVENTA_FE '01','10',42640
SP_FACTURAVENTA_FE '01','10',42648
SP_VENTA_SFS '01','10',7113
SP_NOTACREDITO_SFS '01','1',10
SP_NOTACREDITO_SFS '01','1',13
SP_NOTACREDITO_SFS '01','1',14
*/
ALTER PROC [dbo].[SP_NOTACREDITO_SFS]
    @CODCIA CHAR(2) ,
    @SERIE CHAR(3) ,
    @NUMERO BIGINT
AS
    SET NOCOUNT ON
    
    DECLARE @mIGV INT ,
        @vIGV1 DECIMAL(8, 2) ,
        @vIGV2 DECIMAL(8, 2)
    SELECT TOP 1
            @mIGV = G.GEN_IGV
    FROM    dbo.GENERAL g
    SET @vIGV1 = ( @mIGV / 100.00 ) + 1
    SET @vIGV2 = @mIGV / 100.00
    
    --SELECT @vIGV2
--1. CAB
    SELECT TOP 1
            --CASE WHEN F.FAR_FBG = 'F' THEN '01'
            --     ELSE '03'
            --END 
            '07'+ '-' + 'F' + RIGHT('000' + RTRIM(LTRIM(F.FAR_NUMSER)),
                                          3) + '-'
            + CAST(F.FAR_NUMFAC AS VARCHAR(20)) AS 'NOMBRE' ,'0101' AS 'CAMPO1',
            CAST(YEAR(f.FAR_FECHA_COMPRA) AS VARCHAR(4)) + '-' + RIGHT('00'
                                                              + CAST(MONTH(F.FAR_FECHA_COMPRA) AS VARCHAR(2)),
                                                              2) + '-'
            + RIGHT('00' + CAST(DAY(F.FAR_FECHA_COMPRA) AS VARCHAR(2)), 2) AS 'FECEMISION' ,
            --CONVERT(VARCHAR(8), CONVERT(DATETIME, CONVERT(VARCHAR(20), FAR_FECHA_COMPRA, 103)
            --+ ' ' + RTRIM(LTRIM(REPLACE(FAR_HORA, '.', '')))), 108) AS 'HORA' ,
            CONVERT(VARCHAR(20),GETDATE(),108) AS 'HORA',
             '0000' AS 'CAMPO2','6' AS 'CAMPO3',
            RTRIM(LTRIM(C.CLI_RUC_ESPOSO)) AS 'RUC',
            RIGHT('00' + CAST(F.FAR_NUM_LOTE AS VARCHAR(2)), 2) AS 'CODMOTIVO' ,
            ( SELECT TOP 1
                        A.ALL_SUBTRA
              FROM      dbo.ALLOG a
              WHERE     A.ALL_CODCIA = F.FAR_CODCIA
                        AND A.ALL_NUMSER = F.FAR_NUMSER
                        AND A.ALL_NUMFAC = F.FAR_NUMFAC
                        AND A.ALL_CODTRA = 2412
            ) AS 'DESCMOTIVO' ,
            CASE WHEN LEFT(RTRIM(LTRIM(F.FAR_CONCEPTO)),1) = 'F' THEN '01'
            ELSE '03' END  AS 'TIPODOCAFECTADO' ,
            RTRIM(LTRIM(F.FAR_CONCEPTO)) AS 'NUMDOCAFECTADO' ,
            CASE WHEN F.FAR_COD_SUNAT = 1 THEN--FACTURA
                      '6'
                 ELSE '1'
            END AS 'TIPDOCUSUARIO' ,
            CASE WHEN F.FAR_COD_SUNAT = 1 THEN--FACTURA
                      RTRIM(LTRIM(C.CLI_RUC_ESPOSO))
                 ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
            END AS 'NUMDOCUSUARIO' ,
            RTRIM(LTRIM(REPLACE(c.CLI_NOMBRE,'&','Y'))) AS 'CLI1' ,
            CASE WHEN F.FAR_MONEDA = 'S' THEN 'PEN'
                 ELSE 'USD'
            END AS 'TIPMONEDA' ,
            '0.00' AS 'SUMOTROSCARGOS' , --SUMATORIA DE OTROS CARGOS
            CASE WHEN f.far_impto = 0 THEN 0
                 ELSE ROUND((F.FAR_BRUTO - F.FAR_TOT_DESCTO),2)
            END AS 'MTOOPERGRAVADAS' ,--OPERACIONES GRABADAS
            '0.00' AS 'MTOOPERINAFECTAS' ,--OPERACIONES INFECTAS
            CASE WHEN f.far_impto = 0 THEN ( F.FAR_BRUTO - F.FAR_TOT_DESCTO )
                 ELSE 0.00
            END AS 'MTOOPEREXONERADAS' ,--OPERACIONES EXONERADAS
            ROUND(F.FAR_IMPTO,2) AS 'MTOIGV' ,
            '0.00' AS 'MTOISC' ,--SUMATORIA ISC
            '0.00' AS 'MTOOTROSTRIBUTOS' , --SUMATORIA DE OTROS TRIBUTOS
            ( SELECT TOP 1
                        A.ALL_NETO
              FROM      dbo.ALLOG a
              WHERE     A.ALL_NUMFAC = @NUMERO
                        AND A.ALL_NUMSER = @SERIE
                        AND A.ALL_TIPMOV = 97
                        AND A.ALL_CODTRA = '2412'
                        AND A.ALL_FBG IN ( 'N' )
            ) AS 'MTOIMPVENTA',
            '2.1' AS 'CAMPO5',
            '2.0' AS 'CAMPO6',
            '' AS 'ACA1','000' AS 'ACA2','0' AS 'ACA3','0.00' AS 'ACA4','' AS 'ACA5',
            'PE' AS 'PAIS','130101' AS 'UBIGEO',
            COALESCE(c.CLI_CASA_DIREC,'') AS 'DIR',
            ' ' AS 'PAIS1',
            ' ' AS 'UBIGEO1',' ' AS 'DIR1'
    FROM    dbo.FACART f
            INNER JOIN dbo.CLIENTES c ON F.FAR_CODCIA = C.CLI_CODCIA
                                         AND F.FAR_CODCLIE = C.CLI_CODCLIE
            LEFT JOIN dbo.ARTI a2 ON F.FAR_CODCIA = A2.ART_CODCIA
                                     AND F.FAR_CODART = A2.ART_KEY
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 97
            AND f.FAR_FBG IN ( 'N' )
            
           

--2. DET

    SELECT  'NIU' AS 'CODUNIDADMEDIDA' ,
          CASE WHEN  F.FAR_CANTIDAD = 0 THEN 1 ELSE (F.FAR_CANTIDAD/F.FAR_EQUIV) END AS 'CTDUNIDADITEM' ,
            CASE WHEN F.FAR_CODART = 0 THEN ''
                 ELSE CAST(F.FAR_CODART AS VARCHAR(20))
            END AS 'CODPRODUCTO' ,
            '-' AS 'CODPRODUCTOSUNAT' ,
            COALESCE(RTRIM(LTRIM(A2.ART_NOMBRE)),
                     ( SELECT TOP 1
                                COALESCE(RTRIM(LTRIM(ax.ALL_CONCEPTO)), '')
                       FROM     dbo.ALLOG ax
                       WHERE    AX.ALL_TIPMOV = 97
                                AND AX.ALL_NUMSER = f.FAR_NUMSER
                                AND ax.ALL_NUMFAC = f.FAR_NUMFAC
                     )) AS 'DESITEM' ,
               CAST(CASE WHEN F.FAR_PRECIO = 0 THEN F.FAR_BRUTO
                      ELSE ROUND(F.FAR_PRECIO / @vIGV1, 2)
                 END AS MONEY) AS 'MTOVALORUNITARIO' ,--SIN IGV
          
             CASE WHEN FAR_JABAS = 0
                 THEN CASE WHEN f.far_impto = 0 THEN 0
                           ELSE CAST(ROUND(( ( (F.FAR_CANTIDAD/F.FAR_EQUIV) * F.FAR_PRECIO )
                                             - F.FAR_DESCTO ) / @VIGV1, 2)
                                * @VIGV2 AS DECIMAL(16, 2))
                      END
                 ELSE CASE WHEN f.far_impto = 0 THEN 0
                           ELSE CAST(ROUND(( ( (F.FAR_CANTIDAD/F.FAR_EQUIV) * F.FAR_PRECIO )
                                             - F.FAR_DESCTO ) / @VIGV1, 2)
                                * @VIGV2 AS DECIMAL(16, 2))
                      END
            END AS 'MTOIGVITEM' ,--SUMATORIA TRIBUTOS POR ITEM
            
            CASE WHEN f.far_impto = 0 THEN '9997'
                 ELSE '1000'
            END AS 'CODTIPTRIBUTOIGV',
        --'1000' AS 'CODTIPTRIBUTOIGV' ,--CODIGO DE TIPO DE TRIBUTO IGV
            CASE WHEN FAR_JABAS = 0
                 THEN CASE WHEN f.far_impto = 0 THEN 0
                           ELSE CAST(ROUND(( ( (F.FAR_CANTIDAD/F.FAR_EQUIV) * F.FAR_PRECIO )
                                             - F.FAR_DESCTO ) / @VIGV1, 2)
                                * @VIGV2 AS DECIMAL(16, 2))
                      END
                 ELSE CASE WHEN f.far_impto = 0 THEN 0
                           ELSE CAST(ROUND(( ( (F.FAR_CANTIDAD/F.FAR_EQUIV) * F.FAR_PRECIO )
                                             - F.FAR_DESCTO ) / @VIGV1, 2)
                                * @VIGV2 AS DECIMAL(16, 2))
                      END
            END AS 'MTOIGVITEM1' ,--MONTO IGV POR ITEM
            CAST(( CASE WHEN f.far_impto = 0
                        THEN 0
                        ELSE ROUND(CAST(F.FAR_PRECIO AS MONEY)
                                   / CAST(( ( SELECT TOP 1
                                                        GEN_IGV
                                              FROM      dbo.GENERAL
                                            ) / 100 ) + 1 AS MONEY), 2)
                   END ) * (F.FAR_CANTIDAD/F.FAR_EQUIV) AS DECIMAL(16, 2)) AS 'BASEIMPIGV' ,
            
            --'0.00' AS 'BASEIMPIGV', --BASE IMPONIBLE IGV
            
            CASE WHEN f.far_impto = 0 THEN 'EXO'
            ELSE 'IGV' 
            END AS 'NOMTRIBITEM' , --NOMBRE DE TRIBUTO POR ITEM
            'VAT' AS 'CODTIPTRIBUTOITEM' ,--CODIGO DE TIPO DE TRIBUTO POR ITEM
            CASE WHEN f.far_impto = 0 THEN '20'
                 ELSE '10'
            END AS 'TIPAFEIGV' ,--TIPO AFECTA IGV7
            --'18.00' AS 'PORCIGV', -- PORCENTAJE DE IGV
            CAST(( SELECT TOP 1
                            GEN_IGV
                   FROM     GENERAL
                 ) AS DECIMAL(16, 2)) AS 'PORCIGV' ,
            '-' AS 'CODISC' , --CODIGO ISC
            --'-' AS 'CODOTROITEM', --CODIGO DE OTRO TRIBUTOS
            '0.00' AS 'MONTOISC' ,
            '0.00' AS 'BASEIMPONIBLEISC' ,
            '' AS 'NOMBRETRIBITEM' ,
            '' AS 'CODTRIBITEM' ,
            '' AS 'CODSISISC' ,
            '15.00' AS 'PORCISC' ,
            /*
             SP_VENTA_BOLETA_SFS '01','1',5
             */
            '-' AS 'CODTRIBOTO' ,
            '0.00' AS 'MONTOTRIBOTO' ,
            '0.00' AS 'BASEIMPONIBLEOTO' ,
            '' AS 'TIPSISISC' , --CAMPO POR DEMAS CONSULTAR
            '' AS 'NOMBRETRIBOTO' ,
            '-' AS 'CODTRIBOTO' ,
            '15.00' AS 'PORCOTO' ,
            /*
SP_VENTA_FACTURA_SFS '01','1',7
*/
            CASE WHEN A2.ART_CALIDAD = 0 THEN '7152' ELSE '-' END AS 'CODIGOICBPER',
            CASE WHEN A2.ART_CALIDAD = 0 THEN (SELECT TOP 2 COALESCE(G.GEN_ICBPER,0) FROM GENERAL G) ELSE 0 END AS 'IMPORTEICBPER',
            CASE WHEN A2.ART_CALIDAD = 0 THEN CAST((F.FAR_CANTIDAD/F.FAR_EQUIV)AS DECIMAL(8,2))  ELSE 0 END AS 'CANTIDADICBPER',
            CASE WHEN A2.ART_CALIDAD = 0 THEN 'ICBPER' ELSE '' END AS 'TITULOICBPER',
            CASE WHEN A2.ART_CALIDAD = 0 THEN 'OTH' ELSE '' END AS 'IDEICBPER',
            CASE WHEN A2.ART_CALIDAD = 0 THEN F.FAR_ICBPER ELSE 0 END AS 'MONTOICBPER',
            CAST(FAR_PRECIO AS DECIMAL(16, 2)) AS 'PRECIOVTAUNITARIO' ,
            CAST(CASE WHEN f.far_impto = 0
                      THEN ( SELECT TOP 1
                                    A.ALL_NETO
                             FROM   dbo.ALLOG a
                             WHERE  A.ALL_NUMFAC = @NUMERO
                                    AND A.ALL_NUMSER = @SERIE
                                    AND A.ALL_TIPMOV = 10
                                    AND A.ALL_CODTRA = '2401'
                                    AND A.ALL_FBG = 'F'
                                    AND A.ALL_CODCIA = @CODCIA
                           )
                      ELSE ROUND(CAST(F.FAR_PRECIO AS MONEY)
                                 / CAST(( ( SELECT TOP 1
                                                    GEN_IGV
                                            FROM    dbo.GENERAL
                                          ) / 100 ) + 1 AS MONEY), 2)
                 END * (F.FAR_CANTIDAD/F.FAR_EQUIV) AS DECIMAL(16, 2)) AS 'VALORVTAXITEM' ,
            --'-' AS 'GRATUITO'
            '0.00' AS 'GRATUITO'
    FROM    dbo.FACART f
            INNER JOIN dbo.CLIENTES c ON F.FAR_CODCIA = C.CLI_CODCIA
                                         AND F.FAR_CODCLIE = C.CLI_CODCLIE
            LEFT JOIN dbo.ARTI a2 ON F.FAR_CODCIA = A2.ART_CODCIA
                                     AND F.FAR_CODART = A2.ART_KEY
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 97
            AND f.FAR_FBG IN ( 'N' )
            
            --TRI
    SELECT TOP 1
            '1000' AS 'CODIGO' ,
            'IGV' AS 'NOMBRE' ,
            'VAT' AS 'COD',
            CASE WHEN f.far_impto = 0 THEN 0
                 ELSE ROUND((F.FAR_BRUTO-F.FAR_TOT_DESCTO),2)
            END AS 'BASEIMPONIBLE' ,
            ROUND(F.FAR_IMPTO,2) AS 'TRIBUTO'
    FROM    dbo.FACART f
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 97
            AND f.FAR_FBG = 'N'
            and f.FAR_CODCIA = @CODCIA
    UNION
   SELECT  TOP 1
            '9997' AS 'CODIGO' ,
            'EXO' AS 'NOMBRE' ,
            'VAT' AS 'COD',
            CASE WHEN f.far_impto = 0 THEN ROUND((F.FAR_BRUTO-F.FAR_TOT_DESCTO),2)
                 ELSE 0.00
            END AS 'BASEIMPONIBLE' ,
            0.00 AS 'TRIBUTO'
    FROM    dbo.FACART f
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 97
            AND f.FAR_FBG = 'N'
            and f.FAR_CODCIA = @CODCIA   
    UNION
    SELECT  
            '9998' AS 'CODIGO' ,
            'INA' AS 'NOMBRE' ,
            'FRE' AS 'COD',
            0.00 AS 'BASEIMPONIBLE' ,
            0.0 AS 'TRIBUTO'
	UNION
	SELECT '7152' AS 'CODIGO',
	'ICBPER' AS 'NOMBRE',
	'OTH' AS 'COD',
	'0' AS 'VASEIMPOBIBLE',
	(SELECT TOP 1 COALESCE(A.ALL_ICBPER,0) FROM dbo.ALLOG a WHERE A.ALL_CODCIA = @CODCIA AND A.ALL_NUMSER = @SERIE AND A.ALL_NUMFAC = @NUMERO AND A.ALL_CODTRA = 2412 AND A.ALL_FBG ='N') AS 'TRIBUTO'

    SELECT  '1000' AS 'COD' 
