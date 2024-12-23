USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_VENTA_FACTURA_SFS]    Script Date: 10/24/2021 18:50:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--exec SP_VENTA_FACTURA_SFS '01','1',411

ALTER PROC [dbo].[SP_VENTA_FACTURA_SFS]
    @CODCIA CHAR(2) ,
    @SERIE CHAR(3) ,
    @NUMERO BIGINT
AS
    SET NOCOUNT ON
    
    DECLARE @mIGV INT ,
        @vIGV1 DECIMAL(16, 2) ,
        @vIGV2 DECIMAL(16, 2)
    --SELECT TOP 1
    --        @mIGV = G.GEN_IGV
    --FROM    dbo.GENERAL g
	select top 1 @mIGV = coalesce(p.PAR_IGV,0) from PARGEN p where p.PAR_CODCIA = @CODCIA
	
    SET @vIGV1 = ( @mIGV / 100.00 ) + 1 --1.18
    SET @vIGV2 = @mIGV / 100.00 --0.18
    
    --select @mIGV,@vIGV1
--1. CAB
    SELECT TOP 1
            CASE WHEN F.FAR_FBG = 'F' THEN '01'
                 ELSE '03'
            END + '-' + F.FAR_FBG + RIGHT('000' + RTRIM(LTRIM(F.FAR_NUMSER)),
                                          3) + '-'
            + CAST(F.FAR_NUMFAC AS VARCHAR(20)) AS 'NOMBRE' ,
            '0101' AS 'TIPOPERACION' ,
             DBO.UFN_FECHASUNAT(F.FAR_FECHA_COMPRA) AS 'FECEMISION',
            CONVERT(VARCHAR(20),GETDATE(),108) AS 'HORA',
            case when f.far_num_lote = 4 then
           dbo.UFN_FECHASUNAT( DATEADD(day,f.FAR_DIAS,f.FAR_FECHA_COMPRA))
           else
           '-' end as 'FECHAVENC' ,
            --dbo.UFN_FECHASUNAT( DATEADD(day,f.FAR_DIAS,f.FAR_FECHA_COMPRA)) AS 'FECHAVENC' ,
            CASE WHEN @CODCIA = '01' THEN '0000'
                 ELSE 
					CASE WHEN @CODCIA = '02' THEN '0000'
					ELSE 
						CASE WHEN @CODCIA = '03' THEN '0000'
					ELSE 
						CASE WHEN @CODCIA = '04' THEN '0000'	
						ELSE '0000'
						END
					END
			END
            END AS 'CODLOCALEMISOR' ,
            CASE WHEN F.FAR_FBG = 'F' THEN '6'
                 ELSE CASE WHEN F.FAR_FBG = 'B' THEN '1'
                           ELSE ''
                      END
            END AS 'TIPDOCUSUARIO' ,
            CASE WHEN F.FAR_FBG = 'F' THEN RTRIM(LTRIM(C.CLI_RUC_ESPOSO))
                 ELSE CASE WHEN F.FAR_FBG = 'B'
                                AND RTRIM(LTRIM(( C.CLI_RUC_ESPOSA ))) = ''
                           THEN '11111111'
                           ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                      END
            END AS 'NUMDOCUSUARIO' ,
            --RTRIM(LTRIM(REPLACE(c.CLI_NOMBRE_ESPOSO, '&', 'Y'))) AS 'RZNSOCIALUSUARIO' ,
            RTRIM(LTRIM(REPLACE(REPLACE(c.CLI_NOMBRE_ESPOSO,CHAR(10),' ') , '&', 'Y'))) AS 'RZNSOCIALUSUARIO',
            CASE WHEN F.FAR_MONEDA = 'S' THEN 'PEN'
                 ELSE 'USD'
            END AS 'TIPMONEDA' ,
            --MODIFICADO
            F.FAR_IMPTO AS 'MTOIGV' , --SUMATORIA DE TRIBUTOS
            --CASE WHEN f.far_impto = 0 THEN 0
           --      ELSE ( F.FAR_BRUTO - F.FAR_TOT_DESCTO )
           -- END AS 'MTOOPERGRAVADAS' ,--OPERACIONES GRABADAS     TOTAL VALOR DE VENTA
           ( F.FAR_BRUTO - F.FAR_TOT_DESCTO )AS 'MTOOPERGRAVADAS',
            CAST(( SELECT TOP 1
                            A.ALL_NETO
                   FROM     dbo.ALLOG a
                   WHERE    A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_NUMSER = @SERIE
                            AND A.ALL_TIPMOV = 10
                            AND A.ALL_CODTRA = '2401'
                            AND A.ALL_FBG = 'F'
                            AND A.ALL_CODCIA = @CODCIA
                 ) AS DECIMAL(16, 2)) AS 'MTOIMPVENTA' , --TOTAL PRECIO DE VVENTA
            F.FAR_TOT_DESCTO AS 'SUMDSCTOGLOBAL' , --SUMATORIA DE OTROS DESCUENTOS     TOTAL DESCUENTOS
            '0.00' AS 'SUMOTROSCARGOS' , --SUMATORIA DE OTROS CARGOS
            '0.00' AS 'TOTANTICIPOS' , --TOTAL ANTICIPOS
            CAST(( ( SELECT TOP 1
                            A.ALL_NETO
                     FROM   dbo.ALLOG a
                     WHERE  A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_NUMSER = @SERIE
                            AND A.ALL_TIPMOV = 10
                            AND A.ALL_CODTRA = '2401'
                            AND A.ALL_FBG = 'F'
                            AND A.ALL_CODCIA = @CODCIA
                   ) - F.FAR_TOT_DESCTO ) AS DECIMAL(16, 2)) AS 'IMPTOTALVENTA' , --IMPORTE TOTAL DE LA VENTA
            '2.1' AS 'UBL' ,
            '2.0' AS 'CUSTOMDOC',
            '' AS 'ACA1','000' AS 'ACA2','0' AS 'ACA3','0.00' AS 'ACA4','' AS 'ACA5',
            'PE' AS 'PAIS','130101' AS 'UBIGEO',
            --COALESCE(c.CLI_CASA_DIREC,'') AS 'DIR',
            RTRIM(LTRIM(REPLACE(COALESCE(c.CLI_CASA_DIREC,''),CHAR(10),' '))) AS 'DIR',
            ' ' AS 'PAIS1',
            ' ' AS 'UBIGEO1',' ' AS 'DIR1'
    FROM    dbo.FACART f
            INNER JOIN dbo.CLIENTES c ON F.FAR_CODCIA = C.CLI_CODCIA
                                         AND F.FAR_CODCLIE = C.CLI_CODCLIE
            INNER JOIN dbo.ARTI a2 ON F.FAR_CODCIA = A2.ART_CODCIA
                                      AND F.FAR_CODART = A2.ART_KEY
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 10
            AND f.FAR_FBG = 'F'
            and f.FAR_CODCIA= @CODCIA

--2. DET

    SELECT  'NIU' AS 'CODUNIDADMEDIDA' ,
            (F.FAR_CANTIDAD/F.FAR_EQUIV) AS 'CTDUNIDADITEM' ,
            CASE WHEN F.FAR_CODART = 0 THEN ''
                 ELSE F.FAR_CODART
            END AS 'CODPRODUCTO' ,
            '-' AS 'CODPRODUCTOSUNAT' ,
            CASE WHEN F.FAR_CODART = 0 THEN F.FAR_CONCEPTO
                 ELSE RTRIM(LTRIM(A2.ART_NOMBRE))
            END AS 'DESITEM' ,
            cast(CASE WHEN f.far_impto = 0 THEN 
						ROUND(CAST(F.FAR_PRECIO AS MONEY), 4)
                 ELSE ROUND(CAST(F.FAR_PRECIO AS MONEY)/ @vIGV1, 4)
                                     
            END as money) AS 'MTOVALORUNITARIO' ,
            --MODIFICADO
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
                        THEN ROUND(CAST(F.FAR_PRECIO AS MONEY)
                                   , 4)
                        ELSE ROUND(CAST(F.FAR_PRECIO AS MONEY)/ CAST(@vIGV1 AS MONEY), 4)
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
             CAST(( @mIGV) AS DECIMAL(16, 2)) AS 'PORCIGV' ,
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
                      THEN ROUND(CAST(F.FAR_PRECIO AS MONEY)
                                 , 4)
                      ELSE ROUND(CAST(F.FAR_PRECIO AS MONEY)
                                  / CAST( @vIGV1 AS MONEY), 4)
                 END * (F.FAR_CANTIDAD/F.FAR_EQUIV) AS DECIMAL(16, 2)) AS 'VALORVTAXITEM' ,
            --'-' AS 'GRATUITO'
            '0.00' AS 'GRATUITO'
    FROM    dbo.FACART f
            INNER JOIN dbo.CLIENTES c ON F.FAR_CODCIA = C.CLI_CODCIA
                                         AND F.FAR_CODCLIE = C.CLI_CODCLIE and c.CLI_CP = 'C'
            INNER JOIN dbo.ARTI a2 ON F.FAR_CODCIA = A2.ART_CODCIA
                                      AND F.FAR_CODART = A2.ART_KEY
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 10
            AND f.FAR_FBG = 'F'
            and f.FAR_CODCIA = @CODCIA
            AND f.FAR_PRECIO <> 0
            
--TRI
    SELECT TOP 1
            '1000' AS 'CODIGO' ,
            'IGV' AS 'NOMBRE' ,
            'VAT' AS 'COD',
            CASE WHEN f.far_impto = 0 THEN 0
                 ELSE ( F.FAR_BRUTO - F.FAR_TOT_DESCTO )
            END AS 'BASEIMPONIBLE' ,
            F.FAR_IMPTO AS 'TRIBUTO'
    FROM    dbo.FACART f
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 10
            AND f.FAR_FBG = 'F'
            and f.FAR_CODCIA = @CODCIA
    UNION
   SELECT  TOP 1
            '9997' AS 'CODIGO' ,
            'EXO' AS 'NOMBRE' ,
            'VAT' AS 'COD',
            CASE WHEN f.far_impto = 0 THEN ( F.FAR_BRUTO - F.FAR_TOT_DESCTO )
                 ELSE 0.00
            END AS 'BASEIMPONIBLE' ,
            0.00 AS 'TRIBUTO'
    FROM    dbo.FACART f
    WHERE   F.FAR_NUMSER = @SERIE
            AND F.FAR_NUMFAC = @NUMERO
            AND f.FAR_TIPMOV = 10
            AND f.FAR_FBG = 'F'
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
	(SELECT TOP 1 COALESCE(A.ALL_ICBPER,0) FROM dbo.ALLOG a WHERE A.ALL_CODCIA = @CODCIA AND A.ALL_NUMSER = @SERIE AND A.ALL_NUMFAC = @NUMERO AND A.ALL_CODTRA = 2401 AND A.ALL_FBG ='F') AS 'TRIBUTO'

    SELECT  '1000' AS 'COD' 
/*
exec SP_VENTA_FACTURA_SFS '01','1',345
exec SP_VENTA_FACTURA_SFS '01','1',344
*/
--PARA LA FORMA DE PAGO 2021-09-01
--PAG
select CASE ALL_SECUENCIA WHEN 4 THEN 'Credito' ELSE 'Contado' END AS 'FORMAPAGO', 
CASE ALL_SECUENCIA WHEN 4 THEN cast(cast(coalesce(ALL_IMPORTE_AMORT,0) as numeric(16,2)) as varchar(20)) ELSE cast(cast(coalesce(ALL_IMPORTE_AMORT,0) as numeric(16,2)) as varchar(20)) END AS 'PENDIENTEPAGO',
case when ALL_SECUENCIA = 4 then
 CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                 ELSE 'USD'
            END 
            else 'PEN' end AS 'TIPMONEDA'
from ALLOG
WHERE ALL_CODCIA = @CODCIA AND ALL_NUMSER = @SERIE AND ALL_NUMFAC = @NUMERO and ALL_FBG = 'F'


--DPA
SELECT CASE ALL_SECUENCIA WHEN 4 THEN cast(cast(coalesce(ALL_IMPORTE_AMORT,0) as numeric(16,2)) as varchar(20)) ELSE '-' END AS 'CUOTAPAGO',
CASE ALL_SECUENCIA WHEN 4 THEN DBO.UFN_FECHASUNAT(ALL_FECHA_VCTO) ELSE '-' END AS 'FECHAVCTO',
case ALL_SECUENCIA when 4 then
 CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                 ELSE 'USD'
            END else '-' end AS 'TIPMONEDA'
FROM ALLOG
WHERE ALL_CODCIA = @CODCIA AND ALL_NUMSER = @SERIE AND ALL_NUMFAC = @NUMERO and ALL_FBG = 'F'

--RTN
SELECT '-' AS 'IMPOPERACION',
'-' AS 'PORRETENCION',
'-' AS 'IMPRETENCION'

--SELECT * FROM ALLOG WHERE ALL_NUMFAC IN(1,2)

