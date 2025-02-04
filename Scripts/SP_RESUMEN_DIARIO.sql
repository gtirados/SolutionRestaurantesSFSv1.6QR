USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_RESUMEN_DIARIO]    Script Date: 02/02/2023 16:31:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
SP_RESUMEN_DIARIO '03','20211104','20211104'
SP_RESUMEN_DIARIO '01','20191012','20191012'
SP_RESUMEN_DIARIO '01','20230130','20230130',1
exec SP_RESUMEN_DIARIO '01','2020-10-07 00:00:00','2020-10-07 00:00:00'
*/
ALTER PROC [dbo].[SP_RESUMEN_DIARIO]
    @CODCIA CHAR(2) ,
    @FECHA DATE ,
    @FECHAACTUAL DATE,
    @ANULADAS BIT = 0
AS
    SET NOCOUNT ON

--OBTENIENDO IGV

DECLARE @IGV MONEY
SELECT TOP 1 @IGV = COALESCE(P.PAR_IGV,0) FROM PARGEN p WHERE P.PAR_CODCIA = @CODCIA

SET @IGV = (@IGV/ 100) + 1
--SELECT @IGV
  
    DECLARE @TBLDATA TABLE
        (
          SERIE CHAR(3) ,
          NUMERO BIGINT ,
          INDICE INT IDENTITY
        )
    DECLARE @SERIE CHAR(3) ,
        @NUMERO BIGINT ,
        @CODTRA CHAR(4) ,
        @SECUENCIA int
        
        DECLARE @TBLEXTORNADOS TABLE
        (
          SERIE CHAR(3) ,
          NUMERO BIGINT ,
          CODTRA CHAR(4) ,
          SECUENCIA INT ,
          INDICE INT IDENTITY
        )
          DECLARE @MIN INT ,
        @MAX INT
        
         DECLARE @TBLPRINCIPAL TABLE
        (
          FECHADOCTO DATE ,
          FECHACTUAL DATE ,
          TIPODOCTO CHAR(2) ,
          IDOCTO VARCHAR(20) ,
          TDI CHAR(1) ,
          NRODOCUSUARIO VARCHAR(8) ,
          CLIENTE VARCHAR(300) ,
          MONEDA CHAR(3) ,
          CAMPO1 MONEY,
          TOTAL MONEY ,
          EXO MONEY ,
          INA MONEY ,
          GRA MONEY ,
          ICBPER MONEY ,
          --TISC MONEY ,
          --TIGV MONEY ,
          --OTROTRIB MONEY ,
          TOTALVTA MONEY ,
          TIPDOCTOMODIFICA CHAR(2) ,
          SERIEBOLMODIFICA CHAR(4) ,
          NROBOLMODIFICA VARCHAR(20) ,
          REGPERCEPCION VARCHAR(1) ,
          PORCPERCEPCION VARCHAR(1) ,
          BASEIMPERCEPCION VARCHAR(1) ,
          MONTOPERCEPCION VARCHAR(1) ,
          MONTOTOTINCPERCEPCION VARCHAR(1) ,
          ESTADO CHAR(1) ,
          INDICE INT IDENTITY
        )

IF @ANULADAS = 1
BEGIN
--ELIMINADOS
  
    
    INSERT  INTO @TBLEXTORNADOS
            ( SERIE ,
              NUMERO ,
              CODTRA ,
              SECUENCIA
            )
            SELECT  ALL_NUMSER ,  -- SERIE - char(3)
                    ALL_NUMFAC , -- NUNERO - bigint
                    ALL_CODTRA , -- CODTRA - char(4)
                    ALL_SECUENCIA  -- SECUENCIA - int
            FROM    dbo.ALLOG a
            WHERE   A.ALL_TIPMOV = 10
                    AND A.ALL_FBG IN ( 'B' )
                    AND A.ALL_FECHA_DIA = @FECHA
                    AND A.ALL_CODCIA = @CODCIA
                    AND A.ALL_FLAG_EXT = 'E'
  
--   SELECT * FROM @TBLEXTORNADOS t

   
    SET @MIN = 1
    SET @MAX = 1
   
    SELECT  @MIN = MIN(T.INDICE)
    FROM    @TBLEXTORNADOS t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLEXTORNADOS t
   
    WHILE @MIN <= @MAX
        BEGIN
            SELECT  @SERIE = T.SERIE ,
                    @NUMERO = T.NUMERO ,
                    @CODTRA = T.CODTRA ,
                    @SECUENCIA = T.SECUENCIA
            FROM    @TBLEXTORNADOS t
            WHERE   T.INDICE = @MIN
	
            INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,campo1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
                      --TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
/*
SP_RESUMEN_DIARIO '03','20211104','20211104'
*/
                    SELECT  
                    --@FECHA ,
                    (select top 1 a2.ALL_FECHA_DIA  from allog a2 
                       WHERE   A2.ALL_NUMSER = a.ALL_NUMSER
                            AND A2.ALL_NUMFAC = A.ALL_NUMFAC
                            AND A2.ALL_CODTRA = 2401 
                            AND A2.ALL_CODCIA = A.ALL_CODCIA)   ,     
                            @FECHAACTUAL ,
                            '03' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                 THEN '11111111'
                                 ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                            END ,
                            RTRIM(LTRIM(c.CLI_NOMBRE)) ,
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                          ROUND((ALL_IMPORTE_AMORT - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            A.ALL_IMPORTE_AMORT ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                           '' ,
                           '' ,
                           '' ,
                           '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                                                         
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_CODTRA = @CODTRA
                            AND A.ALL_SECUENCIA = @SECUENCIA
                            AND A.ALL_FBG IN ('B')
                            AND A.ALL_CODCIA = @CODCIA
	
            SET @MIN = @MIN + 1
        END
        DELETE FROM @TBLPRINCIPAL WHERE ESTADO = 1
END
ELSE
BEGIN
 INSERT  INTO @TBLDATA
            ( SERIE ,
              NUMERO
            )
            SELECT DISTINCT
                    ALL_NUMSER ,
                    ALL_NUMFAC
            FROM    dbo.ALLOG a
            WHERE   A.ALL_TIPMOV = 10
                    AND A.ALL_FBG = 'B'
                    AND A.ALL_FECHA_DIA = @FECHA
                    AND A.ALL_CODCIA = @CODCIA
                    AND A.ALL_FLAG_EXT = 'N'
                                
  
   
  
  
    SELECT  @MIN = MIN(T.INDICE)
    FROM    @TBLDATA t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLDATA t
  
  
  
    WHILE @MIN <= @MAX
        BEGIN
        
         SELECT TOP 1
                    @SERIE = T.SERIE ,
                    @NUMERO = T.NUMERO
            FROM    @TBLDATA t
            WHERE   T.INDICE = @MIN
            
     IF ( SELECT COUNT(A.ALL_FBG) FROM dbo.ALLOG a
         WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_FBG IN ('B')) = 1
                            BEGIN
                            
                            INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,CAMPO1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
                      --TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
                    SELECT  
                    @FECHA ,
                            @FECHAACTUAL ,
                            '03' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            
                            CASE WHEN A.ALL_CODCLIE = 99999 THEN
                            (SELECT TOP 1 CASE WHEN  LEN(RTRIM(LTRIM(f.FAR_DOCCLI))) = 0 THEN '11111111' ELSE RTRIM(LTRIM(f.FAR_DOCCLI)) END  FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG)
                            ELSE 
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                                                THEN '11111111'
                                                                ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                                                                END
                            END
                            ,

                            CASE WHEN a.all_codclie = 99999 THEN (SELECT TOP 1 RTRIM(LTRIM(f.FAR_CLIENTE)) FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG) ELSE RTRIM(LTRIM(c.CLI_NOMBRE)) END ,  
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                            --0,
                            ROUND((ALL_NETO - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            --@igv,
                            --A.ALL_BRUTO ,
                            --ROUND((A.ALL_NETO / @IGV),2),
                            --ALL_NETO - COALESCE(ALL_ICBPER,0),
                            '0.00',
                            --1000,
                            /*
SP_RESUMEN_DIARIO '01','20190724','20190724'
*/
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            --'0.00' ,
                            --A.ALL_IMPTO ,
                            --A.ALL_NETO - ROUND((A.ALL_NETO / @IGV),2),
                            --'0.00' ,
                            A.ALL_NETO ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            --AND A.ALL_SECUENCIA = 1 --PREGUNTAR SI ESTE WHERE ES CORRECTO CUANDO TIENEN DOS FORMAS DE PAGO
                            AND A.ALL_FBG ='B'
                            AND a.ALL_CODCIA = @CODCIA	
                            END
  ELSE
  BEGIN
  
  INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,CAMPO1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
--                      TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
	                /*
  SP_RESUMEN_DIARIO '03','20211104','20211104'
  */
                    SELECT TOP 1  
		                    @FECHA ,
		                   
                            @FECHAACTUAL ,
                            '03' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            CASE WHEN A.ALL_CODCLIE = 99999 THEN
                            (SELECT TOP 1 CASE WHEN  LEN(RTRIM(LTRIM(f.FAR_DOCCLI))) = 0 THEN '11111111' ELSE RTRIM(LTRIM(f.FAR_DOCCLI)) END  FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG)
                            ELSE 
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                                                THEN '11111111'
                                                                ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                                                                END
                            END,
                            --RTRIM(LTRIM(c.CLI_NOMBRE)) ,
                            CASE WHEN a.all_codclie = 99999 THEN (SELECT TOP 1 RTRIM(LTRIM(f.FAR_CLIENTE)) FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG) ELSE RTRIM(LTRIM(c.CLI_NOMBRE)) END ,  
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                            ROUND((ALL_NETO - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            --A.ALL_BRUTO ,
                            --ROUND((A.ALL_NETO / @IGV),2),
                            --ALL_NETO - COALESCE(ALL_ICBPER,0),
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            --'0.00' ,
                            --A.ALL_IMPTO ,
                            --A.ALL_NETO - ROUND((A.ALL_NETO / @IGV),2),
                            --'0.00' ,
                            A.ALL_NETO ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            --AND A.ALL_SECUENCIA = 1 --PREGUNTAR SI ESTE WHERE ES CORRECTO CUANDO TIENEN DOS FORMAS DE PAGO
                            AND A.ALL_FBG IN ('B')
                            AND A.ALL_CODCIA = @CODCIA
  END
           
  
            
            SET @MIN = @MIN + 1
        END
        
      
--INICIO NOTAS DE CREDITO
DELETE FROM @TBLDATA;
 INSERT  INTO @TBLDATA
            ( SERIE ,
              NUMERO
            )
            SELECT DISTINCT
                    ALL_NUMSER ,
                    ALL_NUMFAC
            FROM    dbo.ALLOG a
            WHERE   A.ALL_TIPMOV = 97
                    AND A.ALL_FBG = 'N'
                    AND A.ALL_FECHA_DIA = @FECHA
                    AND A.ALL_CODCIA = @CODCIA
                    AND A.ALL_FLAG_EXT = '' AND A.ALL_CODSUNAT = 3

--INICIO RECORRE
SET @MIN = NULL
SET @MAX = NULL
SELECT  @MIN = MIN(T.INDICE)
    FROM    @TBLDATA t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLDATA t
  
  
  
    WHILE @MIN <= @MAX
        BEGIN
        
         SELECT TOP 1
                    @SERIE = T.SERIE ,
                    @NUMERO = T.NUMERO
            FROM    @TBLDATA t
            WHERE   T.INDICE = @MIN
            
     IF ( SELECT COUNT(A.ALL_FBG) FROM dbo.ALLOG a
         WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_FBG IN ('N')) = 1
                            BEGIN
                            
                            INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,CAMPO1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
                      --TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
                    SELECT  
                    @FECHA ,
                            @FECHAACTUAL ,
                            '07' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            
                            CASE WHEN A.ALL_CODCLIE = 99999 THEN
                            (SELECT TOP 1 CASE WHEN  LEN(RTRIM(LTRIM(f.FAR_DOCCLI))) = 0 THEN '11111111' ELSE RTRIM(LTRIM(f.FAR_DOCCLI)) END  FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG)
                            ELSE 
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                                                THEN '11111111'
                                                                ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                                                                END
                            END
                            ,

                            CASE WHEN a.all_codclie = 99999 THEN (SELECT TOP 1 RTRIM(LTRIM(f.FAR_CLIENTE)) FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG) ELSE RTRIM(LTRIM(c.CLI_NOMBRE)) END ,  
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                            --0,
                            ROUND((ALL_NETO - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            --@igv,
                            --A.ALL_BRUTO ,
                            --ROUND((A.ALL_NETO / @IGV),2),
                            --ALL_NETO - COALESCE(ALL_ICBPER,0),
                            '0.00',
                            --1000,
                            /*
SP_RESUMEN_DIARIO '01','20190724','20190724'
*/
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            --'0.00' ,
                            --A.ALL_IMPTO ,
                            --A.ALL_NETO - ROUND((A.ALL_NETO / @IGV),2),
                            --'0.00' ,
                            A.ALL_NETO ,
                            '03' , --TIPDOCTOMODIFICA
                            (SELECT TOP 1 LEFT(COALESCE(FAR_CONCEPTO,''),4) FROM FACART WHERE FAR_FBG = 'N' AND FAR_NUMSER = A.ALL_NUMSER AND FAR_NUMFAC = A.ALL_NUMFAC AND FAR_CODCIA = A.ALL_CODCIA) , --SERIEBOLMODIFICA
                            (SELECT TOP 1 rtrim(SUBSTRING(COALESCE(FAR_CONCEPTO,''),6,LEN(LTRIM(RTRIM(COALESCE(FAR_CONCEPTO,'')))))) FROM FACART WHERE FAR_FBG = 'N' AND FAR_NUMSER = A.ALL_NUMSER AND FAR_NUMFAC = A.ALL_NUMFAC AND FAR_CODCIA = A.ALL_CODCIA) , --NROBOLMODIFICA ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            --AND A.ALL_SECUENCIA = 1 --PREGUNTAR SI ESTE WHERE ES CORRECTO CUANDO TIENEN DOS FORMAS DE PAGO
                            AND A.ALL_FBG ='N'
                            AND a.ALL_CODCIA = @CODCIA	
                            END
  ELSE
  BEGIN
  
  INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,CAMPO1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
--                      TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
	                /*
  SP_RESUMEN_DIARIO '03','20211104','20211104'
  */
                    SELECT TOP 1  
		                    @FECHA ,
		                   
                            @FECHAACTUAL ,
                            '07' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            CASE WHEN A.ALL_CODCLIE = 99999 THEN
                            (SELECT TOP 1 CASE WHEN  LEN(RTRIM(LTRIM(f.FAR_DOCCLI))) = 0 THEN '11111111' ELSE RTRIM(LTRIM(f.FAR_DOCCLI)) END  FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG)
                            ELSE 
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                                                THEN '11111111'
                                                                ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                                                                END
                            END,
                            --RTRIM(LTRIM(c.CLI_NOMBRE)) ,
                            CASE WHEN a.all_codclie = 99999 THEN (SELECT TOP 1 RTRIM(LTRIM(f.FAR_CLIENTE)) FROM dbo.FACART f WHERE f.FAR_CODCIA = @CODCIA AND F.FAR_NUMSER = A.ALL_NUMSER AND F.FAR_NUMFAC = A.ALL_NUMFAC AND F.FAR_FBG = A.ALL_FBG) ELSE RTRIM(LTRIM(c.CLI_NOMBRE)) END ,  
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                            ROUND((ALL_NETO - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            --A.ALL_BRUTO ,
                            --ROUND((A.ALL_NETO / @IGV),2),
                            --ALL_NETO - COALESCE(ALL_ICBPER,0),
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            --'0.00' ,
                            --A.ALL_IMPTO ,
                            --A.ALL_NETO - ROUND((A.ALL_NETO / @IGV),2),
                            --'0.00' ,
                            A.ALL_NETO ,
                           '03' , --TIPDOCTOMODIFICA
                            (SELECT TOP 1 LEFT(COALESCE(FAR_CONCEPTO,''),4) FROM FACART WHERE FAR_FBG = 'N' AND FAR_NUMSER = A.ALL_NUMSER AND FAR_NUMFAC = A.ALL_NUMFAC AND FAR_CODCIA = A.ALL_CODCIA) , --SERIEBOLMODIFICA
                            (SELECT TOP 1 rtrim(SUBSTRING(COALESCE(FAR_CONCEPTO,''),6,LEN(LTRIM(RTRIM(COALESCE(FAR_CONCEPTO,'')))))) FROM FACART WHERE FAR_FBG = 'N' AND FAR_NUMSER = A.ALL_NUMSER AND FAR_NUMFAC = A.ALL_NUMFAC AND FAR_CODCIA = A.ALL_CODCIA) , --NROBOLMODIFICA ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            --AND A.ALL_SECUENCIA = 1 --PREGUNTAR SI ESTE WHERE ES CORRECTO CUANDO TIENEN DOS FORMAS DE PAGO
                            AND A.ALL_FBG IN ('N')
                            AND A.ALL_CODCIA = @CODCIA
  END
           
  
            
            SET @MIN = @MIN + 1
        END
--FIN RECORRE
            
   --                SELECT * FROM @TBLDATA
   --SELECT * FROM @TBLPRINCIPAL
   --                 RETURN
                    /*
                    SP_RESUMEN_DIARIO '01','20220711','20220711'
                    */
--FIN NOTAS DE CREDITO
  
  --ELIMINADOS
  
  
    INSERT  INTO @TBLEXTORNADOS
            ( SERIE ,
              NUMERO ,
              CODTRA ,
              SECUENCIA
            )
            SELECT  ALL_NUMSER ,  -- SERIE - char(3)
                    ALL_NUMFAC , -- NUNERO - bigint
                    ALL_CODTRA , -- CODTRA - char(4)
                    ALL_SECUENCIA  -- SECUENCIA - int
            FROM    dbo.ALLOG a
            WHERE   A.ALL_TIPMOV = 10
                    AND A.ALL_FBG IN ( 'B' )
                    AND A.ALL_FECHA_DIA = @FECHA
                    AND A.ALL_CODCIA = @CODCIA
                    AND A.ALL_FLAG_EXT = 'E'
  
--   SELECT * FROM @TBLEXTORNADOS t

   
    SET @MIN = 1
    SET @MAX = 1
   
    SELECT  @MIN = MIN(T.INDICE)
    FROM    @TBLEXTORNADOS t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLEXTORNADOS t
   
    WHILE @MIN <= @MAX
        BEGIN
            SELECT  @SERIE = T.SERIE ,
                    @NUMERO = T.NUMERO ,
                    @CODTRA = T.CODTRA ,
                    @SECUENCIA = T.SECUENCIA
            FROM    @TBLEXTORNADOS t
            WHERE   T.INDICE = @MIN
	
            INSERT  INTO @TBLPRINCIPAL
                    ( FECHADOCTO ,
                      FECHACTUAL ,
                      TIPODOCTO ,
                      IDOCTO ,
                      TDI ,
                      NRODOCUSUARIO ,
                      CLIENTE ,
                      MONEDA ,campo1,
                      TOTAL ,
                      EXO ,
                      INA ,
                      GRA ,
                      ICBPER ,
                      --TISC ,
                      --TIGV ,
                      --OTROTRIB ,
                      TOTALVTA ,
                      TIPDOCTOMODIFICA ,
                      SERIEBOLMODIFICA ,
                      NROBOLMODIFICA ,
                      REGPERCEPCION ,
                      PORCPERCEPCION ,
                      BASEIMPERCEPCION ,
                      MONTOPERCEPCION ,
                      MONTOTOTINCPERCEPCION ,
                      ESTADO
	                )
/*
SP_RESUMEN_DIARIO '03','20211104','20211104'
*/
                    SELECT  
                    --@FECHA ,
                    (select top 1 a2.ALL_FECHA_DIA  from allog a2 
                       WHERE   A2.ALL_NUMSER = a.ALL_NUMSER
                            AND A2.ALL_NUMFAC = A.ALL_NUMFAC
                            AND A2.ALL_CODTRA = 2401 
                            AND A2.ALL_CODCIA = A.ALL_CODCIA)   ,     
                            @FECHAACTUAL ,
                            '03' ,
                            'B' + RIGHT('000' + RTRIM(LTRIM(A.ALL_NUMSER)), 3)
                            + '-' + CAST(A.ALL_NUMFAC AS VARCHAR(20)) AS 'IDOCTO' ,
                            '1' ,
                            CASE WHEN LEN(RTRIM(LTRIM(COALESCE(C.CLI_RUC_ESPOSA,
                                                              '')))) = 0
                                 THEN '11111111'
                                 ELSE RTRIM(LTRIM(C.CLI_RUC_ESPOSA))
                            END ,
                            RTRIM(LTRIM(c.CLI_NOMBRE)) ,
                            CASE WHEN ALL_MONEDA_CAJA = 'S' THEN 'PEN'
                                 ELSE 'USD'
                            END ,
                          ROUND((ALL_IMPORTE_AMORT - COALESCE(ALL_ICBPER,0)) / @igv,2),
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            '0.00' ,
                            COALESCE(ALL_ICBPER,0) ,--ICBPER
                            A.ALL_IMPORTE_AMORT ,
                            '' ,
                            '' ,
                            '' ,
                            '' ,
                           '' ,
                           '' ,
                           '' ,
                           '' ,
                            CASE WHEN A.ALL_CODTRA = 1111 THEN '3'
                                 ELSE '1'
                            END
                    FROM    dbo.ALLOG a
                            INNER JOIN dbo.CLIENTES c ON A.ALL_CODCLIE = C.CLI_CODCLIE
                                                         AND A.ALL_CODCIA = C.CLI_CODCIA
                                                         AND C.CLI_CP = 'C'
                                                         
                    WHERE   A.ALL_NUMSER = @SERIE
                            AND A.ALL_NUMFAC = @NUMERO
                            AND A.ALL_CODTRA = @CODTRA
                            AND A.ALL_SECUENCIA = @SECUENCIA
                            AND A.ALL_FBG IN ('B')
                            AND A.ALL_CODCIA = @CODCIA
	
            SET @MIN = @MIN + 1
        END
   
   
   
END

   
    --SELECT PRINCIPAL
 SELECT  FECHADOCTO ,
            FECHACTUAL ,
            TIPODOCTO ,
            IDOCTO ,
            TDI ,
            NRODOCUSUARIO ,
            CLIENTE ,
            MONEDA ,CAMPO1,
            CAST(TOTAL AS VARCHAR(20)) AS 'TOTAL' ,
            EXO ,
            INA ,
            GRA ,
            ICBPER ,
            --TISC ,
            --CAST(TIGV AS VARCHAR(20)) AS 'TIGV' ,
            --OTROTRIB ,
            CAST(TOTALVTA AS VARCHAR(20)) AS 'TOTALVTA' ,
            RTRIM(LTRIM(TIPDOCTOMODIFICA)) AS 'TIPDOCTOMODIFICA',
            RTRIM(LTRIM(SERIEBOLMODIFICA)) AS 'SERIEBOLMODIFICA',
            NROBOLMODIFICA ,
            REGPERCEPCION ,
            PORCPERCEPCION ,
            BASEIMPERCEPCION ,
            MONTOPERCEPCION ,
            MONTOTOTINCPERCEPCION ,
            ESTADO ,
            INDICE
    FROM    @TBLPRINCIPAL t ORDER BY IDOCTO
    
    
    --ARCHIVO TRD
    DECLARE @TBLTRD TABLE(C0 int,C1 INT, C2 VARCHAR(10),C3 VARCHAR(10),C4 MONEY, C5 MONEY,C6 CHAR(2),C7 VARCHAR(20), ESTADO CHAR(1))

    
    DECLARE @C1 int, @C2 int
    
    SELECT @C1 = MIN(INDICE) FROM @TBLPRINCIPAL
    SELECT @C2 = MAX(INDICE) FROM @TBLPRINCIPAL
    
    DECLARE @tVALOR MONEY,@TVALOR2 MONEY,@tTIPODOCTO CHAR(2), @tIDOCTO VARCHAR(20),@tESTADO CHAR(1)
    
    WHILE @C1 <= @C2
    BEGIN
    SELECT @TVALOR = t.CAMPO1,@TVALOR2 = t.ICBPER,@tTIPODOCTO = t.TIPODOCTO, @tIDOCTO = t.IDOCTO, @tESTADO = T.ESTADO FROM @TBLPRINCIPAL t WHERE t.INDICE = @C1
    INSERT INTO @TBLTRD
            ( C0, C1, C2, C3, C4, C5,C6,C7,ESTADO )
    VALUES  ( @C1, -- C0 - tinyint
              1000, -- C1 - int
              'IGV', -- C2 - varchar(10)
              'VAT', -- C3 - varchar(10)
              @tVALOR, -- C4 - money
              ROUND(@TVALOR * (@IGV - 1),2),  -- C5 - money
              @TTIPODOCTO, 
              @TIDOCTO,@tESTADO
              )
              INSERT INTO @TBLTRD
                      ( C0, C1, C2, C3, C4, C5 ,C6,C7,ESTADO )
              VALUES  ( @C1, -- C0 - tinyint
                        9997, -- C1 - int
                        'EXO', -- C2 - varchar(10)
                        'VAT', -- C3 - varchar(10)
                        0, -- C4 - money
                        0 ,  -- C5 - money
              @TTIPODOCTO, 
              @TIDOCTO,@tESTADO
                        )
                        INSERT INTO @TBLTRD
                                ( C0, C1, C2, C3, C4, C5 ,C6,C7 ,ESTADO)
                        VALUES  ( @C1, -- C0 - tinyint
                                  9998, -- C1 - int
                                  'INA', -- C2 - varchar(10)
                                  'FRE', -- C3 - varchar(10)
                                  0, -- C4 - money
                                  0,  -- C5 - money
              @TTIPODOCTO, 
              @TIDOCTO,@TESTADO
                                  )
                                  INSERT INTO @TBLTRD
                                          ( C0, C1, C2, C3, C4, C5 ,C6,C7,ESTADO )
                                  VALUES  ( @C1, -- C0 - tinyint
                                            7152, -- C1 - int
                                            'ICBPER', -- C2 - varchar(10)
                                            'OTH', -- C3 - varchar(10)
                                            0, -- C4 - money
                                            @TVALOR2  ,  -- C5 - money
              @TTIPODOCTO, 
              @TIDOCTO,@TESTADO
                                            )
		SET @C1 = @C1 + 1	
    END
    
    SELECT * FROM @TBLTRD t
   /*
SP_RESUMEN_DIARIO '01','20190725','20190725'
*/
