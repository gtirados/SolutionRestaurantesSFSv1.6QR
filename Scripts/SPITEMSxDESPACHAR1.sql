/*
exec SPITEMSxDESPACHAR1 '01','20221022',1
exec SPITEMSxDESPACHAR1 '01','20221022',2
exec SPITEMSxDESPACHAR1 '01','20221022',3
*/

ALTER PROC [dbo].[SPITEMSxDESPACHAR1]
    @CODCIA CHAR(2) ,
    @FECHA DATETIME ,
    @TIPO TINYINT --1: COCINA, 2: BAR, 3:TODOS
AS
    SET NOCOUNT ON 
    DECLARE @TBLPEDIDOS TABLE
        (
          REGISTRO VARCHAR(20) ,
          NUMERO INT ,
          SERIE CHAR(3) ,
          TIEMPO VARCHAR(20) ,
          MOZO VARCHAR(30) ,
          MESA VARCHAR(50) ,
          ENTRADAS INT ,
          SEGUNDOS INT ,
          AHORA INT
        )
        
        DECLARE @TBLFINAL TABLE(CANTIDAD NUMERIC(18,2),PRODUCTO VARCHAR(60),PRODUCTOTOOL VARCHAR(60),DETALLE VARCHAR(300),SERIE CHAR(3),NUMERO BIGINT,FAMILIA INT,CODIGO BIGINT,SEC INT,MARCA BIT,ADICIONAL VARCHAR(4000),CC BIT,NOMFAMILIA VARCHAR(40))
        DECLARE @TBLCOMBOS TABLE(INDICE INT IDENTITY,CANTIDAD NUMERIC(18,2),PRODUCTO VARCHAR(60),PRODUCTOTOOL VARCHAR(60),DETALLE VARCHAR(300),SERIE CHAR(3),NUMERO BIGINT,FAMILIA INT,CODIGO BIGINT,SEC INT,MARCA BIT,ADICIONAL VARCHAR(4000))
                        DECLARE @MIN INT, @MAX INT
                        DECLARE @CANTIDAD NUMERIC(18,2),@PRODUCTO VARCHAR(60),@PRODUCTOTOOL VARCHAR(60),@DETALLE VARCHAR(300),@SERIE CHAR(3),@NUMERO BIGINT,@FAMILIA INT,@CODIGO BIGINT,@SEC INT,@MARCA BIT,@ADICIONAL VARCHAR(4000)
                        
                        
    IF @TIPO = 1 --COCINA
        BEGIN
            INSERT  INTO @TBLPEDIDOS
                    ( REGISTRO ,
                      NUMERO ,
                      SERIE ,
                      TIEMPO ,
                      MOZO ,
                      MESA ,
                      ENTRADAS ,
                      SEGUNDOS ,
                      ahora
                    )
                    SELECT  DBO.FnDevuelveHora(pc.FECHAREG) ,
                            PC.NUMFAC ,
                            PC.NUMSER ,
                            CONVERT(VARCHAR(10), GETDATE() - PC.FECHAREG, 108) ,
                            ( SELECT    RTRIM(LTRIM(V.VEM_NOMBRE))
                              FROM      dbo.VEMAEST v
                              WHERE     V.VEM_CODCIA = @CODCIA
                                        AND V.VEM_CODVEN = ( SELECT TOP 1
                                                              P.PED_CODVEN
                                                             FROM
                                                              dbo.PEDIDOS p
                                                             WHERE
                                                              P.PED_CODCIA = @CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMFAC = PC.NUMFAC
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                           )
                            ) ,
                            RTRIM(LTRIM(ISNULL(M.MES_DESCRIP, ''))) ,
                           dbo.FnDevuelve_Cantidad_items(@CODCIA,pc.FECHA,pc.NUMFAC,pc.NUMSER,1),--COCINA
                            0 ,
                            ( SELECT    COUNT(p.PED_CODART)
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND P.PED_CANATEN <> P.PED_CANTIDAD
                                                              AND P.PED_ESTADO = 'N'
                                                              AND ( p.PED_ENVIAR_EN IS NULL
                                                              OR GETDATE() >= P.PED_ENVIAR_EN
                                                              )
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_FECHA = PC.FECHA
                                        AND P.PED_NUMSER = PC.NUMSER
                                        AND P.PED_NUMFAC = PC.NUMFAC
                            ) --AHORA
                    FROM    dbo.PEDIDOS_CABECERA pc
                            INNER JOIN dbo.MESAS m ON PC.CODCIA = M.MES_CODCIA
                                                      AND PC.CODMESA = M.MES_CODMES
                    WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
              --  AND pc.FACTURADO = 0
                    ORDER BY PC.NUMFAC
  update @TBLPEDIDOS set mozo = 'DELIVERY' WHERE MOZO IS NULL      
            SELECT  *
            FROM    @TBLPEDIDOS t
            WHERE   T.ahora <> 0
                    AND ( T.ENTRADAS <> 0
                          OR T.SEGUNDOS <> 0
                        )
    
    /*
exec SPITEMSxDESPACHAR1 '01','20141009',3
*/
INSERT INTO @TBLFINAL
            SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS NUMERIC(18,4)),-- AS 'CANTIDAD' ,
                    LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39),-- AS 'PRODUCTO' ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,--AS 'PRODUCTOtool' ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22),-- AS 'DETALLE' ,
                    p.PED_NUMSER ,--AS 'SERIE' ,
                    P.PED_NUMFAC ,--AS 'NUMERO' ,
                    P.PED_FAMILIA2 ,--AS 'FAMILIA' ,
                    A.ART_KEY ,--AS 'CODIGO' ,
                    P.PED_NUMSEC ,--AS 'SEC' ,
                    ISNULL(P.PED_MARCADOS, 0) ,--AS 'MARCA' ,
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart) ,--AS 'ADICIONAL'
                                                 1
                                                 ,RTRIM(LTRIM(f.TAB_NOMLARGO))
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N'
                                                INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                    AND P.PED_FAMILIA = 1 --COCINA
                    AND A.ART_FLAG_STOCK <> 'C'
                
                  --NUEVO
                        INSERT INTO @TBLCOMBOS
                        SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS NUMERIC(18,2)) ,
                    LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    P.PED_FAMILIA2 ,
                    A.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.ART_FLAG_STOCK='C'
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                
                
                /*
exec SPITEMSxDESPACHAR1 '01','20141224',1
exec SPITEMSxDESPACHAR1 '01','20141224',2
exec SPITEMSxDESPACHAR1 '01','20141224',3
*/


 SELECT @MIN = MIN(INDICE) FROM @TBLCOMBOS
                        SELECT @MAX = MAX(INDICE) FROM @TBLCOMBOS
                        --SELECT * FROM @TBLCOMBOS
                        
                        WHILE @MIN <= @MAX
                        BEGIN
							SELECT @CANTIDAD =CANTIDAD,@PRODUCTO = PRODUCTO,@PRODUCTOTOOL = PRODUCTOTOOL,@DETALLE =DETALLE,@SERIE =SERIE,@NUMERO  = NUMERO,@FAMILIA = @FAMILIA
							,@CODIGO  = CODIGO,@SEC = SEC,@MARCA = @MARCA,@ADICIONAL = @ADICIONAL FROM @TBLCOMBOS WHERE INDICE = @MIN
							
							--SELECT * FROM PAQUETES WHERE PA_CODPA=@CODIGO
							INSERT INTO @TBLFINAL
							  SELECT  CAST(A.PA_PROM * @CANTIDAD AS NUMERIC(18,2)) ,
                    LEFT(RTRIM(LTRIM(AA.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(AA.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    AA.ART_FAMILIA ,
                    AA.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
                                                 ,0
                                                 ,RTRIM(LTRIM(f.TAB_NOMLARGO))
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.PAQUETES a ON P.PED_CODCIA = A.PA_CODCIA
                                             AND P.PED_CODART = A.PA_CODPA
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.PA_CODPA = @CODIGO
                                             INNER JOIN ARTI aa ON A.PA_CODART = aa.ART_KEY AND A.PA_CODCIA = aa.ART_CODCIA
                                              INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
            AND aa.ART_FAMILIA=1
            
            
							SET @MIN = @MIN + 1
                        END
                        
 INSERT INTO @TBLFINAL
 SELECT CANTIDAD,' ' +PRODUCTO,PRODUCTOTOOL,DETALLE,SERIE,NUMERO,1,CODIGO,SEC,MARCA,ADICIONAL,1,'' FROM @TBLCOMBOS
    SELECT * FROM @TBLFINAL ORDER BY FAMILIA,SEC,PRODUCTO

        END 
       
    IF @TIPO = 2 --BAR
        BEGIN
            INSERT  INTO @TBLPEDIDOS
                    ( REGISTRO ,
                      NUMERO ,
                      SERIE ,
                      TIEMPO ,
                      MOZO ,
                      MESA ,
                      ENTRADAS ,
                      SEGUNDOS ,
                      ahora
                    )
                    SELECT  DBO.FnDevuelveHora(pc.FECHAREG) ,
                            PC.NUMFAC ,
                            PC.NUMSER ,
                            CONVERT(VARCHAR(10), GETDATE() - PC.FECHAREG, 108) ,
                      ( SELECT    RTRIM(LTRIM(V.VEM_NOMBRE))
                              FROM      dbo.VEMAEST v
                              WHERE     V.VEM_CODCIA = @CODCIA
                                        AND V.VEM_CODVEN = ( SELECT TOP 1
                                                              P.PED_CODVEN
                                                             FROM
                                                              dbo.PEDIDOS p
                                                             WHERE
                                                              P.PED_CODCIA = @CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                    AND P.PED_NUMFAC = PC.NUMFAC
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                           )
                            ) ,
                            RTRIM(LTRIM(ISNULL(M.MES_DESCRIP, ''))) ,
                            0 ,
                             dbo.FnDevuelve_Cantidad_items(@CODCIA,pc.FECHA,pc.NUMFAC,pc.NUMSER,2) ,--BAR
                            ( SELECT    COUNT(p.PED_CODART)
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND P.PED_CANATEN <> P.PED_CANTIDAD
                                                              AND P.PED_ESTADO = 'N'
                                                              AND ( p.PED_ENVIAR_EN IS NULL
                                                              OR GETDATE() >= P.PED_ENVIAR_EN
                                                              )
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_FECHA = PC.FECHA
                                        AND P.PED_NUMSER = PC.NUMSER
                                        AND P.PED_NUMFAC = PC.NUMFAC
                            ) --AHORA
                    FROM    dbo.PEDIDOS_CABECERA pc
                            INNER JOIN dbo.MESAS m ON PC.CODCIA = M.MES_CODCIA
                                                      AND PC.CODMESA = M.MES_CODMES
                    WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
              --  AND pc.FACTURADO = 0
                    ORDER BY PC.NUMFAC
  update @TBLPEDIDOS set mozo = 'DELIVERY' WHERE MOZO IS NULL      
            SELECT  *
            FROM    @TBLPEDIDOS t
            WHERE   T.ahora <> 0
                    AND ( T.ENTRADAS <> 0
                          OR T.SEGUNDOS <> 0
                        )

INSERT INTO @TBLFINAL
            SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS numeric(18,4)),-- AS 'CANTIDAD' ,
                    LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39) ,--AS 'PRODUCTO' ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,--AS 'PRODUCTOtool' ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,--AS 'DETALLE' ,
                    p.PED_NUMSER ,--AS 'SERIE' ,
                    P.PED_NUMFAC ,--AS 'NUMERO' ,
                    P.PED_FAMILIA2 ,--AS 'FAMILIA' ,
                    A.ART_KEY ,--AS 'CODIGO' ,
                    P.PED_NUMSEC ,--AS 'SEC' ,
                    ISNULL(P.PED_MARCADOS, 0) ,--AS 'MARCA' ,
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart) ,--AS 'ADICIONAL'
                                                 1,
                                                 RTRIM(LTRIM(f.TAB_NOMLARGO))
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
         AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N'
                                              INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                    AND P.PED_FAMILIA = 2 --BAR 
                    AND A.ART_FLAG_STOCK<> 'C'
                    
                       --NUEVO
                        INSERT INTO @TBLCOMBOS
                        SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS NUMERIC(18,2)) ,
                    ' ' + LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    P.PED_FAMILIA2 ,
                    A.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.ART_FLAG_STOCK='C'
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                
                                    /*
exec SPITEMSxDESPACHAR1 '01','20141224',2
*/
                    SELECT @MIN = MIN(INDICE) FROM @TBLCOMBOS
                        SELECT @MAX = MAX(INDICE) FROM @TBLCOMBOS
                        
                        WHILE @MIN <= @MAX
                        BEGIN
							SELECT @CANTIDAD =CANTIDAD,@PRODUCTO = PRODUCTO,@PRODUCTOTOOL = PRODUCTOTOOL,@DETALLE =DETALLE,@SERIE =SERIE,@NUMERO  = NUMERO,@FAMILIA = @FAMILIA
							,@CODIGO  = CODIGO,@SEC = SEC,@MARCA = @MARCA,@ADICIONAL = @ADICIONAL FROM @TBLCOMBOS WHERE INDICE = @MIN
							
							--SELECT * FROM PAQUETES WHERE PA_CODPA=@CODIGO
							INSERT INTO @TBLFINAL
							  SELECT  CAST(A.PA_PROM * @CANTIDAD AS NUMERIC(18,2)) ,
                    LEFT(RTRIM(LTRIM(AA.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(AA.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    AA.ART_FAMILIA ,
                    AA.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
                              ,0,
                              RTRIM(LTRIM(f.TAB_NOMLARGO))
      FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.PAQUETES a ON P.PED_CODCIA = A.PA_CODCIA
                                             AND P.PED_CODART = A.PA_CODPA
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.PA_CODPA = @CODIGO
                                             INNER JOIN ARTI aa ON A.PA_CODART = aa.ART_KEY AND A.PA_CODCIA = aa.ART_CODCIA
                                              INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
            AND   aa.ART_FAMILIA=2
            
            
            
							SET @MIN = @MIN + 1
                        END
                        
 
 INSERT INTO @TBLFINAL
 SELECT CANTIDAD,' ' +PRODUCTO,PRODUCTOTOOL,DETALLE,SERIE,NUMERO,2,CODIGO,SEC,MARCA,ADICIONAL,1,'' FROM @TBLCOMBOS
 
    SELECT * FROM @TBLFINAL ORDER BY FAMILIA,SEC,PRODUCTO
    
    
        END
       
    IF @TIPO = 3 --TODOS
        BEGIN
            INSERT  INTO @TBLPEDIDOS
                    ( REGISTRO ,
                      NUMERO ,
                      SERIE ,
                      TIEMPO ,
                      MOZO ,
                      MESA ,
                      ENTRADAS ,
                      SEGUNDOS ,
                      ahora
                    )
                    SELECT  DBO.FnDevuelveHora(pc.FECHAREG) ,
                            PC.NUMFAC ,
                            PC.NUMSER ,
                            CONVERT(VARCHAR(10), GETDATE() - PC.FECHAREG, 108) ,
                            ( SELECT    RTRIM(LTRIM(V.VEM_NOMBRE))
                              FROM      dbo.VEMAEST v
                              WHERE     V.VEM_CODCIA = @CODCIA
                                        AND V.VEM_CODVEN = ( SELECT TOP 1
                                                              P.PED_CODVEN
                                                             FROM
                                                              dbo.PEDIDOS p
                                                             WHERE
                                                              P.PED_CODCIA = @CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMFAC = PC.NUMFAC
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                              --and p.PED_FAMILIA in (1,2)
                                                           )
                            ) ,
                            RTRIM(LTRIM(ISNULL(M.MES_DESCRIP, ''))) ,
                              dbo.FnDevuelve_Cantidad_items(@CODCIA,pc.FECHA,pc.NUMFAC,pc.NUMSER,1),--COCINA
                            --dbo.FnDevuelve_Cantidad_items(@CODCIA,pc.FECHA,pc.NUMFAC,pc.NUMSER,2),--BAR
                            0,
/*
exec SPITEMSxDESPACHAR1 '01','20221022',3
*/

                          ( SELECT    COUNT(p.PED_CODART)
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND P.PED_CANATEN <> P.PED_CANTIDAD
                                                              AND P.PED_ESTADO = 'N'
                                                              AND ( p.PED_ENVIAR_EN IS NULL
        OR GETDATE() >= P.PED_ENVIAR_EN
                                                     )
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_FECHA = PC.FECHA
                                        AND P.PED_NUMSER = PC.NUMSER
                                        AND P.PED_NUMFAC = PC.NUMFAC 
                                        --and p.PED_FAMILIA in (1,2) --NUEVO
                            ) --AHORA
                    FROM    dbo.PEDIDOS_CABECERA pc
                            INNER JOIN dbo.MESAS m ON PC.CODCIA = M.MES_CODCIA
                                                      AND PC.CODMESA = M.MES_CODMES
                    WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
                    ORDER BY PC.NUMFAC

                    
          update @TBLPEDIDOS set mozo = 'DELIVERY' WHERE MOZO IS NULL      
          
            SELECT  *
            FROM    @TBLPEDIDOS t
            WHERE   T.ahora <> 0
                    AND ( T.ENTRADAS <> 0
                          OR T.SEGUNDOS <> 0
                        )
    



INSERT INTO @TBLFINAL
            SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS NUMERIC(18,2)),-- AS 'CANTIDAD' ,
                    LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39),-- AS 'PRODUCTO' ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,--AS 'PRODUCTOtool' ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22),-- AS 'DETALLE' ,
                    p.PED_NUMSER ,--AS 'SERIE' ,
                    P.PED_NUMFAC ,--AS 'NUMERO' ,
                    P.PED_FAMILIA2 ,--AS 'FAMILIA' ,
                    A.ART_KEY ,--AS 'CODIGO' ,
                    P.PED_NUMSEC ,--AS 'SEC' ,
                    ISNULL(P.PED_MARCADOS, 0) ,--AS 'MARCA' ,
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart) --AS 'ADICIONAL'
                                                 ,1
                                                 ,RTRIM(LTRIM(f.TAB_NOMLARGO))
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.ART_FLAG_STOCK <> 'C'
                                             INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                        --and p.PED_FAMILIA in(1,2)--NUEVO
                        
                        --NUEVO
                        
                        
                        INSERT INTO @TBLCOMBOS
                        SELECT  CAST(P.PED_CANTIDAD - P.PED_CANATEN AS NUMERIC(18,2)) ,
                    LEFT(RTRIM(LTRIM(A.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(A.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    P.PED_FAMILIA2 ,
                    A.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                 AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                             AND P.PED_CODART = A.ART_KEY
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.ART_FLAG_STOCK='C'
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
          --  AND pc.FACTURADO = 0
                    AND p.PED_ESTADO = 'N'
                    AND ( p.PED_ENVIAR_EN IS NULL
                          OR GETDATE() >= PED_ENVIAR_EN
                        )
                        
                        SELECT @MIN = MIN(INDICE) FROM @TBLCOMBOS
                        SELECT @MAX = MAX(INDICE) FROM @TBLCOMBOS
                        --SELECT * FROM @TBLCOMBOS
                        
                        WHILE @MIN <= @MAX
                        BEGIN
							SELECT @CANTIDAD =CANTIDAD,@PRODUCTO = PRODUCTO,@PRODUCTOTOOL = PRODUCTOTOOL,@DETALLE =DETALLE,@SERIE =SERIE,@NUMERO  = NUMERO,@FAMILIA = @FAMILIA
							,@CODIGO  = CODIGO,@SEC = SEC,@MARCA = @MARCA,@ADICIONAL = @ADICIONAL FROM @TBLCOMBOS WHERE INDICE = @MIN
							
							--SELECT * FROM PAQUETES WHERE PA_CODPA=@CODIGO
							INSERT INTO @TBLFINAL
							  SELECT  CAST(A.PA_PROM * @CANTIDAD AS NUMERIC(18,2)) ,
                    LEFT(RTRIM(LTRIM(AA.ART_NOMBRE)), 39) ,
                    RTRIM(LTRIM(AA.ART_NOMBRE)) ,
                    LEFT(ISNULL(P.PED_OFERTA, ''), 22) ,
                    p.PED_NUMSER ,
                    P.PED_NUMFAC ,
                    AA.ART_FAMILIA ,
                    AA.ART_KEY ,
                    P.PED_NUMSEC ,
                    ISNULL(P.PED_MARCADOS, 0),
                    dbo.FnDevuelveCaracteristica(pc.CODCIA, pc.FECHA,
                                                 p.ped_numfac, p.ped_numser,
                                                 p.ped_numsec, p.ped_codart)
                                                 ,0
                                                 ,RTRIM(LTRIM(f.TAB_NOMLARGO))
            FROM    dbo.PEDIDOS_CABECERA pc
                    INNER JOIN dbo.PEDIDOS p ON pc.CODCIA = p.PED_CODCIA
                                                AND pc.FECHA = p.PED_FECHA
                                                AND pc.NUMFAC = p.PED_NUMFAC
                                                AND pc.NUMSER = p.PED_NUMSER
                    INNER JOIN dbo.PAQUETES a ON P.PED_CODCIA = A.PA_CODCIA
                                             AND P.PED_CODART = A.PA_CODPA
                                             AND P.PED_CANATEN <> P.PED_CANTIDAD
                                             AND p.PED_ESTADO = 'N' AND A.PA_CODPA = @CODIGO
                                             INNER JOIN ARTI aa ON A.PA_CODART = aa.ART_KEY AND A.PA_CODCIA = aa.ART_CODCIA
                                              INNER JOIN TABLAS f ON f.TAB_CODCIA = p.PED_CODCIA AND f.TAB_TIPREG=122 AND f.TAB_NUMTAB = p.PED_FAMILIA2
            WHERE   CONVERT(VARCHAR(8), pc.FECHA, 112) = @FECHA
            --and p.PED_FAMILIA in(1,2)--NUEVO
            
            
							SET @MIN = @MIN + 1
                        END
                        
                        
         
  if @TIPO = 3 --todos
    begin
    update @TBLFINAL set familia = 1
    end
    
    
    
     INSERT INTO @TBLFINAL
 SELECT CANTIDAD,' ' +PRODUCTO,PRODUCTOTOOL,DETALLE,SERIE,NUMERO,2,CODIGO,SEC,MARCA,ADICIONAL,1,'' FROM @TBLCOMBOS
 
 INSERT INTO @TBLFINAL
 SELECT CANTIDAD,' ' +PRODUCTO,PRODUCTOTOOL,DETALLE,SERIE,NUMERO,1,CODIGO,SEC,MARCA,ADICIONAL,1,'' FROM @TBLCOMBOS
 
    SELECT * FROM @TBLFINAL ORDER BY FAMILIA,SEC,PRODUCTO
    
   
 
/*
exec SPITEMSxDESPACHAR1 '01','20141224',3
*/

        END
