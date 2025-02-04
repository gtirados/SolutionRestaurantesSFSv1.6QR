USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SPPRINTDESPACHO]    Script Date: 03/11/2024 11:07:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--exec SPPRINTDESPACHO '01','20210630','100',20152
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/*
exec SPPRINTDESPACHO '01','20200602','100',91019
select ped_codart,* from pedidos where ped_numser = '100' and ped_numfac=20150
select  * from pedidos_cabecera where numser = '100' and numfac=20150
exec SPPRINTDESPACHO '01','20220308','100',346344
*/
ALTER PROC [dbo].[SPPRINTDESPACHO]
    @CODCIA CHAR(2) ,
    @FECHA DATETIME ,
    @NUMSER CHAR(3) ,
    @NUMFAC BIGINT
AS 
    SET NOCOUNT ON 
    DECLARE @IDCLIENTE BIGINT

    SELECT TOP 1
            @IDCLIENTE = PC.IDCLIENTE
    FROM    dbo.PEDIDOS_CABECERA pc
    WHERE   PC.CODCIA = @CODCIA
            AND PC.FECHA = @FECHA
            AND PC.NUMSER = @NUMSER
            AND PC.NUMFAC = @NUMFAC

    DECLARE @TBLFONOS TABLE
        (
          FONO VARCHAR(20) ,
          INDICE INT IDENTITY
        )
    INSERT  INTO @TBLFONOS
            ( FONO
            )
            SELECT  CT.FONO
            FROM    dbo.CLIENTES_TELEFONOS ct
            WHERE   CT.CODCIA = @CODCIA
                    AND CT.IDCLIENTE = @IDCLIENTE
    INSERT  INTO @TBLFONOS
            ( FONO
            )
            SELECT  C.CLI_TELEF1
            FROM    dbo.CLIENTES c
            WHERE   C.CLI_CODCIA = @CODCIA
                    AND C.CLI_CODCLIE = @IDCLIENTE


    DECLARE @MIN INT ,
        @MAX INT

    SELECT  @MIN = MIN(t.INDICE)
    FROM    @TBLFONOS t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLFONOS t
    
        DECLARE @REFERENCIA VARCHAR(100)
        DECLARE @IDDIRECCION INT 
        
        
        SELECT TOP 1 @IDDIRECCION = PC.IDDIRECCION FROM dbo.PEDIDOS_CABECERA pc WHERE CODCIA = @CODCIA AND NUMSER = @NUMSER AND NUMFAC = @NUMFAC
        
        IF @IDDIRECCION = 0
        BEGIN
        	SELECT TOP 1 @REFERENCIA = COALESCE(c.CLI_OBS,'') FROM dbo.CLIENTES c WHERE C.CLI_CODCLIE = @IDCLIENTE
        END
        ELSE
        BEGIN
        	SELECT TOP 1 @REFERENCIA = COALESCE(CD.REFERENCIA,'') FROM dbo.CLIENTES_DIRECCIONES cd WHERE CODCIA = @CODCIA AND CD.IDCLIENTE = @IDCLIENTE and IDDIRECCION = @IDDIRECCION
        END
        
--SELECT * FROM CLIENTES
--SELECT * FROM CLIENTES_DIRECCIONES


    DECLARE @STRFONO VARCHAR(300) ,
        @FONO VARCHAR(20)
    WHILE @MIN <= @MAX 
        BEGIN
	
            SELECT  @FONO = T.FONO
            FROM    @TBLFONOS t
            WHERE   T.INDICE = @MIN
            SET @STRFONO = @STRFONO + ',' + @FONO
            SET @MIN = @MIN + 1
        END
        
        
                DECLARE @ICBPER DECIMAL(8,2)
    
    SELECT TOP 1 @ICBPER = COALESCE(G.GEN_ICBPER,0) FROM dbo.GENERAL g
    
             DECLARE @TBLICBPER TABLE
                (
                  CODART BIGINT ,
                  ICBPER MONEY
                )
    
    
    INSERT  INTO @TBLICBPER
                    ( CODART ,
                      ICBPER
                    )
                    SELECT  p.PA_CODPA ,
                           SUM( p.PA_PROM*@ICBPER) AS 'Importe'
                    FROM    dbo.PAQUETES p
                            INNER JOIN dbo.ARTI art ON P.PA_CODCIA = art.ART_CODCIA
                                                       AND P.PA_CODART = art.art_key
                    WHERE   PA_CODPA IN (
                            SELECT  A2.ART_KEY
                            FROM    pedidos p
                                    INNER JOIN dbo.ARTI a2 ON p.PED_CODCIA = a2.ART_CODCIA
                                                              AND p.PED_CODART = a2.ART_KEY
                                    INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
                                                          AND A.ALL_CODCIA = @CODCIA
                                                          AND A.ALL_FECHA_DIA = @FECHA
                                                          AND A.ALL_FLAG_EXT = 'N'
                            WHERE  
                            -- ped_codclie = @CodMesa
--                                    AND 
                                    ped_codcia = @CodCia
                                    AND ped_fecha = @Fecha
                                    AND ped_estado = 'N'
                                    AND ped_situacion <> 'A'
                                    AND ped_fac <> ped_Cantidad
                                    AND ped_numfac = @NUMFAC
                                    AND a2.ART_FLAG_STOCK = 'C' )
                            AND art.ART_CALIDAD = 0 GROUP BY p.PA_CODPA
    
        
/*
exec SPPRINTDESPACHO '01','20210630','100',20152
*/
declare @total money
declare @cant int


select @total = sum(PED_SUBTOTAL) from PEDIDOS where PED_NUMFAC=@NUMFAC AND PED_ESTADO ='N'

select @ICBPER = icbper from @TBLICBPER
select @cant = sum(ped_cantidad) from PEDIDOS where PED_NUMFAC=  @NUMFAC  and PED_CODART in (select codart from @TBLICBPER)

set @total = @total --+ (@ICBPER * @cant)
--select *,@ICBPER,@total,@cant from @TBLICBPER

    SELECT  A2.ALL_FBG + '/' + RTRIM(LTRIM(a2.ALL_NUMSER)) + '-'
            + RTRIM(LTRIM(CAST(a2.ALL_NUMFAC AS VARCHAR(20)))) AS 'NRODOCTO' ,
            p.PED_NUMSER + '-' + RTRIM(LTRIM(CAST(P.PED_NUMFAC AS VARCHAR(20)))) AS 'NROCOMANDA' ,
            c.CLI_NOMBRE AS 'CLIENTE' ,
            PC.DIRECCION AS 'DIRECCION' ,
            ISNULL(@FONO, '') AS 'FONO' ,
            ISNULL(ZR.DENOMINACION, '') AS 'ZONA' ,
            R.REPARTIDOR AS 'REPARTIDOR' ,
            RTRIM(LTRIM(a.ART_NOMBRE)) AS 'PRODUCTO' ,
            P.PED_CANTIDAD AS 'CANTIDAD' ,
            --( SELECT    SUM(PE.ped_SUBTOTAL)
            --  FROM      dbo.pedidos pe
            --  WHERE     PE.ped_CODCIA = PC.CODCIA
            --            AND PE.ped_NUMSER = PC.NUMSER
            --            AND PE.ped_NUMFAC = PC.NUMFAC
            --            AND PE.ped_FECHA = PC.FECHA
            --            AND PE.ped_ESTADO = 'N'
            --            --aqui
            --            --exec SPPRINTDESPACHO '01','20210630','100',20152
            --) + 
            -- COALESCE(t.ICBPER*ped_cantidad, 0) 
            @total AS 'TOTAL' ,
            PC.PAGO ,
            --PC.PAGO - ( SELECT  SUM(PE.PED_SUBTOTAL)
            --            FROM    dbo.PEDIDOS pe
            --            WHERE   PE.PED_CODCIA = PC.CODCIA
            --                    AND PE.PED_NUMSER = PC.NUMSER
            --                    AND PE.PED_NUMFAC = PC.NUMFAC
            --                    AND PE.PED_FECHA = PC.FECHA
            --                    AND PE.PED_ESTADO = 'N'
            --          ) AS 'VUELTO' ,
            ISNULL(pc.vuelto,0) AS 'VUELTO',
            ISNULL(PC.OBS, '') AS 'OBS' ,
            p.PED_FECHA AS 'FECHA' ,
            pc.FECHAREG AS 'HORA' ,
            ISNULL(P.PED_OFERTA, '') AS 'DETALLE' ,
  DBO.FnDevuelveCaracteristica(@CODCIA, @FECHA, P.PED_NUMFAC,
                                         P.PED_NUMSER, P.PED_NUMSEC,
                                         P.PED_CODART) AS 'CARACTERISTICA' ,
            ISNULL(PC.MONTO_ENVIO, 0) AS 'ENVIO',
            ISNULL(u.NOMBRE,'') AS 'URB',
            @REFERENCIA AS 'REF'
            --, COALESCE(t.ICBPER*ped_cantidad, 0) AS 'COMBO_ICBPER'
    FROM    dbo.PEDIDOS p
            LEFT JOIN dbo.PEDIDOS_CABECERA pc ON p.PED_CODCIA = pc.codcia
                                                 AND p.PED_NUMFAC = pc.NUMFAC
                                                 AND p.PED_NUMSER = pc.NUMSER
                                                 AND p.PED_FECHA = pc.FECHA
                                                 AND p.PED_ESTADO = 'N'
            LEFT JOIN dbo.CLIENTES c ON pc.IDCLIENTE = c.CLI_CODCLIE
                                        AND pc.CODCIA = c.CLI_CODCIA
            LEFT JOIN dbo.ZONAS_REPARTO zr ON ZR.CODCIA = C.CLI_CODCIA
                                              AND ZR.IDREPARTO = C.CLI_ZONADELIVERY
            LEFT JOIN dbo.REPARTIDORES r ON PC.CODCIA = R.CODCIA
                                            AND PC.IDREPARTIDOR = R.IDREPARTIDOR
            LEFT JOIN dbo.ALLOG a2 ON a2.ALL_CODCIA = p.PED_CODCIA
                                      AND a2.ALL_NUMFAC = @NUMFAC
                                      AND A2.ALL_NUMSER = @NUMSER
            LEFT JOIN dbo.ARTI a ON p.PED_CODCIA = a.ART_CODCIA
                                    AND p.PED_CODART = a.ART_KEY
                                    LEFT JOIN dbo.URBANIZACION u ON c.CLI_CODCIA = u.CODCIA AND c.CLI_URB = u.IDURBANIZACION
                                    LEFT JOIN @TBLICBPER t ON P.PED_CODART = t.CODART
    WHERE   p.PED_NUMSER = @NUMSER
            AND P.PED_NUMFAC = @NUMFAC
            AND P.PED_CODCIA = @CODCIA
            AND P.PED_FECHA = @FECHA
            AND p.PED_ESTADO = 'N'

----SELECT ALL_NUMFAC, * FROM dbo.ALLOG a WHERE a.ALL_FECHA_DIA='20140128' AND ALL_CODCLIE=52 AND ALL_NUMFAC_C=79972

