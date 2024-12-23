USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SPITEMSDESPACHADOS]    Script Date: 07/19/2022 16:58:39 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
exec SPITEMSDESPACHADOS '01','20130618','20130618',-1
*/
--SELECT * FROM dbo.PEDIDOS p WHERE p.PED_NUMFAC=14 AND PED_ESTADO='N'
ALTER PROC [dbo].[SPITEMSDESPACHADOS]
    @CODCIA CHAR(2) ,
    @FECHA1 DATETIME ,
    @FECHA2 DATETIME ,
    @IDFAMILIA INT = -1
AS 
    SET NOCOUNT ON 

--SELECT PED_HORA,PED_FECHAREG, * FROM dbo.PEDIDOS p

    SELECT  P.PED_FECHA AS 'FECHA' ,
            RTRIM(LTRIM(P.PED_NUMSER)) + '-'
            + CAST(P.PED_NUMFAC AS VARCHAR(20)) AS 'COMANDA' ,
    RTRIM(LTRIM(A.ART_NOMBRE)) AS 'PRODUCTO' ,
            P.PED_CANTIDAD AS 'CANTIDAD' ,
            P.PED_NOMCLIE AS 'MESA' ,
            RTRIM(LTRIM(V.VEM_NOMBRE)) AS 'MOZO' ,
            ISNULL(P.HORADESPACHO,'') AS 'HORA' ,
            P.PED_OFERTA AS 'DETALLE' ,
            P.PED_NUMFAC AS 'NUMERO' ,
            ISNULL(CONVERT(VARCHAR(10), GETDATE() - P.FECHADESPACHO, 108),'') AS 'TIEMPO' ,
            P.PED_NUMSEC AS 'SEC' ,
            PED_NUMSER AS 'SERIE' ,
            P.PED_CODART AS 'CODART' ,
            CONVERT(VARCHAR(10),p.FECHADESPACHO - p.PED_FECHAREg ,108) AS 'TIEMPO2'
            /*
exec SPITEMSDESPACHADOS '01','20130618','20130618',-1
*/
            --DBO.FnDevuelveColor(@CODCIA, P.PED_NUMSER, P.PED_NUMFAC,
            --                    P.PED_FECHA, P.PED_NUMSEC, P.PED_CODART,
            --                    CONVERT(VARCHAR(10), GETDATE()
            --                    - P.PED_FECHAREG, 108), @IDFAMILIA) AS 'COLOR'
    FROM    dbo.PEDIDOS p
            INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                     AND P.PED_CODART = A.ART_KEY
                                     AND P.PED_CANATEN = P.PED_CANTIDAD
                                     AND p.PED_ESTADO = 'N'
            INNER JOIN dbo.VEMAEST v ON P.PED_CODVEN = V.VEM_CODVEN
                                        AND V.VEM_CODCIA = P.PED_CODCIA
    WHERE   P.PED_CODCIA = @CODCIA
            AND P.PED_FECHA BETWEEN @FECHA1 AND @FECHA2
            --AND P.PED_CANATEN = 0
            AND P.PED_ESTADO = 'N'
            AND (A.ART_FAMILIA = @IDFAMILIA
            OR @IDFAMILIA = -1)
    ORDER BY P.HORADESPACHO
