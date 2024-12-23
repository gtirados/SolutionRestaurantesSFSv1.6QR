USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_CONSULTA_EXTORNADOS]    Script Date: 05/21/2022 11:54:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
exec SP_CONSULTA_EXTORNADOS '01','20140823','20140923','-1'
*/
ALTER PROC [dbo].[SP_CONSULTA_EXTORNADOS]
    @CODCIA CHAR(2) ,
    @DESDE DATE ,
    @HASTA DATE ,
    @CODUSUARIO VARCHAR(10) = '-1' ,
    @CODMESA VARCHAR(10) = '-1'
AS
    SET NOCOUNT ON 

    SELECT  P.PED_FECHA AS 'FECHA' ,
            RTRIM(LTRIM(P.PED_NUMSER)) + '-'
            + CAST(P.PED_NUMFAC AS VARCHAR(20)) AS 'COMANDA' ,
            RTRIM(LTRIM(A.ART_NOMBRE)) AS 'PRODUCTO' ,
            P.PED_CANTIDAD AS 'CANTIDAD' ,
            p.PED_PRECIO as 'PRECIO',
            RTRIM(LTRIM(COALESCE(M.MES_DESCRIP,'') )) AS 'MESA' ,
            RTRIM(LTRIM(V.VEM_NOMBRE)) AS 'MOZO' ,
            P.USUARIO_ELIMINA AS 'USUARIO',
            RTRIM(LTRIM(MA.DESCRIPCION)) AS 'MOTIVO'
    FROM    dbo.PEDIDOS p
            INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                     AND P.PED_CODART = A.ART_KEY
            LEFT JOIN dbo.MESAS m ON P.PED_CODCIA = M.MES_CODCIA
                                     AND P.PED_CODCLIE = M.MES_CODMES
            INNER JOIN dbo.VEMAEST v ON P.PED_CODCIA = V.VEM_CODCIA
                                        AND P.PED_CODVEN = V.VEM_CODVEN
                                        INNER JOIN dbo.MOTIVO_ANULACION ma ON P.IDMOTIVO_ELIMINA= MA.IDMOTIVO AND P.PED_CODCIA = MA.CODCIA
    WHERE   p.PED_CODCIA = @CODCIA
            AND P.PED_FECHA BETWEEN @DESDE AND @HASTA
            AND ISNULL(P.PED_CODCLIE, '-1') = CASE WHEN @CODMESA = '-1'
                                                   THEN ISNULL(P.PED_CODCLIE,
                                                              '-1')
                                                   ELSE @CODMESA
                                              END
            AND ISNULL(P.PED_CODUSU, '-1') = CASE WHEN @CODUSUARIO = '-1'
                                                  THEN ISNULL(P.PED_CODUSU,
                                                              '-1')
                                                  ELSE @CODUSUARIO
                                             END
            AND PED_ESTADO = 'E'                                  



