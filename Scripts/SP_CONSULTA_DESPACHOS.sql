/*
exec SP_CONSULTA_DESPACHOS '01','20221022','-1'
*/

ALTER PROC [dbo].[SP_CONSULTA_DESPACHOS]
    (
      @CODCIA CHAR(2) ,
      @FECHA DATETIME ,
      @MESA VARCHAR(10) = '-1'
    )
AS 
    SET NOCOUNT ON 

    SELECT  PC.NUMFAC AS 'COMANDA' ,
            RTRIM(LTRIM(M.MES_DESCRIP)) AS 'MESA' ,
            DBO.FnDevuelveHora(PC.FECHAREG) AS 'HEE' ,
            DBO.FnDevuelvePlatosConsulta(@CODCIA,
                                         CONVERT(VARCHAR(8), PC.FECHA, 112),
                                         PC.NUMSER, PC.NUMFAC) AS 'ENTRADAS' ,
            ISNULL(CASE WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,7,8,9,10,11,12)
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) BETWEEN 1
                                         AND     11
                        THEN CONVERT(CHAR(8), ( SELECT TOP 1
                                                        P.FECHADESPACHO
                                                FROM    dbo.PEDIDOS p
                                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                                WHERE   P.PED_CODCIA = PC.CODCIA
                                                        AND P.PED_FECHA = PC.FECHA
                                                        AND P.PED_NUMSER = PC.NUMSER
                                                        AND P.PED_NUMFAC = pc.numfac
                                                        AND P.PED_ESTADO = 'N'
                                                ORDER BY P.FECHADESPACHO DESC
                                              ), 108) + ' a.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) BETWEEN 13
                                         AND     23
                        THEN CONVERT(CHAR(8), DATEADD(HOUR, -12,
                                                      ( SELECT TOP 1
                                                              P.FECHADESPACHO
                                                        FROM  dbo.PEDIDOS p
                                                              INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                                        WHERE P.PED_CODCIA = PC.CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                              AND P.PED_NUMFAC = pc.numfac
                                                              AND P.PED_ESTADO = 'N'
                                                        ORDER BY P.FECHADESPACHO DESC
                                                      )), 108) + ' p.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) = 12
                        THEN CONVERT(CHAR(8), ( SELECT TOP 1
                                                        P.FECHADESPACHO
                                                FROM    dbo.PEDIDOS p
                                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                                WHERE   P.PED_CODCIA = PC.CODCIA
                                                        AND P.PED_FECHA = PC.FECHA
                                                        AND P.PED_NUMSER = PC.NUMSER
                                                        AND P.PED_NUMFAC = pc.numfac
                                                        AND P.PED_ESTADO = 'N'
                                                ORDER BY P.FECHADESPACHO DESC
                                              ), 108) + ' p.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) = 0
                        THEN CONVERT(CHAR(8), DATEADD(HOUR, 12,
                                                      ( SELECT TOP 1
                                                              P.FECHADESPACHO
                                                        FROM  dbo.PEDIDOS p
                                                              INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA in (1,2,3,4,5,6,7,8,9,10,11,12)
                                                        WHERE P.PED_CODCIA = PC.CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                              AND P.PED_NUMFAC = pc.numfac
                                                              AND P.PED_ESTADO = 'N'
                                                        ORDER BY P.FECHADESPACHO DESC
                                                      )), 108) + ' a.m'
                   END, '') AS 'HSE' ,
                   isnull(DBO.FnDevuelveHora((SELECT TOP 1 SEG.PED_FECHACANTADO FROM dbo.PEDIDOS seg WHERE seg.PED_CODCIA = @CODCIA AND SEG.PED_FECHA = @FECHA AND SEG.PED_NUMSER = PC.NUMSER AND SEG.PED_NUMFAC = PC.NUMFAC AND SEG.PED_FAMILIA = '2' ORDER BY SEG.PED_FECHACANTADO)),'') AS 'HES',
            DBO.FnDevuelvePlatosConsulta(@CODCIA,
                                         CONVERT(VARCHAR(8), PC.FECHA, 112),
                                         PC.NUMSER, PC.NUMFAC) AS 'SEGUNDOS' ,
            ISNULL(CASE WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) BETWEEN 1
                                         AND     11
                        THEN CONVERT(CHAR(8), ( SELECT TOP 1
                                                        P.FECHADESPACHO
                                                FROM    dbo.PEDIDOS p
                                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                                WHERE   P.PED_CODCIA = PC.CODCIA
                                                        AND P.PED_FECHA = PC.FECHA
                                                        AND P.PED_NUMSER = PC.NUMSER
                                                        AND P.PED_NUMFAC = pc.numfac
                                                        AND P.PED_ESTADO = 'N'
                                                ORDER BY P.FECHADESPACHO DESC
                                              ), 108) + ' a.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) BETWEEN 13
                                         AND     23
                        THEN CONVERT(CHAR(8), DATEADD(HOUR, -12,
                                                      ( SELECT TOP 1
                                                              P.FECHADESPACHO
                                                        FROM  dbo.PEDIDOS p
                                                              INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                                        WHERE P.PED_CODCIA = PC.CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                              AND P.PED_NUMFAC = pc.numfac
                                                              AND P.PED_ESTADO = 'N'
                                                        ORDER BY P.FECHADESPACHO DESC
                                                      )), 108) + ' p.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) = 12
                        THEN CONVERT(CHAR(8), ( SELECT TOP 1
                                                        P.FECHADESPACHO
                                                FROM    dbo.PEDIDOS p
                                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                                WHERE   P.PED_CODCIA = PC.CODCIA
                                                        AND P.PED_FECHA = PC.FECHA
                                                        AND P.PED_NUMSER = PC.NUMSER
                                                        AND P.PED_NUMFAC = pc.numfac
                                                        AND P.PED_ESTADO = 'N'
                                                ORDER BY P.FECHADESPACHO DESC
                                              ), 108) + ' p.m'
                        WHEN DATEPART(HOUR,
                                      ( SELECT TOP 1
                                                P.FECHADESPACHO
                                        FROM    dbo.PEDIDOS p
                                                INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                        WHERE   P.PED_CODCIA = PC.CODCIA
                                                AND P.PED_FECHA = PC.FECHA
                                                AND P.PED_NUMSER = PC.NUMSER
                                                AND P.PED_NUMFAC = pc.numfac
                                                AND P.PED_ESTADO = 'N'
                                        ORDER BY P.FECHADESPACHO DESC
                                      )) = 0
                        THEN CONVERT(CHAR(8), DATEADD(HOUR, 12,
                                                      ( SELECT TOP 1
                                                              P.FECHADESPACHO
                                                        FROM  dbo.PEDIDOS p
                                                              INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY
                                                              AND A.ART_FAMILIA IS NULL
                                                        WHERE P.PED_CODCIA = PC.CODCIA
                                                              AND P.PED_FECHA = PC.FECHA
                                                              AND P.PED_NUMSER = PC.NUMSER
                                                              AND P.PED_NUMFAC = pc.numfac
                                                              AND P.PED_ESTADO = 'N'
                                                        ORDER BY P.FECHADESPACHO DESC
                                                      )), 108) + ' a.m'
                   END, '') AS 'HSS' ,
            ( SELECT TOP 1
                        ISNULL(P.PED_COMENSALES, 0)
              FROM      dbo.PEDIDOS p
              WHERE     P.PED_CODCIA = @CODCIA
                        AND P.PED_FECHA = PC.FECHA
                        AND P.PED_NUMSER = PC.NUMSER
                        AND P.PED_NUMFAC = PC.NUMFAC
                        AND P.PED_ESTADO = 'N'
            ) AS 'PERS'
    FROM    dbo.PEDIDOS_CABECERA pc
            INNER JOIN dbo.MESAS m ON PC.CODCIA = M.MES_CODCIA
                                      AND PC.CODMESA = M.MES_CODMES-- AND pc.FACTURADO = 0
    WHERE   PC.CODCIA = @CODCIA
            AND CONVERT(VARCHAR(8), PC.FECHA, 112) = CONVERT(VARCHAR(8), @FECHA, 112)
            AND ISNULL(PC.CODMESA, '') = CASE WHEN @MESA = '-1'
                                              THEN ISNULL(PC.CODMESA, '')
                                              ELSE @MESA
                                         END
