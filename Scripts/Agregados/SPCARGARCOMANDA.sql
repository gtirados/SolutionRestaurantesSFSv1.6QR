IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'dbo'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'SPCARGARCOMANDA' )		-- Validación del nombre
    BEGIN
        DROP PROC [dbo].[SPCARGARCOMANDA]
    END
GO
/*
exec SpCargarComanda '01','26','2025-07-10 00:00:00'
exec SpCargarComanda '01','26','2025-07-10 00:00:00',1
exec SpCargarComanda '01','5','2025-07-10 00:00:00'
exec SpCargarComanda '01','20','2025-07-10 00:00:00'
*/
CREATE PROCEDURE [dbo].[SPCARGARCOMANDA]
    @CodCia CHAR(2),
    @CodMesa VARCHAR(10),
    @Fecha DATE,
    @Fac BIT = NULL
--@NumSer char(3) out,
--@NumFac int out,
--@CodMozo int out,
--@Mozo varchar(50) out
--With encryption
AS --Obtener la ultima comanda de acuerdo al valor maximo de ped_numfac
--Consultar con el ing. si sabe una mejor forma de obtener el pedido actual de la mesa

DECLARE @tt INT;
/*    SELECT  @tt = ISNULL(MAX(ped_numfac), 0)
    FROM    pedidos
    WHERE   ped_codclie = @codmesa
            AND ped_codcia = @codcia
            AND ped_fecha = @fecha */
SELECT TOP 1
       @tt = pc.NUMFAC
FROM dbo.PEDIDOS_CABECERA pc
WHERE pc.CODCIA = @CodCia
      AND CONVERT(VARCHAR(8), pc.FECHA, 112) = @Fecha
      AND pc.CODMESA = @CodMesa
      AND pc.FACTURADO = 0
ORDER BY pc.NUMFAC DESC;

DECLARE @ICBPER DECIMAL(8, 2);

SELECT TOP 1
       @ICBPER = COALESCE(g.GEN_ICBPER, 0)
FROM dbo.GENERAL g;

DECLARE @TBLICBPER TABLE
(
    CODART BIGINT,
    ICBPER MONEY
);

IF @Fac IS NULL
BEGIN
    SELECT A.ALL_NUMFAC AS numfac,
           PED_NUMFAC,
           PED_NUMSER,
           PED_CODVEN,
           --dbo.FnDevuelveMozo(@CodCia, PED_CODVEN) AS 'mozo' ,
           v.VEM_NOMBRE AS 'MOZO',
           PED_CODART AS 'CodPlato',
		   CASE WHEN PED_PADRE IS NULL THEN '' else '=> ' END +  dbo.FnDevuelvePlato(@CodCia, PED_CODART) AS 'Plato',
           PED_OFERTA AS 'Detalle',
           PED_PRECIO AS 'Precio',
           PED_CANTIDAD AS 'Cantidad',
           PED_SUBTOTAL AS 'Importe',
           PED_NUMSEC AS 'sec',
           PED_APROBADO AS 'apro',
           PED_CANATEN AS 'aten',
           PED_CTA AS 'cuenta',
           p.PED_CLIENTE AS 'CLIENTE',
           p.PED_COMENSALES AS 'COMENSALES',
           CAST(p.CANTADO AS INT) AS 'CANTADO',
           CASE
               WHEN p.PED_ENVIAR_EN IS NULL THEN
                   'AHORA'
               ELSE
                   CAST(DATEDIFF(MINUTE, p.PED_FECHAREG, p.PED_ENVIAR_EN) AS VARCHAR(20)) + ' min'
           END AS 'ENVIAR',
           p.PED_FAMILIA2 AS 'FAM'
		   ,CASE WHEN p.PED_PADRE IS NULL THEN 'SI' ELSE 'NO' END AS 'padre'
    FROM PEDIDOS p
        INNER JOIN dbo.PEDIDOS_CABECERA pc
            ON p.PED_FECHA = pc.FECHA
               AND p.PED_CODCIA = pc.CODCIA
               AND p.PED_NUMSER = pc.NUMSER
               AND p.PED_NUMFAC = pc.NUMFAC
        INNER JOIN dbo.VEMAEST v
            ON pc.CODCIA = v.VEM_CODCIA
               AND pc.CODMOZO = v.VEM_CODVEN
        INNER JOIN ALLOG A
            ON p.PED_TRANSP = A.ALL_NUMOPER
               AND A.ALL_CODCIA = @CodCia
               AND A.ALL_FECHA_DIA = @Fecha
               AND A.ALL_FLAG_EXT = 'N'
    --inner join allog on pedidos.ped_codcia = allog.all_codcia and pedidos.ped_fecha = allog.all_fecha_dia
    --and pedidos.ped_codart = allog.all_codclie
    WHERE PED_CODCLIE = @CodMesa
          AND PED_CODCIA = @CodCia
          AND PED_FECHA = @Fecha
          AND PED_ESTADO = 'N'
          AND PED_SITUACION <> 'A' --and ped_fac <> ped_Cantidad
          AND PED_NUMFAC = @tt
		  ORDER BY 
    -- Agrupa por el padre (si es hijo, toma PED_PADRE, si es padre su propio PED_NUMSEC)
    CASE 
        WHEN p.PED_PADRE IS NULL THEN p.PED_NUMSEC
        ELSE p.PED_PADRE
    END,
    -- Ordena padre primero, luego hijos
    CASE 
        WHEN p.PED_PADRE IS NULL THEN 0
        ELSE 1
    END,
    p.PED_NUMSEC;

END;
ELSE
BEGIN
    /*
        exec SpCargarComanda '01','7','2020-08-17 00:00:00',1
        */
    --SELECT * FROM @TBLICBPER t
    INSERT INTO @TBLICBPER
    (
        CODART,
        ICBPER
    )
    SELECT p.PA_CODPA,
           SUM(p.PA_PROM * @ICBPER) AS 'Importe'
    FROM dbo.PAQUETES p
        INNER JOIN dbo.ARTI art
            ON p.PA_CODCIA = art.ART_CODCIA
               AND p.PA_CODART = art.ART_KEY
    WHERE PA_CODPA IN
          (
              SELECT a2.ART_KEY
              FROM PEDIDOS p
                  INNER JOIN dbo.ARTI a2
                      ON p.PED_CODCIA = a2.ART_CODCIA
                         AND p.PED_CODART = a2.ART_KEY
                  INNER JOIN ALLOG A
                      ON p.PED_TRANSP = A.ALL_NUMOPER
                         AND A.ALL_CODCIA = @CodCia
                         AND A.ALL_FECHA_DIA = @Fecha
                         AND A.ALL_FLAG_EXT = 'N'
              WHERE PED_CODCLIE = @CodMesa
                    AND PED_CODCIA = @CodCia
                    AND PED_FECHA = @Fecha
                    AND PED_ESTADO = 'N'
                    AND PED_SITUACION <> 'A'
                    AND PED_FAC <> PED_CANTIDAD
                    AND PED_NUMFAC = @tt
                    AND a2.ART_FLAG_STOCK = 'C'
          )
          AND art.ART_CALIDAD = 0
    GROUP BY p.PA_CODPA;


    SELECT
        --dbo.FnDevuelveNumOper(@CodCia,@Fecha,ped_transp,ped_codart) as numfac, 
        A.ALL_NUMFAC AS numfac,
        PED_CODART AS 'CodPlato',
        --dbo.FnDevuelvePlato(@CodCia, PED_CODART) AS 'Plato',
		CASE WHEN PED_PADRE IS NULL THEN '' else '=> ' END +  dbo.FnDevuelvePlato(@CodCia, PED_CODART) AS 'Plato',
        PED_PRECIO AS 'Precio',
        --( ped_cantidad - ped_FAC ) AS 'CantTotal' ,
        CASE
            WHEN CANTIDAD_DELIVERY IS NULL THEN
                PED_CANTIDAD - PED_FAC
            ELSE
                CANTIDAD_DELIVERY
        END AS 'CantTotal',
        /*
exec SpCargarComanda '01','22','20140923',1
*/
        CASE
            WHEN CANTIDAD_DELIVERY IS NULL THEN
        (PED_CANTIDAD - PED_FAC)
            ELSE
                CANTIDAD_DELIVERY - PED_FAC
        END AS 'Faltante',
        CASE
            WHEN CANTIDAD_DELIVERY IS NULL THEN
        (PED_CANTIDAD - PED_FAC) * PED_PRECIO
            ELSE
                CANTIDAD_DELIVERY * PED_PRECIO
        END AS 'Importe',
        --( Ped_Cantidad - ped_FAC ) * Ped_Precio AS 'Importe' ,
        PED_NUMSEC AS 'sec',
        PED_APROBADO AS 'apro',
        PED_FAC AS 'aten',
        PED_UNIDAD AS 'uni',
        PED_NUMSEC,
        PED_CTA AS 'CUENTA',
        CASE
            WHEN
            (
                SELECT COALESCE(xa.ART_CALIDAD, 1)
                FROM dbo.ARTI xa
                WHERE xa.ART_CODCIA = @CodCia
                      AND xa.ART_KEY = p.PED_CODART
            ) = 0 THEN
                1
            ELSE
                0
        END AS 'ICBPER',
        @ICBPER AS 'GEN_ICBPER',
        COALESCE(t.ICBPER * PED_CANTIDAD, 0) AS 'COMBO_ICBPER'
    FROM PEDIDOS p
        INNER JOIN ALLOG A
            ON p.PED_TRANSP = A.ALL_NUMOPER
               AND A.ALL_CODCIA = @CodCia
               AND A.ALL_FECHA_DIA = @Fecha
               AND A.ALL_FLAG_EXT = 'N'
        LEFT JOIN @TBLICBPER t
            ON p.PED_CODART = t.CODART
    WHERE PED_CODCLIE = @CodMesa
          AND PED_CODCIA = @CodCia
          AND PED_FECHA = @Fecha
          AND PED_ESTADO = 'N'
          AND PED_SITUACION <> 'A'
          AND PED_FAC <> PED_CANTIDAD
          AND PED_NUMFAC = @tt
		  	  ORDER BY 
    -- Agrupa por el padre (si es hijo, toma PED_PADRE, si es padre su propio PED_NUMSEC)
    CASE 
        WHEN p.PED_PADRE IS NULL THEN p.PED_NUMSEC
        ELSE p.PED_PADRE
    END,
    -- Ordena padre primero, luego hijos
    CASE 
        WHEN p.PED_PADRE IS NULL THEN 0
        ELSE 1
    END,
    p.PED_NUMSEC;




END;
GO