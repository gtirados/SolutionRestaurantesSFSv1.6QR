/*
exec SpPrintComanda2 '01','100',11,'157,111,','0,1,'
exec SpPrintComanda2 '01','100',12,'7,157,','0,1,'
exec SpPrintComanda2 '01','100',13,'157,','0,'
exec SpPrintComanda2 '01','100',17,'33,157,112,','0,1,2,'
*/
IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPPRINTCOMANDA2'
)
BEGIN
    DROP PROC [dbo].[SPPRINTCOMANDA2];
END;
GO
/*

*/
CREATE PROCEDURE [dbo].[SPPRINTCOMANDA2]
    @CodCia CHAR(2),
    @NumSer CHAR(3),
    @NumFac INT,
    @xdet VARCHAR(4000) = NULL,
    @xnumsec VARCHAR(4000) = NULL,
    @precuenta BIT = NULL,
    @CTA CHAR(1) = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @tbltmp TABLE
(
    cp INT
);
DECLARE @idoc INT;

--productos en general
DECLARE @tbldata TABLE
(
    PED_FECHA DATETIME,
    NROCOMANDA VARCHAR(10),
    PED_CANTIDAD MONEY,
    PED_PRECIO MONEY,
    PED_IGV MONEY,
    PED_BRUTO MONEY,
    PED_HORA VARCHAR(15),
    PED_MONEDA CHAR(1),
    PED_SUBTOTAL MONEY,
    ART_NOMBRE VARCHAR(80),
    CLI_NOMBRE VARCHAR(80),
    VEM_NOMBRE VARCHAR(60),
    PED_OFERTA VARCHAR(300),
    PED_CLIENTE VARCHAR(120),
    ped_familia INT,
    codprod BIGINT,
    flag CHAR(1),
    actual DATETIME,
    FAMILIA VARCHAR(100),
    PED_SEC tinyint,
    CARACTERISTICAS VARCHAR(4000)
);

DECLARE @fecha DATETIME,
        @nrocomanda VARCHAR(15),
        @moneda CHAR(1),
        @mesa VARCHAR(50),
        @mozo VARCHAR(40);
--PARCHE ICBPER

DECLARE @ICBPER DECIMAL(8, 2);

SELECT TOP 1
       @ICBPER = COALESCE(g.GEN_ICBPER, 0)
FROM dbo.GENERAL g;
--FIN PARCHE

DECLARE @TBLICBPER TABLE
(
    CODART BIGINT,
    ICBPER MONEY
);


IF @precuenta IS NULL
BEGIN
    IF @xdet IS NULL
    BEGIN
        INSERT INTO @tbldata
        SELECT PEDIDOS.PED_FECHA,
               PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
               PEDIDOS.PED_CANTIDAD,
               PEDIDOS.PED_PRECIO,
               PEDIDOS.PED_IGV,
               PEDIDOS.PED_BRUTO,
               PEDIDOS.PED_HORA,
               PEDIDOS.PED_MONEDA,
               PEDIDOS.PED_SUBTOTAL,
               ARTI.ART_NOMBRE,
               RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
               VEMAEST.VEM_NOMBRE,
               PEDIDOS.PED_OFERTA,
               PEDIDOS.PED_CLIENTE,
               PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
               ARTI.ART_KEY,
               ARTI.ART_FLAG_STOCK,
               GETDATE(),
               (
                   SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
                   FROM dbo.TABLAS t
                   WHERE t.TAB_TIPREG = 122
                         AND t.TAB_CODCIA = @CodCia
                         AND t.TAB_NUMTAB = PEDIDOS.PED_FAMILIA2
               ),
               pedidos.ped_numsec,
               dbo.FnDevuelveCaracteristica(
                                               PEDIDOS.PED_CODCIA,
                                               PEDIDOS.PED_FECHA,
                                               PEDIDOS.PED_NUMFAC,
                                               PEDIDOS.PED_NUMSER,
                                               PEDIDOS.PED_NUMSEC,
                                               PEDIDOS.PED_CODART
                                           )
        FROM dbo.PEDIDOS PEDIDOS
            INNER JOIN dbo.MESAS CLIENTES
                ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
                   AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
            INNER JOIN dbo.VEMAEST VEMAEST
                ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
                   AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
            INNER JOIN dbo.ARTI ARTI
                ON PEDIDOS.PED_CODART = ARTI.ART_KEY
                   AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
        WHERE PEDIDOS.PED_NUMSER = @NumSer
              AND PEDIDOS.PED_NUMFAC = @NumFac
              AND PEDIDOS.PED_CODCIA = @CodCia;

    END;
    ELSE
    BEGIN
        INSERT INTO @tbldata
        SELECT PEDIDOS.PED_FECHA,
               PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
               PEDIDOS.PED_CANTIDAD,
               PEDIDOS.PED_PRECIO,
               PEDIDOS.PED_IGV,
               PEDIDOS.PED_BRUTO,
               PEDIDOS.PED_HORA,
               PEDIDOS.PED_MONEDA,
               PEDIDOS.PED_SUBTOTAL,
               ARTI.ART_NOMBRE,
               RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
               VEMAEST.VEM_NOMBRE,
               PEDIDOS.PED_OFERTA,
               PEDIDOS.PED_CLIENTE,
               PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
               ARTI.ART_KEY,
               ARTI.ART_FLAG_STOCK,
               GETDATE(),
               (
                   SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
                   FROM dbo.TABLAS t
                   WHERE t.TAB_TIPREG = 122
                         AND t.TAB_CODCIA = @CodCia
                         AND t.TAB_NUMTAB = PEDIDOS.PED_FAMILIA2
               ),
               pedidos.PED_numsec,
               dbo.FnDevuelveCaracteristica(
                                               PEDIDOS.PED_CODCIA,
                                               PEDIDOS.PED_FECHA,
                                               PEDIDOS.PED_NUMFAC,
                                               PEDIDOS.PED_NUMSER,
                                               PEDIDOS.PED_NUMSEC,
                                               PEDIDOS.PED_CODART
                                           )
        FROM dbo.PEDIDOS PEDIDOS
            INNER JOIN dbo.MESAS CLIENTES
                ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
                   AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
            INNER JOIN dbo.VEMAEST VEMAEST
                ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
                   AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
            INNER JOIN dbo.ARTI ARTI
                ON PEDIDOS.PED_CODART = ARTI.ART_KEY
                   AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
        WHERE PEDIDOS.PED_NUMSER = @NumSer
              AND PEDIDOS.PED_NUMFAC = @NumFac
              AND PEDIDOS.PED_CODCIA = @CodCia
              AND PEDIDOS.PED_CODART IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
                  )
              AND PEDIDOS.PED_NUMSEC IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
                  )
        ORDER BY PEDIDOS.PED_FECHAREG;

        --SELECT * FROM @tbldata

        --obtengo datos para impresion
        SELECT @fecha = PEDIDOS.PED_FECHA,
               @nrocomanda = PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))),
               @moneda = PEDIDOS.PED_MONEDA,
               @mesa = RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON),
               @mozo = VEMAEST.VEM_NOMBRE
        FROM dbo.PEDIDOS PEDIDOS
            INNER JOIN dbo.MESAS CLIENTES
                ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
                   AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
            INNER JOIN dbo.VEMAEST VEMAEST
                ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
                   AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
            INNER JOIN dbo.ARTI ARTI
                ON PEDIDOS.PED_CODART = ARTI.ART_KEY
                   AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
        WHERE PEDIDOS.PED_NUMSER = @NumSer
              AND PEDIDOS.PED_NUMFAC = @NumFac
              AND PEDIDOS.PED_CODCIA = @CodCia
              AND PEDIDOS.PED_CODART IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
                  )
              AND PEDIDOS.PED_NUMSEC IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
                  );


    END;

	
    --actualizo los pedidos de acuerdo a impresion
    UPDATE PEDIDOS
    SET PED_APROBADO = '1'
    WHERE PEDIDOS.PED_CODART IN
          (
              SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
          )
          AND PEDIDOS.PED_NUMSEC IN
              (
                  SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
              );


    --Select * from @tbldata
    --tabla para los productos que son combos
    DECLARE @tblcombos TABLE
    (
        codcombo BIGINT,
        cant BIGINT,
        sec TINYINT ,
        hora VARCHAR(15)
    );

    INSERT INTO @tblcombos
    SELECT codprod,
           PED_CANTIDAD,
           ped_sec,
           PED_HORA
    FROM @tbldata
    WHERE flag = 'C';
    

    

    --SELECT * FROM @tblcombos
    DECLARE @hora VARCHAR(15);

    IF EXISTS (SELECT TOP 1 codcombo FROM @tblcombos)
    BEGIN --ENTRA AQUI ES PORQUE TIENE COMBOS

        DECLARE @tbltmpCombos TABLE
        (
            PED_FECHA DATETIME,
            NROCOMANDA VARCHAR(10),
            PED_CANTIDAD MONEY,
            PED_PRECIO MONEY,
            PED_IGV MONEY,
            PED_BRUTO MONEY,
            PED_HORA VARCHAR(15),
            PED_MONEDA CHAR(1),
            PED_SUBTOTAL MONEY,
            ART_NOMBRE VARCHAR(80),
            CLI_NOMBRE VARCHAR(80),
            VEM_NOMBRE VARCHAR(60),
            ped_oferta VARCHAR(300),
            ped_cliente VARCHAR(120),
            ped_familia2 INT,
            codprod BIGINT,
            flag CHAR(1),
            FAMILIA VARCHAR(100),
			CARACTERISTICA VARCHAR(200),
            num TINYINT IDENTITY
        );

        INSERT INTO @tbltmpCombos
        SELECT PED_FECHA,
               NROCOMANDA,
               PED_CANTIDAD,
               PED_PRECIO,
               PED_IGV,
               PED_BRUTO,
               PED_HORA,
               PED_MONEDA,
               PED_SUBTOTAL,
               ART_NOMBRE,
               CLI_NOMBRE,
               VEM_NOMBRE,
               PED_OFERTA,
               PED_CLIENTE,
               ped_familia,
               codprod,
               flag,
               FAMILIA,CARACTERISTICAS
        FROM @tbldata
        WHERE flag = 'C';



        DELETE FROM @tbldata
        WHERE flag = 'C';





        DECLARE @codcombo BIGINT,
                @cant BIGINT,
                @num TINYINT,
                @po VARCHAR(300);
        --aqui entra el cursor
        DECLARE cCombos CURSOR FOR
        SELECT codcombo,
               cant,
               sec,
               hora
        FROM @tblcombos;

        OPEN cCombos;

        FETCH cCombos
        INTO @codcombo,
             @cant,
             @num,
             @hora;

        WHILE (@@Fetch_Status = 0)
        BEGIN

            INSERT INTO @tbldata
            SELECT PED_FECHA,
                   NROCOMANDA,
                   PED_CANTIDAD,
                   PED_PRECIO,
                   PED_IGV,
                   PED_BRUTO,
                   PED_HORA,
                   PED_MONEDA,
                   PED_SUBTOTAL,
                   ART_NOMBRE,
                   CLI_NOMBRE,
                   VEM_NOMBRE,
                   ped_oferta,
                   ped_cliente,
                   ped_familia2,
                   codprod,
                   flag,
                   GETDATE(),
                   FAMILIA,
                   @num,
                   CARACTERISTICA
            FROM @tbltmpCombos
            WHERE codprod = @codcombo
                  AND num = @num;


            INSERT INTO @tbldata
            SELECT @fecha,
                   @nrocomanda,
                   pa.PA_PROM * @cant, --ped_cantidad  GTS ACA CANTIDAD DE COMBOS
                   0,
                   0,
                   0,
                   @hora,
                   @moneda,
                   0,
                   ' * ' + ar.ART_NOMBRE,
                   @mesa,
                   @mozo,
                   ISNULL(@po, ''),
                   '',
                   ar.ART_FAMILIA,
                   pa.PA_CODART,
                   ar.ART_FLAG_STOCK,
                   GETDATE(),
                   (
                       SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
                       FROM dbo.TABLAS t
                       WHERE t.TAB_TIPREG = 122
                             AND t.TAB_CODCIA = @CodCia
                             AND t.TAB_NUMTAB = ar.ART_FAMILIA
                   ),
                   @num, --validar
                   dbo.FnDevuelveCaracteristica(@CodCia,@fecha,@NumFac,@NumSer,@num,pa.PA_CODART)
            FROM PAQUETES pa
                INNER JOIN ARTI ar
                    ON pa.PA_CODCIA = ar.ART_CODCIA
                       AND PA_CODART = ar.ART_KEY
            WHERE pa.PA_CODPA = @codcombo
                  AND pa.PA_CODCIA = @CodCia;

            FETCH cCombos
            INTO @codcombo,
                 @cant,
                 @num,
                 @hora;
        END;

        CLOSE cCombos;
        DEALLOCATE cCombos;


    END;
    



    --PARCHE PARA AGREGAR FAMILIA AL COMBO FALTANTE
    DECLARE @MIN INT,
            @MAX INT;
    DECLARE @TBLFAMILIA TABLE
    (
        IDFAMILIA INT,
        INDICE INT IDENTITY
    );
    INSERT INTO @TBLFAMILIA
    (
        IDFAMILIA
    )
    SELECT DISTINCT
           ped_familia
    FROM @tbldata;

    SELECT @MIN = MIN(t.INDICE)
    FROM @TBLFAMILIA t;
    SELECT @MAX = MAX(t.INDICE)
    FROM @TBLFAMILIA t;
    
    --        select * from @tblcombos
    --select * from @tbldata
    --select * from @tbltmpCombos
    --select * from @TBLFAMILIA


/*
exec SpPrintComanda2 '01','100',13,'157,','0,'
*/


    WHILE @MIN <= @MAX
    BEGIN
        IF NOT EXISTS
        (
            SELECT TOP 1
                   NROCOMANDA
            FROM @tbldata
            WHERE ped_familia =
            (
                SELECT TOP 1 t.IDFAMILIA FROM @TBLFAMILIA t WHERE t.INDICE = @MIN
            )
                  AND flag = 'C'
        )
        BEGIN

            INSERT INTO @tbldata
            SELECT top 1  PED_FECHA,
                   NROCOMANDA,
                   PED_CANTIDAD,
                   PED_PRECIO,
                   PED_IGV,
                   PED_BRUTO,
                   dbo.FnDevuelveHora(GETDATE()),
                   PED_MONEDA,
                   PED_SUBTOTAL,
                   ART_NOMBRE,
                   CLI_NOMBRE,
                   VEM_NOMBRE,
                   PED_OFERTA,
                   PED_CLIENTE,
                   (
                       SELECT TOP 1 t.IDFAMILIA FROM @TBLFAMILIA t WHERE t.INDICE = @MIN
                   ),
                   codprod,
                   'C',
                   --actual,
                   getdate(),
                   (
                       SELECT TOP 1
                              t.TAB_NOMLARGO
                       FROM dbo.TABLAS t
                       WHERE t.TAB_TIPREG = 122
                             AND t.TAB_NUMTAB =
                             (
                                 SELECT TOP 1 t.IDFAMILIA FROM @TBLFAMILIA t WHERE t.INDICE = @MIN
                             )
                   ),
                   1, --validar
                   ''
            FROM @tbltmpCombos
            --WHERE ped_familia =
            --(
            --    --SELECT TOP 1 t.ped_familia FROM @tbldata t WHERE t.flag = 'C'
            --     SELECT TOP 1 t.ped_familia2 FROM @tbltmpCombos t WHERE t.flag = 'C'
            --)
            --      AND flag = 'C';


        END;


        SET @MIN = @MIN + 1;
    END;



    SELECT  PED_FECHA ,
    NROCOMANDA,
    PED_CANTIDAD ,
    PED_PRECIO ,
    PED_IGV ,
    PED_BRUTO ,
    PED_HORA ,
    PED_MONEDA ,
    PED_SUBTOTAL ,
    ART_NOMBRE ,
    CLI_NOMBRE ,
    VEM_NOMBRE ,
    PED_OFERTA ,
    PED_CLIENTE ,
    ped_familia ,
    codprod ,
    flag ,
    actual ,
    FAMILIA ,
    CARACTERISTICAS,ped_sec
    FROM @tbldata
    ORDER BY ped_familia,
             flag,
             ART_NOMBRE;

END;
ELSE
BEGIN

    SELECT TOP 1
           @fecha = PED_FECHA
    FROM PEDIDOS
    WHERE PED_CODCIA = @CodCia
          AND PED_NUMSER = @NumSer
          AND PED_NUMFAC = @NumFac;

    INSERT INTO @TBLICBPER
    (
        CODART,
        ICBPER
    )
    SELECT p.PA_CODPA,
           SUM(p.PA_PROM * @ICBPER)
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
                         AND A.ALL_FECHA_DIA = @fecha
                         AND A.ALL_FLAG_EXT = 'N'
              WHERE PED_CODCIA = @CodCia
                    AND PED_FECHA = @fecha
                    AND PED_ESTADO = 'N'
                    AND PED_SITUACION <> 'A'
                    AND PED_FAC <> PED_CANTIDAD
                    AND PED_NUMFAC = @NumFac
                    AND a2.ART_FLAG_STOCK = 'C'
          )
          AND art.ART_CALIDAD = 0
    GROUP BY p.PA_CODPA;


    DECLARE @total MONEY;
    DECLARE @cant2 INT;



    SELECT @total = SUM(PED_SUBTOTAL)
    FROM PEDIDOS
    WHERE PED_NUMFAC = @NumFac;

    --exec SpPrintComanda2 '01','100',20141,'35,27,','0,1,',1,' '
    --SELECT * FROM PEDIDOS WHERE PED_NUMFAC=20141
    --SELECT * FROM @TBLICBPER


    SELECT @ICBPER = ICBPER
    FROM @TBLICBPER;
    SELECT @cant2 = SUM(PED_CANTIDAD)
    FROM PEDIDOS
    WHERE PED_NUMFAC = @NumFac
          AND PED_CODART IN
              (
                  SELECT CODART FROM @TBLICBPER
              );

    SET @total = @total + (@ICBPER * @cant2);

    --entra aqui cuando es precuenta
    IF @xdet IS NULL
    BEGIN
        SELECT PEDIDOS.PED_FECHA,
               PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
               --PEDIDOS.PED_CANTIDAD ,
               CASE
                   WHEN CANTIDAD_DELIVERY IS NULL THEN
                       PEDIDOS.PED_CANTIDAD
                   ELSE
                       PEDIDOS.CANTIDAD_DELIVERY
               END AS 'PED_CANTIDAD',
               PEDIDOS.PED_PRECIO,
               PEDIDOS.PED_IGV,
               PEDIDOS.PED_BRUTO,
               PEDIDOS.PED_HORA,
               PEDIDOS.PED_MONEDA,
               (CASE
                    WHEN
                    (
                        SELECT COALESCE(xa.ART_CALIDAD, 1)
                        FROM dbo.ARTI xa
                        WHERE xa.ART_CODCIA = @CodCia
                              AND xa.ART_KEY = PEDIDOS.PED_CODART
                    ) = 0 THEN
               (CASE
                    WHEN CANTIDAD_DELIVERY IS NULL THEN
                        PEDIDOS.PED_CANTIDAD
                    ELSE
                        PEDIDOS.CANTIDAD_DELIVERY
                END * @ICBPER
               )
                    ELSE
                        0
                END
               ) + PEDIDOS.PED_SUBTOTAL + COALESCE(t.ICBPER, 0) AS 'PED_SUBTOTAL',
               CASE
                   WHEN PEDIDOS.CANTIDAD_DELIVERY IS NOT NULL THEN
                       '1/2  '
                   ELSE
                       ''
               END + ARTI.ART_NOMBRE AS 'ART_NOMBRE',
               RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
               --VEMAEST.VEM_NOMBRE ,
               VEMAEST.VEM_NOMBRE,
               PEDIDOS.PED_CTA,
               dbo.FnDevuelveEstadoMesa(@CodCia, MES_CODMES) AS 'ESTADO_MESA',
               PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
               @total AS 'TOTAL'
        FROM dbo.PEDIDOS PEDIDOS
            INNER JOIN dbo.MESAS CLIENTES
                ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
                   AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
            INNER JOIN dbo.PEDIDOS_CABECERA pc
                ON PEDIDOS.PED_CODCIA = pc.CODCIA
                   AND PEDIDOS.PED_NUMFAC = pc.NUMFAC
                   AND PEDIDOS.PED_NUMSER = pc.NUMSER
                   AND PEDIDOS.PED_FECHA = pc.FECHA
            INNER JOIN dbo.VEMAEST VEMAEST
                ON pc.CODMOZO = VEMAEST.VEM_CODVEN
                   AND pc.CODCIA = VEMAEST.VEM_CODCIA
            INNER JOIN dbo.ARTI ARTI
                ON PEDIDOS.PED_CODART = ARTI.ART_KEY
                   AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
            LEFT JOIN @TBLICBPER t
                ON PEDIDOS.PED_CODART = t.CODART
        WHERE PEDIDOS.PED_NUMSER = @NumSer
              AND PEDIDOS.PED_NUMFAC = @NumFac
              AND PEDIDOS.PED_CODCIA = @CodCia
              AND ISNULL(PEDIDOS.PED_CTA, '') = @CTA
        ORDER BY PED_NUMSEC; --CAMBIADO
    END;
    ELSE
    BEGIN

        SELECT PEDIDOS.PED_FECHA,
               PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
               CASE
                   WHEN CANTIDAD_DELIVERY IS NULL THEN
                       PEDIDOS.PED_CANTIDAD
                   ELSE
                       PEDIDOS.CANTIDAD_DELIVERY
               END AS 'PED_CANTIDAD',
               PEDIDOS.PED_PRECIO,
               PEDIDOS.PED_IGV,
               PEDIDOS.PED_BRUTO,
               PEDIDOS.PED_HORA,
               PEDIDOS.PED_MONEDA,
               (CASE
                    WHEN
                    (
                        SELECT COALESCE(xa.ART_CALIDAD, 1)
                        FROM dbo.ARTI xa
                        WHERE xa.ART_CODCIA = @CodCia
                              AND xa.ART_KEY = PEDIDOS.PED_CODART
                    ) = 0 THEN
               (CASE
                    WHEN CANTIDAD_DELIVERY IS NULL THEN
                        PEDIDOS.PED_CANTIDAD
                    ELSE
                        PEDIDOS.CANTIDAD_DELIVERY
                END * @ICBPER
               )
                    ELSE
                        0
                END
               ) + PEDIDOS.PED_SUBTOTAL + COALESCE(t.ICBPER, 0) AS 'PED_SUBTOTAL',
               --PEDIDOS.PED_SUBTOTAL ,
               CASE
                   WHEN PEDIDOS.CANTIDAD_DELIVERY IS NOT NULL THEN
                       '1/2  '
                   ELSE
                       ''
               END + ARTI.ART_NOMBRE AS 'ART_NOMBRE',
               RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
               VEMAEST.VEM_NOMBRE,
               PEDIDOS.PED_CTA,
               dbo.FnDevuelveEstadoMesa(@CodCia, MES_CODMES) AS 'ESTADO_MESA',
               PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
               dbo.FnDevuelveCaracteristica(
                                               PEDIDOS.PED_CODCIA,
                                               PEDIDOS.PED_FECHA,
                                               PEDIDOS.PED_NUMFAC,
                                               PEDIDOS.PED_NUMSER,
                                               PEDIDOS.PED_NUMSEC,
                                               PEDIDOS.PED_CODART
                                           ) AS 'CARACTERISTICA',
               @total AS 'TOTAL'
        --,coalesce(t.ICBPER,0)

        FROM dbo.PEDIDOS PEDIDOS
            INNER JOIN dbo.MESAS CLIENTES
                ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
                   AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
            INNER JOIN dbo.PEDIDOS_CABECERA pc
                ON PEDIDOS.PED_CODCIA = pc.CODCIA
                   AND PEDIDOS.PED_NUMFAC = pc.NUMFAC
                   AND PEDIDOS.PED_NUMSER = pc.NUMSER
                   AND PEDIDOS.PED_FECHA = pc.FECHA
            INNER JOIN dbo.VEMAEST VEMAEST
                ON pc.CODMOZO = VEMAEST.VEM_CODVEN
                   AND pc.CODCIA = VEMAEST.VEM_CODCIA
            INNER JOIN dbo.ARTI ARTI
                ON PEDIDOS.PED_CODART = ARTI.ART_KEY
                   AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
            LEFT JOIN @TBLICBPER t
                ON PEDIDOS.PED_CODART = t.CODART
        /*
exec SpPrintComanda2 '01','100',46,'2050,','2,',1,' '
*/
        WHERE PEDIDOS.PED_NUMSER = @NumSer
              AND PEDIDOS.PED_NUMFAC = @NumFac
              AND PEDIDOS.PED_CODCIA = @CodCia
              AND PEDIDOS.PED_CODART IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
                  )
              AND PEDIDOS.PED_NUMSEC IN
                  (
                      SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
                  )
              AND ISNULL(PEDIDOS.PED_CTA, '') = @CTA
        ORDER BY PED_NUMSEC; --CAMBIADO

    END;
END;

GO