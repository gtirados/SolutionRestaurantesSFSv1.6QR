IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'dbo' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_COMANDA_LISTAGREGADOSxFAMILIA'
) -- Validación del nombre
BEGIN
    DROP PROC [dbo].[USP_COMANDA_LISTAGREGADOSxFAMILIA];
END;
GO
/*
USP_COMANDA_LISTAGREGADOSxFAMILIA '01',1
USP_COMANDA_LISTAGREGADOSxFAMILIA '01',5
*/
CREATE PROCEDURE [dbo].[USP_COMANDA_LISTAGREGADOSxFAMILIA]
    @codcia CHAR(2),
    @idfamilia INT
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @LISTA TINYINT;

    --OBTENER LAS SUBFAMILIAS DE AGREGADOS DE LA FAMILIA - inicio
    DECLARE @TBLSUBFAMILIA TABLE
    (
        IDSUBFAM INT
    );
    INSERT INTO @TBLSUBFAMILIA
    (
        IDSUBFAM
    )
    SELECT TAB_NUMTAB A
    FROM [dbo].[TABLAS]
    WHERE TAB_TIPREG = '123'
          AND TAB_CODCIA = @codcia
          AND TAB_CODART = @idfamilia
          AND COALESCE(TAB_AGREGADO, 0) = 1
    ORDER BY TAB_NOMLARGO;
    --OBTENER LAS SUBFAMILIAS DE AGREGADOS DE LA FAMILIA - fin

    --LISTANDO LOS PLATOS QUE PERTENECEN A AGREGADOS - inicio

    SELECT @LISTA = p.PAR_LISTA
    FROM dbo.PARGEN p
    WHERE p.PAR_CODCIA = @codcia;

    SELECT a.ART_KEY AS 'Codigo',
           a.ART_NOMBRE AS 'Plato',
           --p.Pre_Pre2 AS 'Precio' ,
           CASE
               WHEN @LISTA = 1 THEN
                   p.PRE_PRE1
               ELSE
                   CASE
                       WHEN @LISTA = 2 THEN
                           p.PRE_PRE2
                       ELSE
                           CASE
                               WHEN @LISTA = 3 THEN
                                   p.PRE_PRE3
                               ELSE
                                   CASE
                                       WHEN @LISTA = 4 THEN
                                           p.PRE_PRE4
                                       ELSE
                                           CASE
                                               WHEN @LISTA = 5 THEN
                                                   p.PRE_PRE5
                                               ELSE
                                                   p.PRE_PRE6
                                           END
                                   END
                           END
                   END
           END AS 'Precio',
           a.ART_FAMILIA AS 'CodFam',
           a.ART_SUBFAM AS 'CodSubFam',
           a.ART_ALTERNO AS 'alt'
    FROM [dbo].[ARTI] a
        inner JOIN [dbo].[PRECIOS] p
            ON a.ART_KEY = p.PRE_CODART
               AND a.ART_CODCIA = @codcia
               AND p.PRE_CODCIA = @codcia
			   AND a.ART_SUBFAM IN (SELECT IDSUBFAM FROM @TBLSUBFAMILIA)
    WHERE a.ART_SITUACION = 0
    ORDER BY a.ART_NOMBRE;
--LISTANDO LOS PLATOS QUE PERTENECEN A AGREGADOS - fin
END;
GO